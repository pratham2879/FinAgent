# FinAgent — Personal Expense Analyst

As a student, manually logging every transaction and keeping track of where your money goes gets tedious fast. I built this because I wanted to ask plain-English questions about my spending without manually digging through spreadsheets. You point it at any `.xlsx` expense file, and it figures out the structure on its own — sheet names, column layout, categories, budget, everything. Then you just ask things like "was I over budget in March?" and it answers.

Under the hood it uses Claude's **tool-use API**: instead of dumping the whole spreadsheet into a prompt, Claude calls typed Python functions to fetch exactly the data it needs, gets back focused JSON, and writes a plain-English answer. It's the same pattern Plaid and Yodlee use — expose financial data as structured endpoints, let the AI query them.

**Stack:** Python 3.9+ · Anthropic SDK · pandas · openpyxl · python-dotenv

---

## How it works

```
┌─────────────────────────────────────────────────────────────────────┐
│                            FinAgent                                 │
│                                                                     │
│   Any expense .xlsx file                                            │
│           │                                                         │
│           ▼  on first run                                           │
│   ┌───────────────────────────────────────────────────────┐        │
│   │              inspector.py  (auto-detection)            │        │
│   │                                                       │        │
│   │  1. Find transaction sheet  (name + column scoring)   │        │
│   │  2. Detect header row       (keyword set matching)    │        │
│   │  3. Map column roles        (date/amount/category…)   │        │
│   │  4. Find dashboard sheet    (month×category grid)     │        │
│   │  5. Detect budget           (scan for "budget" label) │        │
│   │                                                       │        │
│   │  → returns ExcelSchema  (cached for the session)      │        │
│   └───────────────────────┬───────────────────────────────┘        │
│                           │  ExcelSchema                            │
│                           ▼                                         │
│   ┌───────────────────────────────────────────────────────┐        │
│   │              tools.py  (5 tool functions)              │        │
│   │                                                       │        │
│   │  _load_df()              reads Transactions sheet     │        │
│   │  _load_dashboard_data()  reads Dashboard sheet        │        │
│   │                          (data_only=True)             │        │
│   │                          falls back to transactions   │        │
│   │                          if no dashboard detected     │        │
│   │                                                       │        │
│   │  get_spending_by_category  ──► Dashboard / fallback   │        │
│   │  get_monthly_summary       ──► Dashboard / fallback   │        │
│   │  compare_months            ──► Dashboard / fallback   │        │
│   │  get_top_merchants         ──► Transactions           │        │
│   │  get_transactions          ──► Transactions           │        │
│   └───────────────────────┬───────────────────────────────┘        │
│                           │  JSON results                           │
│                           ▼                                         │
│   ┌───────────────────────────────────────────────────────┐        │
│   │              agent.py  (agentic loop)                  │        │
│   │                                                       │        │
│   │  System prompt built from schema                      │        │
│   │  (file name, date range, budget, categories)          │        │
│   │                                                       │        │
│   │  User query                                           │        │
│   │      │                                               │        │
│   │      ▼                                               │        │
│   │  Claude API ──► tool_use blocks? ──► execute tools   │        │
│   │      ▲                                   │           │        │
│   │      └────── tool_result messages ◄──────┘           │        │
│   │                                                       │        │
│   │  Repeats up to MAX_ITERATIONS = 5                     │        │
│   │      ▼  stop_reason = "end_turn"                      │        │
│   │  Plain-English answer                                 │        │
│   └───────────────────────────────────────────────────────┘        │
│                                                                     │
│   main.py  —  CLI loop + startup inspection banner                 │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Project structure

```
FinAgent/
├── inspector.py     Auto-detects any Excel file's structure → ExcelSchema
├── tools.py         5 tool functions + JSON schema builder + dual-source loader
├── agent.py         Agentic loop — dynamic system prompt, tool dispatch
├── config.py        API key, model, Excel path, fallback budget
├── main.py          CLI loop — inspection banner, reload command
├── .env             Secrets (not committed — copy from .env.example)
├── .env.example     Required environment variable template
├── requirements.txt
└── README.md
```

---

## The inspector (inspector.py)

The first time you ask a question, the inspector scans the file once and caches an `ExcelSchema`. Nothing else in the codebase has hardcoded sheet names, column positions, or category lists — they all come from this.

**Step 1 — Find the transaction sheet**
Each sheet gets a score: +2 per keyword match in the name (`transactions`, `data`, `expenses`, …) plus a content score from scanning the first 20 rows for recognisable column headers. Highest score wins.

**Step 2 — Detect the header row**
Rows 1–20 are scored by how many of 7 keyword sets they cover (date, month, description, category, amount, payment, notes). Bonus points for having both `amount` and `category`. The top-scoring row becomes the header.

**Step 3 — Map column roles**
Each header cell is fuzzy-matched against keyword sets — so `"Amount ($)"` strips the `($)` and matches `amount`. The original column name is kept so pandas can rename it to the internal standard.

**Step 4 — Find the dashboard**
Sheets named `dashboard`, `summary`, `overview`, etc. are checked for a column with ≥ 3 values matching the month-string pattern (`Aug-25`, `Mar-26`, …). Columns are collected left-to-right until the first `TOTAL` column or a None gap — this prevents the scanner from picking up unrelated tables sitting to the right on the same sheet.

**Step 5 — Detect budget**
Scans the dashboard for a cell labelled "budget" and returns the nearest adjacent number.

---

## Why two data sources

This was the trickiest design decision. I initially read everything from the Transactions sheet, but monthly totals kept coming out wrong. It turned out my spreadsheet had fixed costs like Rent hardcoded directly in Dashboard cells — they were never logged as individual transactions.

The fix was to treat the two sheets like separate database layers:

```
Transactions sheet  →  get_top_merchants(), get_transactions()
                        row-level detail: merchant, date, individual amount

Dashboard sheet     →  get_spending_by_category(), get_monthly_summary(), compare_months()
    read with           pre-aggregated totals that include fixed costs not in
    data_only=True      the transaction log — this is the authoritative source
```

If no dashboard exists, all five tools fall back to aggregating from transactions automatically.

This is the same OLTP/OLAP separation used in production data warehouses — transactions for row-level queries, a pre-aggregated layer for summary queries.

---

## The agentic loop (agent.py)

```python
# System prompt is built from the detected schema — not hardcoded
system = _build_system_prompt()   # includes file name, date range, categories, budget

messages = [{"role": "user", "content": user_question}]

for _ in range(MAX_ITERATIONS):          # hard cap = 5

    response = claude_api.messages.create(
        model    = MODEL,
        system   = system,
        tools    = build_tool_definitions(schema.categories),  # category list injected
        messages = messages,
    )

    tool_use_blocks = [b for b in response.content if b.type == "tool_use"]

    if not tool_use_blocks:
        return response.text             # Claude is done — final answer

    # Execute ALL tool calls from this turn (Claude may call 2+ at once)
    tool_results = []
    for block in tool_use_blocks:
        result = execute_tool(block.name, block.input)
        tool_results.append({
            "type":        "tool_result",
            "tool_use_id": block.id,
            "content":     json.dumps(result),
        })

    messages.append({"role": "assistant", "content": response.content})
    messages.append({"role": "user",      "content": tool_results})
    # Claude sees results and either calls more tools or gives a final answer
```

One thing worth noting: Claude can issue multiple tool calls in a single turn. A question like "did I spend more on food in March than February?" requires data from two months — Claude requests both `get_spending_by_category` calls at once, the loop executes them together, and both results go back in one round-trip.

---

## The 5 tools

| Tool | Data source | What it returns |
|---|---|---|
| `get_spending_by_category(month, year)` | Dashboard → Transactions | Total per category, sorted by spend |
| `get_monthly_summary(month, year)` | Dashboard → Transactions | Total, budget, over/under, % utilisation |
| `compare_months(month1, year1, month2, year2)` | Dashboard → Transactions | Category delta, biggest movers |
| `get_top_merchants(month, year, n=5)` | Transactions | Top N stores by total spend |
| `get_transactions(month, year, category?, min_amount?)` | Transactions | Filtered line-item list |

Tool JSON schemas are generated at runtime via `build_tool_definitions(categories)` — the detected category list is injected so Claude always knows the valid values for the current file.

---

## Why function calling instead of dumping the spreadsheet into the prompt

1. **Token efficiency.** 100+ transactions × every message = thousands of tokens per query. Tool calls fetch only the slice each question actually needs.

2. **Correct arithmetic.** LLMs make arithmetic mistakes on raw data. Tool functions run in Python — the numbers are always right.

3. **Scales cleanly.** A 500-row sheet works the same as a 50-row sheet when data comes through typed functions. Prompt-stuffing gets worse as the file grows.

4. **Matches how real finance APIs work.** Plaid's `/transactions/get`, Yodlee's `/transactions`, MX's `/transactions` — the AI calls a typed endpoint and gets structured JSON back. FinAgent is the same pattern at personal scale.

---

## Setup

**Requirements:** Python 3.9+, pip

```bash
pip install -r requirements.txt
```

Copy `.env.example` to `.env` and fill in your values:
```
ANTHROPIC_API_KEY=sk-ant-...
EXCEL_PATH=C:\path\to\your_expenses.xlsx
```

`EXCEL_PATH` is required — there's no built-in default. Get your API key from [console.anthropic.com](https://console.anthropic.com).

**Model:** defaults to `claude-sonnet-4-5`, set in `config.py` → `MODEL`. Any Anthropic model with tool use support will work.

```bash
python main.py
```

On startup it prints everything it detected about your file:
```
  File    : Student_Expense_Tracker.xlsx
  Sheet   : Transactions  (header row 4)
  Columns : {'date': 'Date', 'month': 'Month', 'description': 'Description', ...}
  Range   : Jan 2025 → May 2026
  Budget  : $650/month
  Categories (9): Dine Out, Entertainment, Groceries, ...
  Dashboard: 'Dashboard' — 8 category columns detected
```

Type `reload` at the prompt to re-inspect the file after making edits.

---

## Example queries

| Query | Tool(s) called |
|---|---|
| `Give me a summary of March 2026` | `get_monthly_summary` |
| `How much did I spend on dine out in November?` | `get_spending_by_category` |
| `Compare February and March 2026` | `compare_months` |
| `What were my top 5 spending spots in March?` | `get_top_merchants` |
| `Show me all transactions over $30 in March` | `get_transactions(min_amount=30)` |
| `Was I over budget in January?` | `get_monthly_summary` |
| `Where did my grocery spend go in March?` | `get_transactions(category='Groceries')` |
