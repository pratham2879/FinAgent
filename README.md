# FinAgent — Personal Expense Analyst

FinAgent is a conversational AI agent that lets you ask natural-language questions about any personal expense spreadsheet. It is built on Claude's **tool-use (function calling) API**: rather than dumping the entire spreadsheet into a prompt, Claude decides at runtime which typed data-retrieval function to call, gets back a focused JSON result, and synthesises a plain-English answer. Point it at any Excel file — it auto-detects the sheet name, column positions, date format, expense categories, and whether a pre-aggregated dashboard exists. The architecture directly mirrors how production fintech platforms (Plaid, Yodlee, MX) expose financial data — as structured, queryable endpoints — making this a realistic demonstration of agentic reasoning over financial data.

**Stack:** Python 3.9+ · Anthropic SDK · pandas · openpyxl · python-dotenv

---

## Full architecture

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

## How the inspector works (inspector.py)

On first query the inspector scans the file once and caches an `ExcelSchema`. No column names, sheet names, or row positions are hardcoded anywhere else.

**Step 1 — Find the transaction sheet**
Each sheet is scored: +2 per keyword match in the sheet name (`transactions`, `data`, `expenses`, …) plus a content score from scanning the first 20 rows for recognisable column headers. Highest score wins.

**Step 2 — Detect the header row**
Rows 1–20 are scored by how many of 7 keyword sets they cover (date, month, description, category, amount, payment, notes). Bonus points for having both `amount` and `category`. The top-scoring row is the header.

**Step 3 — Map column roles**
Each header cell is fuzzy-matched against keyword sets (e.g. `"Amount ($)"` → strips `($)` → matches `amount`). Original column name is stored so pandas `rename()` can map it to the internal standard name.

**Step 4 — Find the dashboard**
Sheets named `dashboard`, `summary`, `overview`, etc. are checked for a column where ≥ 3 values match the month-string pattern (`Aug-25`, `Mar-26`, …). The row above the first month value is the header row. Columns are collected left-to-right until the first `TOTAL` column or the first None gap — this stops the scanner from leaking into unrelated tables on the same sheet (e.g. a "Utilities Breakdown" to the right).

**Step 5 — Detect budget**
Scans the dashboard for a cell labelled "budget" and returns the nearest adjacent number.

---

## Dual data-source design (the non-obvious part)

The agent reads from two different sheets for different query types:

```
Transactions sheet  →  get_top_merchants(), get_transactions()
                        needed for row-level detail (merchant, date, amount)

Dashboard sheet     →  get_spending_by_category(), get_monthly_summary(), compare_months()
    read with           needed because some spreadsheets hardcode fixed costs
    data_only=True      (e.g. Rent = $420) directly in Dashboard cells rather
                        than as individual transactions — reading only the
                        Transactions sheet silently under-counts those months
```

If no dashboard is detected, all five tools fall back to aggregating from the Transactions sheet automatically.

**Interview talking point:** "Reading only the Transactions sheet gave wrong monthly totals — Rent was hardcoded in the Dashboard. I fixed it by treating the Dashboard as the OLAP layer (pre-aggregated, authoritative totals read with `data_only=True`) and the Transactions sheet as the OLTP layer (row-level detail for filtering). That's the same separation used in production financial data warehouses."

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

**Why multiple tool calls per turn matter:** "Did I spend more on food in March than February?" requires data from two months. Claude issues both `get_spending_by_category` calls in a single response. The loop executes them together and sends both results back in one round-trip.

---

## The 5 tools

| Tool | Data source | What it returns |
|---|---|---|
| `get_spending_by_category(month, year)` | Dashboard → Transactions | Total per category, sorted by spend |
| `get_monthly_summary(month, year)` | Dashboard → Transactions | Total, budget, over/under, % utilisation |
| `compare_months(month1, year1, month2, year2)` | Dashboard → Transactions | Category delta, biggest movers |
| `get_top_merchants(month, year, n=5)` | Transactions | Top N stores by total spend |
| `get_transactions(month, year, category?, min_amount?)` | Transactions | Filtered line-item list |

Tool JSON schemas are generated at runtime with `build_tool_definitions(categories)` — the detected category list is injected into each schema description so Claude always knows the valid values for the current file.

---

## Why function calling instead of prompt-stuffing

1. **Token efficiency.** 100+ transactions × every message = thousands of tokens wasted. Tool calls retrieve only the slice each question needs.

2. **Deterministic arithmetic.** LLMs mis-add numbers when given raw data. Tool functions run in Python — the maths is always correct.

3. **Scales with data.** A 500-row sheet works identically to a 100-row sheet when queried via typed functions. Prompt-stuffing degrades linearly with size.

4. **Production pattern.** Plaid's `/transactions/get`, Yodlee's `/transactions`, and MX's `/transactions` all work this way — the AI calls a typed endpoint and gets structured JSON back. FinAgent uses the identical pattern at personal scale.

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

`EXCEL_PATH` is required — there is no safe default. `ANTHROPIC_API_KEY` must be a valid key from [console.anthropic.com](https://console.anthropic.com).

**Model:** `claude-sonnet-4-5` (set in `config.py` → `MODEL`). Change this to any Anthropic model that supports tool use.

Run:
```bash
python main.py
```

On startup FinAgent prints everything it detected about your file:
```
  File    : Student_Expense_Tracker.xlsx
  Sheet   : Transactions  (header row 4)
  Columns : {'date': 'Date', 'month': 'Month', 'description': 'Description', ...}
  Range   : Jan 2025 → May 2026
  Budget  : $650/month
  Categories (9): Dine Out, Entertainment, Groceries, ...
  Dashboard: 'Dashboard' — 8 category columns detected
```

Type `reload` at the prompt to re-inspect the file after edits.

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
