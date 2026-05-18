import json
import anthropic

from config import ANTHROPIC_API_KEY, MODEL, MONTHLY_BUDGET
from tools import execute_tool, get_schema, build_tool_definitions

MAX_ITERATIONS = 5


def _build_system_prompt() -> str:
    schema = get_schema()

    cats      = ", ".join(schema.categories) if schema.categories else "see file"
    budget    = schema.monthly_budget or MONTHLY_BUDGET
    date_info = (
        f"{schema.date_range_start} → {schema.date_range_end}"
        if schema.date_range_start else "see file"
    )
    file_name = schema.file_path.replace("\\", "/").split("/")[-1]
    source_note = (
        "Monthly totals come from the pre-aggregated dashboard (includes all fixed "
        "costs that may be hardcoded outside the transaction log)."
        if schema.has_dashboard
        else "Monthly totals are computed directly from individual transactions."
    )

    return f"""You are FinAgent, a personal finance assistant with direct access to the \
user's transaction data from '{file_name}'.

Data summary:
- Date range: {date_info}
- Monthly budget: ${budget:,.0f}
- Expense categories: {cats}
- {source_note}

Rules:
1. Always cite specific dollar figures from tool results — never estimate or guess.
2. If a question needs data from multiple months or tools, call all tools first, \
then compose one answer.
3. Keep answers conversational and concise — plain English prose, not bullet lists.
4. If a tool returns no data for a month, say so naturally.
5. Only answer questions about personal finance, spending, and budgeting."""


_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


def run_agent(user_message: str) -> str:
    """Run the agentic loop for a single user query. Returns the final text answer."""
    schema   = get_schema()
    tools    = build_tool_definitions(schema.categories)
    system   = _build_system_prompt()
    messages = [{"role": "user", "content": user_message}]

    for _ in range(MAX_ITERATIONS):
        response = _client.messages.create(
            model=MODEL,
            max_tokens=4096,
            system=system,
            tools=tools,
            messages=messages,
        )

        tool_use_blocks = [b for b in response.content if b.type == "tool_use"]

        if not tool_use_blocks:
            text_blocks = [b.text for b in response.content if b.type == "text"]
            return "\n".join(text_blocks) if text_blocks else "(No response generated.)"

        tool_results = []
        for block in response.content:
            if block.type != "tool_use":
                continue
            args = ", ".join(f"{k}={repr(v)}" for k, v in block.input.items())
            print(f"  Calling {block.name}({args})...")
            result = execute_tool(block.name, block.input)
            tool_results.append({
                "type":        "tool_result",
                "tool_use_id": block.id,
                "content":     json.dumps(result),
            })

        messages.append({"role": "assistant", "content": response.content})
        messages.append({"role": "user",      "content": tool_results})

    last_text = next((b.text for b in response.content if b.type == "text"), None)
    return last_text or "(Reached the maximum number of tool-call iterations.)"
