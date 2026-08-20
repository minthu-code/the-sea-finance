import os
import json
import logging
from typing import List, Dict, Any
import requests

import exhibitledger as el

logger = logging.getLogger(__name__)

# Use OpenAI for the analyst, similar to the OCR logic
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")

def build_analyst_system_prompt(exhibition_code: str) -> str:
    """Build a professional system prompt with live data context."""
    try:
        ex = el.get_exhibition(exhibition_code)
        metrics = el.calculate_inventory_metrics(exhibition_code)
        report = el.calculate_report(exhibition_code)
        artist_roi = el.calculate_artist_roi(exhibition_code)
        forecast = el.get_forecast_metrics(exhibition_code)
        
        ex_name = ex["name"] if ex else exhibition_code
        
        # Prepare ROI summary for context
        roi_summary = "\n".join([f"- {r['artist']}: {r['sold_artworks']}/{r['total_artworks']} works, Contribution: {el.money(r['net_contribution'])}" for r in artist_roi[:5]])
        
        prompt = f"""You are "SEA", an elite Gallery Finance Analyst for THE SEA ART GALLERY.
Your partner is Mike, an independent curator. You are more powerful and proactive than previous iterations.

TONE & STYLE:
- Calm, concise, warm, and professional.
- Never flatter; challenge weak ideas respectfully.
- If you don't know something or data is missing, say "I don't know".
- Present trade-offs before recommending options.

EXHIBITION DATA: {ex_name} ({exhibition_code})
- Net Profit: {el.money(report['totals']['net_profit'])}
- Revenue: {el.money(report['totals']['gallery_revenue'])}
- Expenses: {el.money(report['totals']['direct_costs'] + report['totals']['operating_expenses'] + report['totals']['allocated_overhead'])}
- Daily Burn Rate: {el.money(forecast['daily_burn_rate'])}
- Projected Final Profit: {el.money(forecast['projected_net_profit'])}
- Sell-through Rate: {metrics['sell_through_rate_pct']:.1f}%

TOP ARTIST PERFORMANCE:
{roi_summary}

YOUR CAPABILITIES:
- You provide proactive insights. Don't just answer; suggest optimizations.
- You analyze ROI per artist and identify "cost centers" in the exhibition.
- You forecast final outcomes based on current burn rates and sales pace.
- You help Mike maintain the "Elite" status of the gallery by ensuring financial discipline.

Always respond in Markdown. Use bold for key figures. Be precise.
"""
        return prompt
    except Exception as e:
        logger.error(f"Failed to build system prompt: {e}")
        return "You are SEA, an elite Gallery Finance Analyst. Provide concise, data-driven insights."

def generate_fallback_analysis(exhibition_code: str, user_query: str) -> str:
    """Generate a rich, data-driven Markdown response using live SQLite metrics when no OpenAI API key is set."""
    try:
        ex = el.get_exhibition(exhibition_code)
        metrics = el.calculate_inventory_metrics(exhibition_code)
        report = el.calculate_report(exhibition_code)
        artist_roi = el.calculate_artist_roi(exhibition_code)
        forecast = el.get_forecast_metrics(exhibition_code)
        ex_name = ex["name"] if ex else exhibition_code
        q = user_query.lower()

        if "artist" in q or "roi" in q or "pay" in q:
            lines = [f"### 🎨 Artist ROI & Contribution Analysis — **{ex_name}**\n"]
            lines.append(f"**Total Artworks:** {metrics['total_artworks']} | **Sold:** {metrics['sold_artworks']} ({metrics['sell_through_rate_pct']:.1f}% Sell-Through)\n")
            lines.append("| Artist | Sold / Total | Gross Sales | Gallery Revenue | Net Contribution |")
            lines.append("|---|---|---|---|---|")
            for r in artist_roi:
                lines.append(f"| **{r['artist']}** | {r['sold_artworks']}/{r['total_artworks']} | {el.money(r['gross_sales'])} | {el.money(r['gallery_revenue'])} | **{el.money(r['net_contribution'])}** |")
            lines.append("\n> **Proactive Insight:** Prioritize promotion for artists with high unsolds to maximize gallery revenue before exhibition close.")
            return "\n".join(lines)

        elif "burn" in q or "forecast" in q or "project" in q:
            return f"""### 🔮 Financial Forecast & Burn Rate — **{ex_name}**

- **Daily Burn Rate:** `{el.money(forecast['daily_burn_rate'])}/day`
- **Projected Final Net Profit:** **{el.money(forecast['projected_net_profit'])}**
- **Current Cash Collected:** **{el.money(metrics['cash_collected_thb'])}**
- **Pending Receivables:** **{el.money(metrics['receivables_thb'])}**

> 💡 **Recommendation:** To achieve target net profit, review remaining uncollected receivables ({el.money(metrics['receivables_thb'])}) and follow up with buyers.
"""

        elif "expense" in q or "cost" in q or "head" in q:
            confirmed = el.list_confirmed_expenses(exhibition_code)
            breakdown: Dict[str, float] = {}
            for row in confirmed:
                head = row["account_head"]
                breakdown[head] = breakdown.get(head, 0.0) + float(row["amount_thb"] or 0)
            sorted_exp = sorted(breakdown.items(), key=lambda x: x[1], reverse=True)
            
            exp_lines = [f"### 💸 Expense Breakdown — **{ex_name}**\n"]
            exp_lines.append(f"- **Direct Costs:** {el.money(report['totals']['direct_costs'])}")
            exp_lines.append(f"- **Operating Expenses:** {el.money(report['totals']['operating_expenses'])}")
            exp_lines.append(f"- **Total Confirmed Expenses:** **{el.money(report['totals']['direct_costs'] + report['totals']['operating_expenses'])}**\n")
            exp_lines.append("**Top Cost Centers:**")
            for head, amt in sorted_exp[:6]:
                exp_lines.append(f"- **{head}:** {el.money(amt)}")
            return "\n".join(exp_lines)

        else:
            return f"""### 🌊 Executive Financial Summary — **{ex_name}**

- **Gross Sales:** **{el.money(report['totals']['gross_sales'])}**
- **Gallery Net Revenue:** **{el.money(report['totals']['gallery_revenue'])}**
- **Total Expenses:** **{el.money(report['totals']['direct_costs'] + report['totals']['operating_expenses'] + report['totals']['allocated_overhead'])}**
- **Net Profit:** **{el.money(report['totals']['net_profit'])}** (`{report['totals']['net_margin_pct']:.1f}%` Margin)
- **Artist Outstanding Payables:** **{el.money(report['totals']['artist_outstanding_total'])}**
- **Sell-Through Rate:** **{metrics['sell_through_rate_pct']:.1f}%** ({metrics['sold_artworks']}/{metrics['total_artworks']} works sold)

#### 🚀 Key Proactive Recommendations:
1. **Follow up on Receivables:** Outstanding balance of **{el.money(metrics['receivables_thb'])}** ready for collection.
2. **Review Burn Rate:** Daily operational burn rate is **{el.money(forecast['daily_burn_rate'])}**.
"""
    except Exception as e:
        logger.error(f"Fallback analysis error: {e}")
        return f"### 🌊 Gallery Financial Assistant\n\n- **Exhibition Code:** `{exhibition_code}`\n- **System Status:** Ready\n\nHow can I help you optimize gallery finances today?"

async def chat_with_analyst(exhibition_code: str, messages: List[Dict[str, str]]) -> str:
    """Send chat history to LLM or use smart fallback analyst."""
    user_query = messages[-1]["content"] if messages else "Executive Summary"

    if not OPENAI_API_KEY:
        return generate_fallback_analysis(exhibition_code, user_query)

    system_prompt = build_analyst_system_prompt(exhibition_code)
    
    payload = {
        "model": "gpt-4o-mini",
        "messages": [{"role": "system", "content": system_prompt}] + messages,
        "temperature": 0.7,
        "max_tokens": 1000
    }
    
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers=headers,
            json=payload,
            timeout=30
        )
        response.raise_for_status()
        result = response.json()
        return result["choices"][0]["message"]["content"]
    except Exception as e:
        logger.error(f"AI Chat failed, switching to fallback analyst: {e}")
        return generate_fallback_analysis(exhibition_code, user_query)
