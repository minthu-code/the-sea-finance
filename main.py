import os
import logging
import threading
import time
from typing import List, Dict, Any, Optional
from fastapi import FastAPI, Request, HTTPException, Depends, Form, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
import requests
import uvicorn
from contextlib import asynccontextmanager
import exhibitledger as el
import ai_analyst

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=os.environ.get("LOG_LEVEL", "INFO"),
)
logger = logging.getLogger(__name__)

# Authentication
DASHBOARD_PASSWORD = os.environ.get("DASHBOARD_PASSWORD", "")

def verify_password(request: Request):
    if not DASHBOARD_PASSWORD:
        return True
    token = request.cookies.get("auth_token")
    if token == DASHBOARD_PASSWORD:
        return True
    raise HTTPException(status_code=401, detail="Unauthorized")

@asynccontextmanager
async def lifespan(app: FastAPI):
    el.init_db()
    logger.info("ExhibitLedger Database Initialized.")
    # Render's free plan has no persistent disk: the filesystem resets on
    # redeploy AND after ~15 minutes idle. This background ping keeps the
    # service warm so idle spin-down (the more common case) doesn't reset
    # the database. It does NOT protect against a redeploy — use Backup /
    # Restore in the avatar menu for that.
    def _keep_alive():
        url = os.environ.get("RENDER_EXTERNAL_URL")
        if not url:
            return
        if not url.startswith("http"):
            url = f"https://{url}"
        while True:
            time.sleep(600)
            try:
                requests.get(f"{url}/health", timeout=10)
            except Exception:
                pass
    threading.Thread(target=_keep_alive, daemon=True).start()
    yield

app = FastAPI(lifespan=lifespan)

# --- Frontend Routes ---
@app.get("/", response_class=HTMLResponse)
async def serve_dashboard():
    with open("dashboard.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

@app.get("/app.js")
async def serve_app_js():
    from fastapi.responses import Response
    with open("app.js", "r", encoding="utf-8") as f:
        return Response(content=f.read(), media_type="application/javascript")

@app.get("/health")
async def health():
    return {"status": "ok"}

@app.post("/api/login")
async def login(password: str = Form(...)):
    if not DASHBOARD_PASSWORD or password == DASHBOARD_PASSWORD:
        response = JSONResponse(content={"status": "ok"})
        response.set_cookie(key="auth_token", value=password, max_age=30 * 24 * 3600, httponly=True, samesite="lax")
        return response
    raise HTTPException(status_code=401, detail="Invalid password")

@app.get("/api/session")
async def api_session(request: Request):
    """Lets the frontend check auth state without triggering a 401-driven redirect loop."""
    return {"password_required": bool(DASHBOARD_PASSWORD), "authenticated": verify_password_soft(request)}

def verify_password_soft(request: Request) -> bool:
    if not DASHBOARD_PASSWORD:
        return True
    return request.cookies.get("auth_token") == DASHBOARD_PASSWORD

# --- Exhibitions ---

@app.get("/api/exhibitions", dependencies=[Depends(verify_password)])
async def api_get_exhibitions():
    rows = el.list_exhibitions()
    return {"exhibitions": [dict(r) for r in rows]}

@app.post("/api/exhibitions", dependencies=[Depends(verify_password)])
async def api_create_exhibition(request: Request):
    data = await request.json()
    try:
        row = el.create_exhibition(
            code=data.get("code"),
            name=data.get("name"),
            location=data.get("location"),
            start_date=data.get("start_date"),
            end_date=data.get("end_date"),
            notes=data.get("notes"),
        )
        return {"status": "success", "exhibition": dict(row)}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.post("/api/exhibitions/close", dependencies=[Depends(verify_password)])
async def api_close_exhibition(request: Request):
    data = await request.json()
    code = data.get("exhibition_code")
    try:
        el.close_exhibition_in_db(code)
        return {"status": "success"}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.get("/api/readiness", dependencies=[Depends(verify_password)])
async def api_readiness(exhibition_code: str):
    try:
        checks = el.data_quality_checks(exhibition_code)
        blocking = [c for c in checks if "No blocking issues" not in c]
        return {"checks": checks, "ready": len(blocking) == 0}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.get("/api/closeout", dependencies=[Depends(verify_password)])
async def api_closeout(exhibition_code: str):
    try:
        return el.get_exhibition_closeout_status(exhibition_code)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

# --- Commission splits ---

@app.get("/api/splits", dependencies=[Depends(verify_password)])
async def api_get_splits(exhibition_code: str):
    rows = el.get_split_rules(exhibition_code)
    return {"splits": [dict(r) for r in rows]}

@app.post("/api/splits", dependencies=[Depends(verify_password)])
async def api_set_splits(request: Request):
    data = await request.json()
    code = data.get("exhibition_code")
    entries = data.get("entries", [])
    try:
        rows = el.set_commission_splits(code, entries)
        return {"status": "success", "splits": [dict(r) for r in rows]}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

# --- Artworks / Inventory ---

@app.get("/api/artworks", dependencies=[Depends(verify_password)])
async def api_get_artworks(exhibition_code: str):
    rows = el.list_artworks(exhibition_code)
    return {"artworks": [dict(r) for r in rows]}

@app.post("/api/artworks", dependencies=[Depends(verify_password)])
async def api_add_artwork(request: Request):
    data = await request.json()
    code = data.get("exhibition_code")
    title = data.get("title")
    artist = data.get("artist")
    price = data.get("price")
    try:
        row = el.add_artwork(code, title, artist, float(price))
        return {"status": "success", "artwork": dict(row)}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.post("/api/artworks/bulk", dependencies=[Depends(verify_password)])
async def api_bulk_import_artworks(request: Request):
    data = await request.json()
    code = data.get("exhibition_code")
    raw_data = data.get("raw_data", "")
    artworks = []
    for line in raw_data.strip().split('\n'):
        if not line.strip():
            continue
        parts = [p.strip() for p in line.split(',')]
        if len(parts) >= 3:
            try:
                artworks.append({"title": parts[0], "artist": parts[1], "price": float(parts[2].replace(",", ""))})
            except Exception:
                continue
    if not code or not artworks:
        raise HTTPException(status_code=400, detail="Invalid data. Use: Title, Artist, Price per line.")
    try:
        ids = el.bulk_add_artworks(code, artworks)
        return {"status": "success", "count": len(ids)}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

# --- Sales & Receivables ---

@app.post("/api/sales", dependencies=[Depends(verify_password)])
async def api_record_sale(request: Request):
    data = await request.json()
    try:
        result = el.record_sale(
            artwork_id=int(data.get("artwork_id")),
            actual_price_thb=float(data.get("price")),
            buyer_name=data.get("buyer"),
            amount_collected_thb=(float(data["collected"]) if data.get("collected") not in (None, "") else None),
            payment_method=data.get("payment_method"),
            notes=data.get("notes"),
        )
        return {"status": "success", "sale": result["sale"]}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.get("/api/sales", dependencies=[Depends(verify_password)])
async def api_list_sales(exhibition_code: str):
    rows = el.list_sales(exhibition_code)
    return {"sales": [dict(r) for r in rows]}

@app.get("/api/receivables", dependencies=[Depends(verify_password)])
async def api_receivables(exhibition_code: str):
    return {"receivables": el.get_receivables_report(exhibition_code)}

@app.post("/api/sales/collect", dependencies=[Depends(verify_password)])
async def api_collect_payment(request: Request):
    data = await request.json()
    try:
        row = el.collect_sale_payment(int(data.get("sale_id")), float(data.get("amount")), data.get("notes"))
        return {"status": "success", "sale": dict(row)}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

# --- Expenses ---

@app.get("/api/account_heads", dependencies=[Depends(verify_password)])
async def api_account_heads():
    return {"account_heads": el.account_head_names()}

@app.post("/api/expenses", dependencies=[Depends(verify_password)])
async def api_add_expense(request: Request):
    """Quick-add: creates a pending expense and immediately confirms it."""
    data = await request.json()
    code = data.get("exhibition_code")
    amount = data.get("amount")
    description = data.get("description")
    artist_tag = data.get("artist_tag") or None
    recipient = data.get("recipient") or None
    category = data.get("category") or None
    try:
        row = el.create_pending_expense(code, f"{amount} {description}", artist_tag=artist_tag, recipient=recipient, category=category)
        if data.get("account_head"):
            el.update_pending_account(row["id"], data["account_head"])
        el.confirm_pending_expense(row["id"])
        return {"status": "success"}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.get("/api/expenses/pending", dependencies=[Depends(verify_password)])
async def api_pending_expenses(exhibition_code: str):
    rows = el.list_pending_expenses(exhibition_code)
    return {"pending": [dict(r) for r in rows]}

@app.get("/api/expenses/confirmed", dependencies=[Depends(verify_password)])
async def api_confirmed_expenses(exhibition_code: str):
    rows = el.list_confirmed_expenses(exhibition_code)
    return {"confirmed": [dict(r) for r in rows]}

@app.post("/api/expenses/confirm", dependencies=[Depends(verify_password)])
async def api_confirm_expense(request: Request):
    data = await request.json()
    try:
        pending_id = int(data.get("pending_id"))
        if data.get("account_head"):
            el.update_pending_account(pending_id, data["account_head"])
        if data.get("amount") not in (None, ""):
            el.update_pending_amount(pending_id, float(data["amount"]))
        row = el.confirm_pending_expense(pending_id)
        return {"status": "success", "expense": dict(row)}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.post("/api/expenses/ignore", dependencies=[Depends(verify_password)])
async def api_ignore_expense(request: Request):
    data = await request.json()
    try:
        row = el.ignore_pending_expense(int(data.get("pending_id")))
        return {"status": "success", "expense": dict(row)}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

# --- Budgets ---

@app.get("/api/budgets", dependencies=[Depends(verify_password)])
async def api_budgets(exhibition_code: str):
    return {"budgets": el.calculate_budget_report(exhibition_code)}

@app.post("/api/budgets", dependencies=[Depends(verify_password)])
async def api_set_budget(request: Request):
    data = await request.json()
    try:
        row = el.set_expense_budget(data.get("exhibition_code"), data.get("account_head"), float(data.get("budget")), data.get("notes"))
        return {"status": "success", "budget": dict(row)}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

# --- Artist ROI & Payables ---

@app.get("/api/artist_roi", dependencies=[Depends(verify_password)])
async def api_artist_roi(exhibition_code: str):
    return {"artist_roi": el.calculate_artist_roi(exhibition_code)}

@app.get("/api/payables", dependencies=[Depends(verify_password)])
async def api_payables(exhibition_code: str):
    rows = el.get_artist_payables(exhibition_code)
    return {"payables": [dict(r) for r in rows]}

@app.post("/api/payables/pay", dependencies=[Depends(verify_password)])
async def api_pay_artist(request: Request):
    data = await request.json()
    try:
        row = el.record_artist_payment(int(data.get("payable_id")), float(data.get("amount")))
        return {"status": "success", "payable": dict(row)}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

# --- Dashboard aggregate ---

@app.get("/api/dashboard", dependencies=[Depends(verify_password)])
async def api_dashboard(exhibition_code: str):
    try:
        metrics = el.calculate_inventory_metrics(exhibition_code)
        report = el.calculate_report(exhibition_code)
        raw_cashflow = el.get_cash_flow_timeline(exhibition_code)
        artist_roi = el.calculate_artist_roi(exhibition_code)
        forecast = el.get_forecast_metrics(exhibition_code)
        checks = el.data_quality_checks(exhibition_code)
        blocking = [c for c in checks if "No blocking issues" not in c]
        budgets = el.calculate_budget_report(exhibition_code)

        labels = [item["date"] for item in raw_cashflow]
        running = 0.0
        values = []
        for item in raw_cashflow:
            running += item["amount"]
            values.append(round(running, 2))

        # Expense breakdown by account head (for the donut) — confirmed expenses only.
        confirmed = el.list_confirmed_expenses(exhibition_code)
        breakdown: Dict[str, float] = {}
        for row in confirmed:
            breakdown[row["account_head"]] = breakdown.get(row["account_head"], 0.0) + float(row["amount_thb"] or 0)
        expense_breakdown = sorted(
            [{"account_head": k, "amount": v} for k, v in breakdown.items()],
            key=lambda x: x["amount"], reverse=True
        )

        return {
            "metrics": {
                "available_count": metrics["available_artworks"],
                "available_value": metrics["unsold_asking_value_thb"],
                "total_count": metrics["total_artworks"],
                "sold_count": metrics["sold_artworks"],
                "sell_through_rate": metrics["sell_through_rate_pct"],
                "gross_sales": metrics["gross_sales_thb"],
                "cash_collected": metrics["cash_collected_thb"],
                "receivables": metrics["receivables_thb"],
                "pending_receipts": metrics["pending_receipts"],
                "average_sale_price": metrics["average_sale_price_thb"],
            },
            "report": {
                "revenue": report["totals"]["gallery_revenue"],
                "gross_sales": report["totals"]["gross_sales"],
                "direct_costs": report["totals"]["direct_costs"],
                "operating_expenses": report["totals"]["operating_expenses"],
                "expenses": report["totals"]["direct_costs"] + report["totals"]["operating_expenses"] + report["totals"]["allocated_overhead"],
                "net_profit": report["totals"]["net_profit"],
                "net_margin_pct": report["totals"]["net_margin_pct"],
                "artist_outstanding_total": report["totals"]["artist_outstanding_total"],
            },
            "artist_roi": artist_roi,
            "forecast": {
                **forecast,
                "projection_labels": labels,
                "projection_values": values,
            },
            "readiness": {"checks": checks, "ready": len(blocking) == 0, "blocking_count": len(blocking)},
            "expense_breakdown": expense_breakdown,
            "budget_alerts": [b for b in budgets if b["budget_thb"] > 0 and b["variance_thb"] < 0],
        }
    except Exception as e:
        logger.exception("Dashboard API failed")
        raise HTTPException(status_code=400, detail=str(e))

# --- SEA AI ---

@app.post("/api/ai/analyze", dependencies=[Depends(verify_password)])
async def api_ai_analyze(request: Request):
    data = await request.json()
    code = data.get("exhibition_code")
    query = data.get("query")
    history = data.get("history", [])
    messages = [{"role": h["role"], "content": h["content"]} for h in history]
    messages.append({"role": "user", "content": query})
    analysis = await ai_analyst.chat_with_analyst(code, messages)
    return {"analysis": analysis}

# --- Export ---

@app.get("/api/export", dependencies=[Depends(verify_password)])
async def api_export(exhibition_code: str):
    try:
        export_dir = os.environ.get("EXPORT_DIR", "./exports")
        file_path = el.export_report_xlsx(exhibition_code, export_dir)
        return FileResponse(file_path, filename=os.path.basename(file_path))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# --- Backup / restore (Render free plan has no persistent disk) ---

@app.get("/api/backup/download", dependencies=[Depends(verify_password)])
async def api_backup_download():
    try:
        path = el.backup_db_file_path()
        if not os.path.exists(path):
            raise HTTPException(status_code=404, detail="No database file found yet.")
        stamp = time.strftime("%Y%m%d_%H%M%S")
        return FileResponse(path, filename=f"exhibitledger_backup_{stamp}.db", media_type="application/octet-stream")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/backup/restore", dependencies=[Depends(verify_password)])
async def api_backup_restore(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        el.restore_db_file(contents)
        return {"status": "success"}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))
