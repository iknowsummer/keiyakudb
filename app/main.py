# FastAPIアプリ本体
import anyio
import os
from fastapi import FastAPI, Request, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from .template_service import generate_docx_stream, convert_stream_docx2pdf

app = FastAPI()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "../templates"))


@app.get("/", response_class=HTMLResponse)
async def top(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/download/outsourcing")
async def download_contract_info(
    company_name: str = Form(...),
    company_address: str = Form(...),
    representative_name: str = Form(...),
    client_name: str = Form(...),
    client_address: str = Form(...),
    service_description: str = Form(...),
    price: str = Form(...),
    payment_method: str = Form(...),
    start_date: str = Form(...),
    end_date: str = Form(...),
    year: str = Form(...),
    month: str = Form(...),
    day: str = Form(...),
):
    context = {
        "company_name": company_name,
        "company_address": company_address,
        "representative_name": representative_name,
        "client_name": client_name,
        "client_address": client_address,
        "service_description": service_description,
        "price": price,
        "payment_method": payment_method,
        "start_date": start_date,
        "end_date": end_date,
        "year": year,
        "month": month,
        "day": day,
    }
    template_path = os.path.join(BASE_DIR, "../contract_templates", "outsourcing.docx")
    if not os.path.exists(template_path):
        raise HTTPException(400, "template not found")
    docx_like = generate_docx_stream(template_path, context)
    try:
        pdf_like = await anyio.to_thread.run_sync(convert_stream_docx2pdf, docx_like)
    except Exception:
        raise HTTPException(502, "PDF変換に失敗しました")
    pdf_like.seek(0)
    headers = {"Content-Disposition": 'attachment; filename="outsourcing.pdf"'}
    return StreamingResponse(
        pdf_like,
        media_type="application/pdf",
        headers=headers,
    )


@app.get("/template-download/outsourcing")
async def download_outsourcing():
    file_path = os.path.join(BASE_DIR, "../contract_templates", "outsourcing.docx")
    file_like = open(file_path, "rb")
    headers = {"Content-Disposition": 'attachment; filename="outsourcing.docx"'}
    return StreamingResponse(
        file_like,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)
