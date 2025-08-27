from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
import os
from io import StringIO

app = FastAPI()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))


@app.get("/", response_class=HTMLResponse)
async def top(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


# generate_docxでWordファイルを一時生成し、ダウンロードさせる（/download/outsourcing）
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
    content = f"""会社名: {company_name}\n会社住所: {company_address}\n代表者名: {representative_name}\nクライアント名: {client_name}\nクライアント住所: {client_address}\n業務内容: {service_description}\n金額: {price}\n支払方法: {payment_method}\n開始日: {start_date}\n終了日: {end_date}\n契約日: {year}年{month}月{day}日\n"""
    file_like = StringIO(content)
    headers = {"Content-Disposition": 'attachment; filename="contract_info.txt"'}
    return StreamingResponse(file_like, media_type="text/plain", headers=headers)


# outsourcing.docxをダウンロードさせるエンドポイント
@app.get("/template-download/outsourcing")
async def download_outsourcing():
    file_path = os.path.join(BASE_DIR, "contract_templates", "outsourcing.docx")
    file_like = open(file_path, "rb")
    headers = {"Content-Disposition": 'attachment; filename="outsourcing.docx"'}
    return StreamingResponse(
        file_like,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )


# サーバ起動用（開発用）
if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
