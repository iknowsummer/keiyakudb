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


@app.get("/hello", response_class=HTMLResponse)
async def hello(request: Request):
    return templates.TemplateResponse(
        "hello.html", {"request": request, "message": "Hello, World!"}
    )


# POSTでnameを受け取り、10回書いたテキストファイルをダウンロードさせる
@app.post("/dltest")
async def download_file(name: str = Form(...)):
    content = "\n".join([name] * 20)
    file_like = StringIO(content)
    headers = {"Content-Disposition": 'attachment; filename="names.txt"'}
    return StreamingResponse(file_like, media_type="text/plain", headers=headers)


# サーバ起動用（開発用）
if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
