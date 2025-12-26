
import sys
import os
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass
import asyncio
import uuid
import sys
from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional

# Import the wrapper
from crawlers.wrapper import (
    run_crawler_task, log_queues, get_log_queue, clear_log_queue,
    get_crawl_result, clear_crawl_result, WConceptCrawler, set_stop_signal
)

app = FastAPI(title="Lotte On Sourcing Helper")

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Path resolution
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
RESULTS_DIR = os.path.join(BASE_DIR, "results")
UPLOADS_DIR = os.path.join(BASE_DIR, "uploads")

# Create directories
for d in [RESULTS_DIR, UPLOADS_DIR]:
    if not os.path.exists(d):
        os.makedirs(d)

# Mount static if exists
if os.path.exists(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# Models
class CrawlRequest(BaseModel):
    crawler_type: str  # musinsa, wconcept, 29cm
    category: Optional[str] = None
    keyword: Optional[str] = None
    count: int = 10
    headless: bool = True

class CrawlResponse(BaseModel):
    request_id: str
    message: str

# Global State
active_tasks = {} # request_id -> task_info

@app.get("/")
async def read_root():
    index_path = os.path.join(TEMPLATES_DIR, "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path)
    return {"message": "System is running. UI not found."}

@app.post("/api/crawl", response_model=CrawlResponse)
async def start_crawl(req: CrawlRequest, background_tasks: BackgroundTasks):
    request_id = str(uuid.uuid4())
    
    # Store task info (could be expanded)
    active_tasks[request_id] = {
        "type": req.crawler_type,
        "status": "running"
    }
    
    # Init log queue
    get_log_queue(request_id)
    
    # Add background task
    background_tasks.add_task(
        run_crawler_task, 
        req.crawler_type, 
        req.dict(), 
        request_id
    )
    
    return CrawlResponse(request_id=request_id, message="Crawler started")

@app.get("/api/status/{request_id}")
async def get_status(request_id: str):
    q = get_log_queue(request_id)
    logs = []
    while not q.empty():
        try:
            logs.append(q.get_nowait())
        except:
            break
            
    return {"logs": logs, "status": active_tasks.get(request_id, {}).get("status", "unknown")}

@app.get("/api/files")
async def list_files():
    if not os.path.exists(RESULTS_DIR):
        return {"files": []}
    
    files = []
    for f in os.listdir(RESULTS_DIR):
        if f.endswith(".xlsx") or f.endswith(".csv"):
            path = os.path.join(RESULTS_DIR, f)
            stat = os.stat(path)
            files.append({
                "filename": f,
                "size": stat.st_size,
                "created": stat.st_ctime
            })
    
    # Sort by created desc
    files.sort(key=lambda x: x["created"], reverse=True)
    return {"files": files}

@app.get("/api/download/{filename}")
async def download_file(filename: str):
    file_path = os.path.join(RESULTS_DIR, filename)
    if os.path.exists(file_path):
        return FileResponse(file_path, filename=filename)
    raise HTTPException(status_code=404, detail="File not found")

@app.post("/api/stop/{request_id}")
async def stop_crawl(request_id: str):
    # Wrapper의 전역 중지 신호 설정
    set_stop_signal(request_id)
    if request_id in active_tasks:
        active_tasks[request_id]["status"] = "stopping"
    return {"message": "중지 요청이 전송되었습니다. 현재 진행 중인 작업만 정지됩니다."}

class SaveRequest(BaseModel):
    request_id: str
    save_path: str
    filename: str

@app.post("/api/save_result")
async def save_result(req: SaveRequest):
    """크롤링 결과를 사용자 지정 경로에 저장"""
    try:
        # 크롤링 결과 가져오기
        result_data = get_crawl_result(req.request_id)
        if not result_data:
            raise HTTPException(status_code=404, detail="No crawl result found for this request_id")
        
        crawler_type = result_data.get("crawler_type")
        data = result_data.get("data")
        
        if not data:
            raise HTTPException(status_code=400, detail="No data to save")
        
        # 저장 경로 생성
        save_dir = req.save_path
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        
        # 파일명 처리 (.xlsx 확장자 자동 추가)
        filename = req.filename
        if not filename.endswith('.xlsx'):
            filename += '.xlsx'
        
        filepath = os.path.join(save_dir, filename)
        
        # 크롤러 타입에 따라 저장
        import pandas as pd
        products = data.get("products", [])
        if not products:
             raise HTTPException(status_code=400, detail="No data found in result")
             
        df = pd.DataFrame(products)
        # 모든 값을 텍스트로 처리 (Excel 자동 변환 방지)
        df = df.astype(str)
        df.to_excel(filepath, index=False, engine='openpyxl')
        
        return {"message": "File saved successfully", "filepath": filepath}
        
        # 결과 삭제 (선택사항)
        # clear_crawl_result(req.request_id)
        
        return {"message": "File saved successfully", "filepath": filepath}
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
