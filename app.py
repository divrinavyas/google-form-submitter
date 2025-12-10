import os
import tempfile
from fastapi import FastAPI, File, UploadFile, Form, HTTPException, BackgroundTasks, Request
from fastapi.responses import JSONResponse
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
import uuid

from submit_google_form import GoogleFormSubmitter

app = FastAPI(
    title="Google Form Auto-Submitter API",
    description="API to automatically submit Google Forms using data from Excel files",
    version="1.0.0"
)

# Setup templates
templates = Jinja2Templates(directory="templates")

# Store job status in memory (use Redis/DB for production)
job_status = {}


class JobStatus(BaseModel):
    job_id: str
    status: str  # "pending", "running", "completed", "failed"
    total_rows: int = 0
    success_count: int = 0
    fail_count: int = 0
    current_row: int = 0
    message: str = ""
    errors: list = []


class SubmissionResponse(BaseModel):
    job_id: str
    message: str


def run_submission_task(job_id: str, form_url: str, excel_path: str):
    """
    Background task to run form submissions
    """
    try:
        job_status[job_id]["status"] = "running"
        
        # Use headless mode for deployment
        submitter = GoogleFormSubmitter(form_url, excel_path, headless=True)
        
        # Callback to update progress
        def progress_callback(current_row, total_rows, success, fail, message=""):
            job_status[job_id].update({
                "current_row": current_row,
                "total_rows": total_rows,
                "success_count": success,
                "fail_count": fail,
                "message": message
            })
        
        result = submitter.run(progress_callback=progress_callback)
        
        job_status[job_id].update({
            "status": "completed",
            "success_count": result["success_count"],
            "fail_count": result["fail_count"],
            "total_rows": result["total_rows"],
            "message": f"Completed: {result['success_count']} successful, {result['fail_count']} failed",
            "errors": result.get("errors", [])
        })
        
    except Exception as e:
        job_status[job_id].update({
            "status": "failed",
            "message": str(e)
        })
    finally:
        # Clean up temp file
        if os.path.exists(excel_path):
            os.remove(excel_path)


@app.post("/submit-form", response_model=SubmissionResponse)
async def submit_form(
    background_tasks: BackgroundTasks,
    form_url: str = Form(..., description="Google Form URL"),
    excel_file: UploadFile = File(..., description="Excel file with form data")
):
    """
    Submit a Google Form using data from an uploaded Excel file.
    
    - **form_url**: The Google Form URL (must be a viewform link)
    - **excel_file**: Excel file (.xlsx or .xls) containing the data to submit
    
    Returns a job_id that can be used to check the submission status.
    """
    # Validate form URL
    if "docs.google.com/forms" not in form_url:
        raise HTTPException(status_code=400, detail="Invalid Google Form URL")
    
    # Validate file type
    if not excel_file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="File must be an Excel file (.xlsx or .xls)")
    
    # Save uploaded file to temp location
    try:
        temp_dir = tempfile.gettempdir()
        temp_file_path = os.path.join(temp_dir, f"{uuid.uuid4()}_{excel_file.filename}")
        
        with open(temp_file_path, "wb") as f:
            content = await excel_file.read()
            f.write(content)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to save uploaded file: {str(e)}")
    
    # Create job
    job_id = str(uuid.uuid4())
    job_status[job_id] = {
        "job_id": job_id,
        "status": "pending",
        "total_rows": 0,
        "success_count": 0,
        "fail_count": 0,
        "current_row": 0,
        "message": "Job queued",
        "errors": []
    }
    
    # Start background task
    background_tasks.add_task(run_submission_task, job_id, form_url, temp_file_path)
    
    return SubmissionResponse(
        job_id=job_id,
        message="Form submission job started. Use /status/{job_id} to check progress."
    )


@app.get("/status/{job_id}", response_model=JobStatus)
async def get_job_status(job_id: str):
    """
    Get the status of a form submission job.
    
    - **job_id**: The job ID returned from /submit-form
    """
    if job_id not in job_status:
        raise HTTPException(status_code=404, detail="Job not found")
    
    return JobStatus(**job_status[job_id])


@app.post("/submit-form-sync")
async def submit_form_sync(
    form_url: str = Form(..., description="Google Form URL"),
    excel_file: UploadFile = File(..., description="Excel file with form data")
):
    """
    Submit a Google Form synchronously (waits for completion).
    Use this for small datasets. For large datasets, use /submit-form instead.
    
    - **form_url**: The Google Form URL
    - **excel_file**: Excel file with form data
    """
    # Validate form URL
    if "docs.google.com/forms" not in form_url:
        raise HTTPException(status_code=400, detail="Invalid Google Form URL")
    
    # Validate file type
    if not excel_file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="File must be an Excel file (.xlsx or .xls)")
    
    # Save uploaded file to temp location
    temp_file_path = None
    try:
        temp_dir = tempfile.gettempdir()
        temp_file_path = os.path.join(temp_dir, f"{uuid.uuid4()}_{excel_file.filename}")
        
        with open(temp_file_path, "wb") as f:
            content = await excel_file.read()
            f.write(content)
        
        # Run submission
        submitter = GoogleFormSubmitter(form_url, temp_file_path)
        result = submitter.run()
        
        return JSONResponse(content={
            "status": "completed",
            "total_rows": result["total_rows"],
            "success_count": result["success_count"],
            "fail_count": result["fail_count"],
            "errors": result.get("errors", [])
        })
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if temp_file_path and os.path.exists(temp_file_path):
            os.remove(temp_file_path)


@app.get("/")
async def root(request: Request):
    """Serve the main UI page"""
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/api")
async def api_info():
    """API Health Check"""
    return {
        "status": "running",
        "message": "Google Form Auto-Submitter API",
        "endpoints": {
            "POST /submit-form": "Submit form asynchronously (returns job_id)",
            "GET /status/{job_id}": "Check job status",
            "POST /submit-form-sync": "Submit form synchronously (waits for completion)"
        }
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
