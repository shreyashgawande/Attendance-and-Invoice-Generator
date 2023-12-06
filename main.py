from fastapi import (FastAPI,UploadFile,File,Query)
import uvicorn
import io
from fastapi.responses import FileResponse,StreamingResponse
from generate_invoice import *
from generate_attendance import *

app = FastAPI(title='Automate Invoice and Attendance',
              summary='This is an api to automate the manual work of creating invoice and attendance every month',
              version='0.0.1')


@app.post('/generate_invoice')
async def invoice_generator(leavesTaken:int = Query(0,description="Number of leaves takes"),
                     salary:int = Query(40000,description="Monthly salary in INR"),
                     file:UploadFile = File(...)):
    invoice_file = file.file
    result_file_path,invoice_bytes = generate_invoice_excel(invoice_file,leavesTaken,salary)
    invoice_bytes.seek(0)
    return StreamingResponse(io.BytesIO(invoice_bytes.read()),media_type="xlsx",headers={"Content-Disposition":f"attachment; filename={result_file_path}"})


@app.post('/generate_attendance')
async def attendance_generator(holidays:list[int] = Query([],description="Number of Holidays taken"),
                               name:str = Query("",description="Please enter your name")):
    attendance_sheet_path,excel_bytes = generate_attendance(holidays,name)
    print(type(excel_bytes))
    excel_bytes.seek(0)
    return StreamingResponse(io.BytesIO(excel_bytes.read()), media_type="xlsx", headers={"Content-Disposition": f"attachment; filename={attendance_sheet_path}"})
    




if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000)
