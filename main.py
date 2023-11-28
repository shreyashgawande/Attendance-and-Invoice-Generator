from fastapi import (FastAPI,UploadFile,File,Query)
import uvicorn
from fastapi.responses import FileResponse
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
    # invoice_file_path = os.path.join(os.getcwd(),file.filename)
    # with open(invoice_file_path,"wb") as f:
    #     f.write(file.file.read())
    #     f.close() 
    result_file_path = generate_invoice_excel(invoice_file,leavesTaken,salary)
    return FileResponse(result_file_path,media_type="application/pdf",filename=f'{result_file_path}')


@app.post('/generate_attendance')
async def attendance_generator(holidays:list[int] = Query([],description="Number of Holidays taken"),
                               name:str = Query("",description="Please enter your name")):
    
    attendance_sheet_path = generate_attendance(holidays,name)
    return FileResponse(attendance_sheet_path,media_type="xlsx",filename=f'{attendance_sheet_path}')

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
  
