FROM python:3.10.9
WORKDIR /

COPY . /

RUN pip install -r requirements.txt

EXPOSE 8900
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8900"]