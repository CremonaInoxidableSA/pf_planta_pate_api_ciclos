FROM python:3.12

RUN apt-get update && apt-get install -y --no-install-recommends \
    tzdata \
    && rm -rf /var/lib/apt/lists/*

ENV TZ=America/Santiago
WORKDIR /app
COPY . .

RUN pip install "fastapi[standard]"
RUN pip install PyMySQL
RUN pip install sqlalchemy
RUN pip install cryptography
RUN pip install opcua
RUN pip install python-jose
RUN pip install pandas openpyxl
RUN pip install bcrypt
RUN pip install aiosmtplib
RUN pip install Pillow
RUN pip install --upgrade pip

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]