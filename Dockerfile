# 1. Lightweight Python Linux OS
FROM python:3.11-slim

# 2. Install LibreOffice for headless .doc -> .docx conversion
RUN apt-get update && \
    apt-get install -y libreoffice --no-install-recommends && \
    rm -rf /var/lib/apt/lists/*

# 3. Set Working Directory
WORKDIR /app

# 4. Install Python Dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 5. Copy Application Files
COPY app.py .

# 6. Expose Port
EXPOSE 8501

# 7. Launch Command
CMD ["streamlit", "run", "app.py", "--server.address=0.0.0.0"]
