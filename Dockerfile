FROM python:3.11-slim

WORKDIR /app

# 필요한 파일들 복사
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# 스트림릿 기본 포트 설정 (8080으로 설정해야 구글 클라우드와 연동됨)
EXPOSE 8080

CMD ["streamlit", "run", "Meeting_assistant_external.py", "--server.port=8080", "--server.address=0.0.0.0"]