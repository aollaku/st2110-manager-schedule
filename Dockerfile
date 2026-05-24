FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV FLASK_ENV=production
ENV ST2110_SECRET_KEY=1312-12312312312-123124312312-3123123123123
RUN chmod 777 run_mac_linux.sh
EXPOSE 5005

ENTRYPOINT ["./run_mac_linux.sh"]