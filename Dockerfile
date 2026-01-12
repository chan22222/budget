FROM node:20-slim

# Python 설치
RUN apt-get update && apt-get install -y python3 python3-pip && \
    pip3 install msoffcrypto-tool --break-system-packages && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 의존성 설치
COPY package*.json ./
RUN npm install --omit=dev

# 소스 복사
COPY . .

# import/export 폴더 생성
RUN mkdir -p import export

EXPOSE 8080

CMD ["node", "server.js"]
