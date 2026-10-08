FROM python:3.11-slim
WORKDIR /app
COPY . .
EXPOSE 8000
ENV PORT=8000
# 本地 dsh 模型需通过宿主机网关访问，容器内 localhost 指向自己
# 运行示例：
#   docker build -t llm-orchestrator .
#   docker run -p 8000:8000 -e OPENAI_API_KEY=sk-xxxx -e PORT=8000 llm-orchestrator
CMD ["python", "app.py"]
