import os
from dotenv import load_dotenv
from langchain_google_genai import ChatGoogleGenerativeAI

# .env 파일의 환경변수를 불러옴
load_dotenv()

# LLM 객체 생성 (모델명: gemini-1.5-flash)
llm = ChatGoogleGenerativeAI(model="gemini-1.5-flash")

# 질문
response = llm.invoke("AI 엔지니어가 되기 위한 가장 중요한 습관 알려줘.")

print(response.content)