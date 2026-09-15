import json
import os
from datetime import datetime, timezone
from typing import Literal
from uuid import uuid4

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Path, Query, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field
from azure.cosmos import CosmosClient

load_dotenv()

app = FastAPI(
    title="RAG Search API",
    version="1.0.0",
    description="東方Projectに関する質問をRAG検索し、回答と情報源を返すAPI",
)

QuestionStatus = Literal["processing", "completed", "failed"]
Evaluation = Literal["correct", "partially_correct", "incorrect"]


class ErrorResponse(BaseModel):
    error: str
    message: str


class QuestionCreate(BaseModel):
    question: str = Field(min_length=1, max_length=1000)
    status: QuestionStatus = "processing"


class Question(BaseModel):
    id: str
    question: str
    answer: str = ""
    sources: list[str] = Field(default_factory=list)
    status: QuestionStatus
    created_at: datetime
    updated_at: datetime


class QuestionCreated(BaseModel):
    id: str
    message: str


class EvaluationCreate(BaseModel):
    question_id: str
    evaluation: Evaluation
    reason: str | None = None
    correct_answer: str | None = None


class MessageResponse(BaseModel):
    message: str


class SearchResponse(BaseModel):
    answer: str
    sources: list[str]


questions: dict[str, Question] = {}
evaluations: list[EvaluationCreate] = []


@app.exception_handler(RequestValidationError)
async def validation_exception_handler(
    request: Request, exc: RequestValidationError
) -> JSONResponse:
    return JSONResponse(
        status_code=400,
        content={
            "error": "INVALID_PARAMETER",
            "message": "リクエストパラメータが不正です",
        },
    )


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


def _error(status_code: int, error: str, message: str) -> HTTPException:
    return HTTPException(
        status_code=status_code,
        detail={"error": error, "message": message},
    )


@app.exception_handler(HTTPException)
async def http_exception_handler(request: Request, exc: HTTPException) -> JSONResponse:
    if isinstance(exc.detail, dict) and "error" in exc.detail:
        content = exc.detail
    else:
        content = {"error": "REQUEST_ERROR", "message": str(exc.detail)}
    return JSONResponse(status_code=exc.status_code, content=content)


@app.post("/qa/{id}", response_model=QuestionCreated, status_code=201)
async def create_question(
    payload: QuestionCreate,
    id: str = Path(min_length=1, description="質問のID"),
) -> QuestionCreated:
    now = _utc_now()
    existing_question = questions.get(id)
    if existing_question is None:
        question = Question(
            id=id,
            question=payload.question,
            status=payload.status,
            created_at=now,
            updated_at=now,
        )
    else:
        question = existing_question.model_copy(
            update={
                "question": payload.question,
                "status": payload.status,
                "updated_at": now,
            }
        )
    questions[id] = question
    return QuestionCreated(id=id, message="質問の登録が成功しました")


@app.get("/qa/{id}", response_model=Question)
async def get_question(id: str = Path(min_length=1, description="質問のID")) -> Question:
    question = questions.get(id)
    if question is None:
        raise _error(404, "NOT_FOUND", "指定されたIDの質問が見つかりませんでした")
    return question


@app.post("/evaluate", response_model=MessageResponse, status_code=201)
async def evaluate_question(payload: EvaluationCreate) -> MessageResponse:
    if payload.question_id not in questions:
        raise _error(404, "NOT_FOUND", "指定されたIDの質問が見つかりませんでした")
    evaluations.append(payload)
    return MessageResponse(message="評価の投稿が成功しました")


def _search_with_azure(question: str) -> SearchResponse | None:
    required = (
        "AZURE_SEARCH_API_KEY",
        "AZURE_SEARCH_SERVICE_ENDPOINT",
        "AZURE_SEARCH_INDEX_NAME",
        "AZURE_OPENAI_ENDPOINT",
        "AZURE_OPENAI_API_KEY",
        "AZURE_OPENAI_API_VERSION",
        "AZURE_OPENAI_CHAT_DEPLOYMENT",
    )
    if not all(os.getenv(name) for name in required):
        return None

    from azure.core.credentials import AzureKeyCredential
    from azure.search.documents import SearchClient
    from openai import AzureOpenAI

    search_client = SearchClient(
        endpoint=os.environ["AZURE_SEARCH_SERVICE_ENDPOINT"],
        index_name=os.environ["AZURE_SEARCH_INDEX_NAME"],
        credential=AzureKeyCredential(os.environ["AZURE_SEARCH_API_KEY"]),
    )
    results = list(search_client.search(question, top=4))
    sources = [str(result.get("source", result.get("title", ""))) for result in results]
    context = "\n\n".join(
        json.dumps(dict(result), ensure_ascii=False, default=str) for result in results
    )
    client = AzureOpenAI(
        azure_endpoint=os.environ["AZURE_OPENAI_ENDPOINT"],
        api_key=os.environ["AZURE_OPENAI_API_KEY"],
        api_version=os.environ["AZURE_OPENAI_API_VERSION"],
    )
    response = client.chat.completions.create(
        model=os.environ["AZURE_OPENAI_CHAT_DEPLOYMENT"],
        temperature=0.2,
        messages=[
            {"role": "system", "content": "検索結果だけを根拠に日本語で回答してください。"},
            {"role": "user", "content": f"検索結果:\n{context}\n\n質問: {question}"},
        ],
    )
    return SearchResponse(
        answer=response.choices[0].message.content or "回答を生成できませんでした。",
        sources=[source for source in sources if source],
    )


@app.get("/search", response_model=SearchResponse)
async def search(
    question: str = Query(min_length=1, max_length=1000, description="RAGに問い合わせる質問文"),
) -> SearchResponse:
    result = _search_with_azure(question)
    if result is not None:
        return result
    return SearchResponse(
        answer="検索サービスが設定されていないため、回答を生成できません。",
        sources=[],
    )