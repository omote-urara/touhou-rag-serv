import html
import io
import os
from html.parser import HTMLParser
from pathlib import PurePosixPath
from typing import Any
from urllib.parse import quote

import requests
from azure.identity import AzureCliCredential, DefaultAzureCredential, InteractiveBrowserCredential
from dotenv import load_dotenv
from langchain_chroma import Chroma
from langchain_openai import AzureChatOpenAI, AzureOpenAIEmbeddings
from pypdf import PdfReader
from pydantic import SecretStr
from docx import Document


load_dotenv()

GRAPH_URL = "https://graph.microsoft.com/v1.0"
SUPPORTED_EXTENSIONS = {".txt", ".md", ".html", ".htm", ".pdf", ".docx"}


class _HTMLTextParser(HTMLParser):
    """HTML タグを除去し、本文だけを集めるためのパーサー。"""

    def __init__(self) -> None:
        """本文を保存する空のリストを初期化する。"""
        print("[HTMLParser] 初期化します")
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        """HTML 内で見つかったテキストを本文リストに追加する。"""
        print(f"[HTMLParser] テキストを取得しました: {len(data)} 文字")
        self.parts.append(data)


class SharePointLoader:
    """Microsoft Graph 経由で SharePoint の文書を読み込むクラス。"""

    def __init__(self) -> None:
        """認証情報、HTTP セッション、サイトとドライブの ID を準備する。"""
        print("[SharePointLoader] 初期化します")
        self.credential =  InteractiveBrowserCredential(
            client_id="bdcdbb62-44ae-46c3-8db8-54eba336eb96",
            tenant_id="def60ec5-4fd5-436e-bd4d-ccccdeb590ee",
        )
        self.session = requests.Session()
        # ドライブ ID があればサイトの表示名や URL パスを解決する必要はない。
        configured_drive_id = os.getenv("SHAREPOINT_DRIVE_ID")
        self.site_id = ""
        self.drive_id = configured_drive_id or self._resolve_drive_id()
        print("[SharePointLoader] 初期化が完了しました")

    def _headers(self) -> dict[str, str]:
        """Graph API の呼び出しに使うアクセストークン付きヘッダーを作る。"""
        print("[SharePointLoader] アクセストークンを取得します")
        token = self.credential.get_token("https://graph.microsoft.com/.default")
        print("[SharePointLoader] アクセストークンを取得しました")
        return {"Authorization": f"Bearer {token.token}"}

    def _get(self, url: str) -> dict[str, Any]:
        """Graph API に GET リクエストを送り、JSON レスポンスを返す。"""
        print(f"[Graph API] GET {url}")
        response = self.session.get(url, headers=self._headers(), timeout=60)
        if not response.ok:
            if response.status_code == 403:
                raise requests.HTTPError(
                    "Graph API のアクセス権がありません。"
                    "DefaultAzureCredential で使用されたユーザー、サービス プリンシパル、"
                    "または Managed Identity に Sites.Read.All / Files.Read.All の"
                    f"管理者同意を付与してください。URL: {url}; 詳細: {response.text}",
                    response=response,
                )
            raise requests.HTTPError(
                f"{response.status_code} for {url}: {response.text}",
                response=response,
            )
        print(f"[Graph API] 成功: {response.status_code}")
        return response.json()

    def _get_site_id(self) -> str:
        """環境変数で指定された SharePoint サイトの ID を取得する。"""
        print("[SharePointLoader] サイト ID を取得します")
        hostname = os.environ["SHAREPOINT_HOSTNAME"]
        site_path = os.environ["SHAREPOINT_SITE_PATH"].strip("/")
        url = f"{GRAPH_URL}/sites/{hostname}:/{quote(site_path, safe='/')}"
        site_id = self._get(url)["id"]
        print("[SharePointLoader] サイト ID を取得しました")
        return site_id

    def _get_drive_id(self) -> str:
        """サイト内のドキュメント ライブラリ名からドライブ ID を探す。"""
        print("[SharePointLoader] ドキュメント ライブラリを検索します")
        drives = self._get(f"{GRAPH_URL}/sites/{self.site_id}/drives")["value"]
        drive_name = os.getenv("SHAREPOINT_DRIVE_NAME", "Documents").casefold()
        for drive in drives:
            if drive["name"].casefold() == drive_name:
                print(f"[SharePointLoader] ライブラリを発見しました: {drive['name']}")
                return drive["id"]
        raise RuntimeError(f"SharePoint drive not found: {drive_name}")

    def _resolve_drive_id(self) -> str:
        """サイト ID を取得してから、対象のドライブ ID を解決する。"""
        print("[SharePointLoader] ドライブ ID を解決します")
        self.site_id = os.getenv("SHAREPOINT_SITE_ID") or self._get_site_id()
        drive_id = self._get_drive_id()
        print("[SharePointLoader] ドライブ ID を解決しました")
        return drive_id

    def _items(self, url: str):
        """Graph API のページネーションをたどってアイテムを順番に返す。"""
        print(f"[SharePointLoader] アイテム一覧を取得します: {url}")
        item_count = 0
        while url:
            page = self._get(url)
            items = page.get("value", [])
            item_count += len(items)
            print(f"[SharePointLoader] ページ取得: {len(items)} 件")
            yield from items
            url = page.get("@odata.nextLink", "")
        print(f"[SharePointLoader] アイテム一覧の取得完了: 合計 {item_count} 件")

    def _files(self, url: str):
        """フォルダーを再帰的にたどり、対応形式のファイルだけを返す。"""
        print(f"[SharePointLoader] ファイルを探索します: {url}")
        for item in self._items(url):
            if "folder" in item:
                print(f"[SharePointLoader] フォルダーに入ります: {item['name']}")
                yield from self._files(
                    f"{GRAPH_URL}/drives/{self.drive_id}/items/{item['id']}/children"
                )
            elif "file" in item and PurePosixPath(item["name"]).suffix.casefold() in SUPPORTED_EXTENSIONS:
                print(f"[SharePointLoader] 対応ファイルを発見: {item['name']}")
                yield item

    def load(self) -> list[tuple[str, str, dict[str, Any]]]:
        """SharePoint のファイルをダウンロードし、テキスト化して返す。"""
        print("[SharePointLoader] 文書の読み込みを開始します")
        root_path = os.getenv("SHAREPOINT_ROOT_PATH", "").strip("/")
        if root_path:
            encoded_path = quote(root_path, safe="/")
            url = f"{GRAPH_URL}/drives/{self.drive_id}/root:/{encoded_path}:/children"
        else:
            url = f"{GRAPH_URL}/drives/{self.drive_id}/root/children"

        documents = []
        file_count = 0
        for item in self._files(url):
            file_count += 1
            print(f"[SharePointLoader] ダウンロードします: {item['name']}")
            download_url = item.get("@microsoft.graph.downloadUrl")
            if not download_url:
                continue
            response = self.session.get(download_url, timeout=120)
            response.raise_for_status()
            text = self._extract_text(item["name"], response.content)
            if text.strip():
                documents.append((text, item["name"], item))
                print(f"[SharePointLoader] テキスト化完了: {item['name']} ({len(text)} 文字)")
        print(f"[SharePointLoader] 文書の読み込み完了: {len(documents)} 件 / ファイル {file_count} 件")
        return documents

    @staticmethod
    def _extract_text(name: str, content: bytes) -> str:
        """拡張子に応じてバイト列をテキストへ変換する。"""
        suffix = PurePosixPath(name).suffix.casefold()
        print(f"[SharePointLoader] テキスト抽出: {name} ({suffix})")
        if suffix in {".txt", ".md"}:
            return content.decode("utf-8", errors="replace")
        if suffix in {".html", ".htm"}:
            parser = _HTMLTextParser()
            parser.feed(content.decode("utf-8", errors="replace"))
            return html.unescape(" ".join(parser.parts))
        if suffix == ".pdf":
            return "\n".join(page.extract_text() or "" for page in PdfReader(io.BytesIO(content)).pages)
        if suffix == ".docx":
            return "\n".join(paragraph.text for paragraph in Document(io.BytesIO(content)).paragraphs)
        return ""


def chunk_text(text: str, size: int = 1200, overlap: int = 200) -> list[str]:
    """検索精度を保つため、文書を重なり付きの固定長チャンクへ分割する。"""
    print(f"[Chunk] 分割開始: {len(text)} 文字, size={size}, overlap={overlap}")
    chunks = []
    start = 0
    while start < len(text):
        end = min(start + size, len(text))
        chunks.append(text[start:end])
        if end == len(text):
            break
        start = end - overlap
    print(f"[Chunk] 分割完了: {len(chunks)} チャンク")
    return chunks


def answer_question(question: str, results: list[Any]) -> str:
    """検索結果を根拠として Azure OpenAI に質問し、回答文を返す。"""
    print(f"[OpenAI] 回答生成を開始します: 検索結果 {len(results)} 件")
    context = "\n\n".join(
        f"出典: {result.metadata.get('source', '不明')}\n{result.page_content}"
        for result in results
    )
    prompt = f"""以下の SharePoint 文書を根拠に質問へ回答してください。
文書に書かれていない内容は推測せず、「文書からは確認できません」と答えてください。
回答は日本語で簡潔にまとめ、必要なら出典ファイル名を示してください。

## 参照文書
{context}

## 質問
{question}
"""
    chat_model = AzureChatOpenAI(
        azure_endpoint=os.environ["AZURE_OPENAI_ENDPOINT"],
        api_key=SecretStr(os.environ["AZURE_OPENAI_API_KEY"]),
        api_version=os.environ["AZURE_OPENAI_API_VERSION"],
        azure_deployment=os.environ["AZURE_OPENAI_CHAT_DEPLOYMENT"],
    )
    response = chat_model.invoke(prompt)
    print("[OpenAI] 回答生成が完了しました")
    return str(response.content)



def main() -> None:
    """SharePoint 文書を Chroma に登録し、質問に近い文書を検索して表示する。"""
    print("[Main] RAG 処理を開始します")
    # Azure OpenAI の埋め込みモデルを初期化する。
    embeddings = AzureOpenAIEmbeddings(
        azure_endpoint=os.environ["AZURE_OPENAI_ENDPOINT"],
        api_key=SecretStr(os.environ["AZURE_OPENAI_API_KEY"]),
        api_version=os.environ["AZURE_OPENAI_API_VERSION"],
        azure_deployment=os.environ["AZURE_OPENAI_DEPLOYMENT"],
    )
    print("[Main] Azure OpenAI Embeddings を初期化しました")
    vectorstore = Chroma(
        collection_name="test_collection",
        embedding_function=embeddings,
        persist_directory="./chroma_db",
    )
    print("[Main] Chroma DB を初期化しました")

    # SharePoint から文書を取得し、チャンク単位で登録するデータを作る。
    texts: list[str] = []
    metadatas: list[dict[str, str]] = []
    ids: list[str] = []
    for text, name, item in SharePointLoader().load():
        print(f"[Main] 登録データを作成します: {name}")
        vectorstore.delete(where={"source_id": item["id"]})
        for index, chunk in enumerate(chunk_text(text)):
            texts.append(chunk)
            metadatas.append({"source": name, "source_id": item["id"]})
            ids.append(f"{item['id']}-{index}")

    # 同じファイルを再取得した場合も、古いチャンクを残さず更新する。
    if texts:
        print(f"[Main] Chroma DB へ {len(texts)} チャンクを登録します")
        vectorstore.add_texts(texts, metadatas=metadatas, ids=ids)
        print(f"[Main] Chroma DB へ {len(texts)} チャンクを登録しました")
    else:
        print("[Main] 登録する文書がありません")

    # 登録したベクトルから質問に近い上位 2 件を検索する。
    question = "幻想郷の巫女は誰？"
    results = vectorstore.similarity_search(question, k=2)
    print(f"[Main] 類似検索が完了しました: {len(results)} 件")
    for result in results:
        print(f"[{result.metadata.get('source', 'unknown')}]\n{result.page_content}\n")

    # 検索結果をコンテキストとして Azure OpenAI に渡し、最終回答を生成する。
    print("回答:")
    print(answer_question(question, results))
        
    


if __name__ == "__main__":
    main()