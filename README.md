## 起動

依存関係を同期して API サーバーを起動します。

```powershell
uv sync
uv run uvicorn t_api_v1.main:app --reload
```

起動後のドキュメントは `http://127.0.0.1:8000/docs` で確認できます。

## エンドポイント

- `POST /qa/{id}`: 質問を登録
- `GET /qa/{id}`: 質問を取得
- `GET /search?question=...`: RAG 検索
- `POST /evaluate`: 検索結果を評価

Azure AI Search と Azure OpenAI の環境変数が設定されている場合、`/search` はそれらを使って回答を生成します。未設定の場合も API は起動し、設定不足を示す回答を返します。
