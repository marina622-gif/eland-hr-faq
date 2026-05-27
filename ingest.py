"""
HR 문서를 ChromaDB에 인덱싱하는 스크립트.
Render 시작 시 gunicorn 실행 전에 먼저 실행됨.
app.py와 동일한 DefaultEmbeddingFunction 사용.
"""
from pathlib import Path

import chromadb
import openpyxl
from chromadb.utils.embedding_functions import DefaultEmbeddingFunction

BASE       = Path(__file__).parent
FAQ_FILE   = BASE / "docs" / "eland_hr_faq.xlsx"
CHROMA_DIR = BASE / "chroma_db"
SHEET      = "FAQ"
B_COL, Q_COL, A_COL = 1, 2, 3
DATA_ROW   = 2


def load_chunks():
    chunks = []

    if FAQ_FILE.exists():
        try:
            wb = openpyxl.load_workbook(str(FAQ_FILE), data_only=True)
            ws = wb[SHEET]
            for i, row in enumerate(list(ws.iter_rows(values_only=True))[DATA_ROW:]):
                if len(row) <= max(Q_COL, A_COL):
                    continue
                cat = str(row[B_COL]).strip() if row[B_COL] else "기타"
                q   = str(row[Q_COL]).strip() if row[Q_COL] else ""
                a   = str(row[A_COL]).strip() if row[A_COL] else ""
                if q and a and cat not in ("카테고리", "None"):
                    chunks.append({
                        "id":   f"faq_{i}",
                        "text": f"질문: {q}\n답변: {a}",
                        "meta": {"source": FAQ_FILE.name, "cat": cat, "type": "faq"},
                    })
        except Exception as e:
            print(f"[ingest] Excel 오류: {e}")

    for fp in sorted((BASE / "docs").rglob("*.txt")):
        try:
            text = fp.read_text(encoding="utf-8", errors="ignore")
            for j, para in enumerate(
                p.strip() for p in text.split("\n\n") if len(p.strip()) > 50
            ):
                chunks.append({
                    "id":   f"txt_{fp.stem}_{j}",
                    "text": para,
                    "meta": {"source": fp.name, "type": "text"},
                })
        except Exception as e:
            print(f"[ingest] {fp.name} 오류: {e}")

    for fp in sorted((BASE / "docs").rglob("*.pdf")):
        try:
            import pdfplumber
            with pdfplumber.open(str(fp)) as pdf:
                for pg, page in enumerate(pdf.pages):
                    t = (page.extract_text() or "").strip()
                    if t:
                        chunks.append({
                            "id":   f"pdf_{fp.stem}_{pg}",
                            "text": t,
                            "meta": {"source": fp.name, "type": "pdf", "page": pg + 1},
                        })
        except ImportError:
            pass
        except Exception as e:
            print(f"[ingest] {fp.name} PDF 오류: {e}")

    for fp in sorted((BASE / "docs").rglob("*.docx")):
        try:
            from docx import Document
            doc = Document(str(fp))
            for j, p in enumerate(doc.paragraphs):
                if p.text.strip() and len(p.text.strip()) > 30:
                    chunks.append({
                        "id":   f"docx_{fp.stem}_{j}",
                        "text": p.text.strip(),
                        "meta": {"source": fp.name, "type": "docx"},
                    })
        except ImportError:
            pass
        except Exception as e:
            print(f"[ingest] {fp.name} DOCX 오류: {e}")

    return chunks


def main():
    print("[ingest] 벡터 인덱스 구축 시작...")
    CHROMA_DIR.mkdir(exist_ok=True)

    ef     = DefaultEmbeddingFunction()
    client = chromadb.PersistentClient(path=str(CHROMA_DIR))
    col    = client.get_or_create_collection(
        name="hr_docs",
        embedding_function=ef,
        metadata={"hnsw:space": "cosine"},
    )

    chunks = load_chunks()
    if not chunks:
        print("[ingest] 인덱싱할 문서 없음")
        return

    existing = col.get()
    if existing["ids"]:
        col.delete(ids=existing["ids"])

    BATCH = 64
    for i in range(0, len(chunks), BATCH):
        b = chunks[i : i + BATCH]
        col.add(
            ids       = [c["id"]   for c in b],
            documents = [c["text"] for c in b],
            metadatas = [c["meta"] for c in b],
        )

    print(f"[ingest] 완료: {len(chunks)}개 청크 인덱싱")


if __name__ == "__main__":
    main()
