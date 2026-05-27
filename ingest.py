"""
HR 문서를 읽어 ChromaDB 벡터DB에 저장하는 스크립트.
지원 형식: PDF, Excel(.xlsx/.xls), 텍스트(.txt/.md/.csv)

실행: python ingest.py
"""
from pathlib import Path

import chromadb
from sentence_transformers import SentenceTransformer
from PyPDF2 import PdfReader
import openpyxl

BASE = Path(__file__).parent
DOCS_DIR = BASE / "docs"
CHROMA_DIR = BASE / "chroma_db"
COLLECTION = "hr_docs"
EMBED_MODEL = "paraphrase-multilingual-MiniLM-L12-v2"
CHUNK_SIZE = 600
OVERLAP = 100


def read_pdf(path: Path) -> str:
    reader = PdfReader(str(path))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def read_excel(path: Path) -> str:
    wb = openpyxl.load_workbook(str(path), data_only=True)
    lines = []
    for ws in wb.worksheets:
        lines.append(f"[시트: {ws.title}]")
        for row in ws.iter_rows(values_only=True):
            line = "\t".join(str(c) if c is not None else "" for c in row).strip()
            if line.replace("\t", ""):
                lines.append(line)
    return "\n".join(lines)


def read_text(path: Path) -> str:
    for enc in ("utf-8", "cp949", "euc-kr"):
        try:
            return path.read_text(encoding=enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return path.read_text(errors="replace")


def make_chunks(text: str) -> list:
    text = text.strip()
    chunks, i = [], 0
    while i < len(text):
        chunk = text[i : i + CHUNK_SIZE].strip()
        if chunk:
            chunks.append(chunk)
        i += CHUNK_SIZE - OVERLAP
    return chunks


def main():
    if not DOCS_DIR.exists():
        DOCS_DIR.mkdir()
        print("docs/ 폴더를 생성했습니다.")
        print("HR 문서(PDF / Excel / TXT)를 docs/ 폴더에 넣고 다시 실행하세요.")
        return

    files = [f for f in sorted(DOCS_DIR.rglob("*")) if f.is_file()]
    if not files:
        print("docs/ 폴더가 비어 있습니다. 문서를 추가하고 다시 실행하세요.")
        return

    print(f"임베딩 모델 로딩 중: {EMBED_MODEL}")
    model = SentenceTransformer(EMBED_MODEL)

    client = chromadb.PersistentClient(path=str(CHROMA_DIR))
    try:
        client.delete_collection(COLLECTION)
        print("기존 벡터DB 초기화 완료")
    except Exception:
        pass
    collection = client.create_collection(
        COLLECTION, metadata={"hnsw:space": "cosine"}
    )

    all_texts, all_ids, all_metas = [], [], []
    chunk_id = 0

    for f in files:
        ext = f.suffix.lower()
        print(f"  처리 중: {f.name}", end=" ... ")
        try:
            if ext == ".pdf":
                text = read_pdf(f)
            elif ext in (".xlsx", ".xls"):
                text = read_excel(f)
            elif ext in (".txt", ".md", ".csv"):
                text = read_text(f)
            else:
                print("지원하지 않는 형식, 건너뜀")
                continue
        except Exception as e:
            print(f"오류: {e}")
            continue

        chunks = make_chunks(text)
        for idx, chunk in enumerate(chunks):
            all_texts.append(chunk)
            all_ids.append(f"c{chunk_id}")
            all_metas.append({"source": f.name, "chunk": idx})
            chunk_id += 1
        print(f"{len(chunks)}개 청크")

    if not all_texts:
        print("\n처리된 텍스트가 없습니다.")
        return

    print(f"\n총 {len(all_texts)}개 청크 임베딩 중...")
    embeddings = model.encode(
        all_texts, batch_size=32, show_progress_bar=True
    ).tolist()

    BATCH = 500
    for i in range(0, len(all_texts), BATCH):
        collection.add(
            documents=all_texts[i : i + BATCH],
            embeddings=embeddings[i : i + BATCH],
            ids=all_ids[i : i + BATCH],
            metadatas=all_metas[i : i + BATCH],
        )

    print(f"\n완료! {len(all_texts)}개 청크가 벡터DB에 저장되었습니다.")
    print(f"저장 위치: {CHROMA_DIR}")


if __name__ == "__main__":
    main()
