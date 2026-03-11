import json
from datetime import datetime
from io import BytesIO

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import JSONResponse, StreamingResponse

from services.report_service import (
    build_rag_chunks_from_texts,
    load_section_config,
    process_document,
    validate_payload,
)

app = FastAPI(title="GuoTou Report API", version="1.0.0")


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/validate")
async def validate(
    policy_json: UploadFile = File(...),
    template_docx: UploadFile = File(...),
    section_config_json: UploadFile | None = File(default=None),
):
    json_data = json.loads((await policy_json.read()).decode("utf-8"))
    template_bytes = await template_docx.read()

    if section_config_json:
        section_config = json.loads((await section_config_json.read()).decode("utf-8"))
    else:
        section_config = load_section_config()

    result = validate_payload(json_data, template_bytes, section_config)
    return JSONResponse(result)


@app.post("/generate")
async def generate(
    policy_json: UploadFile = File(...),
    template_docx: UploadFile = File(...),
    api_key: str = Form(default=""),
    enable_rag: bool = Form(default=False),
    rag_top_k: int = Form(default=3),
    include_rag_snippets: bool = Form(default=False),
    section_config_json: UploadFile | None = File(default=None),
    rag_files: list[UploadFile] | None = File(default=None),
):
    json_data = json.loads((await policy_json.read()).decode("utf-8"))
    template_bytes = await template_docx.read()

    if section_config_json:
        section_config = json.loads((await section_config_json.read()).decode("utf-8"))
    else:
        section_config = load_section_config()

    rag_chunks = []
    if enable_rag and rag_files:
        rag_texts = []
        for f in rag_files:
            raw = await f.read()
            try:
                txt = raw.decode("utf-8")
            except UnicodeDecodeError:
                txt = raw.decode("utf-8", errors="ignore")
            rag_texts.append({"name": f.filename or "unknown", "text": txt})
        rag_chunks = build_rag_chunks_from_texts(rag_texts)

    output, meta = process_document(
        template_file=BytesIO(template_bytes),
        json_data=json_data,
        api_key=api_key,
        section_config=section_config,
        rag_chunks=rag_chunks,
        rag_top_k=rag_top_k,
        include_rag_snippets=include_rag_snippets,
    )

    headers = {
        "X-Not-Found-Sections": json.dumps(meta.get("not_found_sections", []), ensure_ascii=False),
        "X-Generated-At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )
