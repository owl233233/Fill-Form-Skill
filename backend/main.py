"""填表工具 FormFiller —— FastAPI 后端"""
import os

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

import extractor
import filler
import storage

app = FastAPI(title="智填 FormFiller", docs_url=None, redoc_url=None)

MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB


@app.exception_handler(Exception)
async def _unhandled_exception_handler(request, exc):
    """未捕获异常也返回可读信息。

    默认情况下前端只会看到一句 "Internal Server Error"，排查无从下手
    （曾出现数据目录被删导致上传 500，但界面上看不出任何线索）。
    这里把异常类型与信息回传给前端，同时把完整堆栈打到服务控制台。
    """
    import traceback
    traceback.print_exc()
    return JSONResponse(
        status_code=500,
        content={"detail": f"{type(exc).__name__}: {exc}"},
    )


# ---------------- 资料库：文件上传 ----------------

@app.post("/api/library/upload")
async def library_upload(file: UploadFile = File(...)):
    content = await file.read()
    if len(content) > MAX_FILE_SIZE:
        raise HTTPException(413, "文件过大（超过 50MB）")
    if not content:
        raise HTTPException(400, "空文件")
    category = extractor.file_category(file.filename)
    if category == "other":
        raise HTTPException(400, f"暂不支持该文件类型：{os.path.splitext(file.filename)[1]}（支持 docx/xlsx/pdf/txt/md/csv/json 及常见图片）")

    stored, path = storage.save_upload(storage.LIBRARY_DIR, file.filename, content)
    fid = storage.new_id()

    def _op(db):
        kv, preview = extractor.extract_info(path, stored)
        if category == "image":
            db["images"].append({
                "id": fid, "filename": stored, "path": path,
                "uploaded_at": storage.now_str(),
            })
            return {"id": fid, "filename": stored, "category": "image",
                    "extracted_count": 0, "preview": ""}
        # 文档类：提取信息并入库（来源标记为文件名）
        added = []
        for nk, item in kv.items():
            if nk not in db["entries"]:
                db["entries"][nk] = {"key": item["key"], "value": item["value"],
                                     "source": stored, "updated_at": storage.now_str()}
                added.append(item["key"])
        db["library_files"].append({
            "id": fid, "filename": stored, "path": path,
            "type": os.path.splitext(stored)[1].lower().lstrip("."),
            "uploaded_at": storage.now_str(),
            "extracted": {item["key"]: item["value"] for item in kv.values()},
            "preview": preview,
        })
        return {"id": fid, "filename": stored, "category": "document",
                "extracted": {item["key"]: item["value"] for item in kv.values()},
                "extracted_count": len(added), "added_keys": added, "preview": preview}

    result = storage.db_write(_op)
    return result


@app.get("/api/library")
def library_list():
    db = storage.db_read()
    return {
        "entries": [{"name": nk, "key": e["key"], "value": e["value"], "source": e["source"]}
                    for nk, e in sorted(db["entries"].items(), key=lambda x: x[1]["key"])],
        "files": [{k: f[k] for k in ("id", "filename", "type", "uploaded_at", "preview")}
                  for f in db["library_files"]],
        "images": [{"id": i["id"], "filename": i["filename"], "name": os.path.splitext(i["filename"])[0],
                    "uploaded_at": i["uploaded_at"]} for i in db["images"]],
    }


@app.delete("/api/library/file/{fid}")
def library_file_delete(fid: str):
    def _op(db):
        for coll in ("library_files", "images"):
            for i, item in enumerate(db[coll]):
                if item["id"] == fid:
                    storage.remove_quiet(item["path"])
                    db[coll].pop(i)
                    # 若无其他文件提供该来源信息，则移除对应 entries
                    if coll == "library_files":
                        src = item.get("filename")
                        others = [f for f in db["library_files"] if f.get("extracted")]
                        for nk in [k for k, e in db["entries"].items() if e.get("source") == src]:
                            if not any(src in str(f.get("filename")) for f in others):
                                db["entries"].pop(nk, None)
                    return {"ok": True}
        return {"ok": False}

    return storage.db_write(_op)


# ---------------- 资料库：键值信息管理 ----------------

@app.post("/api/library/entry")
async def library_entry_add(key: str = Form(...), value: str = Form(...)):
    key, value = key.strip(), value.strip()
    if not key or not value:
        raise HTTPException(400, "字段名和内容都不能为空")
    nk = extractor.normalize_key(key)
    if not nk:
        raise HTTPException(400, "无效的字段名")

    def _op(db):
        existed = nk in db["entries"]
        db["entries"][nk] = {"key": key, "value": value, "source": "手动添加",
                             "updated_at": storage.now_str()}
        return {"ok": True, "name": nk, "existed": existed}

    return storage.db_write(_op)


@app.delete("/api/library/entry/{name}")
def library_entry_delete(name: str):
    def _op(db):
        removed = db["entries"].pop(name, None)
        return {"ok": removed is not None}

    return storage.db_write(_op)


# ---------------- 模板管理 ----------------

@app.post("/api/templates/upload")
async def template_upload(file: UploadFile = File(...), naming_pattern: str = Form("")):
    content = await file.read()
    if len(content) > MAX_FILE_SIZE:
        raise HTTPException(413, "文件过大（超过 50MB）")
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in filler.SUPPORTED_TEMPLATE_EXTS:
        raise HTTPException(400, f"暂不支持 {ext} 模板（目前支持 .docx / .xlsx）")

    stored, path = storage.save_upload(storage.TEMPLATES_DIR, file.filename, content)
    try:
        text_fields, img_fields = filler.parse_template(path, stored)
    except Exception as e:
        storage.remove_quiet(path)
        raise HTTPException(400, f"模板解析失败：{e}")

    tid = storage.new_id()

    def _op(db):
        db["templates"].append({
            "id": tid, "filename": stored, "path": path, "type": ext.lstrip("."),
            "naming_pattern": (naming_pattern or "").strip(),
            "text_fields": sorted(text_fields),
            "img_fields": sorted(img_fields),
            "uploaded_at": storage.now_str(),
        })
        return _template_brief(db["templates"][-1])

    return storage.db_write(_op)


def _template_brief(t):
    return {k: t[k] for k in ("id", "filename", "type", "naming_pattern",
                              "text_fields", "img_fields", "uploaded_at")}


@app.get("/api/templates")
def templates_list():
    db = storage.db_read()
    return [_template_brief(t) for t in db["templates"]]


@app.delete("/api/templates/{tid}")
def template_delete(tid: str):
    def _op(db):
        for i, t in enumerate(db["templates"]):
            if t["id"] == tid:
                storage.remove_quiet(t["path"])
                db["templates"].pop(i)
                return {"ok": True}
        return {"ok": False}

    return storage.db_write(_op)


# ---------------- 填充：预览（缺失检测）与生成 ----------------

@app.post("/api/fill/preview")
async def fill_preview(template_id: str = Form(...)):
    db = storage.db_read()
    tpl = next((t for t in db["templates"] if t["id"] == template_id), None)
    if not tpl:
        raise HTTPException(404, "模板不存在")

    entries = db["entries"]
    fields = []
    for name in tpl["text_fields"]:
        hit = extractor.lookup_field(name, entries)
        dyn = name in filler.DYNAMIC_FIELDS or filler.normalize(name) in {"当前日期", "当前时间"}
        val = hit[1]["value"] if hit else None
        if dyn and not val:
            val = filler.dynamic_value(name)
        fields.append({
            "name": name,
            "value": val,
            "source": hit[1]["source"] if hit else None,
            "dynamic": dyn,
            "missing": hit is None and not dyn,
        })

    images = []
    img_map_by_name = {extractor.image_name_key(i["filename"]): i for i in db["images"]}
    for name in tpl["img_fields"]:
        nk = extractor.normalize_key(name)
        img = img_map_by_name.get(nk)
        images.append({"name": name, "matched": img["filename"] if img else None})

    available_images = [{"name": os.path.splitext(i["filename"])[0], "filename": i["filename"]}
                        for i in db["images"]]

    suggested = filler.render_filename(tpl["naming_pattern"], tpl["filename"],
                                       {name: (extractor.lookup_field(name, db["entries"]) or (None, {"value": ""}))[1]["value"]
                                        for name in tpl["text_fields"]},
                                       "." + tpl["type"])

    return {
        "template": _template_brief(tpl),
        "fields": fields,
        "images": images,
        "available_images": available_images,
        "suggested_filename": suggested,
    }


@app.post("/api/fill/generate")
async def fill_generate(template_id: str = Form(...),
                        filename: str = Form(""),
                        values: str = Form("{}"),
                        save_to_library: bool = Form(False)):
    import json
    db = storage.db_read()
    tpl = next((t for t in db["templates"] if t["id"] == template_id), None)
    if not tpl:
        raise HTTPException(404, "模板不存在")
    try:
        vals = json.loads(values or "{}")
        if not isinstance(vals, dict):
            raise ValueError
    except ValueError:
        raise HTTPException(400, "values 必须是 JSON 对象")

    # 缺失校验：文本字段既无资料库值也无用户输入、且非动态字段 → 报缺失，让前端询问用户
    missing = []
    for name in tpl["text_fields"]:
        hit = extractor.lookup_field(name, db["entries"])
        user_val = vals.get(name)
        dyn = filler.normalize(name) in {"当前日期", "当前时间"} or name in filler.DYNAMIC_FIELDS
        if user_val is None or str(user_val).strip() == "":
            if hit is None and not dyn:
                missing.append(name)

    if missing:
        return JSONResponse(status_code=422, content={
            "detail": "存在缺失信息，需要补充后才能生成", "missing": missing})

    # 组装填充值：资料库值（智能匹配，统一以模板字段名的归一化键存放）+ 用户覆盖值
    text_values = {}
    for name in tpl["text_fields"]:
        hit = extractor.lookup_field(name, db["entries"])
        if hit:
            text_values[filler.normalize(name)] = hit[1]["value"]
    for k, v in vals.items():
        if str(v).strip():
            text_values[filler.normalize(k) or k] = str(v).strip()

    # 图片映射：按资料库图片名自动匹配
    image_map = {}
    for i in db["images"]:
        image_map[extractor.image_name_key(i["filename"])] = i["path"]

    # 输出文件名
    ext = "." + tpl["type"]
    out_name = filler.render_filename(filename, tpl["filename"], text_values, ext)
    out_path = os.path.join(storage.OUTPUT_DIR, out_name)
    base, e = os.path.splitext(out_name)
    n = 1
    while os.path.exists(out_path):
        out_path = os.path.join(storage.OUTPUT_DIR, f"{base}({n}){e}")
        n += 1
    out_name = os.path.basename(out_path)

    try:
        result = filler.fill_template(tpl["path"], tpl["filename"], text_values, image_map, out_path)
    except Exception as e:
        raise HTTPException(500, f"生成失败：{e}")

    # 可选：将用户补充的新信息保存到资料库（仅保存资料库中确实不存在对应信息的项目）
    saved_keys = []
    if save_to_library:
        def _save(db2):
            for k, v in vals.items():
                if not str(v).strip():
                    continue
                # 如果该字段名（含别名/后缀）在资料库中已能找到匹配，则不再重复保存
                if extractor.lookup_field(k, db2["entries"]) is not None:
                    continue
                nk = extractor.normalize_key(k)
                if nk and nk not in db2["entries"]:
                    db2["entries"][nk] = {"key": k, "value": str(v).strip(),
                                          "source": "生成时补充", "updated_at": storage.now_str()}
                    saved_keys.append(k)
        storage.db_write(_save)

    rid = storage.new_id()

    def _rec(db):
        db["records"].append({
            "id": rid, "template_id": tpl["id"], "template_name": tpl["filename"],
            "filename": out_name, "path": out_path, "created_at": storage.now_str(),
            "values_count": len(text_values),
        })

    storage.db_write(_rec)
    return {"id": rid, "filename": out_name, "download_url": f"/api/download/output/{out_name}",
            "images_inserted": result["images_inserted"],
            "images_missing": result["images_missing"], "saved_keys": saved_keys}


@app.get("/api/records")
def records_list():
    db = storage.db_read()
    return [{k: r[k] for k in ("id", "template_name", "filename", "created_at")}
            for r in reversed(db["records"][-100:])]


# ---------------- 下载 ----------------

@app.get("/api/download/output/{filename}")
def download_output(filename: str):
    path = os.path.join(storage.OUTPUT_DIR, os.path.basename(filename))
    if not os.path.exists(path):
        raise HTTPException(404, "文件不存在")
    return FileResponse(path, filename=filename,
                        media_type="application/octet-stream")


@app.get("/api/download/template/{tid}")
def download_template(tid: str):
    db = storage.db_read()
    tpl = next((t for t in db["templates"] if t["id"] == tid), None)
    if not tpl or not os.path.exists(tpl["path"]):
        raise HTTPException(404, "模板不存在")
    return FileResponse(tpl["path"], filename=tpl["filename"],
                        media_type="application/octet-stream")


# ---------------- 静态前端 ----------------

# 资料库文件静态访问需先于前端 catch-all 挂载
app.mount("/data/library", StaticFiles(directory=storage.LIBRARY_DIR), name="library-files")

frontend_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "frontend")
app.mount("/", StaticFiles(directory=frontend_dir, html=True), name="frontend")


@app.exception_handler(404)
async def not_found_handler(request, exc):
    return JSONResponse(status_code=404, content={"detail": "接口不存在"})
