# -*- coding: utf-8 -*-
"""智填 FormFiller - 命令行入口

复用 backend 的同一套别名智能匹配逻辑，无需启动 Web 服务即可一次性填表。

用法：
  # 1) 查看模板有哪些占位符
  python cli.py scan --template 员工入职登记表模板.docx

  # 2) 用资料文件填充模板
  python cli.py fill --template 员工入职登记表模板.docx \\
      --source 李云飞个人简历.docx 岗位信息.txt \\
      --out 李云飞_入职登记表.docx

  # 3) 起网页版（资料库可长期累积，推荐日常使用）
  python cli.py serve --port 8765

说明：
  --source 支持 docx/xlsx/pdf/txt/md/csv/json 及常见图片。
  图片文件按文件名匹配模板中的 {{img:名称}} 占位符。
  缺的字段用 --set 补，如 --set 签名=李云飞。
"""
import argparse
import json
import os
import re
import sys
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
BACKEND = os.path.join(BASE, "backend")
if BACKEND not in sys.path:
    sys.path.insert(0, BACKEND)

# Windows 控制台默认 GBK，直接打印中文会炸，统一切 UTF-8
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:
        pass

import extractor      # noqa: E402
import filler         # noqa: E402
import form_analyzer  # noqa: E402


def _build_library(sources):
    """从资料文件构建 entries / images，逻辑与 backend/main.py 的入库流程一致"""
    entries, images = {}, {}
    report = []
    for p in sources:
        if not os.path.exists(p):
            raise SystemExit("错误：资料文件不存在 -> %s" % p)
        name = os.path.basename(p)
        cat = extractor.file_category(name)
        if cat == "image":
            images[extractor.image_name_key(name)] = os.path.abspath(p)
            report.append({"file": name, "category": "image", "count": 0})
        elif cat == "document":
            kv, _preview = extractor.extract_info(p, name)
            entries.update(kv)
            report.append({"file": name, "category": "document", "count": len(kv)})
        else:
            report.append({"file": name, "category": "unsupported", "count": 0})
    return entries, images, report


def _resolve(template, filename, entries, images, overrides):
    """把模板占位符解析成 (text_values, 字段明细, 缺失列表)"""
    text_fields, img_fields = filler.parse_template(template, filename)
    text_values, detail = {}, []

    for name in sorted(text_fields):
        hit = extractor.lookup_field(name, entries)
        dyn = filler.normalize(name) in {"当前日期", "当前时间"} or name in filler.DYNAMIC_FIELDS
        value = hit[1]["value"] if hit else None
        source = hit[1].get("source") if hit else None
        if dyn and not value:
            value = filler.dynamic_value(name)
            source = "动态生成"
        matched_from = hit[0] if hit else None
        detail.append({
            "field": name, "value": value, "source": source,
            "matched_from": matched_from, "dynamic": dyn,
            "missing": hit is None and not dyn,
        })
        if value:
            text_values[filler.normalize(name)] = value

    # 用户覆盖
    for k, v in (overrides or {}).items():
        if str(v).strip():
            text_values[filler.normalize(k) or k] = str(v).strip()

    # 回填明细，避免 --set 补上的字段在报告里仍显示为缺失
    for d in detail:
        eff = text_values.get(filler.normalize(d["field"]))
        if eff:
            if d["missing"]:
                d["source"] = "命令行补充"
                d["missing"] = False
            d["value"] = eff

    missing = [d["field"] for d in detail if d["missing"]
               and not (filler.normalize(d["field"]) in text_values)]

    img_detail = []
    for name in sorted(img_fields):
        nk = extractor.normalize_key(name)
        matched = images.get(nk)
        img_detail.append({"field": name, "matched": os.path.basename(matched) if matched else None,
                           "missing": matched is None})
    return text_values, detail, missing, img_detail, sorted(text_fields), sorted(img_fields)


# ---------------------------------------------------------------- 表格型表单模式
#
# 很多单位/高校的表格不是 {{占位符}} 模板，而是「标签格 + 空格」的传统表单。
# 这类模板用 filler.parse_template 只能识别到 0 个字段，必须走 form_analyzer。


# 复合字段：模板一个空 = 资料好几项拼起来。
# 例：模板「取得专业技术职务情况」要写"讲师（2025年7月15日取得）"，
#     而资料里职称和取得时间是分开的两项。
_COMPOSITES = {
    "最后学历毕业院校及学位": ("{毕业院校} {所学专业} {学位}", ["毕业院校", "所学专业", "学位"]),
    "取得专业技术职务情况": ("{职称}（{职称获得时间}取得）", ["职称", "职称获得时间"]),
    "最后学历取得学位证时间": ("{毕业时间}", ["毕业时间"]),
}


def _derived_values(entries):
    """从已有信息推算/拼合的字段（年龄、复合字段）"""
    out = {}

    def get(*names):
        for n in names:
            hit = extractor.lookup_field(n, entries)
            if hit:
                return hit[1]["value"]
        return None

    birth = get("出生年月", "出生日期")
    if birth:
        m = re.search(r"(19|20)(\d{2})\s*[年./\-]?\s*(\d{1,2})?", str(birth))
        if m:
            y = int(m.group(1) + m.group(2))
            mo = int(m.group(3)) if m.group(3) else 1
            now = datetime.now()
            # 按月份精确计算：生日还没到就不能算满一岁
            age = now.year - y - ((now.month, now.day) < (mo, 1))
            out["年龄"] = str(age)

    # 复合字段：模板一个空要填资料里好几项拼起来
    for target, (tmpl, keys) in _COMPOSITES.items():
        vals = {k: get(k) for k in keys}
        if all(vals.values()):
            out[target] = tmpl.format(**vals)
        elif any(vals.values()):
            out[target] = " ".join(v for v in vals.values() if v)
    return out


def cmd_scan(args):
    if not os.path.exists(args.template):
        raise SystemExit("错误：模板不存在 -> %s" % args.template)
    name = os.path.basename(args.template)
    text_fields, img_fields = filler.parse_template(args.template, name)
    if args.json:
        print(json.dumps({"template": name, "text_fields": sorted(text_fields),
                          "img_fields": sorted(img_fields)}, ensure_ascii=False, indent=2))
        return 0
    print("模板：%s" % name)
    print("文本占位符（%d）：%s" % (len(text_fields), "、".join(sorted(text_fields)) or "无"))
    print("图片占位符（%d）：%s" % (len(img_fields), "、".join(sorted(img_fields)) or "无"))

    if not text_fields:
        # 没有占位符 → 可能是「标签格 + 空格」的传统表单，再按表格型分析一次
        doc = extractor.open_docx_any(args.template)
        fields = form_analyzer.analyze_template_fields(doc)
        lists = form_analyzer.analyze_list_tables(doc)
        print()
        print("未发现 {{占位符}}，已改用表格型表单分析：")
        print("  待填字段（%d）：%s" % (
            len(fields), "、".join(f["name"] for f in fields) or "无"))
        if lists:
            print("  多行列表（%d 处）：" % len(lists))
            for lt in lists:
                print("    表%d  列=%s" % (lt["table"], "/".join(
                    h for h in lt["header"] if h)))
    return 0


def cmd_fill_form(args, tpl_name):
    """表格型表单填充：模板驱动，从资料里找同名标签取值后按坐标写回"""
    doc = extractor.open_docx_any(args.template)
    fields = form_analyzer.analyze_template_fields(doc)

    overrides = {}
    for kv in args.set or []:
        if "=" in kv:
            k, v = kv.split("=", 1)
            overrides[k.strip()] = v.strip()

    # 资料索引：把所有资料的标签汇成一个 dict，再用别名/后缀智能匹配
    entries, merged = {}, {}
    for p in args.source or []:
        if not os.path.exists(p):
            raise SystemExit("错误：资料文件不存在 -> %s" % p)
        name = os.path.basename(p)
        ext = os.path.splitext(p)[1].lower()
        # 路径 A：表格型资料（标签格 + 值格，如应聘信息表），读 doc.tables
        d = extractor.open_docx_any(p) if ext in (".docx", ".docm") else None
        idx = form_analyzer.build_source_index(d) if d is not None else {}
        for k, v in idx.items():
            nk = extractor.normalize_key(k)
            if nk:
                merged.setdefault(nk, {"key": k, "value": v, "source": name})
        # 路径 B：文本型资料（「标签：值」/ 文本框简历 / txt / pdf / xlsx）
        try:
            kv, _ = extractor.extract_info(p, name)
        except Exception:
            kv = {}
        for k, v in kv.items():
            merged.setdefault(k, {"key": v.get("key", k), "value": v["value"],
                                  "source": name})
    entries = merged
    entries.update(overrides)

    values = {}
    derived = _derived_values(entries)
    for f in fields:
        hit = extractor.lookup_field(f["name"], entries)
        if hit:
            values[f["name"]] = hit[1]["value"]
    # 复合/推算字段优先——它们比单项匹配更贴切
    values.update({k: v for k, v in derived.items()})
    # 未解决清单要在补上复合字段之后再算，否则会出现"已填却又说没找到"
    unresolved = [f["name"] for f in fields if f["name"] not in values]

    written = form_analyzer.fill_fields(doc, fields, values)

    # 「本人」定位线索：论文作者排序需要知道本人是谁。优先取 --set 里显式
    # 指定的本人英文姓/拼音，其次用资料里的姓名兜底。
    self_hints = []
    for k in ("本人", "本人姓氏", "姓名拼音", "拼音姓", "姓氏拼音"):
        if k in overrides:
            self_hints.append(overrides[k])
    if "姓名" in entries:
        self_hints.append(entries["姓名"]["value"])

    # 多行列表（学习经历、工作经历、论文、亲属……）
    list_written = []
    if not args.no_lists:
        src_sections = {}
        for p in args.source or []:
            if os.path.splitext(p)[1].lower() not in (".docx", ".docm"):
                continue
            d = extractor.open_docx_any(p)
            for name, sec in form_analyzer.extract_list_sections(d).items():
                src_sections.setdefault(name, sec)
            # 文本型论文引用（简历/信息表里的文本框段落）——表格读不到，
            # 这里单独从全文文本解析，ISSN/影响因子/收录/年卷期页才有来源。
            text = extractor.extract_docx_text(p)
            for name, sec in form_analyzer.extract_papers_from_text(
                    text, self_hints).items():
                src_sections[name] = sec
        for lt in form_analyzer.analyze_list_tables(doc):
            header = lt["header"]
            best_name, best_map, best_score = None, None, 0
            for name, sec in src_sections.items():
                cmap = form_analyzer.match_columns_multi(header, sec["header"])
                score = sum(1 for c in cmap if c is not None)
                if score > best_score:
                    best_name, best_map, best_score = name, cmap, score
            if best_map and best_score:
                n = form_analyzer.fill_list_table(
                    doc, lt["table"], header, list(lt["data_rows"]),
                    src_sections[best_name]["rows"], best_map)
                if n:
                    list_written.append((lt["table"], best_name, n))

    out_path = os.path.abspath(args.out or ("已填_" + tpl_name))
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    doc.save(out_path)

    if args.json:
        print(json.dumps({"status": "ok", "output": out_path,
                          "written": written, "unresolved": unresolved,
                          "lists": list_written}, ensure_ascii=False, indent=2))
        return 0

    print("=== 已填字段（%d）===" % len(written))
    for w in written:
        print("  ✓ %s" % w)
    if list_written:
        print("\n=== 已填多行列表（%d 处）===" % len(list_written))
        for ti, name, n in list_written:
            print("  ✓ 表%d ← 资料「%s」（%d 格）" % (ti, name, n))
    if unresolved:
        print("\n=== 资料里没找到、需你补（%d）===" % len(unresolved))
        for u in unresolved:
            print("  - %s" % u)
    print("\n输出：%s" % out_path)
    print("大小：%d 字节" % os.path.getsize(out_path))
    return 0


def cmd_fill(args):
    if not os.path.exists(args.template):
        raise SystemExit("错误：模板不存在 -> %s" % args.template)
    tpl_name = os.path.basename(args.template)

    # 自动切换：模板里一个 {{占位符}} 都没有，但有「标签格 + 空格」→ 走表格型表单模式
    try:
        _tf, _if = filler.parse_template(args.template, tpl_name)
    except Exception:
        _tf, _if = set(), set()
    if not _tf and not args.no_form_mode:
        _probe = extractor.open_docx_any(args.template)
        if form_analyzer.analyze_template_fields(_probe):
            print("模板没有 {{占位符}}，已自动切换到「表格型表单」模式。\n")
            return cmd_fill_form(args, tpl_name)

    overrides = {}
    for kv in args.set or []:
        if "=" in kv:
            k, v = kv.split("=", 1)
            overrides[k.strip()] = v.strip()

    entries, images, src_report = _build_library(args.source or [])
    text_values, detail, missing, img_detail, text_fields, img_fields = _resolve(
        args.template, tpl_name, entries, images, overrides)

    if missing and not args.force:
        print("资料中缺少以下字段，无法生成（补上后重试，或加 --force 留空）：")
        for m in missing:
            print("  - %s" % m)
        print("\n补值方式：--set 字段名=值")
        if args.json:
            print(json.dumps({"status": "missing", "missing": missing}, ensure_ascii=False))
        return 1

    ext = os.path.splitext(tpl_name)[1].lower()
    out_name = args.name or filler.render_filename("", tpl_name, text_values, ext)
    out_path = os.path.abspath(args.out or out_name)
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)

    result = filler.fill_template(args.template, tpl_name, text_values, images, out_path)

    if args.json:
        print(json.dumps({"status": "ok", "output": out_path, "fields": detail,
                          "images": img_detail}, ensure_ascii=False, indent=2))
        return 0

    print("=== 资料来源 ===")
    for r in src_report:
        tag = {"document": "文档", "image": "图片", "unsupported": "不支持"}[r["category"]]
        print("  %s %s（提取 %d 项）" % (tag, r["file"], r["count"]))
    print("\n=== 字段匹配 ===")
    for d in detail:
        if d["missing"]:
            print("  ✗ %-14s 缺失（已留空）" % d["field"])
        elif d["matched_from"] and d["matched_from"] != d["field"]:
            print("  ✓ %-14s ← %s：%s" % (d["field"], d["matched_from"], str(d["value"])[:40]))
        else:
            print("  ✓ %-14s %s：%s" % (d["field"], d["source"] or "", str(d["value"])[:40]))
    for d in img_detail:
        print("  %s 图片 %-10s %s" % ("✓" if not d["missing"] else "✗", d["field"], d["matched"] or "缺失"))
    print("\n生成完成：%s" % out_path)
    print("大小：%d 字节%s" % (os.path.getsize(out_path),
                             "，插入图片 %d 张" % result.get("images_inserted", 0)
                             if result.get("images_inserted") else ""))
    if missing:
        print("\n注意：%d 个字段缺失并已留空，建议人工补：%s" % (len(missing), "、".join(missing)))
    return 0


def cmd_serve(args):
    import uvicorn
    print("启动智填 FormFiller：http://127.0.0.1:%d" % args.port)
    uvicorn.run("main:app", host="127.0.0.1", port=args.port, reload=False, app_dir=BACKEND)
    return 0


def main():
    p = argparse.ArgumentParser(description="智填 FormFiller 命令行入口")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("scan", help="查看模板占位符")
    s.add_argument("--template", required=True)
    s.add_argument("--json", action="store_true")
    s.set_defaults(func=cmd_scan)

    f = sub.add_parser("fill", help="用资料文件填充模板")
    f.add_argument("--template", required=True, help="模板 .docx 或 .xlsx")
    f.add_argument("--source", nargs="*", default=[], help="资料文件，可多份")
    # 用 append 而非 nargs="*"：后者在重复传参时会互相覆盖，只保留最后一个
    f.add_argument("--set", action="append", default=[], metavar="K=V",
                   help="补充字段，可重复：--set 签名=李云飞 --set 签署日期=2026-08-30")
    f.add_argument("--out", default=None, help="输出路径，默认按模板名+日期命名")
    f.add_argument("--name", default=None, help="输出文件名，如 '{{姓名}}_入职登记表_{{当前日期}}.docx'")
    f.add_argument("--force", action="store_true", help="有缺失字段也生成（留空）")
    f.add_argument("--no-form-mode", action="store_true",
                   help="禁用表格型表单自动切换")
    f.add_argument("--no-lists", action="store_true",
                   help="不填多行列表（学习经历/工作经历/论文等）")
    f.add_argument("--json", action="store_true")
    f.set_defaults(func=cmd_fill)

    v = sub.add_parser("serve", help="启动 Web 界面")
    v.add_argument("--port", type=int, default=8765)
    v.set_defaults(func=cmd_serve)

    args = p.parse_args()
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
