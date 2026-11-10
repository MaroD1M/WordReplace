import copy
import hashlib
import hmac
import io
import logging
import tempfile
import threading
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Any, Iterator
from uuid import uuid4

import pandas as pd
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from app.core.config import settings
from app.core.excel import read_excel

logger = logging.getLogger("wordreplace")

HEADER_ROW_OFFSET = 2
MAX_FILENAME_STEM = 120
RESERVED_FILENAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}

INLINE_RUN_XPATH = (
    "w:r | w:hyperlink/w:r | w:sdt/w:sdtContent/w:r | "
    "w:sdt/w:sdtContent/w:hyperlink/w:r | "
    "w:ins/w:r | w:ins/w:hyperlink/w:r | "
    "w:smartTag/w:r | w:fldSimple/w:r"
)


@dataclass
class ReplacedFile:
    item_id: str
    filename: str
    data: io.BytesIO
    row_idx: int
    replace_count: int


RUN_CACHE: dict[str, dict[str, Any]] = {}
RUN_CACHE_BYTES = 0
_CACHE_LOCK = threading.RLock()


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


def _entry_size(result: dict[str, Any]) -> int:
    total = 0
    for file in result.get("files", []):
        try:
            total += len(file.data.getvalue())
        except Exception:
            continue
    total += len(result.get("details", [])) * 256
    return total


def _prune_cache(now: datetime | None = None) -> None:
    global RUN_CACHE_BYTES

    now = now or _utcnow()
    expired = [
        rid for rid, item in RUN_CACHE.items()
        if item.get("expires_at") and item["expires_at"] <= now
    ]
    for rid in expired:
        item = RUN_CACHE.pop(rid, None)
        if item:
            RUN_CACHE_BYTES -= int(item.get("_size", 0))

    while len(RUN_CACHE) > settings.run_cache_max_entries:
        oldest = min(RUN_CACHE.items(), key=lambda item: item[1].get("created_at", now))
        item = RUN_CACHE.pop(oldest[0], None)
        if item:
            RUN_CACHE_BYTES -= int(item.get("_size", 0))

    while RUN_CACHE_BYTES > settings.run_cache_max_bytes and RUN_CACHE:
        oldest = min(RUN_CACHE.items(), key=lambda item: item[1].get("created_at", now))
        item = RUN_CACHE.pop(oldest[0], None)
        if item:
            RUN_CACHE_BYTES -= int(item.get("_size", 0))


def clean_text(text: str) -> str:
    if not isinstance(text, str):
        return ""
    return " ".join(str(text).strip().split())


def clean_filename(filename: str) -> str:
    bad = '\\/:*?"<>|\x00\x01\x02\x03\x04\x05\x06\x07\x08\x09\x0a\x0b\x0c\x0d\x0e\x0f\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1a\x1b\x1c\x1d\x1e\x1f'
    table = str.maketrans({ch: "_" for ch in bad})
    value = str(filename).translate(table).strip().strip(".")

    stem, dot, ext = value.rpartition(".")
    if not dot:
        stem, ext = value, ""
    stem = stem.strip().strip(".")
    if stem.upper() in RESERVED_FILENAMES:
        stem = f"_{stem}"
    if len(stem) > MAX_FILENAME_STEM:
        stem = stem[:MAX_FILENAME_STEM]
    value = f"{stem}.{ext}" if ext else stem
    return value or "未命名.docx"


def precompute_replace_patterns(
    rules: list[tuple[str, str]],
    excel_row: pd.Series,
    keep_placeholder_on_empty: bool = False,
) -> list[tuple[str, str, str, str]]:
    patterns: list[tuple[str, str, str, str]] = []
    for old_text, col_name in rules:
        replacement = str(excel_row[col_name]).strip() if col_name in excel_row.index else ""
        if keep_placeholder_on_empty and not replacement:
            continue
        cleaned = clean_text(old_text)
        if cleaned:
            patterns.append((old_text, col_name, cleaned, replacement))
    return patterns


def _find_occurrences(text: str, key: str) -> list[int]:
    starts: list[int] = []
    if not key:
        return starts
    index = text.find(key)
    while index != -1:
        starts.append(index)
        index = text.find(key, index + len(key))
    return starts


def _paragraph_runs(paragraph) -> list[Run]:
    """Return runs in the same order used to build the paragraph text.

    Includes inline content controls, tracked insertions, smart tags and
    simple fields so placeholders inside them are replaced too.
    """
    return [Run(element, paragraph) for element in paragraph._p.xpath(INLINE_RUN_XPATH)]


def paragraph_text(paragraph) -> str:
    return "".join(run.text for run in _paragraph_runs(paragraph))


def _has_embedded_object(run: Run) -> bool:
    return bool(run._r.xpath(".//w:drawing | .//w:pict | .//w:object"))


def _replace_span(paragraph, start: int, end: int, replacement: str) -> None:
    """Replace paragraph text[start:end] while preserving run formatting."""
    runs = _paragraph_runs(paragraph)
    spans: list[tuple[int, int, int]] = []
    cursor = 0
    for idx, run in enumerate(runs):
        run_text = run.text
        spans.append((idx, cursor, cursor + len(run_text)))
        cursor += len(run_text)

    start_run = end_run = None
    start_local = end_local = 0
    for idx, run_start, run_end in spans:
        if run_start <= start < run_end:
            start_run, start_local = idx, start - run_start
        if run_start < end <= run_end:
            end_run, end_local = idx, end - run_start

    if start_run is None or end_run is None:
        return

    if start_run == end_run:
        text = runs[start_run].text
        runs[start_run].text = text[:start_local] + replacement + text[end_local:]
        return

    start_text = runs[start_run].text
    runs[start_run].text = start_text[:start_local] + replacement
    for idx in range(start_run + 1, end_run):
        if not _has_embedded_object(runs[idx]):
            runs[idx].text = ""
    end_text = runs[end_run].text
    runs[end_run].text = end_text[end_local:]


def process_paragraph(paragraph, replace_patterns: list[tuple[str, str, str, str]]) -> int:
    if not replace_patterns:
        return 0

    count = 0
    for _, _, key, replacement in sorted(replace_patterns, key=lambda p: len(p[2]), reverse=True):
        if not key:
            continue
        text = paragraph_text(paragraph)
        if not text or key not in text:
            continue
        for start in reversed(_find_occurrences(text, key)):
            _replace_span(paragraph, start, start + len(key), replacement)
            count += 1
    return count


def _iter_header_footer_paragraphs(doc: Document) -> Iterator[Paragraph]:
    for section in doc.sections:
        containers = (
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
            section.even_page_header,
            section.even_page_footer,
        )
        for container in containers:
            try:
                if container.is_linked_to_previous:
                    continue
            except Exception:
                continue
            for p_element in container._element.iter(qn("w:p")):
                yield Paragraph(p_element, None)


def iter_document_paragraphs(doc: Document) -> Iterator[Paragraph]:
    for p_element in doc._body._element.iter(qn("w:p")):
        yield Paragraph(p_element, None)
    yield from _iter_header_footer_paragraphs(doc)


def _replace_document(document: Document, excel_row: pd.Series, rules: list[tuple[str, str]], keep_placeholder_on_empty: bool) -> int:
    patterns = precompute_replace_patterns(rules, excel_row, keep_placeholder_on_empty)
    total_replace = 0
    for paragraph in iter_document_paragraphs(document):
        total_replace += process_paragraph(paragraph, patterns)
    return total_replace


def replace_word_with_format(
    word_bytes: bytes,
    excel_row: pd.Series,
    rules: list[tuple[str, str]],
    keep_placeholder_on_empty: bool = False,
) -> tuple[io.BytesIO, int]:
    doc = Document(io.BytesIO(word_bytes))
    total_replace = _replace_document(doc, excel_row, rules, keep_placeholder_on_empty)
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output, total_replace


def _merge_fallback(replaced_files: list[ReplacedFile]) -> io.BytesIO:
    main_doc = Document(io.BytesIO(replaced_files[0].data.getvalue()))
    main_body = main_doc._body._element

    for file in replaced_files[1:]:
        sub_doc = Document(io.BytesIO(file.data.getvalue()))
        page_break_para = OxmlElement("w:p")
        page_break_run = OxmlElement("w:r")
        page_break = OxmlElement("w:br")
        page_break.set(qn("w:type"), "page")
        page_break_run.append(page_break)
        page_break_para.append(page_break_run)
        main_body.append(page_break_para)

        for element in sub_doc._body._element:
            main_body.append(copy.deepcopy(element))

    output = io.BytesIO()
    main_doc.save(output)
    output.seek(0)
    return output


def merge_word_documents(replaced_files: list[ReplacedFile]) -> io.BytesIO:
    if not replaced_files:
        raise ValueError("没有可合并文档")

    try:
        from docxcompose.composer import Composer
    except ImportError:
        logger.warning("docxcompose is not installed, using basic merge fallback")
        return _merge_fallback(replaced_files)

    composer = Composer(Document(io.BytesIO(replaced_files[0].data.getvalue())))
    for file in replaced_files[1:]:
        composer.append(Document(io.BytesIO(file.data.getvalue())))

    output = io.BytesIO()
    composer.save(output)
    output.seek(0)
    return output


def _build_filename(row: pd.Series, file_name_column: str, row_idx: int) -> str:
    if file_name_column in row.index and str(row[file_name_column]).strip():
        base = str(row[file_name_column]).strip()
    else:
        base = f"文件_{row_idx + 1}"
    if base.lower().endswith(".docx"):
        base = base[: -len(".docx")]
    return clean_filename(f"{base}.docx")


def _apply_filename_options(
    file_name: str,
    seq: int,
    use_seq_prefix: bool,
    seq_format: str,
    name_prefix: str,
    name_suffix: str,
    seq_value: str = "",
) -> str:
    stem, ext = file_name.rsplit(".", 1) if "." in file_name else (file_name, "docx")
    if seq_format == "01":
        seq_text = f"{seq:02d}"
    elif seq_format == "0001":
        seq_text = f"{seq:04d}"
    elif seq_format == "1.":
        seq_text = f"{seq}."
    elif seq_format == "一":
        seq_text = _chinese_number(seq)
    else:
        seq_text = str(seq)
    prefix_value = seq_value.strip() if seq_value.strip() else seq_text
    prefix = f"{prefix_value}_" if use_seq_prefix else ""
    final_name = f"{prefix}{name_prefix}{stem}{name_suffix}.{ext}"
    return clean_filename(final_name)


def _chinese_number(value: int) -> str:
    if value <= 0:
        return str(value)
    digits = "零一二三四五六七八九"
    if value < 10:
        return digits[value]
    if value < 20:
        return "十" + (digits[value % 10] if value % 10 else "")
    if value < 100:
        return digits[value // 10] + "十" + (digits[value % 10] if value % 10 else "")
    if value < 1000:
        hundreds, remainder = divmod(value, 100)
        text = digits[hundreds] + "百"
        if remainder == 0:
            return text
        if remainder < 10:
            return text + "零" + digits[remainder]
        if remainder < 20:
            return text + "一十" + (digits[remainder % 10] if remainder % 10 else "")
        tens, ones = divmod(remainder, 10)
        return text + digits[tens] + "十" + (digits[ones] if ones else "")
    return str(value)


def _dedupe_filename(filename: str, used: dict[str, int]) -> str:
    if filename not in used:
        used[filename] = 0
        return filename

    stem, ext = filename.rsplit(".", 1) if "." in filename else (filename, "docx")
    used[filename] += 1
    candidate = f"{stem}_{used[filename]}.{ext}"
    while candidate in used:
        used[filename] += 1
        candidate = f"{stem}_{used[filename]}.{ext}"
    used[candidate] = 0
    return candidate


def _validate_columns(excel_df: pd.DataFrame, rules: list[dict[str, str]], file_name_column: str) -> None:
    columns = {str(column) for column in excel_df.columns}
    missing = {rule["excel_column"] for rule in rules if rule["excel_column"] not in columns}
    if file_name_column and file_name_column not in columns:
        missing.add(file_name_column)
    if missing:
        raise ValueError("以下 Excel 列不存在：" + "、".join(sorted(missing)))


def run_replace(
    word_bytes: bytes,
    excel_bytes: bytes,
    rules: list[dict[str, str]],
    start_row: int,
    end_row: int,
    file_name_column: str,
    use_seq_prefix: bool = False,
    seq_format: str = "1",
    seq_column: str = "",
    name_prefix: str = "",
    name_suffix: str = "",
    sheet_name: str = "",
    keep_placeholder_on_empty: bool = False,
) -> dict[str, Any]:
    excel_df = read_excel(excel_bytes, sheet_name=sheet_name or 0)
    _validate_columns(excel_df, rules, file_name_column)

    max_row = len(excel_df)
    start_idx = max(0, start_row - 1)
    end_idx = min(end_row - 1, max_row - 1)
    if max_row == 0 or start_idx > end_idx:
        return {"files": [], "total": 0, "success": 0, "failed": 0, "replacements": 0, "details": []}

    rule_pairs = [(rule["keyword"], rule["excel_column"]) for rule in rules]
    out_files: list[ReplacedFile] = []
    total_replace = 0
    details: list[dict[str, Any]] = []
    used_names: dict[str, int] = {}

    template_doc = Document(io.BytesIO(word_bytes))

    for seq, row_idx in enumerate(range(start_idx, end_idx + 1), start=1):
        row = excel_df.iloc[row_idx]
        base_name = _build_filename(row, file_name_column, row_idx)
        seq_value = str(row.get(seq_column, "")).strip() if seq_column and seq_column in excel_df.columns else ""
        filename = _dedupe_filename(
            _apply_filename_options(
                file_name=base_name,
                seq=seq,
                use_seq_prefix=use_seq_prefix,
                seq_format=seq_format,
                name_prefix=name_prefix,
                name_suffix=name_suffix,
                seq_value=seq_value,
            ),
            used_names,
        )
        try:
            doc = copy.deepcopy(template_doc)
            cnt = _replace_document(doc, row, rule_pairs, keep_placeholder_on_empty)
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            item_id = uuid4().hex[:8]
            out_files.append(ReplacedFile(item_id=item_id, filename=filename, data=output, row_idx=row_idx, replace_count=cnt))
            total_replace += cnt
            details.append(
                {
                    "item_id": item_id,
                    "seq": seq,
                    "row_number": row_idx + HEADER_ROW_OFFSET,
                    "file_name": filename,
                    "status": "成功",
                    "replace_count": cnt,
                    "message": "",
                }
            )
        except Exception as exc:
            logger.warning("replace row failed: row_idx=%s err=%s", row_idx, exc)
            details.append(
                {
                    "item_id": uuid4().hex[:8],
                    "seq": seq,
                    "row_number": row_idx + HEADER_ROW_OFFSET,
                    "file_name": filename,
                    "status": "失败",
                    "replace_count": 0,
                    "message": str(exc),
                }
            )
            continue

    return {
        "files": out_files,
        "total": end_idx - start_idx + 1,
        "success": len(out_files),
        "failed": (end_idx - start_idx + 1) - len(out_files),
        "replacements": total_replace,
        "details": details,
    }


def cache_run(result: dict[str, Any]) -> str:
    global RUN_CACHE_BYTES

    run_id = uuid4().hex[:12]
    now = _utcnow()
    size = _entry_size(result)
    if size > settings.run_cache_max_bytes:
        raise ValueError(
            f"本次结果约 {size // (1024 * 1024)}MB，超过缓存上限 "
            f"{settings.run_cache_max_bytes // (1024 * 1024)}MB，请减少行数后重试"
        )

    with _CACHE_LOCK:
        _prune_cache(now)
        RUN_CACHE[run_id] = {
            **result,
            "_size": size,
            "created_at": now,
            "expires_at": now + timedelta(seconds=settings.run_cache_ttl_seconds),
        }
        RUN_CACHE_BYTES += size
        _prune_cache(now)
    return run_id


def get_run(run_id: str) -> dict[str, Any] | None:
    with _CACHE_LOCK:
        _prune_cache()
        return RUN_CACHE.get(run_id)


def get_run_file(run_id: str, item_id: str) -> ReplacedFile | None:
    with _CACHE_LOCK:
        _prune_cache()
        run = RUN_CACHE.get(run_id)
        if not run:
            return None
        for file in run.get("files", []):
            if file.item_id == item_id:
                return file
    return None


def delete_run_file(run_id: str, item_id: str) -> dict[str, Any] | None:
    global RUN_CACHE_BYTES

    with _CACHE_LOCK:
        _prune_cache()
        run = RUN_CACHE.get(run_id)
        if not run:
            return None

        removed = [file for file in run.get("files", []) if file.item_id == item_id]
        run["files"] = [file for file in run.get("files", []) if file.item_id != item_id]
        if not any(detail.get("item_id") == item_id for detail in run.get("details", [])):
            return run

        new_details: list[dict[str, Any]] = []
        seq = 1
        for detail in run.get("details", []):
            if detail.get("item_id") == item_id:
                continue
            detail["seq"] = seq
            seq += 1
            new_details.append(detail)

        removed_size = sum(len(file.data.getvalue()) for file in removed)
        run["details"] = new_details
        run["total"] = len(new_details)
        run["success"] = sum(1 for detail in new_details if detail.get("status") == "成功")
        run["failed"] = run["total"] - run["success"]
        run["replacements"] = sum(int(detail.get("replace_count", 0)) for detail in new_details)
        run["_size"] = max(0, int(run.get("_size", 0)) - removed_size)
        RUN_CACHE_BYTES = max(0, RUN_CACHE_BYTES - removed_size)
        return run


def _token_digest(run_id: str, expires_at: int) -> str:
    message = f"{run_id}:{expires_at}".encode("utf-8")
    return hmac.new(settings.export_token_secret.encode("utf-8"), message, hashlib.sha256).hexdigest()


def sign_export_token(run_id: str, ttl_seconds: int | None = None) -> str:
    ttl = settings.run_cache_ttl_seconds if ttl_seconds is None else ttl_seconds
    expires_at = int(_utcnow().timestamp()) + ttl
    return f"{expires_at}.{_token_digest(run_id, expires_at)}"


def verify_export_token(run_id: str, token: str) -> bool:
    try:
        expires_raw, digest = token.split(".", 1)
        expires_at = int(expires_raw)
    except (AttributeError, TypeError, ValueError):
        return False
    if expires_at < int(_utcnow().timestamp()):
        return False
    expected = _token_digest(run_id, expires_at)
    return hmac.compare_digest(expected, digest)


def export_zip(files: list[ReplacedFile]) -> io.BytesIO:
    zip_buffer = tempfile.SpooledTemporaryFile(max_size=32 * 1024 * 1024, mode="w+b")
    used: dict[str, int] = defaultdict(int)
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for file in files:
            name = file.filename
            used[name] += 1
            if used[name] > 1:
                stem, ext = name.rsplit(".", 1)
                name = f"{stem}_{used[name]-1}.{ext}"
            archive.writestr(name, file.data.getvalue())
    zip_buffer.seek(0)
    return zip_buffer
