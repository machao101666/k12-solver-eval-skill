#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Evaluate existing JSONL solving results against a prepared evaluation set."""

from __future__ import annotations

import argparse
import fcntl
import json
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

import stage7_score_judge
from config import cfg, column


BASE_DIR = Path(__file__).resolve().parent
# 列映射来自 config.yaml 的 excel.columns
COL_IMAGE_PATH = column('image_path')
COL_CONFIDENCE = column('confidence')
COL_PSEUDO_LABEL = column('pseudo_label')
COL_MODEL_OUTPUT = column('model_output')
COL_TIME = column('model_time')
COL_STATUS = column('model_status')
COL_SCORE = column('target_score')
COL_REASON = column('target_reason')
ONLY_CONFIDENCE = str(cfg('evaluation.only_confidence', 'high')).strip().lower()


class RunLock:
    """Prevent two evaluator processes from writing the same run concurrently."""

    def __init__(self, run_dir: Path):
        self.path = run_dir / ".existing_eval.lock"
        self.fd = None

    def __enter__(self):
        self.fd = os.open(self.path, os.O_CREAT | os.O_RDWR, 0o644)
        try:
            fcntl.flock(self.fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError as exc:
            owner = os.read(self.fd, 200).decode("utf-8", errors="replace").strip() or "unknown"
            os.close(self.fd)
            self.fd = None
            raise RuntimeError(f"评测目录正在被其他进程使用: {self.path} (owner={owner})") from exc
        os.ftruncate(self.fd, 0)
        os.write(self.fd, f"pid={os.getpid()}\n".encode("utf-8"))
        os.fsync(self.fd)
        return self

    def __exit__(self, exc_type, exc, tb):
        if self.fd is not None:
            fcntl.flock(self.fd, fcntl.LOCK_UN)
            os.close(self.fd)
            self.fd = None


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="对已有 evaluation.xlsx、题图和待测结果 JSONL 单独执行 high 样本评分与报告生成"
    )
    parser.add_argument("--xlsx", help="包含题目、confidence 和 pseudo_label 的 evaluation.xlsx")
    parser.add_argument("--images-dir", help="原始题图根目录，例如 evaluation/")
    parser.add_argument(
        "--results-jsonl",
        nargs="+",
        help="一个或多个待测结果 JSONL；每条记录需包含 response 和 images[0].path",
    )
    parser.add_argument("--output-root", help="批次输出根目录；默认写入 runs/existing_eval_<timestamp>")
    parser.add_argument("--resume-run", help="已有独立评测目录；只补跑 AI/AJ 非法或为空的 high 行并重新生成报告")
    parser.add_argument("--concurrency", type=int, default=8, help="评分并发数，默认 8")
    parser.add_argument("--invalid-retries", type=int, default=2, help="非法裁判响应额外重试次数，默认 2")
    parser.add_argument("--report-title", help="单个 JSONL 时可指定报告标题")
    parser.add_argument("--cross-model-xlsx", dest="cross_model_xlsx", help="Stage 8 横向评分表的来源 Excel，可选；未指定时用本次评测 Excel")
    parser.add_argument("--prepare-only", action="store_true", help="只准备独立数据目录，不调用评分接口和报告")
    return parser.parse_args()


def resolve_path(value: str, *, base: Path = BASE_DIR) -> Path:
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = base / path
    return path.resolve()


def normalized_relative_path(value: str) -> str:
    path = str(value or "").strip().replace("\\", "/")
    if not path:
        raise ValueError("图片路径为空")
    if path.startswith("/") or re.match(r"^[A-Za-z]:/", path):
        raise ValueError(f"图片路径必须是相对路径: {value}")
    parts = []
    for part in path.split("/"):
        if part in ("", "."):
            continue
        if part == "..":
            raise ValueError(f"图片路径禁止包含 ..: {value}")
        parts.append(part)
    if not parts:
        raise ValueError(f"图片路径无有效内容: {value}")
    return "/".join(parts)


def canonical_image_key(value: str, images_dir_name: str) -> str:
    path = normalized_relative_path(value)
    prefixes = ["evaluation/", f"{images_dir_name}/"]
    changed = True
    while changed:
        changed = False
        for prefix in prefixes:
            if path.startswith(prefix):
                path = path[len(prefix):]
                changed = True
    return normalized_relative_path(path)


def load_result_records(path: Path, images_dir_name: str) -> tuple[list[tuple[str, dict]], dict[str, tuple[str, dict]]]:
    records = []
    by_image = {}
    for line_number, raw_line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        if not raw_line.strip():
            continue
        try:
            record = json.loads(raw_line)
        except json.JSONDecodeError as exc:
            raise ValueError(f"{path} 第 {line_number} 行不是合法 JSON: {exc}") from exc
        images = record.get("images")
        response = record.get("response")
        if not isinstance(images, list) or not images or not isinstance(images[0], dict) or not images[0].get("path"):
            raise ValueError(f"{path} 第 {line_number} 行缺少 images[0].path")
        if not isinstance(response, str) or not response.strip():
            raise ValueError(f"{path} 第 {line_number} 行缺少非空 response")
        key = canonical_image_key(images[0]["path"], images_dir_name)
        if key in by_image:
            raise ValueError(f"{path} 存在重复图片路径: {images[0]['path']}")
        item = (raw_line, record)
        records.append(item)
        by_image[key] = item
    if not records:
        raise ValueError(f"{path} 没有有效 JSONL 记录")
    return records, by_image


def locate_source_image(images_dir: Path, excel_path: str, canonical_key: str) -> Path:
    images_root = images_dir.resolve()
    candidates = [
        images_root / canonical_key,
        images_root / normalized_relative_path(excel_path),
    ]
    seen = set()
    for candidate in candidates:
        candidate = candidate.resolve()
        try:
            candidate.relative_to(images_root)
        except ValueError as exc:
            raise ValueError(f"题图路径逃出图片根目录: {candidate}") from exc
        if candidate in seen:
            continue
        seen.add(candidate)
        if candidate.is_file():
            return candidate
    raise FileNotFoundError(f"找不到题图: Excel={excel_path}, canonical={canonical_key}, images_dir={images_dir}")


def validate_source_workbook(xlsx_path: Path) -> int:
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=False)
    ws = wb.active
    try:
        if ws.max_column < COL_PSEUDO_LABEL:
            raise ValueError(f"Excel 至少需要 {COL_PSEUDO_LABEL} 列（AC pseudo_label），当前只有 {ws.max_column} 列")
        high_rows = 0
        for row in range(2, ws.max_row + 1):
            confidence_cell = ws.cell(row, COL_CONFIDENCE)
            if confidence_cell.data_type == "f":
                raise ValueError(f"Excel 第 {row} 行 confidence 是公式；请先固化为 high/medium/low 文本")
            confidence = str(confidence_cell.value or "").strip().lower()
            if confidence != ONLY_CONFIDENCE:
                continue
            high_rows += 1
            if not ws.cell(row, COL_IMAGE_PATH).value:
                raise ValueError(f"Excel 第 {row} 行 confidence=high 但题目路径为空")
            if not ws.cell(row, COL_PSEUDO_LABEL).value:
                raise ValueError(f"Excel 第 {row} 行 confidence=high 但 pseudo_label 为空")
        if not high_rows:
            raise ValueError("Excel 中没有 confidence=high 的记录")
        return high_rows
    finally:
        wb.close()


def prepare_run(
    xlsx_path: Path,
    images_dir: Path,
    result_path: Path,
    run_dir: Path,
) -> dict:
    _, records_by_image = load_result_records(result_path, images_dir.name)
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    for row in range(ws.max_row, 1, -1):
        confidence_cell = ws.cell(row, COL_CONFIDENCE)
        if confidence_cell.data_type == "f":
            raise ValueError(f"Excel 第 {row} 行 confidence 是公式；请先固化为 high/medium/low 文本")
        confidence = str(confidence_cell.value or "").strip().lower()
        if confidence != ONLY_CONFIDENCE:
            ws.delete_rows(row, 1)

    selected_raw_lines = []
    selected_paths = set()
    run_dir.mkdir(parents=True)

    for row in range(2, ws.max_row + 1):
        excel_image_path = str(ws.cell(row, COL_IMAGE_PATH).value or "")
        key = canonical_image_key(excel_image_path, images_dir.name)
        item = records_by_image.get(key)
        if item is None:
            raise ValueError(f"{result_path} 缺少 high 题目的待测结果: {excel_image_path}")
        raw_line, record = item
        source_image = locate_source_image(images_dir, excel_image_path, key)
        output_image_path = Path("evaluation") / key
        target_image = (run_dir / output_image_path).resolve()
        output_images_root = (run_dir / "evaluation").resolve()
        try:
            target_image.relative_to(output_images_root)
        except ValueError as exc:
            raise ValueError(f"输出题图路径逃出运行目录: {target_image}") from exc
        target_image.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_image, target_image)

        ws.cell(row, COL_IMAGE_PATH).value = output_image_path.as_posix()
        output_cell = ws.cell(row, COL_MODEL_OUTPUT)
        output_cell.value = record["response"]
        output_cell.data_type = "s"
        ws.cell(row, COL_TIME).value = None
        ws.cell(row, COL_STATUS).value = "待评测"
        for column in range(COL_SCORE, ws.max_column + 1):
            ws.cell(row, column).value = None

        selected_raw_lines.append(raw_line)
        selected_paths.add(key)

    extra_paths = set(records_by_image) - selected_paths
    output_xlsx = run_dir / "evaluation.xlsx"
    wb.save(output_xlsx)

    filtered_jsonl = run_dir / f"{result_path.stem}_high_{len(selected_raw_lines)}.jsonl"
    filtered_jsonl.write_text("\n".join(selected_raw_lines) + "\n", encoding="utf-8")

    metadata = {
        "mode": "existing_results_evaluation",
        "created_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "source_xlsx": str(xlsx_path),
        "source_images_dir": str(images_dir),
        "source_results_jsonl": str(result_path),
        "high_count": len(selected_raw_lines),
        "ignored_jsonl_records": len(extra_paths),
        "evaluation": "evaluation.xlsx",
        "filtered_results_jsonl": filtered_jsonl.name,
        "report": "report.md",
        "status": "prepared",
    }
    (run_dir / "run.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return metadata


def validate_scores(run_dir: Path, expected: int) -> dict[int, int]:
    workbook = openpyxl.load_workbook(run_dir / "evaluation.xlsx", read_only=True, data_only=True)
    ws = workbook.active
    distribution = {0: 0, 1: 0, 2: 0, 3: 0}
    invalid = []
    try:
        for row in range(2, ws.max_row + 1):
            score = ws.cell(row, COL_SCORE).value
            reason = str(ws.cell(row, COL_REASON).value or "").strip()
            if score not in distribution or not reason:
                invalid.append((row, ws.cell(row, 1).value, score, str(reason or "")[:200]))
                continue
            distribution[score] += 1
    finally:
        workbook.close()
    valid = sum(distribution.values())
    if valid != expected or invalid:
        raise RuntimeError(f"评分不完整: valid={valid}/{expected}, invalid={invalid[:20]}")
    return distribution


def generate_report(run_dir: Path, title: str, cross_model_xlsx: Path | None) -> None:
    command = [
        sys.executable,
        str(BASE_DIR / "stage8_generate_report.py"),
        str(run_dir / "evaluation.xlsx"),
        "--work-dir",
        str(run_dir),
        "--output",
        "report.md",
        "--title",
        title,
    ]
    if cross_model_xlsx:
        command.extend(["--cross-model-xlsx", str(cross_model_xlsx)])
    subprocess.run(command, cwd=BASE_DIR, check=True)


def update_run_metadata(run_dir: Path, **values) -> None:
    path = run_dir / "run.json"
    metadata = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {
        "mode": "existing_results_evaluation",
        "created_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "run_dir": str(run_dir),
    }
    metadata.update(values)
    metadata["updated_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
    path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def main() -> int:
    args = parse_args()
    cross_model_xlsx = resolve_path(args.cross_model_xlsx) if args.cross_model_xlsx else None
    if cross_model_xlsx and not cross_model_xlsx.is_file():
        raise SystemExit(f"横向评分来源 Excel 不存在: {cross_model_xlsx}")

    if args.resume_run:
        if args.prepare_only:
            raise SystemExit("--resume-run 不能与 --prepare-only 同时使用")
        run_dir = resolve_path(args.resume_run)
        run_metadata_path = run_dir / "run.json"
        if not run_metadata_path.is_file():
            raise SystemExit(f"拒绝补跑非独立评测目录，缺少 run.json: {run_dir}")
        try:
            run_metadata = json.loads(run_metadata_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise SystemExit(f"run.json 不是合法 JSON: {run_metadata_path}") from exc
        if run_metadata.get("mode") != "existing_results_evaluation":
            raise SystemExit(f"拒绝补跑非 existing_results_evaluation 目录: {run_dir}")
        evaluation_path = run_dir / "evaluation.xlsx"
        if not evaluation_path.is_file():
            raise SystemExit(f"续跑目录缺少 evaluation.xlsx: {run_dir}")
        high_count = validate_source_workbook(evaluation_path)
        with RunLock(run_dir):
            update_run_metadata(run_dir, status="scoring_invalid")
            stage7_score_judge.process_excel(
                work_dir=str(run_dir),
                concurrency=args.concurrency,
                targets="target",
                invalid_retries=args.invalid_retries,
                only_invalid=True,
            )
            distribution = validate_scores(run_dir, high_count)
            title = args.report_title or f"{run_dir.name} 高置信度{high_count}题评测报告"
            generate_report(run_dir, title, cross_model_xlsx)
            update_run_metadata(run_dir, status="success", score_distribution=distribution)
        summary = {
            "run_dir": str(run_dir),
            "status": "success",
            "high_count": high_count,
            "score_distribution": distribution,
            "evaluation": str(evaluation_path),
            "report": str(run_dir / "report.md"),
        }
        print(json.dumps({"mode": "existing_results_resume", "runs": [summary]}, ensure_ascii=False, indent=2))
        return 0

    if not args.xlsx or not args.images_dir or not args.results_jsonl:
        raise SystemExit("新评测必须同时提供 --xlsx、--images-dir 和 --results-jsonl；补跑已有目录请使用 --resume-run")
    xlsx_path = resolve_path(args.xlsx)
    images_dir = resolve_path(args.images_dir)
    result_paths = [resolve_path(value) for value in args.results_jsonl]
    if not xlsx_path.is_file():
        raise SystemExit(f"Excel 不存在: {xlsx_path}")
    if not images_dir.is_dir():
        raise SystemExit(f"题图目录不存在: {images_dir}")
    for result_path in result_paths:
        if not result_path.is_file():
            raise SystemExit(f"结果 JSONL 不存在: {result_path}")

    high_count = validate_source_workbook(xlsx_path)
    run_names = [f"{result_path.stem}_high_{high_count}" for result_path in result_paths]
    duplicates = sorted({name for name in run_names if run_names.count(name) > 1})
    if duplicates:
        raise SystemExit(f"结果 JSONL 文件名冲突，会生成相同子目录: {duplicates}")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_root = resolve_path(args.output_root) if args.output_root else BASE_DIR / "runs" / f"existing_eval_{timestamp}"
    if output_root.exists():
        raise SystemExit(f"输出目录已存在，拒绝覆盖: {output_root}")
    output_root.mkdir(parents=True)

    summaries = []
    for result_path in result_paths:
        run_dir = output_root / f"{result_path.stem}_high_{high_count}"
        metadata = prepare_run(xlsx_path, images_dir, result_path, run_dir)
        if args.prepare_only:
            summaries.append({"run_dir": str(run_dir), "status": "prepared", "high_count": high_count})
            continue

        with RunLock(run_dir):
            update_run_metadata(run_dir, status="scoring")
            stage7_score_judge.process_excel(
                work_dir=str(run_dir),
                concurrency=args.concurrency,
                targets="target",
                invalid_retries=args.invalid_retries,
            )
            distribution = validate_scores(run_dir, high_count)
            title = args.report_title if args.report_title and len(result_paths) == 1 else f"{result_path.stem} 高置信度{high_count}题评测报告"
            generate_report(run_dir, title, cross_model_xlsx)
            update_run_metadata(run_dir, status="success", score_distribution=distribution)
        summaries.append({
            "run_dir": str(run_dir),
            "status": "success",
            "high_count": high_count,
            "score_distribution": distribution,
            "evaluation": str(run_dir / "evaluation.xlsx"),
            "report": str(run_dir / "report.md"),
        })

    print(json.dumps({"mode": "existing_results_evaluation", "output_root": str(output_root), "runs": summaries}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
