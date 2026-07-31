#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动化评测总控脚本：每次运行创建独立 run 目录，并把输入、过程文件、日志、Excel、报告保存为一个整体。

用法：
  python3 run_pipeline.py --csv 1.csv
  python3 run_pipeline.py --run-dir runs/run_20260617_183012 --from-stage 5
  python3 run_pipeline.py --run-dir runs/run_20260617_183012 --only-stage 8
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
RUNS_DIR = BASE_DIR / "runs"
INDEX_PATH = RUNS_DIR / "INDEX.md"

STAGES = {
    1: {
        "name": "OCR 切题",
        "command": lambda ctx: [sys.executable, str(BASE_DIR / "stage1_ocr_cut_questions.py"), str(ctx["input_csv_path"]), "--work-dir", str(ctx["run_dir"])],
    },
    2: {
        "name": "有效性判定与分类",
        "command": lambda ctx: [sys.executable, str(BASE_DIR / "stage2_clean_classify.py"), "--work-dir", str(ctx["run_dir"])],
    },
    3: {
        "name": "5 模型解题",
        "command": lambda ctx: [sys.executable, str(BASE_DIR / "stage3_multi_model_solve.py"), "--work-dir", str(ctx["run_dir"])],
    },
    4: {
        "name": "共识裁判",
        "command": lambda ctx: [sys.executable, str(BASE_DIR / "stage4_consensus_judge.py"), "--work-dir", str(ctx["run_dir"])],
    },
    5: {
        "name": "BOS 上传并写 AE",
        "command": lambda ctx: [sys.executable, str(BASE_DIR / "stage5_upload_bos_urls.py"), str(ctx["image_ocr_clean_path"]), "--write-excel", "--work-dir", str(ctx["run_dir"]), "--xlsx", str(ctx["evaluation_path"])],
    },
    6: {
        "name": "待测模型解题",
        "command": lambda ctx: [sys.executable, str(BASE_DIR / "stage6_target_solve.py"), "--work-dir", str(ctx["run_dir"])],
    },
    7: {
        "name": "评分裁判",
        "command": lambda ctx: [sys.executable, str(BASE_DIR / "stage7_score_judge.py"), "--work-dir", str(ctx["run_dir"]), "--concurrency", str(ctx.get("stage7_concurrency", 8))],
    },
    8: {
        "name": "报告生成",
        "command": lambda ctx: [sys.executable, str(BASE_DIR / "stage8_generate_report.py"), str(ctx["evaluation_path"]), "--output", str(ctx["report_path"]), "--work-dir", str(ctx["run_dir"])],
    },
}


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def make_run_id() -> str:
    return "run_" + datetime.now().strftime("%Y%m%d_%H%M%S")


def load_json(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_json(path: Path, data: dict[str, Any]) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def run_artifacts(run_dir: Path) -> dict[str, str]:
    return {
        "evaluation": str(run_dir / "evaluation.xlsx") if (run_dir / "evaluation.xlsx").exists() else "",
        "report": str(run_dir / "report.md") if (run_dir / "report.md").exists() else "",
        "logs": str(run_dir / "logs") if (run_dir / "logs").exists() else "",
        "progress": str(run_dir / "progress.jsonl") if (run_dir / "progress.jsonl").exists() else "",
    }


def emit_progress_event(
    run_dir: Path,
    event: str,
    *,
    stage_no: int | None = None,
    stage_name: str = "",
    status: str = "",
    message: str = "",
    payload: dict[str, Any] | None = None,
) -> dict[str, Any]:
    data = {
        "time": now_iso(),
        "run_id": run_dir.name,
        "run_dir": str(run_dir),
        "event": event,
        "stage": stage_no,
        "stage_name": stage_name,
        "status": status,
        "message": message,
        "artifacts": run_artifacts(run_dir),
    }
    if payload:
        data["payload"] = payload

    progress_path = run_dir / "progress.jsonl"
    with progress_path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(data, ensure_ascii=False) + "\n")

    print("[PROGRESS] " + json.dumps(data, ensure_ascii=False), flush=True)
    return data


def resolve_run_dir(value: str | None) -> Path:
    if value:
        run_dir = Path(value).expanduser()
        if not run_dir.is_absolute():
            run_dir = BASE_DIR / run_dir
        return run_dir
    latest = latest_run_dir()
    if latest is None:
        raise SystemExit("没有找到 runs/run_*/run.json")
    return latest


def latest_run_dir() -> Path | None:
    candidates = [path.parent for path in RUNS_DIR.glob("run_*/run.json")]
    if not candidates:
        return None
    return max(candidates, key=lambda path: (path / "run.json").stat().st_mtime)


def latest_pipeline_log(run_dir: Path) -> Path | None:
    log_candidates = sorted((run_dir / "logs").glob("run_*/pipeline.log"), key=lambda p: p.stat().st_mtime, reverse=True)
    if log_candidates:
        return log_candidates[0]
    direct = run_dir / "logs" / "pipeline.log"
    return direct if direct.exists() else None


def tail_lines(path: Path, limit: int = 40) -> list[str]:
    if not path.exists():
        return []
    lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    return lines[-limit:]


def latest_progress_file(run_dir: Path) -> Path | None:
    path = run_dir / "progress.jsonl"
    return path if path.exists() else None


def print_progress_tail(run_dir: Path, limit: int = 20) -> None:
    progress_path = latest_progress_file(run_dir)
    if progress_path is None:
        return
    print(f"进度事件：{progress_path}")
    for line in tail_lines(progress_path, limit):
        try:
            event = json.loads(line)
            stage = event.get("stage")
            stage_part = f"Stage {stage} " if stage else ""
            print(f"  {event.get('time', '')} {stage_part}{event.get('event', '')}: {event.get('message', '')}")
        except Exception:
            print(line)


def print_run_status(run_dir: Path) -> str:
    run_json_path = run_dir / "run.json"
    meta = load_json(run_json_path)
    if not meta:
        raise SystemExit(f"找不到或无法读取 run.json: {run_json_path}")

    status = meta.get("status", "unknown")
    print(f"Run 目录：{run_dir}")
    print(f"状态：{status}")
    print(f"创建时间：{meta.get('created_at', '')}")
    print(f"更新时间：{meta.get('updated_at', '')}")

    stages = meta.get("stages", {})
    if stages:
        print("阶段状态：")
        for key in sorted(stages, key=lambda name: int(name.replace("stage", "")) if name.replace("stage", "").isdigit() else 999):
            stage = stages[key]
            name = stage.get("name", "")
            stage_status = stage.get("status", "")
            elapsed = stage.get("elapsed_seconds", "")
            assessment = stage.get("assessment", "")
            suffix = f"，耗时 {elapsed}s" if elapsed != "" else ""
            print(f"  {key} {name}: {stage_status}{suffix}")
            if assessment:
                print(f"    {assessment}")

    artifacts = meta.get("artifacts", {})
    if artifacts:
        print("产物：")
        for name, rel_path in artifacts.items():
            if rel_path:
                print(f"  {name}: {run_dir / rel_path}")

    log_path = latest_pipeline_log(run_dir)
    if log_path:
        print(f"最新日志：{log_path}")
    progress_path = latest_progress_file(run_dir)
    if progress_path:
        print(f"进度事件：{progress_path}")
    return status


def print_latest_run() -> None:
    run_dir = latest_run_dir()
    if run_dir is None:
        raise SystemExit("没有找到历史 run")
    print_run_status(run_dir)


def print_run_tail(run_dir: Path, limit: int = 40) -> None:
    print_progress_tail(run_dir, limit=limit)
    log_path = latest_pipeline_log(run_dir)
    if log_path is None:
        print(f"未找到 pipeline.log: {run_dir / 'logs'}")
        return
    print(f"日志：{log_path}")
    for line in tail_lines(log_path, limit):
        print(line)


def watch_run(run_dir: Path, interval: int = 15, tail_limit: int = 20) -> None:
    terminal_status = {"success", "failed", "partial", "dry_run"}
    seen_log_lines = 0
    seen_progress_lines = 0
    while True:
        print("\n" + "=" * 70)
        status = print_run_status(run_dir)
        progress_path = latest_progress_file(run_dir)
        if progress_path:
            progress_lines = progress_path.read_text(encoding="utf-8", errors="replace").splitlines()
            new_progress = progress_lines[seen_progress_lines:]
            if not new_progress and seen_progress_lines == 0:
                new_progress = progress_lines[-tail_limit:]
            if new_progress:
                print("新增进度事件：")
                for line in new_progress[-tail_limit:]:
                    try:
                        event = json.loads(line)
                        stage = event.get("stage")
                        stage_part = f"Stage {stage} " if stage else ""
                        print(f"  {event.get('time', '')} {stage_part}{event.get('event', '')}: {event.get('message', '')}")
                    except Exception:
                        print(line)
            seen_progress_lines = len(progress_lines)
        log_path = latest_pipeline_log(run_dir)
        if log_path:
            lines = log_path.read_text(encoding="utf-8", errors="replace").splitlines()
            new_lines = lines[seen_log_lines:]
            if not new_lines and seen_log_lines == 0:
                new_lines = lines[-tail_limit:]
            if new_lines:
                print("最近日志：")
                for line in new_lines[-tail_limit:]:
                    print(line)
            seen_log_lines = len(lines)
        if status in terminal_status:
            print(f"\n监控结束：{status}")
            break
        time.sleep(max(1, interval))


def prepare_run_dir(args: argparse.Namespace) -> tuple[Path, dict[str, Any]]:
    RUNS_DIR.mkdir(exist_ok=True)

    if args.run_dir:
        run_dir = Path(args.run_dir).expanduser()
        if not run_dir.is_absolute():
            run_dir = BASE_DIR / run_dir
        run_dir.mkdir(parents=True, exist_ok=True)
        run_id = run_dir.name
    else:
        run_id = args.run_id or make_run_id()
        run_dir = RUNS_DIR / run_id
        if run_dir.exists():
            raise SystemExit(f"run 目录已存在：{run_dir}")
        run_dir.mkdir(parents=True)

    run_json_path = run_dir / "run.json"
    run_meta = load_json(run_json_path)
    run_meta.setdefault("run_id", run_dir.name)
    run_meta.setdefault("created_at", now_iso())
    run_meta.setdefault("status", "created")
    run_meta.setdefault("base_dir", str(BASE_DIR))
    run_meta.setdefault("run_dir", str(run_dir))
    run_meta.setdefault("code_mode", "single_source_with_work_dir")
    run_meta.setdefault("stages", {})

    if args.csv:
        csv_path = Path(args.csv).expanduser()
        if not csv_path.is_absolute():
            csv_path = BASE_DIR / csv_path
        if not csv_path.exists():
            raise SystemExit(f"输入 CSV 不存在：{csv_path}")
        shutil.copy2(csv_path, run_dir / "input.csv")
        run_meta["input_csv_source"] = str(csv_path)
        run_meta["input_csv"] = "input.csv"
    elif not (run_dir / "input.csv").exists() and stage_range(args)[0] <= 1:
        raise SystemExit("首次从 Stage 1 运行必须提供 --csv")

    if not (run_dir / "evaluation.xlsx").exists():
        template = Path(args.template).expanduser() if args.template else BASE_DIR / "evaluation.xlsx"
        if not template.is_absolute():
            template = BASE_DIR / template
        if not template.exists():
            raise SystemExit(f"evaluation 模板不存在：{template}")
        shutil.copy2(template, run_dir / "evaluation.xlsx")
        run_meta["evaluation_template"] = str(template)

    run_meta["evaluation"] = "evaluation.xlsx"
    run_meta["report"] = "report.md"
    save_json(run_json_path, run_meta)
    return run_dir, run_meta


def stage_range(args: argparse.Namespace) -> list[int]:
    if args.latest or args.status is not None or args.tail is not None or args.watch is not None:
        return []
    if args.only_stage:
        return [args.only_stage]
    start = args.from_stage or 1
    end = args.to_stage or 8
    if start < 1 or end > 8 or start > end:
        raise SystemExit("阶段范围必须满足 1 <= from-stage <= to-stage <= 8")
    return list(range(start, end + 1))


def update_stage_meta(run_dir: Path, stage_no: int, values: dict[str, Any]) -> dict[str, Any]:
    run_json_path = run_dir / "run.json"
    run_meta = load_json(run_json_path)
    run_meta.setdefault("stages", {})
    key = f"stage{stage_no}"
    current = run_meta["stages"].setdefault(key, {})
    current.update(values)
    run_meta["updated_at"] = now_iso()
    save_json(run_json_path, run_meta)
    return run_meta


def excel_has_value(path: Path, column: int, *, reject_error: bool = False) -> bool:
    if not path.exists():
        return False
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    try:
        for row in range(2, ws.max_row + 1):
            value = ws.cell(row=row, column=column).value
            if value is None or str(value).strip() == "":
                continue
            if reject_error and str(value).strip().upper().startswith("ERROR"):
                continue
            return True
        return False
    finally:
        wb.close()


def assess_required_outputs(run_dir: Path, stage_no: int) -> tuple[str, str]:
    xlsx_path = run_dir / "evaluation.xlsx"
    if stage_no == 6:
        if excel_has_value(xlsx_path, 32, reject_error=False):
            return "success", "Stage 6 已写入 AF 待测模型输出或错误记录，单行失败不阻断流水线"
        return "success", "Stage 6 未产生 AF 内容；按失败写表、不阻断策略继续后续阶段"
    if stage_no == 7:
        if excel_has_value(xlsx_path, 35, reject_error=False) or excel_has_value(xlsx_path, 36, reject_error=False):
            return "success", "Stage 7 已写入评分结果或错误原因，单请求失败不阻断流水线"
        return "success", "Stage 7 未产生评分内容；按失败写表、不阻断策略继续后续阶段"
    if stage_no == 8:
        if not (run_dir / "report.md").exists():
            return "failed", "Stage 8 未生成 report.md"
        return "success", "Stage 8 已生成报告"
    return "success", ""


def run_context(run_dir: Path) -> dict[str, Path]:
    return {
        "run_dir": run_dir,
        "input_csv_path": run_dir / "input.csv",
        "evaluation_path": run_dir / "evaluation.xlsx",
        "image_ocr_clean_path": run_dir / "image_ocr_clean",
        "report_path": run_dir / "report.md",
    }


def run_stage(run_dir: Path, stage_no: int, dry_run: bool = False, stage7_concurrency: int = 8) -> None:
    stage = STAGES[stage_no]
    ctx = run_context(run_dir)
    ctx["stage7_concurrency"] = stage7_concurrency
    command = stage["command"](ctx)

    print(f"\n[Stage {stage_no}] {stage['name']}")
    print("$ " + " ".join(command))
    emit_progress_event(
        run_dir,
        "stage_started",
        stage_no=stage_no,
        stage_name=stage["name"],
        status="running",
        message=f"Stage {stage_no} 开始：{stage['name']}",
        payload={"command": command},
    )

    update_stage_meta(run_dir, stage_no, {
        "name": stage["name"],
        "status": "running",
        "started_at": now_iso(),
        "command": command,
    })

    if dry_run:
        update_stage_meta(run_dir, stage_no, {"status": "dry_run", "ended_at": now_iso()})
        emit_progress_event(
            run_dir,
            "dry_run_stage",
            stage_no=stage_no,
            stage_name=stage["name"],
            status="dry_run",
            message=f"Stage {stage_no} dry-run：{stage['name']}",
        )
        return

    started_at = datetime.now()
    env = os.environ.copy()
    env["AUTO_EVAL_WORK_DIR"] = str(run_dir)
    result = subprocess.run(command, cwd=run_dir, text=True, env=env)
    elapsed = round((datetime.now() - started_at).total_seconds(), 3)

    if result.returncode != 0:
        stage_status = "partial" if stage_no in {6, 7, 8} and (run_dir / "report.md").exists() else "failed"
        update_stage_meta(run_dir, stage_no, {
            "status": stage_status,
            "ended_at": now_iso(),
            "elapsed_seconds": elapsed,
            "returncode": result.returncode,
        })
        emit_progress_event(
            run_dir,
            "stage_failed",
            stage_no=stage_no,
            stage_name=stage["name"],
            status=stage_status,
            message=f"Stage {stage_no} 失败：{stage['name']}，returncode={result.returncode}",
            payload={"elapsed_seconds": elapsed, "returncode": result.returncode},
        )
        raise SystemExit(f"Stage {stage_no} 失败，returncode={result.returncode}，run_dir={run_dir}")

    assessed_status, assessed_reason = assess_required_outputs(run_dir, stage_no)
    update_stage_meta(run_dir, stage_no, {
        "status": assessed_status,
        "ended_at": now_iso(),
        "elapsed_seconds": elapsed,
        "returncode": result.returncode,
        "assessment": assessed_reason,
    })
    emit_progress_event(
        run_dir,
        "stage_completed" if assessed_status == "success" else "stage_partial",
        stage_no=stage_no,
        stage_name=stage["name"],
        status=assessed_status,
        message=f"Stage {stage_no} 完成：{stage['name']}，状态={assessed_status}",
        payload={"elapsed_seconds": elapsed, "returncode": result.returncode, "assessment": assessed_reason},
    )
    if assessed_status != "success":
        raise SystemExit(f"Stage {stage_no} 产物不完整：{assessed_reason}，run_dir={run_dir}")


def ensure_index() -> None:
    RUNS_DIR.mkdir(exist_ok=True)
    if not INDEX_PATH.exists():
        INDEX_PATH.write_text(
            "# Runs Index\n\n"
            "| run_id | status | created_at | updated_at | input | evaluation | report |\n"
            "|---|---|---|---|---|---|---|\n",
            encoding="utf-8",
        )


def refresh_index() -> None:
    ensure_index()
    rows = []
    for run_json in sorted(RUNS_DIR.glob("run_*/run.json"), reverse=True):
        meta = load_json(run_json)
        run_id = meta.get("run_id", run_json.parent.name)
        status = meta.get("status", "")
        created_at = meta.get("created_at", "")
        updated_at = meta.get("updated_at", "")
        input_csv = meta.get("input_csv", "")
        evaluation = meta.get("evaluation", "")
        report = meta.get("report", "") if (run_json.parent / str(meta.get("report", ""))).exists() else ""
        rows.append(f"| {run_id} | {status} | {created_at} | {updated_at} | {input_csv} | {evaluation} | {report} |")

    INDEX_PATH.write_text(
        "# Runs Index\n\n"
        "| run_id | status | created_at | updated_at | input | evaluation | report |\n"
        "|---|---|---|---|---|---|---|\n"
        + "\n".join(rows)
        + ("\n" if rows else ""),
        encoding="utf-8",
    )


def finalize_run(run_dir: Path, status: str) -> None:
    run_json_path = run_dir / "run.json"
    run_meta = load_json(run_json_path)
    if status == "success":
        stage_statuses = [stage.get("status") for stage in run_meta.get("stages", {}).values()]
        if any(stage_status in {"partial", "failed"} for stage_status in stage_statuses):
            status = "partial" if (run_dir / "report.md").exists() else "failed"
    run_meta["status"] = status
    run_meta["updated_at"] = now_iso()
    if status in {"success", "failed", "partial"}:
        run_meta["ended_at"] = now_iso()
    run_meta["artifacts"] = {
        "evaluation": "evaluation.xlsx" if (run_dir / "evaluation.xlsx").exists() else "",
        "report": "report.md" if (run_dir / "report.md").exists() else "",
        "logs": "logs" if (run_dir / "logs").exists() else "",
        "image": "image" if (run_dir / "image").exists() else "",
        "image_ocr": "image_ocr" if (run_dir / "image_ocr").exists() else "",
        "image_ocr_clean": "image_ocr_clean" if (run_dir / "image_ocr_clean").exists() else "",
    }
    save_json(run_json_path, run_meta)
    refresh_index()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="自动化评测总控：每次运行保存到独立 runs/run_<timestamp>/ 目录")
    parser.add_argument("--csv", help="输入 CSV 路径；首次从 Stage 1 运行必填")
    parser.add_argument("--template", help="evaluation.xlsx 模板路径，默认使用根目录 evaluation.xlsx")
    parser.add_argument("--run-id", help="指定 run_id，默认 run_YYYYMMDD_HHMMSS")
    parser.add_argument("--run-dir", help="使用已有 run 目录继续运行")
    parser.add_argument("--from-stage", type=int, help="从指定阶段开始，默认 1")
    parser.add_argument("--to-stage", type=int, help="运行到指定阶段，默认 8")
    parser.add_argument("--only-stage", type=int, help="只运行某个阶段")
    parser.add_argument("--dry-run", action="store_true", help="只打印命令并写 run.json，不实际执行")
    parser.add_argument("--latest", action="store_true", help="查看最近一次 run 的状态")
    parser.add_argument("--status", metavar="RUN_DIR", nargs="?", const="", help="查看指定 run 目录状态；不传值时查看最近一次 run")
    parser.add_argument("--tail", metavar="RUN_DIR", nargs="?", const="", help="输出指定 run 的最近日志；不传值时查看最近一次 run")
    parser.add_argument("--watch", metavar="RUN_DIR", nargs="?", const="", help="持续观察指定 run 直到结束；不传值时观察最近一次 run")
    parser.add_argument("--tail-lines", type=int, default=40, help="--tail 输出行数，默认40")
    parser.add_argument("--watch-interval", type=int, default=15, help="--watch 刷新间隔秒数，默认15")
    parser.add_argument("--stage7-concurrency", type=int, default=8, help="Stage 7 评分请求级并发数，默认8")
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.latest:
        print_latest_run()
        return
    if args.status is not None:
        print_run_status(resolve_run_dir(args.status or None))
        return
    if args.tail is not None:
        print_run_tail(resolve_run_dir(args.tail or None), limit=args.tail_lines)
        return
    if args.watch is not None:
        watch_run(resolve_run_dir(args.watch or None), interval=args.watch_interval, tail_limit=args.tail_lines)
        return

    stages = stage_range(args)
    run_dir, run_meta = prepare_run_dir(args)
    print(f"Run 目录：{run_dir}")
    print(f"执行阶段：{', '.join(str(s) for s in stages)}")

    finalize_run(run_dir, "running")
    emit_progress_event(
        run_dir,
        "run_started",
        status="running",
        message=f"评测启动，执行阶段：{', '.join(str(s) for s in stages)}",
        payload={"stages": stages},
    )
    completed_without_exception = False
    try:
        for stage_no in stages:
            run_stage(run_dir, stage_no, dry_run=args.dry_run, stage7_concurrency=args.stage7_concurrency)
        completed_without_exception = True
    except SystemExit as exc:
        status = "partial" if (run_dir / "report.md").exists() else "failed"
        finalize_run(run_dir, status)
        emit_progress_event(
            run_dir,
            "run_partial" if status == "partial" else "run_failed",
            status=status,
            message=f"运行结束，状态={status}：{exc}",
        )
        if status == "partial":
            print(f"\n运行结束但结果不完整：{exc}")
            print(f"Run 目录：{run_dir}")
            print(f"报告：{run_dir / 'report.md' if (run_dir / 'report.md').exists() else '未生成'}")
            print(f"表格：{run_dir / 'evaluation.xlsx' if (run_dir / 'evaluation.xlsx').exists() else '未生成'}")
            print(f"状态：partial")
            return
        raise
    except BaseException:
        finalize_run(run_dir, "failed")
        emit_progress_event(
            run_dir,
            "run_failed",
            status="failed",
            message="运行异常中断，状态=failed",
        )
        raise

    if completed_without_exception:
        final_status = "dry_run" if args.dry_run else "success"
        finalize_run(run_dir, final_status)
        emit_progress_event(
            run_dir,
            "run_completed" if final_status == "success" else "run_dry_run",
            status=final_status,
            message=f"运行完成，状态={final_status}",
        )

    print("\n运行完成")
    print(f"Run 目录：{run_dir}")
    print(f"报告：{run_dir / 'report.md' if (run_dir / 'report.md').exists() else '未生成'}")
    print(f"表格：{run_dir / 'evaluation.xlsx' if (run_dir / 'evaluation.xlsx').exists() else '未生成'}")
    if (run_dir / "logs").exists():
        print(f"日志：{run_dir / 'logs'}")
    if (run_dir / "progress.jsonl").exists():
        print(f"进度事件：{run_dir / 'progress.jsonl'}")
    print(f"索引：{INDEX_PATH}")


if __name__ == "__main__":
    main()
