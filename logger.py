# -*- coding: utf-8 -*-
"""
共享日志模块 — 全局流水线日志系统
====================================
支持：
  - 每次运行按时间戳创建独立日志目录
  - 阶段级上下文管理
  - API 调用日志（自动截断 prompt/response）
  - 文件流转追踪
  - metadata.json 进度快照
  - 多线程安全

用法：
    from logger import get_logger, PipelineLogger

    logger = get_logger(base_dir=__file__)

    with logger.stage("clean", desc="有效性判定与分类") as stage:
        stage.info("开始处理 image_ocr/ 目录")
        stage.log_api_call("ernie", prompt[:500], response[:500], elapsed=1.2)
        stage.log_file_moved(src="image_ocr/1-1.png", dst="image_ocr_clean/小学/数学/1.png")
        stage.set_progress(current=42, total=100)
"""

import os
import sys
import json
import time
import threading
import traceback
from datetime import datetime

# ============================================
# 全局配置：来自 config.yaml，见 config.py
# logger 被所有阶段导入，因此这里对配置缺失做兜底，不抛异常。
# ============================================
try:
    from config import cfg as _cfg

    LOG_DIR_NAME = _cfg("paths.log_dir", "logs")
    MAX_MESSAGE_LENGTH = int(_cfg("logging.max_message_length", 2000))
    MAX_API_LOG_LENGTH = int(_cfg("logging.max_api_log_length", 800))
except Exception:  # 配置不可用时退回默认值，保证日志始终可写
    LOG_DIR_NAME = "logs"
    MAX_MESSAGE_LENGTH = 2000   # 单条日志截断长度
    MAX_API_LOG_LENGTH = 800    # API prompt/response 截断长度


# ============================================
# 单例 Logger
# ============================================
_logger_instance = None
_logger_lock = threading.Lock()


def get_logger(base_dir=None, run_id=None):
    """获取全局唯一的 PipelineLogger 实例。优先使用 AUTO_EVAL_WORK_DIR 作为运行目录。"""
    global _logger_instance
    work_dir = os.environ.get("AUTO_EVAL_WORK_DIR")
    if work_dir:
        base_dir = work_dir
    if _logger_instance is None:
        with _logger_lock:
            if _logger_instance is None:
                _logger_instance = PipelineLogger(base_dir=base_dir, run_id=run_id)
    return _logger_instance


# ============================================
# PipelineLogger
# ============================================
class PipelineLogger:
    def __init__(self, base_dir=None, run_id=None):
        if base_dir is None:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        elif os.path.isfile(base_dir):
            base_dir = os.path.dirname(os.path.abspath(base_dir))

        self.base_dir = base_dir
        self.run_id = run_id or datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_dir = os.path.join(base_dir, LOG_DIR_NAME, f"run_{self.run_id}")
        self.stages_dir = os.path.join(self.log_dir, "stages")

        os.makedirs(self.stages_dir, exist_ok=True)

        self.log_file = os.path.join(self.log_dir, "pipeline.log")
        self.meta_file = os.path.join(self.log_dir, "metadata.json")

        self._file_lock = threading.Lock()
        self._metadata = self._load_metadata()

        self._write_line("=" * 70)
        self._write_line(f"流水线启动 | run_id={self.run_id} | 目录={self.base_dir}")
        self._write_line("=" * 70)

    # -------------------- 文件写入 --------------------
    def _write_line(self, msg):
        msg = str(msg)[:MAX_MESSAGE_LENGTH]
        with self._file_lock:
            with open(self.log_file, "a", encoding="utf-8") as f:
                f.write(msg + "\n")
            print(msg, flush=True)

    def _write_metadata(self):
        with self._file_lock:
            self._metadata["updated_at"] = datetime.now().isoformat()
            with open(self.meta_file, "w", encoding="utf-8") as f:
                json.dump(self._metadata, f, ensure_ascii=False, indent=2)

    def _load_metadata(self):
        if os.path.exists(self.meta_file):
            try:
                with open(self.meta_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {
            "run_id": self.run_id,
            "base_dir": self.base_dir,
            "started_at": datetime.now().isoformat(),
            "stages": {},
            "summary": {"total": 0, "success": 0, "failed": 0, "skipped": 0},
        }

    # -------------------- 阶段管理 --------------------
    def stage(self, name, desc="", stage_order=None):
        """返回一个 StageContext 上下文管理器"""
        return StageContext(self, name, desc, stage_order)

    def _stage_start(self, name, desc, stage_order):
        order_str = f"[{stage_order}] " if stage_order else ""
        self._write_line(f"\n{'─' * 50}")
        self._write_line(f"▶ {order_str}阶段开始: {name} — {desc}")
        self._write_line(f"{'─' * 50}")

        stage_key = f"{stage_order}_{name}" if stage_order else name
        stage_data = {
            "name": name,
            "desc": desc,
            "order": stage_order,
            "status": "running",
            "started_at": datetime.now().isoformat(),
            "progress": {"current": 0, "total": 0},
            "files_in": [],
            "files_out": [],
            "api_calls": [],
            "errors": [],
            "rows": 0,
        }
        self._metadata["stages"][stage_key] = stage_data
        self._write_metadata()
        return stage_data

    def _stage_end(self, name, stage_order, status="success", summary=""):
        stage_key = f"{stage_order}_{name}" if stage_order else name
        stage_data = self._metadata["stages"].get(stage_key, {})
        stage_data["status"] = status
        stage_data["ended_at"] = datetime.now().isoformat()

        started = stage_data.get("started_at", "")
        duration = ""
        if started:
            try:
                dt_start = datetime.fromisoformat(started)
                duration = f"{(datetime.now() - dt_start).total_seconds():.1f}s"
            except Exception:
                pass

        self._metadata["summary"]["total"] += 1
        if status == "success":
            self._metadata["summary"]["success"] += 1
        elif status == "failed":
            self._metadata["summary"]["failed"] += 1
        else:
            self._metadata["summary"]["skipped"] += 1

        icon = {"success": "✔", "failed": "✖", "skipped": "⊘"}.get(status, "?")
        self._write_line(f"{'─' * 50}")
        self._write_line(f"{icon} 阶段结束: {name} | 状态={status} | 耗时={duration}")
        if summary:
            self._write_line(f"   摘要: {summary}")
        self._write_line(f"{'─' * 50}\n")

        # 保存阶段详情 JSON
        with self._file_lock:
            stage_json_path = os.path.join(self.stages_dir, f"{stage_key}.json")
            with open(stage_json_path, "w", encoding="utf-8") as f:
                json.dump(stage_data, f, ensure_ascii=False, indent=2)

        self._metadata["stages"][stage_key] = stage_data
        self._write_metadata()

    # -------------------- 完成 --------------------
    def finish(self):
        summary = self._metadata["summary"]
        self._write_line("=" * 70)
        self._write_line("流水线结束")
        self._write_line(f"  阶段总计: {summary['total']}")
        self._write_line(f"  成功: {summary['success']}  失败: {summary['failed']}  跳过: {summary['skipped']}")
        self._write_line(f"  日志目录: {self.log_dir}")
        self._write_line("=" * 70)
        self._metadata["ended_at"] = datetime.now().isoformat()
        self._write_metadata()


# ============================================
# StageContext
# ============================================
class StageContext:
    def __init__(self, logger, name, desc="", stage_order=None):
        self.logger = logger
        self.name = name
        self.desc = desc
        self.stage_order = stage_order
        self._started = False

    def __enter__(self):
        self._stage_data = self.logger._stage_start(self.name, self.desc, self.stage_order)
        self._started = True
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if not self._started:
            return
        if exc_type is not None:
            tb = "".join(traceback.format_exception(exc_type, exc_val, exc_tb))
            self.error(f"未捕获异常: {tb}")
            self.logger._stage_end(self.name, self.stage_order, status="failed",
                                   summary=f"异常: {exc_val}")
        else:
            self.logger._stage_end(self.name, self.stage_order, status="success",
                                   summary=self._summary())
        self._started = False
        return False

    def _summary(self):
        stage_key = f"{self.stage_order}_{self.name}" if self.stage_order else self.name
        data = self.logger._metadata["stages"].get(stage_key, {})
        progress = data.get("progress", {})
        rows = data.get("rows", 0)
        api_count = len(data.get("api_calls", []))
        parts = []
        if rows:
            parts.append(f"处理 {rows} 行")
        if api_count:
            parts.append(f"API 调用 {api_count} 次")
        if progress.get("total"):
            parts.append(f"进度 {progress['current']}/{progress['total']}")
        return "，".join(parts) if parts else "完成"

    # -------------------- 日志方法 --------------------
    def _log(self, level, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.logger._write_line(f"  [{ts}] [{level}] {msg}")

    def info(self, msg):
        self._log("INFO", msg)

    def warn(self, msg):
        self._log("WARN", msg)

    def error(self, msg):
        self._log("ERROR", msg)
        self._update_stage_errors(msg)

    # -------------------- 结构化日志方法 --------------------
    def log_api_call(self, model, prompt, response, elapsed, success=True):
        """记录一次 API 调用"""
        record = {
            "time": datetime.now().isoformat(),
            "model": model,
            "prompt_preview": str(prompt)[:MAX_API_LOG_LENGTH],
            "response_preview": str(response)[:MAX_API_LOG_LENGTH],
            "elapsed": elapsed,
            "success": success,
        }
        self._update_stage("api_calls", record)

        icon = "✓" if success else "✗"
        self.info(f"{icon} API 调用: model={model}, 耗时={elapsed}s, "
                  f"prompt_len={len(str(prompt))}, response_len={len(str(response))}")

    def log_file_moved(self, src, dst):
        """记录文件流转"""
        record = {"src": src, "dst": dst}
        self._update_stage("files_moved", record)
        self.info(f"文件流转: {os.path.basename(src)} → {dst}")

    def log_file_output(self, path):
        """记录阶段输出文件"""
        self._update_stage("files_out", path)
        self.info(f"输出文件: {path}")

    def log_file_input(self, path):
        """记录阶段输入文件"""
        self._update_stage("files_in", path)
        self.info(f"输入文件: {path}")

    def set_progress(self, current, total):
        """更新进度"""
        self._update_stage("progress", {"current": current, "total": total})
        if total > 0:
            pct = current * 100 // total if total else 0
            if current % max(1, total // 20) == 0 or current == total:
                self.info(f"进度: {current}/{total} ({pct}%)")

    def set_row_count(self, count):
        """设置处理行数"""
        self._update_stage("rows", count)

    def log_item_result(self, item_id, status, detail=""):
        """记录单项处理结果（行级别）"""
        record = {"item_id": item_id, "status": status, "detail": str(detail)[:200]}
        self._update_stage("item_results", record)

    # -------------------- 内部辅助 --------------------
    def _update_stage(self, key, value):
        stage_key = f"{self.stage_order}_{self.name}" if self.stage_order else self.name
        data = self.logger._metadata["stages"].get(stage_key, {})
        if isinstance(data.get(key), list):
            data[key].append(value)
        elif isinstance(data.get(key), dict):
            data[key].update(value)
        else:
            data[key] = value
        self.logger._metadata["stages"][stage_key] = data

    def _update_stage_errors(self, msg):
        stage_key = f"{self.stage_order}_{self.name}" if self.stage_order else self.name
        data = self.logger._metadata["stages"].get(stage_key, {})
        data["errors"].append(msg)
        self.logger._metadata["stages"][stage_key] = data


# ============================================
# 快捷函数（无阶段上下文时使用）
# ============================================
def info(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] [INFO] {msg}", flush=True)

def warn(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] [WARN] {msg}", flush=True)

def error(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] [ERROR] {msg}", flush=True)