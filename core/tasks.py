# -*- coding: utf-8 -*-
"""任务管理：TaskState / TaskManager / MANAGER 单例 / 状态持久化。"""

import os
import json
import time
import uuid
import threading
from dataclasses import dataclass, field
from typing import Optional

from core import config
from core.utils import _IO_LOCK, _task_disk_size

_STATE_SAVE_TS = {}
_STATE_SAVE_LOCK = threading.RLock()


# ============================================================
# 任务管理
# ============================================================

@dataclass
class TaskState:
    task_id: str
    kind: str
    src_path: str
    src_name: str
    out_dir: str
    work_dir: str
    target_lang: str = "zh-CN"
    created_at: float = field(default_factory=time.time)
    status: str = "queued"
    current: int = 0
    total: int = 0
    label: str = ""
    log: list = field(default_factory=list)
    output_files: list = field(default_factory=list)
    preview_images: list = field(default_factory=list)
    preview_html: str = ""
    error: str = ""
    estimated_size: int = 0
    current_size: int = 0
    last_index: int = 0
    pdf_quality: str = config.DEFAULT_PDF_QUALITY
    make_bilingual: bool = True
    # ★ 新增：运行参数存到 task
    domain: str = "general"
    extra_prompt: str = ""
    trial_pages: int = 5
    _preview_scanned: bool = False
    _last_size_ts: float = 0.0
    stop_event: threading.Event = field(default_factory=threading.Event)
    thread: Optional[threading.Thread] = None
    _lock: threading.RLock = field(
        default_factory=threading.RLock, repr=False, compare=False
    )

    def to_dict(self):
        with self._lock:
            log_snap = list(self.log[-60:])
            files_snap = list(self.output_files or [])
            return {
                "task_id": self.task_id, "kind": self.kind,
                "src_name": self.src_name, "out_dir": self.out_dir,
                "target_lang": self.target_lang, "created_at": self.created_at,
                "status": self.status if self.status != "stopping" else "paused",
                "current": self.current, "total": self.total,
                "label": self.label, "log": log_snap,
                "output_files": files_snap, "error": self.error,
                "estimated_size": self.estimated_size,
                "current_size": self.current_size,
                "last_index": self.last_index,
                "pdf_quality": self.pdf_quality,
                "make_bilingual": self.make_bilingual,
                "domain": self.domain,
                "extra_prompt": self.extra_prompt,
                "trial_pages": self.trial_pages,
            }

    def log_msg(self, m):
        with self._lock:
            # ★ 新增：日志加时间戳
            ts = time.strftime("%H:%M:%S")
            self.log.append(f"[{ts}] {m}")
            if len(self.log) > 200:
                del self.log[:-200]

    def log_tail(self, n=40):
        """★ 修复：锁内取日志尾部快照，供刷新线程安全读取。"""
        with self._lock:
            return "\n".join(self.log[-n:])


class TaskManager:
    def __init__(self):
        self._tasks = {}
        self._lock = threading.Lock()

    def create(self, kind, src_path, src_name, out_dir, work_dir,
               target_lang="zh-CN"):
        tid = uuid.uuid4().hex[:8]
        t = TaskState(task_id=tid, kind=kind, src_path=src_path,
                      src_name=src_name, out_dir=out_dir, work_dir=work_dir,
                      target_lang=target_lang)
        with self._lock:
            self._tasks[tid] = t
        return t

    def get(self, tid):
        with self._lock:
            return self._tasks.get(tid)

    def all_sorted(self):
        with self._lock:
            return sorted(list(self._tasks.values()),
                          key=lambda t: -t.created_at)

    def running(self):
        with self._lock:
            return [t for t in self._tasks.values()
                    if t.status in ("queued", "running", "stopping")]

    def find_active_by_out_dir(self, out_dir):
        target = os.path.normcase(os.path.abspath(out_dir))
        with self._lock:
            for t in self._tasks.values():
                if os.path.normcase(os.path.abspath(t.out_dir)) == target and \
                   t.status in ("queued", "running", "stopping"):
                    return t
        return None

    def load_from_disk(self, state_dict):
        tid = state_dict.get("task_id")
        if not tid:
            return
        with self._lock:
            if tid in self._tasks:
                return
        t = TaskState(
            task_id=tid, kind=state_dict.get("kind", "pdf"),
            src_path="", src_name=state_dict.get("src_name", ""),
            out_dir=state_dict.get("out_dir", ""),
            work_dir=os.path.join(state_dict.get("out_dir", ""), "_work"),
            target_lang=state_dict.get("target_lang", "zh-CN"),
            created_at=state_dict.get("created_at", time.time()),
            status="paused",
            current=state_dict.get("current", 0),
            total=state_dict.get("total", 0),
            label=state_dict.get("label", ""),
            log=state_dict.get("log", []),
            output_files=state_dict.get("output_files", []) or [],
            error=state_dict.get("error", ""),
            estimated_size=state_dict.get("estimated_size", 0),
            current_size=state_dict.get("current_size", 0),
            last_index=state_dict.get("last_index", 0),
            pdf_quality=state_dict.get("pdf_quality", config.DEFAULT_PDF_QUALITY),
            make_bilingual=state_dict.get("make_bilingual", True),
            domain=state_dict.get("domain", "general"),
            extra_prompt=state_dict.get("extra_prompt", ""),
            trial_pages=state_dict.get("trial_pages", 5),
        )
        with self._lock:
            self._tasks[tid] = t


MANAGER = TaskManager()


def save_state(task, force=False):
    if task is None:
        return
    now = time.monotonic()
    with _STATE_SAVE_LOCK:
        last = _STATE_SAVE_TS.get(task.task_id)
        if not force and last is not None and (now - last) < config.STATE_SAVE_INTERVAL:
            return
        _STATE_SAVE_TS[task.task_id] = now

    try:
        path = os.path.join(task.work_dir, "state.json")
        os.makedirs(task.work_dir, exist_ok=True)
        data = task.to_dict()
        with _IO_LOCK:
            tmp = path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            os.replace(tmp, path)
    except Exception as e:
        # ★ 修复：静默失败 → 至少可见
        print(f"[warn] save_state 保存失败：{e}")


def _add_output(task, *paths_):
    if task is None:
        return
    with task._lock:
        for p in paths_:
            if p and p not in task.output_files:
                task.output_files.append(p)


def scan_all_states():
    root = os.path.abspath(config.RESULT_ROOT)
    if not os.path.isdir(root):
        return
    for name in os.listdir(root):
        p = os.path.join(root, name, "_work", "state.json")
        if os.path.isfile(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    MANAGER.load_from_disk(json.load(f))
            except Exception:
                pass


def cleanup_orphan_files():
    root = os.path.abspath(config.RESULT_ROOT)
    if not os.path.isdir(root):
        return 0
    removed = 0
    for name in os.listdir(root):
        d = os.path.join(root, name)
        if not os.path.isdir(d):
            continue
        for f in os.listdir(d):
            fp = os.path.join(d, f)
            try:
                if time.time() - os.path.getmtime(fp) < 86400:
                    continue
            except Exception:
                continue
            if f.endswith(".tmp.pdf") or f.endswith(".bak.pdf"):
                try:
                    os.remove(fp)
                    removed += 1
                    print(f"   [clean] {name}/{f}")
                except Exception:
                    pass
        # ★ 修复：双语 PDF 检查点文件（崩溃遗留）
        work_d = os.path.join(d, "_work")
        if os.path.isdir(work_d):
            for f in os.listdir(work_d):
                if "_bilingual_ckpt." in f and f.endswith(".pdf"):
                    fp = os.path.join(work_d, f)
                    try:
                        if time.time() - os.path.getmtime(fp) < 86400:
                            continue
                    except Exception:
                        continue
                    try:
                        os.remove(fp)
                        removed += 1
                        print(f"   [clean] {name}/_work/{f}")
                    except Exception:
                        pass
    return removed


def refresh_task_size(task, force=False):
    if task is None:
        return
    now = time.time()
    if not force and (now - task._last_size_ts) < config.SIZE_REFRESH_INTERVAL:
        return
    task._last_size_ts = now
    try:
        task.current_size = _task_disk_size(task)
    except Exception as e:
        # ★ 修复：静默失败 → 至少可见
        print(f"[warn] refresh_task_size 失败：{e}")
