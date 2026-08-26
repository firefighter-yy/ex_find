"""Shared task lifecycle and user-facing failure classification.

The UI and services use this small module to make cancellation cooperative and
to guarantee that owned resources are closed only after worker tasks finish.
"""

from __future__ import annotations

import threading
from collections.abc import Callable
from concurrent.futures import Future, ThreadPoolExecutor, wait
from dataclasses import dataclass
from enum import Enum
from typing import Any


class TaskState(str, Enum):
    IDLE = "idle"
    RUNNING = "running"
    CANCELLING = "cancelling"
    SUCCEEDED = "succeeded"
    CANCELLED = "cancelled"
    FAILED = "failed"
    CLOSED = "closed"


class ErrorCategory(str, Enum):
    EXCEL_UNAVAILABLE = "excel_unavailable"
    SOURCE_UNAVAILABLE = "source_unavailable"
    SOURCE_CHANGED = "source_changed"
    ACCESS_DENIED = "access_denied"
    DISK_SPACE = "disk_space"
    CANCELLED = "cancelled"
    UNKNOWN = "unknown"


@dataclass(frozen=True, slots=True)
class ClassifiedError:
    category: ErrorCategory
    message: str
    recoverable: bool = True


def classify_error(error: BaseException) -> ClassifiedError:
    """Map technical exceptions to safe, actionable UI text."""
    name = type(error).__name__.casefold()
    text = str(error).casefold()
    if "cancel" in name or "cancel" in text:
        return ClassifiedError(ErrorCategory.CANCELLED, "操作已取消")
    if "excelunavailable" in name or "com" in text or "excel" in text and "start" in text:
        return ClassifiedError(ErrorCategory.EXCEL_UNAVAILABLE, "无法启动 Microsoft Excel，请检查 Office 安装和 COM 注册")
    if isinstance(error, PermissionError) or "permission" in text or "access denied" in text:
        return ClassifiedError(ErrorCategory.ACCESS_DENIED, "没有权限访问该文件或目录，请选择可写位置")
    if isinstance(error, FileNotFoundError) or "unavailable" in text or "not found" in text:
        return ClassifiedError(ErrorCategory.SOURCE_UNAVAILABLE, "源文件不可用，请确认文件存在且未被移除")
    if "changed" in text or "变化" in text or "stale" in text:
        return ClassifiedError(ErrorCategory.SOURCE_CHANGED, "源文件已变化，请重新准备文件")
    if isinstance(error, OSError) and getattr(error, "winerror", None) == 112:
        return ClassifiedError(ErrorCategory.DISK_SPACE, "磁盘空间不足，请清理空间后重试")
    return ClassifiedError(ErrorCategory.UNKNOWN, "操作失败，请检查文件和目标路径后重试")


class TaskCoordinator:
    """Own one worker executor and one cooperative cancellation token."""

    def __init__(self, *, thread_name_prefix: str = "excel-search", max_workers: int = 1) -> None:
        self._executor = ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix=thread_name_prefix)
        self._cancel_event = threading.Event()
        self._future: Future[Any] | None = None
        self._state = TaskState.IDLE
        self._lock = threading.RLock()

    @property
    def state(self) -> TaskState:
        with self._lock:
            return self._state

    @property
    def future(self) -> Future[Any] | None:
        with self._lock:
            return self._future

    @property
    def cancel_event(self) -> threading.Event:
        return self._cancel_event

    def submit(self, function: Callable[..., Any], *args: Any, **kwargs: Any) -> Future[Any]:
        with self._lock:
            if self._state in {TaskState.CLOSED, TaskState.RUNNING, TaskState.CANCELLING}:
                raise RuntimeError("a task is already active or coordinator is closed")
            self._cancel_event.clear()
            self._state = TaskState.RUNNING
            future = self._executor.submit(function, *args, **kwargs)
            self._future = future
            future.add_done_callback(self._finish)
            return future

    def cancel(self) -> bool:
        with self._lock:
            if self._future is None or self._future.done():
                return False
            self._state = TaskState.CANCELLING
            self._cancel_event.set()
            return self._future.cancel()

    def wait(self, timeout: float | None = None) -> bool:
        future = self.future
        if future is None:
            return True
        done, _ = wait((future,), timeout=timeout)
        return bool(done)

    def close(self, *, wait: bool = True, timeout: float | None = None) -> bool:
        self.cancel()
        completed = self.wait(timeout)
        if not completed and wait:
            self._executor.shutdown(wait=True, cancel_futures=True)
            completed = True
        else:
            self._executor.shutdown(wait=False, cancel_futures=True)
        with self._lock:
            self._state = TaskState.CLOSED
        return completed

    def _finish(self, future: Future[Any]) -> None:
        with self._lock:
            if self._state == TaskState.CLOSED:
                return
            if future.cancelled() or self._cancel_event.is_set():
                self._state = TaskState.CANCELLED
            elif future.exception() is not None:
                self._state = TaskState.FAILED
            else:
                self._state = TaskState.SUCCEEDED
