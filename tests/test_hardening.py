import threading
import time

from ex_transform.hardening import ErrorCategory, TaskCoordinator, TaskState, classify_error


def test_coordinator_cancels_and_waits_before_closing():
    started = threading.Event()
    released = threading.Event()

    def worker(cancel_event):
        started.set()
        while not cancel_event.is_set():
            time.sleep(0.001)
        released.set()
        return "cancelled"

    coordinator = TaskCoordinator()
    future = coordinator.submit(worker, coordinator.cancel_event)
    assert started.wait(1)

    assert coordinator.close(wait=True, timeout=1) is True
    assert released.is_set()
    assert future.result() == "cancelled"
    assert coordinator.state == TaskState.CLOSED


def test_coordinator_can_run_another_task_after_completion():
    coordinator = TaskCoordinator()
    assert coordinator.submit(lambda: 1).result() == 1
    assert coordinator.submit(lambda: 2).result() == 2
    coordinator.close()


def test_error_classification_is_actionable_and_safe():
    assert classify_error(PermissionError("secret.xlsx")).category == ErrorCategory.ACCESS_DENIED
    assert classify_error(FileNotFoundError("secret.xlsx")).category == ErrorCategory.SOURCE_UNAVAILABLE
    assert "secret.xlsx" not in classify_error(RuntimeError("secret.xlsx broke")).message
