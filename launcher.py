from __future__ import annotations

import sys
from collections.abc import Sequence


def launch() -> None:
    # Keep GUI imports out of multiprocessing spawn bootstrap.  PyInstaller's
    # Office COM workers must be able to initialize without creating a Tk root.
    from docuforge.app import launch as launch_app

    launch_app()


def _verify_tk_runtime() -> None:
    import tkinter as tk
    from tkinterdnd2 import DND_FILES, TkinterDnD

    root = TkinterDnD.Tk()
    try:
        root.withdraw()
        drop_probe = tk.Label(root, text="LayoutLoom drag-and-drop self-test")
        drop_probe.pack()
        drop_probe.drop_target_register(DND_FILES)
        root.update_idletasks()
    finally:
        root.destroy()


def _multiprocessing_probe(connection: object) -> None:
    try:
        connection.send("layoutloom-multiprocessing-ok")  # type: ignore[attr-defined]
    finally:
        connection.close()  # type: ignore[attr-defined]


def _verify_multiprocessing_runtime() -> None:
    import multiprocessing

    context = multiprocessing.get_context("spawn")
    parent_connection, child_connection = context.Pipe(duplex=False)
    process = context.Process(
        target=_multiprocessing_probe,
        args=(child_connection,),
        name="layoutloom-self-test-worker",
        daemon=False,
    )
    process.start()
    child_connection.close()
    try:
        if not parent_connection.poll(15):
            raise RuntimeError("LayoutLoom multiprocessing self-test timed out")
        if parent_connection.recv() != "layoutloom-multiprocessing-ok":
            raise RuntimeError(
                "LayoutLoom multiprocessing self-test returned invalid data"
            )
    finally:
        parent_connection.close()
        process.join(3)
        if process.is_alive():
            process.terminate()
            process.join(3)
    if process.exitcode != 0:
        raise RuntimeError(
            f"LayoutLoom multiprocessing self-test failed: {process.exitcode}"
        )


def self_test() -> int:
    import wordninja

    from docuforge.registry import CORE_OPERATION_IDS, get_operations

    operations = get_operations()
    operation_ids = [operation.id for operation in operations]
    if len(operations) != len(CORE_OPERATION_IDS) or len(operation_ids) != len(
        set(operation_ids)
    ):
        raise RuntimeError("LayoutLoom operation catalog self-test failed")
    if wordninja.split("Forthispart") != ["For", "this", "part"]:
        raise RuntimeError("LayoutLoom English boundary model self-test failed")
    if getattr(sys, "frozen", False):
        from docuforge.engines import poppler_bin_path
        from docuforge.processors.video import detect_video_engine

        if not poppler_bin_path():
            raise RuntimeError("Bundled Poppler self-test failed")
        video_engine = detect_video_engine()
        if not video_engine.available or video_engine.ffprobe_executable is None:
            raise RuntimeError(f"Bundled FFmpeg self-test failed: {video_engine.reason}")
        _verify_tk_runtime()
        _verify_multiprocessing_runtime()
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    if "--self-test" in arguments:
        return self_test()
    launch()
    return 0


if __name__ == "__main__":
    import multiprocessing

    multiprocessing.freeze_support()
    raise SystemExit(main())
