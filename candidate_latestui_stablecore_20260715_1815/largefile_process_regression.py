import hashlib
import multiprocessing
import sys
import time
from pathlib import Path

from PyQt6.QtCore import QCoreApplication, QTimer

from main_gui import RollForwardWorker
from roll_forward_core import SubjectConfig, find_prior_file


CORE_SHA256 = "173C1EAA3D27F796BA2238E9BBA8436745EF4D6A9A92CDB05024E10CA4E1B60B"


def main():
    if len(sys.argv) != 4:
        raise SystemExit(
            "Usage: python largefile_process_regression.py SUBJECT PRIOR_DIR OUTPUT_DIR"
        )

    source_dir = Path(__file__).resolve().parent
    root_dir = source_dir.parent
    subject_code = sys.argv[1]
    prior_dir = Path(sys.argv[2]).resolve()
    output_dir = Path(sys.argv[3]).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    core_hash = hashlib.sha256((source_dir / "roll_forward_core.py").read_bytes()).hexdigest().upper()
    if core_hash != CORE_SHA256:
        raise AssertionError(f"Core source changed: {core_hash}")
    if not prior_dir.exists():
        raise FileNotFoundError(prior_dir)

    subject_config = SubjectConfig().get_subject(subject_code)
    matched_prior = find_prior_file(str(prior_dir), subject_code, "2025", subject_config)
    if not matched_prior:
        raise AssertionError(f"No prior workbook matched for {subject_code}: {prior_dir}")
    matched_prior = Path(matched_prior)

    app = QCoreApplication(sys.argv)
    ticks = 0
    results = []
    progress_messages = []
    started_at = time.monotonic()

    timer = QTimer()
    timer.setInterval(100)

    def count_tick():
        nonlocal ticks
        ticks += 1

    def capture_progress(current, total, message):
        progress_messages.append((current, total, message))

    def capture_results(value):
        results.extend(value)
        app.quit()

    timer.timeout.connect(count_tick)
    timer.start()

    worker = RollForwardWorker(
        subject_codes=[subject_code],
        template_dir=str(root_dir / "templates"),
        prior_dir=str(prior_dir),
        company_name="大文件响应回归测试",
        bs_date="2026/12/31",
        output_dir=str(output_dir),
        functional_currency="人民币",
        accounting_standard="企业会计准则",
        pm_value="1000000",
        te_value="750000",
        sad_value="50000",
        cra_records=[],
        roll_forward_wording=False,
        generate_summary=True,
        llm_enhanced=False,
        llm_wording_revision=False,
        llm_options={},
    )
    worker.progress_signal.connect(capture_progress)
    worker.finished_signal.connect(capture_results)
    worker.start()
    app.exec()
    worker.wait()

    elapsed = time.monotonic() - started_at
    if not results or not results[0][1]:
        raise AssertionError(
            f"Isolated roll failed: results={results}, progress={progress_messages}"
        )
    if ticks < 10:
        raise AssertionError(f"GUI event loop did not remain responsive: {ticks} ticks")
    if not results[0][3] or not Path(results[0][3]).is_file():
        raise AssertionError(f"Output file missing: {results[0][3] if results else None}")

    print(f"PASS subject={subject_code}")
    print(f"PASS source_mb={matched_prior.stat().st_size / (1024 * 1024):.2f}")
    print(f"PASS source={matched_prior}")
    print(f"PASS elapsed_seconds={elapsed:.2f}")
    print(f"PASS gui_ticks={ticks}")
    print(f"PASS progress_events={len(progress_messages)}")
    print(f"PASS output={results[0][3]}")
    print(f"PASS core_sha256={core_hash}")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
