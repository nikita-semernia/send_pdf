import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path
import requests
import logging
from logging.handlers import RotatingFileHandler
from winotify import Notification
import argparse


# === CONFIG ===
BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "configs" / "config.json"

def load_config() -> dict:
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(f"config.json not found: {CONFIG_PATH}")
    # utf-8-sig щоб не впасти, якщо BOM випадково з’явиться (як зі students.json) [web:342]
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))

def get_bot_token() -> str:
    cfg = load_config()
    token = str(cfg.get("bot_token", "")).strip()
    if not token:
        raise ValueError("bot_token missing in config.json")
    return token

# === LOGS ===

LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

logger = logging.getLogger("send_pdf")
logger.setLevel(logging.INFO)

_handler = RotatingFileHandler(
    LOG_DIR / "send_pdf.log",
    maxBytes=2 * 1024 * 1024,   # 2 MB
    backupCount=5,              # 5 старих файлів
    encoding="utf-8",
)
_formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
_handler.setFormatter(_formatter)
logger.addHandler(_handler)



def load_students_active_by_id(students_path: Path) -> dict:
    """
    Returns: { "u001": {"id":..., "name":..., "chat_id":..., ...}, ... } only for active students.
    """
    if not students_path.exists():
        raise FileNotFoundError(f"students.json not found: {students_path}")

    data = json.loads(students_path.read_text(encoding="utf-8-sig"))
    if not isinstance(data, list):
        raise ValueError("students.json must be a JSON array (list of objects)")

    out = {}
    for s in data:
        if not isinstance(s, dict):
            continue
        if s.get("active", True) is False:
            continue
        sid = str(s.get("id", "")).strip()
        if not sid:
            continue
        out[sid] = s
    return out


def get_chat_id(student_id: str, students_path: Path) -> str:
    students = load_students_active_by_id(students_path)
    if student_id not in students:
        raise KeyError(f"student id not found or inactive: {student_id}")
    chat_id = students[student_id].get("chat_id", None)
    if chat_id is None or str(chat_id).strip() == "":
        raise KeyError(f"chat_id missing for student id: {student_id}")
    return str(chat_id)



def tg_send_document(pdf_path: str, student_id: str, students_path: Path, caption: str) -> bool:
    chat_id = get_chat_id(student_id, students_path)

    token = get_bot_token()
    url = f"https://api.telegram.org/bot{token}/sendDocument"
    data = {"chat_id": chat_id}
    if caption and caption.strip():
        data["caption"] = caption

    with open(pdf_path, "rb") as f:
        resp = requests.post(
            url,
            data=data,
            files={"document": f},
            timeout=60,
        )

    # Telegram Bot API: JSON response includes boolean field 'ok' and possibly 'description' [web:9]
    try:
        payload = resp.json()
    except Exception:
        print("ERROR: non-JSON response:", resp.status_code, resp.text[:500])
        logger.error("non-JSON response: status=%s body=%s", resp.status_code, resp.text[:500])
        return False

    if payload.get("ok") is True:
        return True

    print("ERROR: telegram failed:", payload.get("description", payload))
    logger.error("telegram failed:", payload.get("description", payload))
    return False


def build_pdf_from_slides(slides_dir: str, output_pdf: str) -> bool:
    try:
        import img2pdf
    except ImportError:
        print("ERROR: img2pdf not installed. Run: python -m pip install img2pdf")
        logger.error("img2pdf not installed. Run: python -m pip install img2pdf")
        return False

    d = Path(slides_dir)
    png_files = sorted(d.glob("slide_*.png"))
    if not png_files:
        print("ERROR: no slide_*.png found in:", slides_dir)
        logger.error("no slide_*.png found in:", slides_dir)
        return False

    with open(output_pdf, "wb") as f:
        f.write(img2pdf.convert([str(p) for p in png_files]))

    return True


def cleanup_temp(slides_dir: str | None, pdf_path: str | None) -> None:
    # delete pdf
    if pdf_path and os.path.isfile(pdf_path):
        try:
            os.remove(pdf_path)
        except Exception as e:
            print("WARN: cannot delete pdf:", e)
            logger.exception("WARN: cannot delete pdf:", e)

    # delete slides directory recursively
    if slides_dir and os.path.isdir(slides_dir):
        try:
            shutil.rmtree(slides_dir)
        except Exception as e:
            print("WARN: cannot delete slides dir:", e)
            logger.exception("WARN: cannot delete slides dir:", e)


def main() -> int:
    logger.info("Started. argv=%r", sys.argv)
    logger.info("Script file: %s", __file__)

    import argparse

    parser = argparse.ArgumentParser(
        prog="send_pdf.py",
        description="Build PDF from slide PNGs (or send ready PDF) and send it via Telegram.",
    )
    parser.add_argument("input_path", help="Slides directory or ready PDF file path")
    parser.add_argument("student_id", help="Student id from students.json (e.g. u001)")
    parser.add_argument("--student-label", default="", help="Human-readable student name for UI/toast")
    parser.add_argument("--profile", default="", help="Quality profile string for UI/toast")
    parser.add_argument("--log-path", default="", help="Optional log file path (key=value lines)")
    parser.add_argument("--students-json", default="", help="Path to students.json (from VBA settings)")
    parser.add_argument("--caption", default="", help="Telegram caption (may contain newlines)")

    try:
        args = parser.parse_args()
    except SystemExit:
        # argparse already printed usage to stderr; we return a standard error code
        logger.info("Argparse exit (bad args).")
        return 2

    input_path = args.input_path
    student_id = args.student_id
    student_label = (args.student_label or "").strip() or student_id
    quality_profile = (args.profile or "").strip()
    caption = (args.caption or "").rstrip()
    log_path = (args.log_path or "").strip()
    students_path = Path(args.students_json).expanduser() if args.students_json else (BASE_DIR / "students.json")

    # Overwrite log file at start (so VBA reads only this run)
    if log_path:
        try:
            with open(log_path, "w", encoding="utf-8") as f:
                f.write("")  # truncate
        except Exception:
            logger.exception("Failed to truncate log-path: %s", log_path)

    def emit(line: str) -> None:
        # Keep console prints for debugging, but in pythonw they are harmless
        try:
            print(line)
        except Exception:
            pass

        if log_path:
            try:
                with open(log_path, "a", encoding="utf-8") as f:
                    f.write(line + "\n")
            except Exception:
                logger.exception("Failed to write log-path: %s", log_path)

    slides_dir = None
    pdf_path = None

    # mode A: directory with slide_*.png -> build PDF (named dd.mm.yyyy.pdf in parent dir)
    if os.path.isdir(input_path):
        slides_dir = input_path
        pdf_path = os.path.join(
            os.path.dirname(input_path),
            datetime.now().strftime("%d.%m.%Y") + ".pdf",
        )

        if not build_pdf_from_slides(slides_dir, pdf_path):
            emit("ERROR: failed to build PDF")
            logger.error("failed to build PDF")
            return 1
    else:
        # mode B: ready pdf path
        pdf_path = input_path
        if not os.path.isfile(pdf_path):
            emit(f"ERROR: PDF not found: {pdf_path}")
            logger.error("PDF not found: %s", pdf_path)
            return 1
    emit("CAPTION=" + caption.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\\n"))

    ok = tg_send_document(pdf_path, student_id, students_path, caption)
    if ok:
        size_bytes = os.path.getsize(pdf_path)
        emit(f"PDF_SIZE_BYTES={size_bytes}")
        emit(f"PDF_PATH={pdf_path}")

        cleanup_temp(slides_dir, pdf_path)

        emit("OK: sent and cleaned")
        logger.info("OK: sent and cleaned")

        # Toast on success only (never fail the whole send because of toast)
        try:
            msg = f"Надіслано: {student_label}"
            if quality_profile:
                msg += f" • {quality_profile}"
            msg += f" • {os.path.basename(pdf_path)}"

            toast = Notification(
                app_id="TutorSendPdf",
                title="Відправка PDF у Telegram",
                msg=msg,
                duration="short",
            )
            toast.show()
        except Exception:
            logger.exception("Toast failed")

        return 0

    emit("ERROR: send failed (no cleanup done)")
    logger.error("send failed (no cleanup done)")
    return 1

if __name__ == "__main__":
    raise SystemExit(main())
