import asyncio
import os
import shutil
import uuid
import time
import threading
import wave
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from fastapi import FastAPI, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import subprocess
import yt_dlp
from google import genai
from google.genai import types
from dotenv import load_dotenv

try:
    import pyaudiowpatch as pyaudio
    PYAUDIO_AVAILABLE = True
except ImportError:
    PYAUDIO_AVAILABLE = False
    print("⚠️  pyaudiowpatch לא מותקן — הקלטת זום לא תהיה זמינה. הרץ: pip install pyaudiowpatch")

try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False
    print("⚠️  psutil לא מותקן — הקלטת זום לא תהיה זמינה. הרץ: pip install psutil")

try:
    import win32gui
    import win32ui
    import win32con
    import numpy as np
    import cv2
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("⚠️  pywin32/opencv לא מותקנו — הקלטת וידאו לא תהיה זמינה. הרץ: pip install pywin32 opencv-python")

import ctypes
from ctypes.wintypes import RECT as WINRECT
# DPI awareness — חובה לפני כל פעולה עם חלונות, אחרת GetWindowRect מחזיר גדלים לוגיים מוקטנים
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

load_dotenv(Path(__file__).parent / ".env")

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])


def _cleanup_temp_dir():
    """מוחק את כל הקבצים הזמניים שנשארו מהפעלות קודמות."""
    if TEMP_DIR.exists():
        for item in TEMP_DIR.iterdir():
            try:
                if item.is_dir():
                    shutil.rmtree(item)
                else:
                    item.unlink()
            except Exception as e:
                print(f"⚠️  לא ניתן למחוק {item}: {e}")
        print(f"🗑️  תיקיית temp נוקתה")

OUTPUT_DIR = Path.home() / "Desktop" / "תמלולים"
TEMP_DIR = Path(__file__).parent / "temp"
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
SERVER_PORT = 7832
DOWNLOAD_TIMEOUT = 600

executor = ThreadPoolExecutor(max_workers=6)
jobs = {}
cancelled = set()
_recording_stop_events = {}  # job_id -> threading.Event
_running_subprocesses = {}   # job_id -> subprocess.Popen

HEBREW_NUMS = [
    "אחד", "שתיים", "שלוש", "ארבע", "חמש", "שש", "שבע", "שמונה", "תשע", "עשר",
    "אחד עשר", "שתים עשר", "שלוש עשר", "ארבע עשר", "חמש עשר",
    "שש עשר", "שבע עשר", "שמונה עשר", "תשע עשר", "עשרים",
    "עשרים ואחד", "עשרים ושתיים", "עשרים ושלוש", "עשרים וארבע", "עשרים וחמש",
    "עשרים ושש", "עשרים ושבע", "עשרים ושמונה", "עשרים ותשע", "שלושים",
]


class JobCancelled(Exception):
    pass


class TranscribeRequest(BaseModel):
    url: str
    title: str = ""
    direct_url: str = ""


class DownloadRequest(BaseModel):
    url: str
    title: str = ""
    format: str = "mp3"


class RecordRequest(BaseModel):
    title: str = ""
    keep_audio: bool = False
    save_video: bool = False


def update_job(job_id, status, message, progress=None, **kwargs):
    if job_id not in jobs:
        jobs[job_id] = {}
    jobs[job_id].update({"status": status, "message": message, **kwargs})
    if progress is not None:
        jobs[job_id]["progress"] = progress
    print(f"[{job_id}] {status} ({jobs[job_id].get('progress', 0)}%): {message}")


def _is_garbage_response(text: str) -> bool:
    """מזהה תגובות זבל מ-Gemini (הזיות של מספרים/תבניות חוזרות)."""
    if not text or len(text) < 10:
        return False
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    if len(lines) < 5:
        return False
    # בדיקה: אם רוב השורות הן תבנית חוזרת (כמו "1 2 3 4 5...")
    sample = lines[:20]
    unique = set(sample)
    if len(unique) <= 3 and len(sample) >= 10:
        return True
    # בדיקה: אם השורות מכילות בעיקר ספרות
    digit_lines = sum(1 for l in sample if sum(c.isdigit() or c == ' ' for c in l) / max(len(l), 1) > 0.7)
    if digit_lines / len(sample) > 0.6:
        return True
    return False


def _check_cancelled(job_id: str):
    if job_id in cancelled:
        raise JobCancelled("בוטל על ידי המשתמש")


def _run_subprocess(job_id: str, cmd: list, **kwargs) -> subprocess.CompletedProcess:
    """מריץ subprocess ומאפשר הריגה מיידית אם ה-job בוטל."""
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, **kwargs)
    _running_subprocesses[job_id] = proc
    try:
        stdout, stderr = proc.communicate()
        _running_subprocesses.pop(job_id, None)
        if job_id in cancelled:
            raise JobCancelled("בוטל על ידי המשתמש")
        if proc.returncode != 0:
            raise subprocess.CalledProcessError(proc.returncode, cmd, stdout, stderr)
        return subprocess.CompletedProcess(cmd, proc.returncode, stdout, stderr)
    except Exception:
        _running_subprocesses.pop(job_id, None)
        raise


def _get_lesson_path(course_folder: Path) -> Path:
    course_folder.mkdir(parents=True, exist_ok=True)
    existing = len(list(course_folder.glob("*.txt")))
    if existing < len(HEBREW_NUMS):
        name = f"שיעור {HEBREW_NUMS[existing]}"
    else:
        name = f"שיעור {existing + 1}"
    return course_folder / f"{name}.txt"


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/restart")
def restart_server():
    import sys
    def _do_restart():
        time.sleep(0.5)
        subprocess.Popen(
            [sys.executable, __file__],
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        os._exit(0)
    threading.Thread(target=_do_restart, daemon=True).start()
    return {"ok": True}


@app.post("/transcribe")
async def transcribe(request: TranscribeRequest, background_tasks: BackgroundTasks):
    job_id = str(uuid.uuid4())[:8]
    update_job(job_id, "pending", "ממתין...", progress=0)
    background_tasks.add_task(process_transcription, job_id, request.url, request.title, request.direct_url)
    return {"job_id": job_id}


@app.post("/download")
async def download(request: DownloadRequest, background_tasks: BackgroundTasks):
    job_id = str(uuid.uuid4())[:8]
    update_job(job_id, "pending", "ממתין...", progress=0)
    background_tasks.add_task(process_download, job_id, request.url, request.title, request.format)
    return {"job_id": job_id}


@app.post("/record/start")
async def record_start(request: RecordRequest, background_tasks: BackgroundTasks):
    if not PYAUDIO_AVAILABLE:
        return {"error": "pyaudiowpatch לא מותקן. הרץ: pip install pyaudiowpatch"}
    if not PSUTIL_AVAILABLE:
        return {"error": "psutil לא מותקן. הרץ: pip install psutil"}
    job_id = str(uuid.uuid4())[:8]
    update_job(job_id, "recording", "ממתין לזום...", progress=5, type="recording")
    if request.save_video and not WIN32_AVAILABLE:
        return {"error": "pywin32/opencv לא מותקנו. הרץ: pip install pywin32 opencv-python"}
    background_tasks.add_task(process_recording, job_id, request.title, request.keep_audio, request.save_video)
    return {"job_id": job_id}


@app.get("/status/{job_id}")
def get_status(job_id: str):
    return jobs.get(job_id, {"status": "not_found", "message": "לא נמצא", "progress": 0})


@app.post("/cancel/{job_id}")
def cancel_job(job_id: str):
    cancelled.add(job_id)
    if job_id in _recording_stop_events:
        _recording_stop_events[job_id].set()
    for key in [job_id, f"{job_id}_video"]:
        if key in _running_subprocesses:
            try:
                _running_subprocesses[key].kill()
            except Exception:
                pass
    if job_id in jobs:
        update_job(job_id, "cancelling", "מבטל...", progress=jobs[job_id].get("progress", 0))
    return {"ok": True}


def _make_postprocessor_hook(job_id):
    def hook(d):
        if d.get("status") == "started":
            pp = d.get("postprocessor", "")
            if "FFmpeg" in pp or "Audio" in pp:
                update_job(job_id, "converting", "ממיר לMP3 (ffmpeg)...", progress=67)
        elif d.get("status") == "finished":
            update_job(job_id, "converting", "המרה הושלמה", progress=70)
    return hook


def _make_progress_hook(job_id):
    def hook(d):
        if job_id in cancelled:
            raise yt_dlp.utils.DownloadError("cancelled")
        if d["status"] == "downloading":
            dl_pct = None
            pct_str = d.get("_percent_str", "").strip().replace("%", "").replace("N/A", "")
            try:
                dl_pct = float(pct_str)
            except Exception:
                pass
            if dl_pct is None:
                downloaded = d.get("downloaded_bytes", 0) or 0
                total = d.get("total_bytes") or d.get("total_bytes_estimate") or 0
                if total > 0:
                    dl_pct = (downloaded / total) * 100
            if dl_pct is not None:
                overall = int(5 + dl_pct * 0.60)
                jobs[job_id]["progress"] = min(overall, 65)
                speed = d.get("_speed_str", "").strip()
                downloaded_mb = (d.get("downloaded_bytes", 0) or 0) / 1024 / 1024
                jobs[job_id]["message"] = f"מוריד... {int(dl_pct)}% ({downloaded_mb:.0f}MB) {speed}"
    return hook


def _yt_dlp_download(url: str, job_id: str, audio_path: Path, fmt: str = "audio") -> str:
    postprocessors = []
    if fmt == "audio":
        postprocessors = [{"key": "FFmpegExtractAudio", "preferredcodec": "mp3", "preferredquality": "128"}]
        outtmpl = str(TEMP_DIR / f"{job_id}.%(ext)s")
    else:
        outtmpl = str(TEMP_DIR / f"{job_id}.%(ext)s")

    base_opts = {
        "outtmpl": outtmpl,
        "format": "bestaudio/best" if fmt == "audio" else "bestvideo+bestaudio/best",
        "quiet": True,
        "no_warnings": True,
        "socket_timeout": 30,
        "progress_hooks": [_make_progress_hook(job_id)],
        "postprocessor_hooks": [_make_postprocessor_hook(job_id)],
        "postprocessors": postprocessors,
    }

    try:
        opts = {**base_opts, "cookiesfrombrowser": ("chrome",)}
        with yt_dlp.YoutubeDL(opts) as ydl:
            info = ydl.extract_info(url, download=True)
            print(f"[yt-dlp] הורד: '{info.get('title')}' מ: {info.get('webpage_url') or info.get('url','')[:80]}")
            return info.get("title", job_id)
    except Exception as e1:
        print(f"yt-dlp+cookies failed: {e1}")

    try:
        with yt_dlp.YoutubeDL(base_opts) as ydl:
            info = ydl.extract_info(url, download=True)
            return info.get("title", job_id)
    except Exception as e2:
        raise Exception(f"yt-dlp נכשל: {e2}")


def _direct_download(url: str, dest: Path, job_id: str):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as resp, open(dest, "wb") as f:
        total = int(resp.headers.get("Content-Length", 0))
        downloaded = 0
        while chunk := resp.read(65536):
            if job_id in cancelled:
                raise JobCancelled("בוטל")
            f.write(chunk)
            downloaded += len(chunk)
            if total:
                jobs[job_id]["progress"] = int(5 + (downloaded / total) * 60)
                jobs[job_id]["message"] = f"מוריד... {int(downloaded/total*100)}%"


def _safe_filename(title: str, fallback: str) -> str:
    safe = "".join(c for c in (title or fallback) if c not in r'\/:*?"<>|').strip()
    return safe or fallback


async def _upload_to_gemini(client, chunk_path: Path, display_name: str, job_id: str = None):
    with open(chunk_path, "rb") as f:
        uploaded = client.files.upload(
            file=f,
            config=types.UploadFileConfig(mime_type="audio/mpeg", display_name=display_name),
        )
    for _ in range(60):
        if uploaded.state.name != "PROCESSING":
            break
        if job_id and job_id in cancelled:
            break
        await asyncio.sleep(2)
        uploaded = client.files.get(name=uploaded.name)
    return uploaded


async def _transcribe_and_save(job_id: str, audio_path: Path, url: str, title: str, loop):
    """Split MP3 into chunks, transcribe with Gemini, save txt. Returns (output_file, course_name)."""
    chunk_dir = TEMP_DIR / f"{job_id}_chunks"
    try:
        _check_cancelled(job_id)
        update_job(job_id, "transcribing", "מכין חלקים לתמלול...", progress=73)
        chunk_dir.mkdir(exist_ok=True)

        split_cmd = [
            "ffmpeg", "-i", str(audio_path),
            "-f", "segment", "-segment_time", "1200",
            "-c", "copy", "-y",
            str(chunk_dir / "chunk_%03d.mp3")
        ]
        await loop.run_in_executor(executor, lambda: _run_subprocess(job_id, split_cmd))
        chunk_files = sorted(chunk_dir.glob("chunk_*.mp3"))
        total_chunks = len(chunk_files)
        print(f"[{job_id}] {total_chunks} chunks")

        if not GEMINI_API_KEY:
            raise Exception("מפתח Gemini API לא הוגדר.")
        client = genai.Client(api_key=GEMINI_API_KEY)

        BATCH_SIZE = 8
        BATCH_DELAY = 60  # שניות בין פעימות

        # העלאת החלקים בפעימות
        update_job(job_id, "transcribing",
                   f"מעלה {total_chunks} חלקים בפעימות של {BATCH_SIZE}...",
                   progress=75, chunk_total=total_chunks)
        uploaded_files = []
        for batch_start in range(0, total_chunks, BATCH_SIZE):
            batch = chunk_files[batch_start:batch_start + BATCH_SIZE]
            batch_tasks = [
                _upload_to_gemini(client, chunk_path, f"{job_id}_{batch_start + i}", job_id)
                for i, chunk_path in enumerate(batch)
            ]
            uploaded_files.extend(await asyncio.gather(*batch_tasks))
            if batch_start + BATCH_SIZE < total_chunks:
                await asyncio.sleep(BATCH_DELAY)

        _check_cancelled(job_id)
        for i, uploaded in enumerate(uploaded_files):
            if uploaded.state.name != "ACTIVE":
                raise Exception(f"Gemini לא עיבד חלק {i+1}")

        def _shift_timestamps(text: str, offset_sec: int) -> str:
            """מוסיף offset לכל timestamp בפורמט [M:SS] או [H:MM:SS] בטקסט."""
            import re
            def replace(m):
                parts = m.group(1).split(":")
                if len(parts) == 2:
                    t = int(parts[0]) * 60 + int(parts[1])
                elif len(parts) == 3:
                    t = int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
                else:
                    return m.group(0)
                t += offset_sec
                h, rem = divmod(t, 3600)
                mn, s = divmod(rem, 60)
                if h > 0:
                    return f"[{h}:{mn:02d}:{s:02d}]"
                return f"[{mn}:{s:02d}]"
            return re.sub(r'\[(\d+:\d{2}(?::\d{2})?)\]', replace, text)

        async def _transcribe_chunk(i, uploaded):
            _check_cancelled(job_id)
            offset_sec = i * 1200
            chunk_num = i + 1
            prompt = (
                f"תמלל את קובץ האודיו בדייקנות מרבית. "
                f"הכנס חותמת זמן בפורמט [M:SS] כל ~10 שניות, החל מ-[0:00]. "
                f"לדוגמה: [0:00] טקסט... [0:10] טקסט... [0:20] טקסט... "
                f"החזר אך ורק את הטקסט המדויק עם חותמות הזמן, ללא כותרות, הסברים, או מידע נוסף. "
                f"שמור על פיסוק תקין. אם האודיו שקט לחלוטין או לא מובן, החזר מחרוזת ריקה בלבד ללא שום טקסט אחר."
            )
            response = await loop.run_in_executor(
                executor,
                lambda up=uploaded, p=prompt: client.models.generate_content(
                    model="models/gemini-2.5-flash",
                    contents=[
                        types.Part.from_uri(file_uri=up.uri, mime_type="audio/mpeg"),
                        p,
                    ],
                )
            )
            text = (response.text or "").strip()
            if _is_garbage_response(text):
                print(f"[{job_id}] chunk {i}: תגובת זבל מ-Gemini — מדלג")
                return ""
            return _shift_timestamps(text, offset_sec)

        # תמלול החלקים בפעימות
        update_job(job_id, "transcribing",
                   f"מתמלל {total_chunks} חלקים בפעימות של {BATCH_SIZE}...",
                   progress=82, chunk_total=total_chunks)
        transcription_parts = []
        for batch_start in range(0, total_chunks, BATCH_SIZE):
            batch = list(enumerate(uploaded_files))[batch_start:batch_start + BATCH_SIZE]
            batch_tasks = [_transcribe_chunk(i, uploaded) for i, uploaded in batch]
            transcription_parts.extend(await asyncio.gather(*batch_tasks))
            if batch_start + BATCH_SIZE < total_chunks:
                await asyncio.sleep(BATCH_DELAY)

        # ניקוי קבצים
        for chunk_path, uploaded in zip(chunk_files, uploaded_files):
            try:
                client.files.delete(name=uploaded.name)
            except Exception:
                pass
            chunk_path.unlink(missing_ok=True)

        try:
            chunk_dir.rmdir()
        except Exception:
            pass

        _check_cancelled(job_id)
        transcription = "\n".join(p for p in transcription_parts if p)

        update_job(job_id, "saving", "שומר תמלול...", progress=95)
        course_name = _safe_filename(title, job_id)
        course_folder = OUTPUT_DIR / course_name
        output_file = _get_lesson_path(course_folder)
        output_file.write_text(
            f"כתובת: {url}\nכותרת: {title}\nתאריך: {time.strftime('%Y-%m-%d %H:%M')}\n{'='*60}\n\n{transcription}",
            encoding="utf-8",
        )
        return output_file, course_name

    except Exception:
        try:
            if chunk_dir.exists():
                for f in chunk_dir.glob("*"):
                    f.unlink(missing_ok=True)
                chunk_dir.rmdir()
        except Exception:
            pass
        raise


# ── Recording helpers ─────────────────────────────────────────────────────────

def _is_zoom_running() -> bool:
    for proc in psutil.process_iter(["name"]):
        try:
            if proc.info["name"] and proc.info["name"].lower() in ("zoom.exe", "zoom"):
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return False


def _record_system_audio(wav_path: Path, stop_event: threading.Event, job_id: str):
    """Record WASAPI loopback (system audio) to wav_path until stop_event is set."""
    p = pyaudio.PyAudio()
    try:
        try:
            wasapi_info = p.get_host_api_info_by_type(pyaudio.paWASAPI)
        except Exception:
            raise RuntimeError("WASAPI לא נמצא במחשב זה.")

        default_speakers = p.get_device_info_by_index(wasapi_info["defaultOutputDevice"])

        if not default_speakers.get("isLoopbackDevice"):
            loopback_device = None
            for loopback in p.get_loopback_device_info_generator():
                if default_speakers["name"] in loopback["name"]:
                    loopback_device = loopback
                    break
            if loopback_device is None:
                raise RuntimeError("לא נמצא מכשיר הקלטה loopback. ודא שהרמקולים מחוברים ופעילים.")
            default_speakers = loopback_device

        channels = default_speakers["maxInputChannels"]
        sample_rate = int(default_speakers["defaultSampleRate"])

        stream = p.open(
            format=pyaudio.paInt16,
            channels=channels,
            rate=sample_rate,
            frames_per_buffer=512,
            input=True,
            input_device_index=default_speakers["index"],
        )

        frames = []
        try:
            while not stop_event.is_set() and job_id not in cancelled:
                data = stream.read(512, exception_on_overflow=False)
                frames.append(data)
        finally:
            stream.stop_stream()
            stream.close()

        if frames:
            with wave.open(str(wav_path), "wb") as wf:
                wf.setnchannels(channels)
                wf.setsampwidth(p.get_sample_size(pyaudio.paInt16))
                wf.setframerate(sample_rate)
                wf.writeframes(b"".join(frames))
    finally:
        p.terminate()


def _monitor_zoom(stop_event: threading.Event, job_id: str):
    """Monitor Zoom process. When Zoom exits, set stop_event to stop recording."""
    zoom_running = _is_zoom_running()

    if not zoom_running:
        update_job(job_id, "recording", "ממתין לזום...", progress=5)
        deadline = time.time() + 600  # wait up to 10 min for Zoom to start
        while time.time() < deadline:
            if stop_event.is_set() or job_id in cancelled:
                return
            if _is_zoom_running():
                zoom_running = True
                break
            time.sleep(2)

    if not stop_event.is_set() and job_id not in cancelled:
        update_job(job_id, "recording", "מקליט... (זום פעיל)", progress=10)

    start_time = time.time()
    while not stop_event.is_set() and job_id not in cancelled:
        if zoom_running and not _is_zoom_running():
            stop_event.set()
            break
        elapsed = int(time.time() - start_time)
        mins, secs = divmod(elapsed, 60)
        update_job(job_id, "recording", f"מקליט... {mins:02d}:{secs:02d}", progress=15)
        time.sleep(3)


def _find_zoom_hwnd():
    """מוצא את חלון פגישת זום — מתעדף 'Zoom Meeting' על פני 'Zoom Workplace'."""
    meeting_windows = []
    fallback_windows = []

    def callback(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if not title:
            return
        tl = title.lower()
        rect = win32gui.GetWindowRect(hwnd)
        w = rect[2] - rect[0]
        h = rect[3] - rect[1]
        if w < 200 or h < 200:
            return
        area = w * h
        if "zoom meeting" in tl or ("zoom" in tl and "workplace" not in tl and "zoom workplace" not in tl):
            meeting_windows.append((hwnd, title, area))
        elif "zoom" in tl:
            fallback_windows.append((hwnd, title, area))

    win32gui.EnumWindows(callback, None)
    meeting_windows.sort(key=lambda x: x[2], reverse=True)
    fallback_windows.sort(key=lambda x: x[2], reverse=True)
    combined = meeting_windows + fallback_windows
    return combined[0][0] if combined else None


def _get_physical_rect(hwnd):
    """מחזיר גבולות פיזיים אמיתיים של החלון (ללא צל DWM וללא בעיות DPI)."""
    DWMWA_EXTENDED_FRAME_BOUNDS = 9
    rect = WINRECT()
    try:
        ctypes.windll.dwmapi.DwmGetWindowAttribute(
            hwnd, DWMWA_EXTENDED_FRAME_BOUNDS,
            ctypes.byref(rect), ctypes.sizeof(rect)
        )
        return rect.left, rect.top, rect.right, rect.bottom
    except Exception:
        return win32gui.GetWindowRect(hwnd)


def _capture_window_frame(hwnd):
    """מצלם פריים מחלון באמצעות PrintWindow (עובד גם מאחורה)."""
    try:
        left, top, right, bottom = _get_physical_rect(hwnd)
        width = (right - left) & ~1
        height = (bottom - top) & ~1
        if width <= 0 or height <= 0:
            return None, 0, 0
        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
        saveDC.SelectObject(saveBitMap)
        ctypes.windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 3)  # PW_RENDERFULLCONTENT
        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)
        img = np.frombuffer(bmpstr, dtype="uint8").reshape((bmpinfo["bmHeight"], bmpinfo["bmWidth"], 4))
        return img[:, :, :3], width, height
    except Exception:
        return None, 0, 0


def _record_window_video(video_path: Path, stop_event: threading.Event, job_id: str, fps: int = 10):
    """מקליט את חלון זום ישירות דרך PrintWindow — עובד גם כשהחלון מאחורה."""
    hwnd = None
    deadline = time.time() + 120
    while time.time() < deadline:
        if stop_event.is_set() or job_id in cancelled:
            return
        hwnd = _find_zoom_hwnd()
        if hwnd:
            break
        time.sleep(2)

    if not hwnd:
        raise RuntimeError("לא נמצא חלון זום להקלטת וידאו")

    left, top, right, bottom = _get_physical_rect(hwnd)
    w = (right - left) & ~1
    h = (bottom - top) & ~1

    cmd = [
        "ffmpeg", "-y",
        "-f", "rawvideo", "-vcodec", "rawvideo",
        "-pix_fmt", "bgr24",
        "-s", f"{w}x{h}",
        "-r", str(fps),
        "-i", "pipe:0",
        "-c:v", "libx264", "-preset", "ultrafast", "-crf", "28",
        "-an", "-f", "matroska", str(video_path)
    ]
    # stdout/stderr חייבים להיות DEVNULL — אחרת ffmpeg נחסם כשה-pipe מתמלא
    proc = subprocess.Popen(cmd, stdin=subprocess.PIPE,
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    _running_subprocesses[f"{job_id}_video"] = proc

    frame_interval = 1.0 / fps
    try:
        while not stop_event.is_set() and job_id not in cancelled:
            t = time.time()
            if proc.poll() is not None:
                break
            frame, fw, fh = _capture_window_frame(hwnd)
            if frame is not None:
                if fw != w or fh != h:
                    frame = cv2.resize(frame, (w, h))
                try:
                    proc.stdin.write(frame.tobytes())
                except (BrokenPipeError, OSError):
                    break
            elapsed = time.time() - t
            sleep_time = frame_interval - elapsed
            if sleep_time > 0:
                time.sleep(sleep_time)
    finally:
        _running_subprocesses.pop(f"{job_id}_video", None)
        try:
            proc.stdin.close()
        except Exception:
            pass
        try:
            proc.wait(timeout=15)
        except Exception:
            proc.kill()


async def process_recording(job_id: str, title: str, keep_audio: bool = False, save_video: bool = False):
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    wav_path = TEMP_DIR / f"{job_id}_rec.wav"
    audio_path = TEMP_DIR / f"{job_id}.mp3"
    video_raw_path = TEMP_DIR / f"{job_id}_raw.mkv"
    temp_files = [wav_path, audio_path, video_raw_path]

    stop_event = threading.Event()
    _recording_stop_events[job_id] = stop_event

    try:
        loop = asyncio.get_event_loop()

        futs = [
            loop.run_in_executor(executor, _record_system_audio, wav_path, stop_event, job_id),
            loop.run_in_executor(executor, _monitor_zoom, stop_event, job_id),
        ]
        if save_video:
            futs.append(loop.run_in_executor(executor, _record_window_video, video_raw_path, stop_event, job_id))
        await asyncio.gather(*futs)

        _recording_stop_events.pop(job_id, None)

        if job_id in cancelled:
            raise JobCancelled("בוטל")

        if not wav_path.exists():
            raise Exception("ההקלטה נכשלה — קובץ לא נוצר.")

        update_job(job_id, "converting", "ממיר לMP3...", progress=30)
        # dynaudnorm מנרמל עוצמה — מונע הקלטה שקטה כשהזום בנמיך
        cmd = ["ffmpeg", "-i", str(wav_path), "-vn", "-acodec", "libmp3lame", "-q:a", "4",
               "-af", "dynaudnorm=p=0.9:s=5", "-y", str(audio_path)]
        await loop.run_in_executor(executor, lambda: _run_subprocess(job_id, cmd))
        wav_path.unlink(missing_ok=True)

        if not title:
            title = f"זום {time.strftime('%Y-%m-%d %H-%M')}"

        # מיזוג וידאו + אודיו
        saved_video = None
        _log = Path(__file__).parent / "video_debug.log"
        _log.write_text(
            f"save_video={save_video}\n"
            f"video_raw_path={video_raw_path}\n"
            f"exists={video_raw_path.exists()}\n"
            f"size={video_raw_path.stat().st_size if video_raw_path.exists() else 'N/A'}\n"
            f"audio_exists={audio_path.exists()}\n",
            encoding="utf-8"
        )
        if save_video and video_raw_path.exists():
            update_job(job_id, "converting", "ממזג וידאו ואודיו...", progress=40)
            course_folder = OUTPUT_DIR / _safe_filename(title, job_id)
            course_folder.mkdir(parents=True, exist_ok=True)
            existing_mp4 = len(list(course_folder.glob("*.mp4")))
            vname = f"שיעור {HEBREW_NUMS[existing_mp4]}" if existing_mp4 < len(HEBREW_NUMS) else f"שיעור {existing_mp4 + 1}"
            saved_video = course_folder / f"{vname}.mp4"
            merge_cmd = [
                "ffmpeg",
                "-i", str(video_raw_path),
                "-i", str(audio_path),
                "-c:v", "copy",
                "-c:a", "aac", "-shortest", "-y", str(saved_video)
            ]
            try:
                await loop.run_in_executor(executor, lambda: _run_subprocess(job_id, merge_cmd))
            except Exception as merge_err:
                _log.write_text(_log.read_text(encoding="utf-8") + f"merge_error={merge_err}\n", encoding="utf-8")
                raise
            video_raw_path.unlink(missing_ok=True)
        elif save_video:
            _log.write_text(_log.read_text(encoding="utf-8") + "SKIPPED: video_raw_path missing\n", encoding="utf-8")

        output_file, course_name = await _transcribe_and_save(
            job_id, audio_path, "הקלטת זום", title, loop
        )

        saved_audio = None
        if keep_audio and audio_path.exists():
            saved_audio = output_file.with_suffix(".mp3")
            audio_path.rename(saved_audio)

        for f in temp_files:
            try:
                Path(f).unlink(missing_ok=True)
            except Exception:
                pass

        update_job(job_id, "done", "הושלם בהצלחה!", progress=100,
                   output_file=str(output_file), title=title,
                   course=course_name, lesson=output_file.stem, type="recording",
                   saved_audio=str(saved_audio) if saved_audio else None,
                   saved_video=str(saved_video) if saved_video else None)

    except JobCancelled:
        stop_event.set()
        _recording_stop_events.pop(job_id, None)
        for f in temp_files:
            try:
                Path(f).unlink(missing_ok=True)
            except Exception:
                pass
        cancelled.discard(job_id)
        update_job(job_id, "cancelled", "בוטל על ידי המשתמש", progress=0)

    except Exception as e:
        stop_event.set()
        _recording_stop_events.pop(job_id, None)
        for f in temp_files:
            try:
                Path(f).unlink(missing_ok=True)
            except Exception:
                pass
        update_job(job_id, "error", str(e), progress=0)


# ── Transcription pipeline ────────────────────────────────────────────────────

async def process_transcription(job_id: str, url: str, title: str, direct_url: str):
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    temp_files = []
    audio_path = TEMP_DIR / f"{job_id}.mp3"

    try:
        _check_cancelled(job_id)
        update_job(job_id, "downloading", "מוריד סרטון...", progress=5)
        actual_title = title
        loop = asyncio.get_event_loop()

        try:
            actual_title = await asyncio.wait_for(
                loop.run_in_executor(executor, _yt_dlp_download, url, job_id, audio_path, "audio"),
                timeout=DOWNLOAD_TIMEOUT,
            )
            if not actual_title:
                actual_title = title or job_id
        except asyncio.TimeoutError:
            raise Exception("ההורדה חרגה מהזמן המותר.")
        except Exception as e:
            if job_id in cancelled:
                raise JobCancelled("בוטל")
            print(f"yt-dlp failed: {e}")

        _check_cancelled(job_id)

        if not audio_path.exists():
            dl_url = direct_url or url
            if not dl_url:
                raise Exception("לא ניתן להוריד את הסרטון.")
            update_job(job_id, "downloading", "מנסה הורדה ישירה...", progress=5)
            raw_path = TEMP_DIR / f"{job_id}_raw"
            await asyncio.wait_for(
                loop.run_in_executor(executor, _direct_download, dl_url, raw_path, job_id),
                timeout=DOWNLOAD_TIMEOUT,
            )
            temp_files.append(raw_path)

            _check_cancelled(job_id)
            update_job(job_id, "converting", "ממיר לMP3...", progress=68)
            cmd = ["ffmpeg", "-i", str(raw_path), "-vn", "-acodec", "libmp3lame", "-q:a", "4", "-y", str(audio_path)]
            await loop.run_in_executor(executor, lambda: _run_subprocess(job_id, cmd))

        temp_files.append(audio_path)

        if not audio_path.exists():
            raise Exception("קובץ האודיו לא נוצר.")

        _check_cancelled(job_id)
        update_job(job_id, "converting", "ממיר לMP3...", progress=70)

        output_file, course_name = await _transcribe_and_save(job_id, audio_path, url, actual_title, loop)

        for f in temp_files:
            try:
                Path(f).unlink(missing_ok=True)
            except Exception:
                pass

        update_job(job_id, "done", "הושלם בהצלחה!", progress=100,
                   output_file=str(output_file), title=actual_title,
                   course=course_name, lesson=output_file.stem)

    except JobCancelled:
        for f in temp_files:
            try:
                Path(f).unlink(missing_ok=True)
            except Exception:
                pass
        cancelled.discard(job_id)
        update_job(job_id, "cancelled", "בוטל על ידי המשתמש", progress=0)

    except Exception as e:
        for f in temp_files:
            try:
                Path(f).unlink(missing_ok=True)
            except Exception:
                pass
        update_job(job_id, "error", str(e), progress=0)


async def process_download(job_id: str, url: str, title: str, fmt: str):
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    temp_files = []

    try:
        _check_cancelled(job_id)
        update_job(job_id, "downloading", f"מוריד {fmt.upper()}...", progress=5)
        actual_title = title
        loop = asyncio.get_event_loop()
        dl_fmt = "audio" if fmt == "mp3" else "video"

        actual_title = await asyncio.wait_for(
            loop.run_in_executor(executor, _yt_dlp_download, url, job_id,
                                 TEMP_DIR / f"{job_id}.{fmt}", dl_fmt),
            timeout=DOWNLOAD_TIMEOUT,
        )
        if not actual_title:
            actual_title = title or job_id

        _check_cancelled(job_id)

        candidates = list(TEMP_DIR.glob(f"{job_id}.*"))
        if not candidates:
            raise Exception("הקובץ לא נוצר.")
        src = candidates[0]
        temp_files.append(src)

        update_job(job_id, "saving", "שומר קובץ...", progress=90)
        safe_title = _safe_filename(actual_title, job_id)
        ext = src.suffix
        dest = OUTPUT_DIR / f"{safe_title}{ext}"
        counter = 1
        while dest.exists():
            dest = OUTPUT_DIR / f"{safe_title}_{counter}{ext}"
            counter += 1

        src.rename(dest)
        update_job(job_id, "done", "הורדה הושלמה!", progress=100,
                   output_file=str(dest), title=actual_title)

    except JobCancelled:
        for f in temp_files:
            try:
                Path(f).unlink(missing_ok=True)
            except Exception:
                pass
        cancelled.discard(job_id)
        update_job(job_id, "cancelled", "בוטל על ידי המשתמש", progress=0)

    except Exception as e:
        for f in temp_files:
            try:
                Path(f).unlink(missing_ok=True)
            except Exception:
                pass
        update_job(job_id, "error", str(e), progress=0)


if __name__ == "__main__":
    import uvicorn
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    _cleanup_temp_dir()
    print(f"שרת תמלול מופעל על http://127.0.0.1:{SERVER_PORT}")
    print(f"תמלולים יישמרו ב: {OUTPUT_DIR}")
    uvicorn.run(app, host="127.0.0.1", port=SERVER_PORT, log_level="warning")
