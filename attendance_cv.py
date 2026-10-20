"""
Face Recognition Attendance System (OpenCV-only, no dlib)

Uses OpenCV's built-in YuNet (detection) and SFace (recognition) models.
The two small model files download automatically on first run.

Install:
    pip install opencv-python numpy

Usage:
    python attendance_cv.py enroll "Kwame Mensah" "10234567"   # capture faces
    python attendance_cv.py run                                 # take attendance
    python attendance_cv.py export                              # today's CSV

Camera: 0 = laptop webcam. For a phone camera, use an IP-camera app
(e.g. IP Webcam on Android) and set CAMERA_SOURCE to its stream URL,
like "http://192.168.1.5:8080/video".
"""
import csv
import sqlite3
import sys
import urllib.request
from datetime import datetime
from pathlib import Path

import cv2
import numpy as np
import csv
import sqlite3
import sys
import time
import urllib.request
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook, load_workbook
import os

import cv2
import numpy as np

CAMERA_SOURCE = 0
DB_PATH = os.path.join(os.environ.get("DATA_DIR", "."), "attendance_cv.db")
MODELS_DIR = Path("models")
SAMPLES_PER_ENROLL = 3        # face captures per enrollment
MATCH_THRESHOLD = 0.65       # cosine similarity; higher = stricter
SCORE_THRESHOLD = 0.8         # detector confidence

YUNET_URL = ("https://github.com/opencv/opencv_zoo/raw/main/models/"
             "face_detection_yunet/face_detection_yunet_2023mar.onnx")
SFACE_URL = ("https://github.com/opencv/opencv_zoo/raw/main/models/"
             "face_recognition_sface/face_recognition_sface_2021dec.onnx")
YUNET_PATH = MODELS_DIR / "face_detection_yunet_2023mar.onnx"
SFACE_PATH = MODELS_DIR / "face_recognition_sface_2021dec.onnx"


def ensure_models():
    MODELS_DIR.mkdir(exist_ok=True)
    for url, path in ((YUNET_URL, YUNET_PATH), (SFACE_URL, SFACE_PATH)):
        if not path.exists():
            print(f"Downloading {path.name} ...")
            urllib.request.urlretrieve(url, path)


def load_models(width, height):
    ensure_models()
    detector = cv2.FaceDetectorYN.create(
        str(YUNET_PATH), "", (width, height), SCORE_THRESHOLD, 0.3, 5000)
    recognizer = cv2.FaceRecognizerSF.create(str(SFACE_PATH), "")
    return detector, recognizer


def open_camera():
    cap = cv2.VideoCapture(CAMERA_SOURCE)
    if not cap.isOpened():
        sys.exit("Cannot open camera. Check CAMERA_SOURCE.")
    ok, frame = cap.read()
    if not ok:
        sys.exit("Cannot read from camera.")
    h, w = frame.shape[:2]
    return cap, w, h


def db():
    con = sqlite3.connect(DB_PATH)
    con.execute(
        """CREATE TABLE IF NOT EXISTS students (
               student_id TEXT PRIMARY KEY,
               name TEXT NOT NULL)""")
    con.execute(
        """CREATE TABLE IF NOT EXISTS samples (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               student_id TEXT NOT NULL,
               feature BLOB NOT NULL)""")
    con.execute(
        """CREATE TABLE IF NOT EXISTS attendance (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               student_id TEXT NOT NULL,
               date TEXT NOT NULL,
               time TEXT NOT NULL,
               UNIQUE(student_id, date))""")
    return con


def get_feature(recognizer, frame, face_row):
    aligned = recognizer.alignCrop(frame, face_row)
    return recognizer.feature(aligned)  # shape (1, 128), float32


def enroll(name, student_id):
    con = db()
    cap, w, h = open_camera()
    detector, recognizer = load_models(w, h)
    detector.setInputSize((w, h))

    # Re-enrolling replaces old samples for this student
    con.execute("DELETE FROM samples WHERE student_id = ?", (student_id,))
    con.execute("INSERT OR REPLACE INTO students VALUES (?,?)", (student_id, name))
    con.commit()

    captured = 0
    print(f"Enrolling {name}. Press SPACE to capture ({SAMPLES_PER_ENROLL} needed), "
          "Q to cancel. Vary your angle/lighting slightly between captures.")
    while captured < SAMPLES_PER_ENROLL:
        ok, frame = cap.read()
        if not ok:
            break
        _, faces = detector.detect(frame)
        display = frame.copy()
        if faces is not None:
            for f in faces:
                x, y, fw, fh = map(int, f[:4])
                cv2.rectangle(display, (x, y), (x + fw, y + fh), (0, 200, 0), 2)
        cv2.putText(display, f"Captured {captured}/{SAMPLES_PER_ENROLL}",
                    (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 0), 2)
        cv2.imshow("Enroll", display)
        key = cv2.waitKey(1) & 0xFF
        if key == ord("q"):
            break
        if key == ord(" "):
            if faces is None or len(faces) != 1:
                print(f"Found {0 if faces is None else len(faces)} faces; need exactly 1.")
                continue
            feat = get_feature(recognizer, frame, faces[0])
            con.execute("INSERT INTO samples (student_id, feature) VALUES (?,?)",
                        (student_id, feat.astype(np.float32).tobytes()))
            con.commit()
            captured += 1
            print(f"Captured sample {captured}/{SAMPLES_PER_ENROLL}")
    cap.release()
    cv2.destroyAllWindows()
    if captured == SAMPLES_PER_ENROLL:
        print(f"Enrolled {name} ({student_id}).")
    else:
        print("Enrollment incomplete.")


def load_known(con):
    names = dict(con.execute("SELECT student_id, name FROM students").fetchall())
    ids, feats = [], []
    for sid, blob in con.execute("SELECT student_id, feature FROM samples"):
        ids.append(sid)
        feats.append(np.frombuffer(blob, dtype=np.float32).reshape(1, 128))
    return names, ids, feats


def best_match(recognizer, feat, ids, feats):
    best_sid, best_score = None, -1.0
    for sid, known in zip(ids, feats):
        score = recognizer.match(feat, known, cv2.FaceRecognizerSF_FR_COSINE)
        if score > best_score:
            best_sid, best_score = sid, score
    if best_score >= MATCH_THRESHOLD:
        return best_sid, best_score
    return None, best_score


def mark(con, student_id):
    now = datetime.now()
    try:
        con.execute(
            "INSERT INTO attendance (student_id, date, time) VALUES (?,?,?)",
            (student_id, now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")))
        con.commit()
        return True
    except sqlite3.IntegrityError:
        return False  # already marked today
   EXCEL_PATH = os.path.join(os.environ.get("DATA_DIR", "."), "attendance.xlsx")


def log_to_excel(name, student_id, now):
    try:
        if Path(EXCEL_PATH).exists():
            wb = load_workbook(EXCEL_PATH)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Attendance"
            ws.append(["Name", "Student ID", "Date", "Time", "Status"])
            for col, width in zip("ABCDE", (25, 15, 14, 12, 10)):
                ws.column_dimensions[col].width = width
        ws.append([name, student_id, now.strftime("%Y-%m-%d"),
                   now.strftime("%H:%M:%S"), "Present"])
        wb.save(EXCEL_PATH)
    except PermissionError:
        print("Could not write to Excel. Close attendance.xlsx and try again.")

def run():
    con = db()
    names, ids, feats = load_known(con)
    if not feats:
        sys.exit('No students enrolled. Run: python attendance_cv.py enroll "NAME" "ID"')
    cap, w, h = open_camera()
    detector, recognizer = load_models(w, h)
    detector.setInputSize((w, h))
    Path("captures").mkdir(exist_ok=True)

    banner_text, banner_color, banner_until = "", (0, 200, 0), 0.0
    print("Taking attendance. Press Q to quit.")
    while True:
        ok, frame = cap.read()
        if not ok:
            break
        raw = frame.copy()  # clean copy for the saved snapshot
        _, faces = detector.detect(frame)
        if faces is not None:
            for f in faces:
                x, y, fw, fh = map(int, f[:4])
                feat = get_feature(recognizer, frame, f)
                sid, score = best_match(recognizer, feat, ids, feats)
                if sid:
                    label = f"{names[sid]} - RECOGNISED"
                    color = (0, 200, 0)
                    if mark(con, sid):
                        now = datetime.now()
                        log_to_excel(names[sid], sid, now)
                        photo = Path("captures") / f"{now:%Y-%m-%d_%H%M%S}_{sid}.jpg"
                        cv2.imwrite(str(photo), raw)
                        banner_text = f"{names[sid]} - PRESENT"
                        banner_color = (0, 200, 0)
                        banner_until = time.time() + 3
                        print(f"Marked present: {names[sid]} ({sid}) -> {photo}")
                    elif time.time() > banner_until:
                        banner_text = f"{names[sid]} - Already marked today"
                        banner_color = (0, 165, 255)
                        banner_until = time.time() + 2
                else:
                    label, color = "Unknown", (0, 0, 255)
                cv2.rectangle(frame, (x, y), (x + fw, y + fh), color, 2)
                cv2.putText(frame, f"{label} {score:.2f}", (x, max(y - 8, 15)),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.6, color, 2)

        if time.time() < banner_until:
            cv2.rectangle(frame, (0, 0), (frame.shape[1], 50), banner_color, -1)
            cv2.putText(frame, banner_text, (15, 34),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.9, (255, 255, 255), 2)

        cv2.imshow("Attendance", frame)
        if cv2.waitKey(1) & 0xFF == ord("q"):
            break
    cap.release()
    cv2.destroyAllWindows()


def export():
    con = db()
    today = datetime.now().strftime("%Y-%m-%d")
    rows = con.execute(
        """SELECT s.name, a.student_id, a.date, a.time
           FROM attendance a JOIN students s ON s.student_id = a.student_id
           WHERE a.date = ? ORDER BY a.time""", (today,)).fetchall()
    out = Path(f"attendance_{today}.csv")
    with out.open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Name", "Student ID", "Date", "Time"])
        w.writerows(rows)
    print(f"Saved {len(rows)} records to {out}")
def remove(student_id):
    con = db()
    row = con.execute("SELECT name FROM students WHERE student_id = ?",
                      (student_id,)).fetchone()
    if not row:
        print(f"No student with ID {student_id}.")
        return
    for table in ("samples", "attendance", "students"):
        con.execute(f"DELETE FROM {table} WHERE student_id = ?", (student_id,))
    con.commit()
    print(f"Removed {row[0]} ({student_id}).")
def list_students():
    con = db()
    rows = con.execute(
        """SELECT s.name, s.student_id, COUNT(p.id)
           FROM students s LEFT JOIN samples p ON p.student_id = s.student_id
           GROUP BY s.student_id ORDER BY s.name""").fetchall()
    if not rows:
        print("No students enrolled yet.")
        return
    print(f"{'Name':<25}{'Student ID':<15}Samples")
    for name, sid, n in rows:
        print(f"{name:<25}{sid:<15}{n}")
    print(f"\nTotal: {len(rows)}")



if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "enroll" and len(sys.argv) == 4:
        enroll(sys.argv[2], sys.argv[3])
    elif cmd == "run":
        run()
    elif cmd == "export":
        export()
    elif cmd == "remove" and len(sys.argv) == 3:
        remove(sys.argv[2])
    elif cmd == "list":
        list_students()
    else:
        print(__doc__)