# FaceAttend - Face Recognition Attendance System

A face recognition attendance system built with Python, OpenCV and Flask.
It recognizes students from a camera, marks them present, and keeps a monthly
Excel register with a dated column for every day.

## Features
- Face detection and recognition with OpenCV (YuNet + SFace). No dlib needed.
- Web app with lecturer login, change password and forgot password (recovery code).
- Enroll students from the browser (3 photos each).
- Live attendance page: green = present, orange = already marked, red = unknown.
- Monthly Excel register: upload a class sheet, each day gets a dated column
  (P = present, A = absent) and a Total column. Upload a new sheet each month.
- Desktop version (`attendance_cv.py`) that works from the terminal.

## Setup
```
pip install -r requirements.txt
```
The two small OpenCV model files download automatically on first run.

## Run the web app
```
set ADMIN_PASSWORD=your-password      (Windows; use export on Mac/Linux)
python app.py
```
Open http://localhost:5000 and sign in. Browsers only allow the camera on
`localhost` or HTTPS.

Forgot the password? Use the recovery code on the login page, or run
`python app.py reset-password` on the computer hosting the app.

## Run the desktop version
```
python attendance_cv.py enroll "Full Name" "StudentID"
python attendance_cv.py run
python attendance_cv.py list
python attendance_cv.py remove StudentID
python attendance_cv.py export
```

## Class sheet format
An .xlsx file with a header row containing **Name** and **Student ID** columns.

## Project structure
```
app.py              Flask web app
attendance_cv.py    Recognition, database and desktop version
templates/          HTML pages
static/             style.css and camera.js
```

## Privacy
The database, captured photos and registers contain students' personal data.
They are excluded by `.gitignore`. Do not upload them to a public repository.
