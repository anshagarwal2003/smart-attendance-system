# Smart Attendance System using Face Recognition

A Python-based smart attendance system that automates classroom attendance using face recognition and blink verification.

## Features
- Face recognition based attendance
- Blink detection for liveness verification
- Teacher-controlled class activation
- 15-minute teacher activation rule
- 13-minute attendance window
- Duplicate attendance prevention
- Excel-based attendance reporting
- Summary sheet with present, absent, and percentage

## Technologies Used
- Python
- OpenCV
- Face Recognition
- Mediapipe
- Pandas
- OpenPyXL
- NumPy

## Project Structure
- `main.py` → main attendance system
- `encode.py` → generates face encodings
- `students.csv` → student name and roll mapping
- `EncodeFile.p` → saved face encodings
- `Attendance.xlsx` → attendance output
- `images/` → student images
- `screenshots/` → project screenshots

## How It Works
1. Student images are stored in the images folder.
2. `encode.py` generates face encodings.
3. Teacher activates the class session.
4. Students verify using live camera and blink detection.
5. Attendance is stored in Excel and summary is generated.

## Team Members
- Harshita Mangal
- Ansh Agarwal
- Darshan Singh

## Note
This project is developed for academic purposes.