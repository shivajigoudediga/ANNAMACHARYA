import cv2
import face_recognition
import os
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import win32com.client
import time

# Initialize text-to-speech engine
speaker = win32com.client.Dispatch("SAPI.SpVoice")

SESSION_TIMINGS = [
    ("12:50", "13:50"), ("13:52", "14:52"), ("14:54", "15:54"), ("15:56", "17:10"),
    ("17:30", "18:30"), ("18:32", "19:32"), ("19:34", "20:34"), ("20:36", "21:00")
]

DATASET_FOLDER = "dataset"
today_date = datetime.now().strftime("%Y-%m-%d")
known_faces = []
known_names = []

def load_known_faces():
    """Load face encodings from dataset."""
    global known_faces, known_names
    if not os.path.exists(DATASET_FOLDER):
        print("⚠️ No dataset found. Please capture faces first.")
        return
    
    for folder in os.listdir(DATASET_FOLDER):
        student_folder = os.path.join(DATASET_FOLDER, folder)
        if not os.path.isdir(student_folder):
            continue
        student_id = folder  # Use folder name as student ID
        for filename in os.listdir(student_folder):
            image_path = os.path.join(student_folder, filename)
            image = face_recognition.load_image_file(image_path)
            encodings = face_recognition.face_encodings(image)
            if encodings:
                known_faces.append(encodings[0])
                known_names.append(student_id)

def initialize_attendance_file(session_time):
    """Create a separate attendance CSV file for each session."""
    session_filename = f"attendance_{today_date}_{session_time.replace(':', '-')}.csv"
    
    if not os.path.exists(session_filename):
        columns = ["Name", "ID", "Status"]
        df = pd.DataFrame(columns=columns)
        df.to_csv(session_filename, index=False)
        print(f"✅ Created attendance file: {session_filename}")
    
    return session_filename  # ✅ Return the session filename

def get_current_class():
    """Get the active class session based on the current time."""
    now = datetime.now().time()
    for start, end in SESSION_TIMINGS:
        start_time = datetime.strptime(start, "%H:%M").time()
        end_time = datetime.strptime(end, "%H:%M").time()
        if start_time <= now <= end_time:
            return f"{start}-{end}", start_time
    return None, None

def mark_attendance(name, student_id):
    """Mark attendance for the current session."""
    current_class, class_start_time = get_current_class()
    if not current_class:
        print("⏳ No active class session. Waiting for next session...")
        return

    # ✅ Ensure attendance file is created for this session
    session_filename = initialize_attendance_file(current_class)
    
    df = pd.read_csv(session_filename)

    now = datetime.now()
    elapsed_time = (now - datetime.combine(now.date(), class_start_time)).total_seconds() / 60

    if student_id in df["ID"].values:
        print(f"⚠️ {name} is already marked.")
        speaker.Speak(f"⚠️ {name}, you have already marked attendance.")
        return

    status = "Present" if 0 <= elapsed_time <= 10 else "Absent"
    print(f"✅ Marking {status} for {name}")
    speaker.Speak(f"✅ Attendance marked for {name}.") if status == "Present" else speaker.Speak("⛔ You are absent for this class. Come next class.")

    df = pd.concat([df, pd.DataFrame([{"Name": name, "ID": student_id, "Status": status}])], ignore_index=True)
    df.to_csv(session_filename, index=False)  # ✅ Save attendance

    # ✅ Automatically transition to the next session after 2 minutes
    if elapsed_time > 60:
        print("🔄 Transitioning to the next class session after 2 minutes...")
        speaker.Speak("🔄 Next session starting in 2 minutes.")
        time.sleep(120)  # 2-minute delay
        recognize_faces()  # Restart face recognition for the next session

def recognize_faces():
    """Run face recognition and update attendance."""
    cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
    if not cap.isOpened():
        print("❌ Error: Unable to access webcam")
        return
    
    while True:
        ret, frame = cap.read()
        if not ret:
            break
        
        rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        face_locations = face_recognition.face_locations(rgb_frame)
        face_encodings = face_recognition.face_encodings(rgb_frame, face_locations)
        
        detected_name, detected_id = "Unknown", ""
        for face_encoding, (top, right, bottom, left) in zip(face_encodings, face_locations):
            face_distances = face_recognition.face_distance(known_faces, face_encoding)
            if len(face_distances) > 0:
                best_match_index = np.argmin(face_distances)
                if face_distances[best_match_index] < 0.6:
                    detected_id = known_names[best_match_index]
                    detected_name = detected_id.split("_")[0]  
            
            color = (0, 255, 0) if detected_id != "Unknown" else (0, 0, 255)
            cv2.rectangle(frame, (left, top), (right, bottom), color, 2)
            cv2.putText(frame, detected_name, (left, top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 2)
        
        cv2.imshow("Face Recognition Attendance", frame)
        key = cv2.waitKey(10) & 0xFF
        
        if key == ord('a') and detected_id != "Unknown":
            mark_attendance(detected_name, detected_id)
        elif key == ord('q'):
            break
    
    cap.release()
    cv2.destroyAllWindows()
    speaker.Speak("✅ Attendance process ended.")

if __name__ == "__main__":
    load_known_faces()
    
    current_class, _ = get_current_class()  # ✅ Get session time
    if current_class:
        initialize_attendance_file(current_class)  # ✅ Pass session time

    recognize_faces()
