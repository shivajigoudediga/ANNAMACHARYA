import cv2
import face_recognition
import os
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import win32com.client
import threading
import time

class FaceAttendanceSystem:
    def __init__(self):
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.dataset_folder = "dataset"
        self.temp_folder = "temp_attendance"
        self.known_faces = []
        self.known_names = []
        self.temp_attendance = {}
        self.late_attendance = {}
        self.start_time = None
        
        self.SESSION_TIMINGS = [
            ("07:26", "08:50"), ("10:17", "11:17"), ("11:19", "12:19"), ("12:21", "13:10"),
            ("14:48", "15:17"), ("15:17", "16:17"), ("21:58", "22:55")
        ]
        
        self.load_known_faces()
        self.create_temp_folder()

    def load_known_faces(self):
        """Load face encodings from dataset folder."""
        if not os.path.exists(self.dataset_folder):
            print("⚠️ Dataset folder not found! Please add student images.")
            return

        for student_folder in os.listdir(self.dataset_folder):
            student_path = os.path.join(self.dataset_folder, student_folder)
            if not os.path.isdir(student_path):
                continue

            student_id = student_folder
            for filename in os.listdir(student_path):
                image_path = os.path.join(student_path, filename)
                image = face_recognition.load_image_file(image_path)
                encodings = face_recognition.face_encodings(image)
                if encodings:
                    self.known_faces.append(encodings[0])
                    self.known_names.append(student_id)
        
        print(f"✅ Loaded {len(self.known_faces)} student face encodings.")

    def create_temp_folder(self):
        """Create temporary folder if it doesn't exist."""
        if not os.path.exists(self.temp_folder):
            os.makedirs(self.temp_folder)
            print(f"✅ Created temporary folder: {self.temp_folder}")

    def get_current_class(self):
        """Get the active class session based on the current time."""
        now = datetime.now().time()
        for start, end in self.SESSION_TIMINGS:
            start_time = datetime.strptime(start, "%H:%M").time()
            end_time = datetime.strptime(end, "%H:%M").time()
            if start_time <= now <= end_time:
                return f"{start}-{end}", datetime.combine(datetime.today(), start_time)
        return None, None

    def mark_temporary_attendance(self, name, student_id):
        """Store attendance temporarily when 'A' is pressed."""
        if not self.start_time:
            print("⚠️ No active class session.")
            return

        now = datetime.now()
        elapsed_minutes = (now - self.start_time).total_seconds() / 60

        if elapsed_minutes > 10:
            print(f"⛔ {name} is late and marked absent in temporary file.")
            self.speaker.Speak(f"{name}, you are late. You are marked absent.")
            self.late_attendance[student_id] = name
            self.save_late_attendance()
            return

        if student_id in self.temp_attendance:
            print(f"⚠️ {name} has already been recorded temporarily.")
            self.speaker.Speak(f"{name}, you have already marked your attendance.")
            return

        self.temp_attendance[student_id] = name
        print(f"✅ Temporary attendance recorded for {name}.")
        self.speaker.Speak(f"✅ Attendance temporarily marked for {name}.")

    def save_late_attendance(self):
        """Save late students' attendance to temporary file."""
        temp_filename = os.path.join(
            self.temp_folder, 
            f"late_attendance_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.csv"
        )
        
        df = pd.DataFrame(columns=["Name", "ID", "Status", "Time"])
        for student_id, name in self.late_attendance.items():
            df = pd.concat([df, pd.DataFrame([{
                "Name": name,
                "ID": student_id,
                "Status": "Absent (Late)",
                "Time": datetime.now().strftime("%H:%M:%S")
            }])], ignore_index=True)
        
        df.to_csv(temp_filename, index=False)
        print(f"✅ Late attendance saved to temporary file: {temp_filename}")

    def confirm_attendance(self):
        leave=input("please to confirm anyone leave the class ")
        if leave==ord('K'):
            print("🔔 Attendance confirmation time reached.")
            self.speaker.Speak("🔔 Attendance confirmation time. Did any student leave? Press 'Y' to enter their ID, or 'N' to continue.")

        while True:
            choice = input("Did any student leave? (Y/N): ").strip().lower()
            if choice == 'y':
                student_id = input("Enter student ID who left: ").strip()
                if student_id in self.temp_attendance:
                    del self.temp_attendance[student_id]
                    print(f"❌ {student_id} marked absent due to leaving the class.")
                    self.speaker.Speak(f"❌ {student_id} is marked absent due to leaving the class.")
                else:
                    print("⚠️ Student not found in temporary records.")
                    self.speaker.Speak("⚠️ Student not found in temporary records.")
            elif choice == 'n':
                break

        self.save_attendance()

    def save_attendance(self):
        """Store attendance permanently by combining temp and late records."""
        if not self.start_time:
            print("⚠️ No active class session.")
            return

        session_filename = f"attendance_{datetime.now().strftime('%Y-%m-%d')}.csv"
        df = pd.DataFrame(columns=["Name", "ID", "Status"])

        # Add present students
        for student_id, name in self.temp_attendance.items():
            df = pd.concat([df, pd.DataFrame([{
                "Name": name,
                "ID": student_id,
                "Status": "Present"
            }])], ignore_index=True)

        # Add late students
        for student_id, name in self.late_attendance.items():
            df = pd.concat([df, pd.DataFrame([{
                "Name": name,
                "ID": student_id,
                "Status": "Absent (Late)"
            }])], ignore_index=True)

        df.to_csv(session_filename, index=False)
        print(f"✅ Final attendance saved to {session_filename}")
        self.speaker.Speak("✅ Attendance saved successfully.")

        # Clean up
        self.temp_attendance.clear()
        self.late_attendance.clear()
        
        # Remove temporary files
        for file in os.listdir(self.temp_folder):
            if file.startswith("late_attendance_"):
                os.remove(os.path.join(self.temp_folder, file))
        print("🧹 Cleaned up temporary attendance files.")

    def recognize_faces(self):
        """Run face recognition and update attendance."""
        current_class, self.start_time = self.get_current_class()
        if not current_class:
            print("⏳ No active class session.")
            return

        print(f"📌 Current session: {current_class}")
        
        cap = cv2.VideoCapture(0)
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
                face_distances = face_recognition.face_distance(self.known_faces, face_encoding)
                if len(face_distances) > 0:
                    best_match_index = np.argmin(face_distances)
                    if face_distances[best_match_index] < 0.6:
                        detected_id = self.known_names[best_match_index]
                        detected_name = detected_id

                # Change rectangle color based on status
                if detected_id in self.temp_attendance:
                    color = (0, 165, 255)  # Orange for marked
                elif detected_id in self.late_attendance:
                    color = (0, 0, 255)    # Red for late
                else:
                    color = (0, 255, 0)     # Green for new

                cv2.rectangle(frame, (left, top), (right, bottom), color, 2)
                cv2.putText(frame, detected_name, (left, top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 2)

            cv2.imshow("Face Recognition Attendance", frame)
            key = cv2.waitKey(10) & 0xFF

            if key == ord('a') and detected_id != "Unknown":
                self.mark_temporary_attendance(detected_name, detected_id)
            elif key == ord('q'):
                break
            elif key == ord('p'):
                confirm_thread = threading.Thread(target=self.confirm_attendance)
                confirm_thread.start()

        cap.release()
        cv2.destroyAllWindows()

if __name__ == "__main__":
    attendance_system = FaceAttendanceSystem()
    attendance_system.recognize_faces()