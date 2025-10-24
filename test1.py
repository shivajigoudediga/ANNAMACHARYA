import cv2
import face_recognition
import numpy as np
import pandas as pd
import os
from datetime import datetime, timedelta
import time
import threading
import win32com.client

class TimedAttendanceSystem:
    def __init__(self):
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.dataset_folder = "dataset"
        self.attendance_folder = "attendance_records"
        self.temp_folder = "temp_attendance"
        
        # Face recognition
        self.known_face_encodings = []
        self.known_face_ids = []
        self.current_student_id = None
        self.min_confidence = 0.55
        
        # Attendance tracking
        self.temp_attendance = {}
        self.session_start = None
        self.current_session = None
        self.running = True
        self.confirmation_thread = None
        
        # Session timings (start, end)
        self.SESSION_TIMINGS = [
            ("05:50", "05:59"), ("06:15", "07:03"),
            ("07:58", "08:10"), ("08:12", "08:20"),
            ("08:20", "08:35"), ("15:17", "16:17"),
            ("23:48", "23:55")
        ]

        self.setup_folders()
        self.load_known_faces()
    
    def setup_folders(self):
        os.makedirs(self.attendance_folder, exist_ok=True)
        os.makedirs(self.temp_folder, exist_ok=True)
    
    def load_known_faces(self):
        """Load multiple images per student for better recognition"""
        if not os.path.exists(self.dataset_folder):
            print("⚠️ Create a 'dataset' folder with student subfolders")
            return
        
        for student_id in os.listdir(self.dataset_folder):
            student_path = os.path.join(self.dataset_folder, student_id)
            if os.path.isdir(student_path):
                for img_file in os.listdir(student_path):
                    if img_file.lower().endswith(('.jpg','.jpeg','.png')):
                        img_path = os.path.join(student_path, img_file)
                        img = face_recognition.load_image_file(img_path)
                        encodings = face_recognition.face_encodings(img)
                        if encodings:
                            self.known_face_encodings.append(encodings[0])
                            self.known_face_ids.append(student_id)
        
        print(f"✅ Loaded {len(set(self.known_face_ids))} students")

    def get_current_session(self):
        """Check if current time falls within any session"""
        now = datetime.now().time()
        for start, end in self.SESSION_TIMINGS:
            start_time = datetime.strptime(start, "%H:%M").time()
            end_time = datetime.strptime(end, "%H:%M").time()
            if start_time <= now <= end_time:
                return f"{start}-{end}", datetime.combine(datetime.today(), start_time)
        return None, None

    def recognize_faces(self, frame):
        """Improved face recognition with resize for better performance"""
        small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
        rgb_small = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)
        
        face_locations = face_recognition.face_locations(rgb_small)
        face_encodings = face_recognition.face_encodings(rgb_small, face_locations)
        
        self.current_student_id = None
        
        for (top, right, bottom, left), face_encoding in zip(face_locations, face_encodings):
            matches = face_recognition.compare_faces(self.known_face_encodings, face_encoding, tolerance=0.5)
            face_distances = face_recognition.face_distance(self.known_face_encodings, face_encoding)
            
            best_match_idx = np.argmin(face_distances)
            confidence = 1 - face_distances[best_match_idx]
            
            if matches[best_match_idx] and confidence >= self.min_confidence:
                self.current_student_id = self.known_face_ids[best_match_idx]
                
                # Scale bounding box back up
                top *= 4; right *= 4; bottom *= 4; left *= 4
                
                # Set box color based on status
                status = self.temp_attendance.get(self.current_student_id, {}).get("status", "New")
                color = (0, 255, 0) if status == "New" else \
                        (0, 165, 255) if status == "Present" else \
                        (0, 0, 255)  # Absent/Late
                
                cv2.rectangle(frame, (left, top), (right, bottom), color, 2)
                cv2.putText(frame, f"{self.current_student_id} ({confidence:.2f})", 
                           (left, top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.7, color, 2)
        
        return frame

    def mark_attendance(self):
        """Mark attendance with timing check"""
        if not self.current_student_id:
            print("⚠️ No student detected")
            self.speaker.Speak("No student detected")
            return
            
        now = datetime.now()
        elapsed_minutes = (now - self.session_start).total_seconds() / 60
        
        if self.current_student_id in self.temp_attendance:
            print(f"⚠️ {self.current_student_id} already marked")
            self.speaker.Speak("Already marked")
            return
            
        status = "Present" if elapsed_minutes <= 10 else "Absent"
        self.temp_attendance[self.current_student_id] = {
            "status": status,
            "timestamp": now.strftime("%Y-%m-%d %H:%M:%S")
        }
        
        # Save to temp file
        self.save_temp_attendance()
        
        print(f"✅ {self.current_student_id} marked {status}")
        self.speaker.Speak(f"{self.current_student_id} marked {status}")

    def save_temp_attendance(self):
        """Save current attendance to temp file"""
        temp_path = os.path.join(self.temp_folder, "current_attendance.csv")
        df = pd.DataFrame([
            {"ID": sid, "Status": data["status"], "Timestamp": data["timestamp"]}
            for sid, data in self.temp_attendance.items()
        ])
        df.to_csv(temp_path, index=False)

    def start_confirmation_timer(self):
        """Start 11-minute timer to check for early leavers"""
        def timer_callback():
            time.sleep(660)  # 11 minutes
            self.check_early_leavers()
        
        self.confirmation_thread = threading.Thread(target=timer_callback)
        self.confirmation_thread.start()

    def check_early_leavers(self):
        """Check for students who left class early"""
        print("\n" + "="*50)
        print("ATTENDANCE CONFIRMATION TIME")
        print("="*50)
        self.speaker.Speak("Attendance confirmation. Did any students leave?")
        
        while True:
            response = input("Did any student leave? (y/n): ").lower()
            if response == 'y':
                student_id = input("Enter student ID: ").strip()
                if student_id in self.known_face_ids:
                    self.temp_attendance[student_id] = {
                        "status": "Absent",
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    self.save_temp_attendance()
                    print(f"❌ {student_id} marked absent")
                    self.speaker.Speak(f"{student_id} marked absent")
                else:
                    print("⚠️ Invalid student ID")
            elif response == 'n':
                break
        
        print("\nPress 'P' to save final attendance")
        self.speaker.Speak("Press P to save final attendance")

    def save_final_attendance(self):
        """Save attendance to a session-specific permanent record"""
        if not self.temp_attendance:
            print("⚠️ No attendance to save")
            return

        today = datetime.now().strftime("%Y-%m-%d")
        session_time = self.current_session.replace(":", "").replace("-", "_")  # Format session time
        filename = f"attendance_{today}_{session_time}.csv"
        final_path = os.path.join(self.attendance_folder, filename)

        final_data = []
        for sid, data in self.temp_attendance.items():
            final_data.append({
                "ID": sid,
                "Status": data["status"],
                "Date": today,
                "Session": self.current_session
            })

        df = pd.DataFrame(final_data)

        # Append to the existing session file if it exists
        if os.path.exists(final_path):
            existing = pd.read_csv(final_path)
            df = pd.concat([existing, df])

        df.to_csv(final_path, index=False)

        # Clean up temp file
        temp_path = os.path.join(self.temp_folder, "current_attendance.csv")
        if os.path.exists(temp_path):
            os.remove(temp_path)

        print(f"\n✅ FINAL ATTENDANCE SAVED: {filename}")
        print(df.to_string(index=False))
        self.speaker.Speak(f"Attendance for session {self.current_session.replace('-', ' to ')} saved successfully")


    def run(self):
        """Main system loop"""
        self.current_session, self.session_start = self.get_current_session()
        if not self.current_session:
            print("⏳ No active session right now")
            return
            
        print(f"\n📝 Session Started: {self.current_session}")
        print(f"⏱️  Early leaver check at {(self.session_start + timedelta(minutes=11)).strftime('%H:%M')}")
        self.speaker.Speak(f"Starting attendance for {self.current_session.replace('-', ' to ')} session")
        
        self.start_confirmation_timer()
        
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            print("❌ Camera error")
            return
            
        try:
            while self.running:
                ret, frame = cap.read()
                if not ret:
                    break
                
                # Process frame
                frame = self.recognize_faces(frame)
                
                # Display instructions
                instr = "A:Mark  P:Save  Q:Quit N:check leavers"
                cv2.putText(frame, instr, (10, 30), 
                          cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2)
                
                cv2.imshow(f"Attendance: {self.current_session}", frame)
                
                key = cv2.waitKey(1) & 0xFF
                if key == ord('a'):
                    self.mark_attendance()
                elif key == ord('p'):
                    self.save_final_attendance()
                    break
                elif key == ord('q'):
                    break
                elif key==ord('n'):
                    self.check_early_leavers()
                    
        finally:
            cap.release()
            cv2.destroyAllWindows()
            if self.confirmation_thread:
                self.confirmation_thread.join()
            print("System stopped")

if __name__ == "__main__":
    system = TimedAttendanceSystem()
    system.run()