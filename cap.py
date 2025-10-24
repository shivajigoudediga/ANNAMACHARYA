import tkinter as tk
from tkinter import simpledialog, messagebox

elif key == ord('n'):
    cv2.destroyAllWindows()  # Close OpenCV to prevent conflict

    root = tk.Tk()
    root.withdraw()  # Hide the root window

    response = messagebox.askquestion("Attendance Confirmation", "Do you want to deny attendance? (Yes/No)")

    if response == "yes":
        reason = simpledialog.askstring("Reason", "Enter reason for denial:")
        print(f"❌ Attendance confirmation denied. Reason: {reason}")
    else:
        print("✅ Attendance confirmed.")

    root.destroy()  # Close Tkinter window
    cv2.imshow("Face Recognition Attendance", frame)  # Reopen OpenCV window if needed
