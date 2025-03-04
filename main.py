import face_recognition  # Library for face detection and recognition
import cv2  # OpenCV for accessing the webcam and image processing
import numpy as np  # NumPy for handling arrays
import csv  # CSV module to store attendance records
from datetime import datetime  # To get the current date and time

# Initialize the webcam
video_capture = cv2.VideoCapture(0)

# ---------------------- LOAD KNOWN FACES ----------------------

# Load an image of a known person and encode their face
sandhya_img = face_recognition.load_image_file("faces/sandhya.png")
sandhya_encoding = face_recognition.face_encodings(sandhya_img)[0]  # Extract face features

rohan_img = face_recognition.load_image_file("faces/rohan.jpg")
rohan_encoding = face_recognition.face_encodings(rohan_img)[0]  # Extract face features

# Store encodings and names of known faces
known_face_encodings = [sandhya_encoding, rohan_encoding]
known_face_names = ["Sandhya", "Rohan"]

# List of employees expected to be marked present
employees = known_face_names.copy()

# ---------------------- ATTENDANCE FILE SETUP ----------------------

# Get current date for the attendance file
now = datetime.now()
current_date = now.strftime('%d-%m-%Y')

# Open a CSV file to store attendance records
f = open(f"{current_date}.csv", "w+", newline="")
lnwriter = csv.writer(f)

# ---------------------- START VIDEO CAPTURE ----------------------

while True:
    # Capture a single frame from the webcam
    _, frame = video_capture.read()

    # Resize frame to 1/4th size for faster processing
    small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)

    # Convert the image from BGR (OpenCV default) to RGB (for face_recognition)
    rgb_small_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)

    # Detect faces in the frame and get their encodings
    face_locations = face_recognition.face_locations(rgb_small_frame)
    face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

    # ---------------------- CHECK FOR KNOWN FACES ----------------------
    
    for face_encoding in face_encodings:
        # Compare the detected face with known faces
        matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
        face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)

        # Find the closest matching face
        best_match_index = np.argmin(face_distances)

        # If a match is found, get the corresponding name
        if matches[best_match_index]:
            name = known_face_names[best_match_index]

            # ---------------------- DISPLAY NAME ON SCREEN ----------------------
            font = cv2.FONT_HERSHEY_SIMPLEX
            position = (10, 100)  # Text position on the screen
            fontScale = 1.5
            fontColor = (255, 0, 0)  # Blue color in BGR format
            thickness = 3
            lineType = 2

            cv2.putText(frame, name + " Present", position, font, fontScale, fontColor, thickness, lineType)

            # ---------------------- MARK ATTENDANCE ----------------------
            if name in employees:
                employees.remove(name)  # Remove from expected employee list
                current_time = now.strftime("%H-%M-%S")  # Get current time
                lnwriter.writerow([name, current_time])  # Write to CSV file

    # Display the video feed with the detected name
    cv2.imshow("Attendance", frame)

    # Press 'q' to exit the program
    if cv2.waitKey(1) & 0xFF == ord("q"):
        break

# ---------------------- CLEANUP ----------------------

video_capture.release()  # Release the webcam
cv2.destroyAllWindows()  # Close all OpenCV windows
f.close()  # Close the attendance file
