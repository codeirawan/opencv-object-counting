import yaml
import numpy as np
import cv2
from datetime import datetime
import openpyxl
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import os

# Initialize variables
start_time = datetime.now()
last_time = datetime.now()
ct = 0
total_output = 0
fastest = 0
ppm = 0
ppm_average = 0
rec_qty = 8
qty = 0

# Prepare for Excel file output
path = "../output/"
output_file = f"output_{start_time.strftime('%d-%m-%Y_%H-%M-%S')}.xlsx"
wb = openpyxl.Workbook()
ws = wb.active
ws.append(("datetime", "total_output", "seconds", "average ppm", "ct", "ppm"))

# File paths
fn_yaml = r"../datasets/area.yml"
fn_out = r"../datasets/output.avi"

# Configuration dictionary
config = {
    'save_video': False,
    'text_overlay': True,
    'object_overlay': True,
    'object_id_overlay': False,
    'object_detection': True,
    'min_area_motion_contour': 60,
    'park_sec_to_wait': 0.001,
    'start_frame': 0
}

# Set capture device
cap = cv2.VideoCapture(0)

# Define the codec and create VideoWriter object if saving video
if config['save_video']:
    fourcc = cv2.VideoWriter_fourcc('D', 'I', 'V', 'X')
    out = cv2.VideoWriter(fn_out, fourcc, 25.0, (640, 480))  # Use appropriate resolution

# Read YAML data
with open(fn_yaml, 'r') as stream:
    object_area_data = yaml.safe_load(stream)

object_contours = []
object_bounding_rects = []
object_mask = []

# Process object areas from YAML
for park in object_area_data:
    points = np.array(park['points'])
    rect = cv2.boundingRect(points)
    points_shifted = points.copy()
    points_shifted[:, 0] -= rect[0]
    points_shifted[:, 1] -= rect[1]
    object_contours.append(points)
    object_bounding_rects.append(rect)
    mask = cv2.drawContours(np.zeros((rect[3], rect[2]), dtype=np.uint8), [points_shifted], contourIdx=-1, color=255, thickness=-1, lineType=cv2.LINE_AA)
    mask = mask == 255
    object_mask.append(mask)

object_status = [False] * len(object_area_data)
object_buffer = [None] * len(object_area_data)

# Main loop
while cap.isOpened():
    try:
        # Capture frame-by-frame
        video_cur_pos = cap.get(cv2.CAP_PROP_POS_MSEC) / 1000.0  # Current position in seconds
        ret, frame = cap.read()    
        if not ret:
            print("Capture Error")
            break

        frame_blur = cv2.GaussianBlur(frame.copy(), (5, 5), 3)
        frame_gray = cv2.cvtColor(frame_blur, cv2.COLOR_BGR2GRAY)
        frame_out = frame.copy()

        if config['object_detection']:
            for ind, park in enumerate(object_area_data):
                points = np.array(park['points'])
                rect = object_bounding_rects[ind]
                roi_gray = frame_gray[rect[1]:rect[1] + rect[3], rect[0]:rect[0] + rect[2]]
                status = np.std(roi_gray) < 20 and np.mean(roi_gray) > 56

                if status != object_status[ind] and object_buffer[ind] is None:
                    object_buffer[ind] = video_cur_pos

                elif status != object_status[ind] and object_buffer[ind] is not None:
                    if video_cur_pos - object_buffer[ind] > config['park_sec_to_wait']:
                        if not status:
                            qty += 1
                            total_output += 1
                            current_time = datetime.now()
                            diff = current_time - last_time
                            ct = diff.total_seconds()
                            ppm = round(1 / ct, 2)  # Calculate PPM (Pieces Per Second)
                            last_time = current_time

                            diff = current_time - start_time
                            seconds = diff.total_seconds()
                            ppm_average = round(total_output / seconds, 2)  # Average PPM (Pieces Per Second)

                            if ppm > fastest:
                                fastest = ppm
                                data = (current_time, total_output, seconds, ppm_average, ct, ppm)
                                ws.append(data)

                            if qty > rec_qty:
                                data = (current_time, total_output, seconds, ppm_average, ct, ppm)
                                ws.append(data)
                                qty = 0

                        object_status[ind] = status
                        object_buffer[ind] = None

                elif status == object_status[ind] and object_buffer[ind] is not None:
                    object_buffer[ind] = None

        if config['object_overlay']:
            for ind, park in enumerate(object_area_data):
                points = np.array(park['points'])
                color = (0, 255, 0) if object_status[ind] else (0, 0, 255)
                cv2.drawContours(frame_out, [points], contourIdx=-1, color=color, thickness=2, lineType=cv2.LINE_AA)
                moments = cv2.moments(points)
                centroid = (int(moments['m10'] / moments['m00']) - 3, int(moments['m01'] / moments['m00']) + 3)
                cv2.putText(frame_out, str(park['id']), centroid, cv2.FONT_HERSHEY_SIMPLEX, 0.4, (0, 0, 0), 1, cv2.LINE_AA)

        if config['text_overlay']:
            cv2.rectangle(frame_out, (1, 5), (350, 70), (0, 255, 0), 2)
            str_on_frame = f"Object Counting: Total Counting = {total_output}, Speed (PPM) = {ppm}"
            cv2.putText(frame_out, str_on_frame, (5, 40), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 255), 2, cv2.LINE_AA)
            str_on_frame = f"Fastest PPM: {fastest}, Average: {ppm_average}"
            cv2.putText(frame_out, str_on_frame, (5, 60), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 0), 1, cv2.LINE_AA)

        # Display the resulting frame
        imS = cv2.resize(frame_out, (960, 720))
        cv2.imshow('Output Counting - by Yaser Ali Husen', imS)
        
        k = cv2.waitKey(1)
        if k == ord('q'):
            break
        elif k == ord('c'):
            cv2.imwrite(f'frame{int(video_cur_pos)}.jpg', frame_out)

    except KeyboardInterrupt:
        data = (datetime.now(), total_output, (datetime.now() - start_time).total_seconds(), ppm_average, ct, ppm)
        ws.append(data)
        wb.save(path + output_file)
        print(f"Actual Speed (PPM): {ppm_average}")
        break

cap.release()
if config['save_video']:
    out.release()
cv2.destroyAllWindows()

# Save the Excel file
wb.save(path + output_file)

# Upload to Google Drive
gauth = GoogleAuth()
gauth.LocalWebserverAuth()  # Creates local webserver and auto handles authentication
drive = GoogleDrive(gauth)

# Create a file instance and set its content from the local file
gfile = drive.CreateFile({'title': os.path.basename(path + output_file)})
gfile.SetContentFile(path + output_file)

# Upload the file
gfile.Upload()

print(f"File uploaded successfully to Google Drive with ID: {gfile['id']}")
