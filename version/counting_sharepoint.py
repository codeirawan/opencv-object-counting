import yaml
import numpy as np
import cv2
from datetime import datetime
import openpyxl
import requests

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
wb = openpyxl.Workbook()
ws = wb.active
ws.append(("datetime", "total_output", "minute", "average ppm", "ct", "ppm"))

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
object_buffer = [None] * len(object_area_data]

# Define functions for SharePoint interaction
def get_access_token(tenant_id, client_id, client_secret):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default'
    }
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    response = requests.post(url, headers=headers, data=payload)
    response.raise_for_status()
    return response.json().get('access_token')

def upload_file_to_sharepoint(access_token, site_id, drive_id, file_path, file_name):
    # Read the file content
    with open(file_path, 'rb') as file:
        file_content = file.read()
    
    # Get upload session URL
    upload_session_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{file_name}:/createUploadSession"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    response = requests.post(upload_session_url, headers=headers)
    response.raise_for_status()
    upload_url = response.json().get('uploadUrl')

    # Upload the file in chunks
    chunk_size = 320 * 1024  # 320 KB
    file_size = len(file_content)
    for i in range(0, file_size, chunk_size):
        chunk_data = file_content[i:i + chunk_size]
        chunk_start = i
        chunk_end = min(i + chunk_size - 1, file_size - 1)
        headers = {
            'Content-Length': f"{chunk_end - chunk_start + 1}",
            'Content-Range': f"bytes {chunk_start}-{chunk_end}/{file_size}",
            'Authorization': f'Bearer {access_token}'
        }
        response = requests.put(upload_url, headers=headers, data=chunk_data)
        response.raise_for_status()

    return response.json()

# Set your SharePoint site and drive information
tenant_id = 'your-tenant-id'
client_id = 'your-client-id'
client_secret = 'your-client-secret'
site_id = 'your-site-id'
drive_id = 'your-drive-id'

# Get access token
access_token = get_access_token(tenant_id, client_id, client_secret)

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
                            ppm = round(60 / ct, 2)
                            last_time = current_time

                            diff = current_time - start_time
                            minutes = diff.total_seconds() / 60
                            ppm_average = round(total_output / minutes, 2)

                            if ppm > fastest:
                                fastest = ppm
                                data = (current_time, total_output, minutes, ppm_average, ct, ppm)
                                ws.append(data)

                            if qty > rec_qty:
                                data = (current_time, total_output, minutes, ppm_average, ct, ppm)
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
        break

# Save final data to Excel
data = (datetime.now(), total_output, (datetime.now() - start_time).total_seconds() / 60, ppm_average, ct, ppm)
ws.append(data)
excel_file_path = path + "output_" + start_time.strftime("%d-%m-%Y %H-%M-%S") + ".xlsx"
wb.save(excel_file_path)
print(f"Actual Speed (PPM): {ppm_average}")

# Upload to SharePoint
upload_file_to_sharepoint(access_token, site_id, drive_id, excel_file_path, "output.xlsx")

cap.release()
if config['save_video']:
    out.release()
cv2.destroyAllWindows()
