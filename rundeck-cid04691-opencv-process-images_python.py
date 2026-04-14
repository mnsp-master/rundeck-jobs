#mnspver 0.0.3

import cv2
import numpy as np
import os
import csv
import sys

# 1. Load the network
net = cv2.dnn.readNetFromCaffe("deploy.prototxt", "res10_300x300_ssd_iter_140000.caffemodel")

# 2. Setup Inputs (Now expecting a file path instead of a directory)
img_path = sys.argv[1]  # Path to the specific image
output_dir = sys.argv[2] # Where to save the CSV
os.makedirs(output_dir, exist_ok=True)

if not os.path.isfile(img_path):
    print(f"Error: File {img_path} not found.")
    sys.exit(1)

# 3. Initialize the CSV log file
log_file = os.path.join(output_dir, "face_metadata.csv")
filename = os.path.basename(img_path)

# 4. Process the single image
img = cv2.imread(img_path)
if img is None:
    print(f"Error: Could not read {img_path}")
    sys.exit(1)

(h, w) = img.shape[:2]
blob = cv2.dnn.blobFromImage(cv2.resize(img, (300, 300)), 1.0, (300, 300), (104.0, 177.0, 123.0))
net.setInput(blob)
detections = net.forward()

# Open CSV and write results
with open(log_file, "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Filename", "CenterX", "CenterY", "Width", "Height", "StartX", "StartY", "EndX", "EndY", "Confidence"])

    for i in range(0, detections.shape[2]):
        confidence = detections[0, 0, i, 2]
        
        if confidence > 0.5:
            box = detections[0, 0, i, 3:7] * np.array([w, h, w, h])
            (sX, sY, eX, eY) = box.astype("int")

            # Ensure coordinates stay within image boundaries
            sX, sY = max(0, sX), max(0, sY)
            eX, eY = min(w - 1, eX), min(h - 1, eY)

            cX, cY = int((sX + eX) / 2), int((sY + eY) / 2)
            fW, fH = eX - sX, eY - sY

            writer.writerow([filename, cX, cY, fW, fH, sX, sY, eX, eY, f"{confidence:.4f}"])

print(f"Detection complete for {filename}. Metadata saved to {log_file}")
