#mnspver 0.0.3_1

import cv2
import numpy as np
import os
import csv
import sys

# 1. Load the network
net = cv2.dnn.readNetFromCaffe("deploy.prototxt", "res10_300x300_ssd_iter_140000.caffemodel")

# 2. Setup Inputs
img_path = sys.argv[1]
output_dir = sys.argv[2]
os.makedirs(output_dir, exist_ok=True)

if not os.path.isfile(img_path):
    sys.exit(1)

# 3. Define Output Filenames
filename = os.path.basename(img_path)
name_only = os.path.splitext(filename)[0]
log_file = os.path.join(output_dir, f"{name_only}.csv")
output_img_path = os.path.join(output_dir, f"detected_{filename}")

# 4. Process Image
img = cv2.imread(img_path)
if img is None:
    sys.exit(1)

(h, w) = img.shape[:2]
blob = cv2.dnn.blobFromImage(cv2.resize(img, (300, 300)), 1.0, (300, 300), (104.0, 177.0, 123.0))
net.setInput(blob)
detections = net.forward()

annotated_img = img.copy()
faces_found = False

# 5. Write CSV (using "w" to ensure 1 CSV per image)
with open(log_file, "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Filename", "CenterX", "CenterY", "Width", "Height", "StartX", "StartY", "EndX", "EndY", "Confidence"])

    for i in range(0, detections.shape[2]):
        confidence = detections[0, 0, i, 2]
        
        if confidence > 0.5:
            faces_found = True
            box = detections[0, 0, i, 3:7] * np.array([w, h, w, h])
            (sX, sY, eX, eY) = box.astype("int")

            sX, sY = max(0, sX), max(0, sY)
            eX, eY = min(w - 1, eX), min(h - 1, eY)

            # Draw Box for ImageMagick visual verification
            cv2.rectangle(annotated_img, (sX, sY), (eX, eY), (0, 255, 0), 2)

            cX, cY = int((sX + eX) / 2), int((sY + eY) / 2)
            fW, fH = eX - sX, eY - sY

            writer.writerow([filename, cX, cY, fW, fH, sX, sY, eX, eY, f"{confidence:.4f}"])

# 6. Save annotated image and exit
if faces_found:
    cv2.imwrite(output_img_path, annotated_img)
    print(f"Success: Metadata in {log_file}")
else:
    print("No faces detected.")
    # Optional: Delete empty CSV if no faces found to keep output_dir clean
    if os.path.exists(log_file):
        os.remove(log_file)
