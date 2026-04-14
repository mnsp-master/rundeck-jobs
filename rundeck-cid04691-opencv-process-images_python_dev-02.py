#mnspver 0.0.3_2

import cv2
import numpy as np
import os
import csv
import sys

# 1. Load the network
# Ensure these files are in the same directory as the script or provide full paths
net = cv2.dnn.readNetFromCaffe("deploy.prototxt", "res10_300x300_ssd_iter_140000.caffemodel")

# 2. Setup Inputs from SysArgs
if len(sys.argv) < 3:
    print("Usage: python script.py <input_image_path> <output_directory>")
    sys.exit(1)

img_path = sys.argv[1]
output_dir = sys.argv[2]
os.makedirs(output_dir, exist_ok=True)

if not os.path.isfile(img_path):
    print(f"Error: File {img_path} not found.")
    sys.exit(1)

# 3. Define Output Filenames
# Creates photo.csv and detected_photo.jpg for easy matching in PowerShell
filename = os.path.basename(img_path)
name_only = os.path.splitext(filename)[0]
log_file = os.path.join(output_dir, f"{name_only}.csv")
output_img_path = os.path.join(output_dir, f"detected_{filename}")

# 4. Process the Image
img = cv2.imread(img_path)
if img is None:
    print(f"Error: Could not read {img_path}")
    sys.exit(1)

(h, w) = img.shape[:2]
blob = cv2.dnn.blobFromImage(cv2.resize(img, (300, 300)), 1.0, (300, 300), (104.0, 177.0, 123.0))
net.setInput(blob)
detections = net.forward()

annotated_img = img.copy()
faces_found = False

# 5. Process Detections and Write CSV
# Using "w" mode creates a fresh CSV for every image processed
with open(log_file, "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Filename", "CenterX", "CenterY", "Width", "Height", "StartX", "StartY", "EndX", "EndY", "Confidence"])

    for i in range(0, detections.shape[2]):
        confidence = detections[0, 0, i, 2]
        
        # Filter out weak detections
        if confidence > 0.5:
            faces_found = True
            box = detections[0, 0, i, 3:7] * np.array([w, h, w, h])
            (sX, sY, eX, eY) = box.astype("int")

            # Ensure coordinates stay within image boundaries
            sX, sY = max(0, sX), max(0, sY)
            eX, eY = min(w - 1, eX), min(h - 1, eY)

            # Calculate Center and Dimensions
            cX, cY = int((sX + eX) / 2), int((sY + eY) / 2)
            fW, fH = eX - sX, eY - sY

            # --- Visualization ---
            color = (0, 255, 0) # Green
            # Draw Face Box
            cv2.rectangle(annotated_img, (sX, sY), (eX, eY), color, 2)
            # Draw Center Dot
            cv2.circle(annotated_img, (cX, cY), 5, color, -1)
            # Add Confidence Label
            label = f"{confidence:.2f}"
            y_label = sY - 10 if sY - 10 > 10 else sY + 10
            cv2.putText(annotated_img, label, (sX, y_label), 
                        cv2.FONT_HERSHEY_SIMPLEX, 0.45, color, 2)

            # --- Log to CSV ---
            writer.writerow([filename, cX, cY, fW, fH, sX, sY, eX, eY, f"{confidence:.4f}"])

# 6. Final Outputs
if faces_found:
    cv2.imwrite(output_img_path, annotated_img)
    print(f"Success: Detected faces saved to {output_img_path} and metadata to {log_file}")
else:
    print(f"No faces detected for {filename}.")
    # Clean up empty CSV to avoid confusing ImageMagick
    if os.path.exists(log_file):
        os.remove(log_file)
    sys.exit(0) # Exit cleanly so PowerShell can continue
