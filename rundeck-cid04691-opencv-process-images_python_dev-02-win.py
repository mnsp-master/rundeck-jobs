#mnsp ver: 0.0.3_2_1

import cv2
import numpy as np
import os
import csv
import sys

# 1. Setup Inputs from SysArgs
if len(sys.argv) < 3:
    print("Usage: python script.py <input_image_path> <output_directory>")
    sys.exit(1)

img_path = sys.argv[1]
output_dir = sys.argv[2]
os.makedirs(output_dir, exist_ok=True)

if not os.path.isfile(img_path):
    print(f"Error: File {img_path} not found.")
    sys.exit(1)

# 2. Load the Image first (needed to initialize YuNet dimensions)
img = cv2.imread(img_path)
if img is None:
    print(f"Error: Could not read {img_path}")
    sys.exit(1)

h, w, _ = img.shape

# 3. Initialize YuNet
# Download face_detection_yunet_2023mar.onnx to your script folder
model_path = "face_detection_yunet_2023mar.onnx"
detector = cv2.FaceDetectorYN.create(
    model=model_path,
    config="",
    input_size=(w, h),
    score_threshold=0.5,  # Confidence threshold
    nms_threshold=0.3,    # Non-maximum suppression threshold
    top_k=5000
)

# 4. Define Output Filenames
filename = os.path.basename(img_path)
name_only = os.path.splitext(filename)[0]
log_file = os.path.join(output_dir, f"{name_only}.csv")
output_img_path = os.path.join(output_dir, f"detected_{filename}")

# 5. Process the Image
# YuNet returns (count, detections)
# detections format: [startX, startY, width, height, right_eye_x, right_eye_y, ...]
_, detections = detector.detect(img)

annotated_img = img.copy()
faces_found = False

# 6. Process Detections and Write CSV
with open(log_file, "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Filename", "CenterX", "CenterY", "Width", "Height", "StartX", "StartY", "EndX", "EndY", "Confidence"])

    if detections is not None:
        faces_found = True
        for det in detections:
            # YuNet bounding box is [x, y, w, h]
            sX, sY, fW, fH = map(int, det[:4])
            confidence = det[-1]
            
            eX, eY = sX + fW, sY + fH
            
            # Ensure coordinates stay within image boundaries
            sX, sY = max(0, sX), max(0, sY)
            eX, eY = min(w - 1, eX), min(h - 1, eY)

            # Calculate Center
            cX, cY = int(sX + (fW / 2)), int(sY + (fH / 2))

            # --- Visualization ---
            color = (0, 255, 0)
            cv2.rectangle(annotated_img, (sX, sY), (eX, eY), color, 2)
            cv2.circle(annotated_img, (cX, cY), 5, color, -1)
            
            label = f"{confidence:.2f}"
            y_label = sY - 10 if sY - 10 > 10 else sY + 10
            cv2.putText(annotated_img, label, (sX, y_label), 
                        cv2.FONT_HERSHEY_SIMPLEX, 0.45, color, 2)

            # --- Log to CSV ---
            writer.writerow([filename, cX, cY, fW, fH, sX, sY, eX, eY, f"{confidence:.4f}"])

# 7. Final Outputs
if faces_found:
    cv2.imwrite(output_img_path, annotated_img)
    print(f"Success: Detected faces saved to {output_img_path} and metadata to {log_file}")
else:
    print(f"No faces detected for {filename}.")
    if os.path.exists(log_file):
        os.remove(log_file)
    sys.exit(0)
