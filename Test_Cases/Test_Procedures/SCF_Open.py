
from pywinauto import Application
from PIL import ImageGrab
import cv2
import numpy as np
import time
import os

# Get current file name with full path
file_full_path = os.path.abspath(__file__)

# Get path of current file
test_dir = os.path.abspath(os.path.dirname(__file__))

# Get path of results
results_dir = os.path.join(test_dir, r"..\Test_Results")

# Get path os test images
test_images_dir = os.path.join(test_dir, r"..\Test_Images")

# Get path os result images
result_images_dir = os.path.join(test_dir, r"..\Result_Images")

# Get just the file name
file_name = os.path.basename(file_full_path)

# Create output markdown file
out_file = file_name.split('.')[0] + '.md'

# Get result file name with full path
full_out_file = os.path.join(results_dir, out_file)

# Open out_file for test reports
out =  open(full_out_file, "w+")
out.write("# Test Report\n")

# Open SC Flight application
app = Application(backend="win32").start(r"C:\Users\Narayana.Chatta\Downloads\SC3_Local\Simulation\flight 30\flight.exe")
main_window = app.window(title="Flight [v0.4.6-rc.3]")

# Wait for the application to open
time.sleep(2)

# Get window coordinates
rect = main_window.rectangle()
bbox = (rect.left, rect.top, rect.right, rect.bottom)

# Capture screenshot using PIL
img1 = ImageGrab.grab(bbox)
img1.save(os.path.join(result_images_dir, "SCFlight_With_UpdateRate_Window.png"))

# Extract file path of original and small images
original_img_path = os.path.join(result_images_dir, "SCFlight_With_UpdateRate_Window.png")
small_img_path  = os.path.join(test_images_dir, "UpdateRate_Window.png")

# Load the original and small images
original_img = cv2.imread(original_img_path)
small_img = cv2.imread(small_img_path)

# Extract file names excluding path
original_img_name = original_img_path.split('\\')[-1]
small_img_name = small_img_path.split('\\')[-1]

# Convert to grayscale
gray_original = cv2.cvtColor(original_img, cv2.COLOR_BGR2GRAY)
gray_small = cv2.cvtColor(small_img, cv2.COLOR_BGR2GRAY)

# Initialize SIFT detector
sift = cv2.SIFT_create()

# Detect keypoints and descriptors
kp1, des1 = sift.detectAndCompute(gray_small, None)
kp2, des2 = sift.detectAndCompute(gray_original, None)

# Match descriptors using FLANN matcher
FLANN_INDEX_KDTREE = 1
index_params = dict(algorithm=FLANN_INDEX_KDTREE, trees=5)
search_params = dict(checks=50)

flann = cv2.FlannBasedMatcher(index_params, search_params)
matches = flann.knnMatch(des1, des2, k=2)

# Apply Lowe's ratio test
good_matches = []
for m, n in matches:
    if m.distance < 0.7 * n.distance:
        good_matches.append(m)

# Proceed if enough good matches are found
if len(good_matches) > 10:
    src_pts = np.float32([kp1[m.queryIdx].pt for m in good_matches]).reshape(-1, 1, 2)
    dst_pts = np.float32([kp2[m.trainIdx].pt for m in good_matches]).reshape(-1, 1, 2)

    # Compute homography
    M, mask = cv2.findHomography(src_pts, dst_pts, cv2.RANSAC, 5.0)

    if M is not None:    
        # Get the bounding box from the small image
        h, w = gray_small.shape
        pts = np.float32([[0, 0], [0, h], [w, h], [w, 0]]).reshape(-1, 1, 2)
        dst = cv2.perspectiveTransform(pts, M)
    
        # Draw polygon on the original image
        matched_img = original_img.copy()
        cv2.polylines(matched_img, [np.int32(dst)], True, (0, 255, 0), 3, cv2.LINE_AA)
    
        # Save the highlighted image
        cv2.imwrite(os.path.join(result_images_dir, 'matched_region_highlighted.png'), matched_img)
        #print(small_img_name, "is matched with part of", original_img_name, "\n")
        #print("Matched part highlighted with polygon in matched_region_highlighted.png\n")
        
        out.write(f"{small_img_name} is matched with part of {original_img_name}\n")
        out.write(f"Matched part highlighted with polygon in matched_region_highlighted.png\n")
        out.write(f"![Test Image]({r"..\Test_Images\UpdateRate_Window.png"})\n")
        out.write(f"Figure1: Test Image for comparing results\n")
        out.write(f"![Result Image captured]({os.path.join(result_images_dir, 'matched_region_highlighted.png')})\n")
        out.write(f"Figure2: Results Image captured to compare with Test Image\n")
        #out.write(f"[{small_img_name}]({os.path.join(test_images_dir, 'matched_region.png')}) is matched with part of [{original_img_name}]({os.path.join(result_images_dir, 'matched_region_highlighted.png')})\n")
    
        # Warp the matched region to extract it
        extracted = cv2.warpPerspective(original_img, np.linalg.inv(M), (w, h))
        cv2.imwrite(os.path.join(result_images_dir, 'matched_region.png'), extracted)
        print("Matched part only captured in matched_region.png\n")
        
    else:
        print("Homography could not be computed.")

else:
    print("Not enough matches found - {}/10".format(len(good_matches)))

# Close the SC Flight
app.kill()
