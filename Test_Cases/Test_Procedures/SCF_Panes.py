
from pywinauto import Application
from PIL import ImageGrab
import cv2
import numpy as np
import time
import os
import sys

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
out.write("# **Test Report**\n")

# Open SC Flight application
app = Application(backend="win32").start(r"C:\Users\Narayana.Chatta\Downloads\SC3_Local\Simulation\flight 30\flight.exe")
main_window = app.window(title="Flight [v0.4.6-rc.3]")

# Wait for the application to open
time.sleep(2)

# Select Fast option from "Select Update Rate And Display Units" window to get required screenshot
# Connect to the window by title
rate_display_window = Application(backend="win32").connect(title="Select Update Rate And Display Units")

# Get the window
dlg = rate_display_window.window(title="Select Update Rate And Display Units")

# Identify the Fast button label
fast_button = dlg.child_window(title="Fast", auto_id="button24", control_type="System.Windows.Forms.Button")

# Perform mouse left click
fast_button.click_input()

# Wait for the window to close
time.sleep(1)

# Get window coordinates
rect = main_window.rectangle()
bbox = (rect.left, rect.top, rect.right, rect.bottom)

# Create folder with file name under Result Image
original_img_folder = os.path.join (result_images_dir, file_name.split('.')[0])
os.makedirs(original_img_folder, exist_ok=True)

# Capture screenshot using PIL
img1 = ImageGrab.grab(bbox)
img1.save(os.path.join(original_img_folder, "SCFlight_Window.png"))

# Extract file path of original image
original_img_path = os.path.join(original_img_folder, "SCFlight_Window.png")

# Load the original
original_img = cv2.imread(original_img_path)

# Extract file names excluding path
original_img_name = original_img_path.split('\\')[-1]

# Initialize SIFT detector
sift = cv2.SIFT_create()
# Match descriptors using FLANN matcher
FLANN_INDEX_KDTREE = 1
index_params = dict(algorithm=FLANN_INDEX_KDTREE, trees=5)
search_params = dict(checks=50)
flann = cv2.FlannBasedMatcher(index_params, search_params)

# Convert to grayscale
gray_original = cv2.cvtColor(original_img, cv2.COLOR_BGR2GRAY)

# Detect keypoints and descriptors
kp2, des2 = sift.detectAndCompute(gray_original, None)

# Images in a list to compare
file_names = [os.path.join(test_images_dir, "RightPane_Map.png"), os.path.join(test_images_dir, "LeftPane_PanelControl.png"), os.path.join(test_images_dir, "LeftPane_PFI.png")]
good_matches = []
i = 0

for filename in file_names:
    # Extract file namly without extension
    small_img_name = file_names[i].split('\\')[-1]
    name = (file_names[i].split('\\')[-1]).split('.')[0]
    i = i+1

    # Load the small image
    small_img = cv2.imread(filename)

    # Convert to grayscale
    gray_small = cv2.cvtColor(small_img, cv2.COLOR_BGR2GRAY)

    # Detect keypoints and descriptors
    kp1, des1 = sift.detectAndCompute(gray_small, None)

    matches = flann.knnMatch(des1, des2, k=2)

    # Apply Lowe's ratio test
    for m, n in matches:
        if m.distance < 0.6 * n.distance:
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
            cv2.imwrite(os.path.join(original_img_folder, 'matched_region_highlighted_' + name + '.png'), matched_img)
            
            out.write(f'| **Title: {small_img_name}** |\n')
            out.write('| :---------------------------- |\n')
            out.write(f'| ![Test Image](../Test_Images/{small_img_name}) |\n')
            out.write(f'| *Figure{i}: Test Image for comparing results* |\n')

            out.write('----------------------------\n')
            out.write(f'**{small_img_name}** is *matched* with part of **{original_img_name}** below: \n\n')
            out.write('----------------------------\n')

            out.write(f'| **Title: {original_img_name}** |\n')
            out.write('| :---------------------------- |\n')
            out.write(f'| ![Result Image captured](../Result_Images/{file_name.split('.')[0]}/{original_img_name}) |\n')
            out.write(f'| *Figure{i+1}: Results Image captured to check Test Image* |\n')

            out.write('----------------------------\n')
            out.write(f'Matched part identical to **{small_img_name}** *highlighted* with polygon in **{'matched_region_highlighted_' + name + '.png'}** below: \n\n')
            out.write('----------------------------\n')

            out.write(f'| **Title: {'matched_region_highlighted_' + name + '.png'}** |\n')
            out.write('| :---------------------------- |\n')
            out.write(f'| ![Captured Image against Test Image](../Result_Images/{file_name.split('.')[0]}/{'matched_region_highlighted_' + name + '.png'}) |\n')
            out.write(f'| *Figure{i+2}: Test Image is identified and marked with polygon* |\n')

        else:
            out.write('Homography could not be computed.\n')
            out.write('----------------------------\n')
            out.write('**Test Result**: *FAIL*\n')
            out.write('----------------------------\n')
            # Close the application and terminate execution
            app.close
            sys.exit()

    else:
        out.write("Not enough matches found - {}/10".format(len(good_matches)))
        out.write('\n----------------------------\n')
        out.write('**Test Result**: *FAIL*\n')
        out.write('----------------------------\n')
        # Close the application and terminate execution
        app.close
        sys.exit()

# Test Result        
out.write('----------------------------\n')
out.write('**Test Result**: *PASS*\n')
out.write('----------------------------\n')

# Close the SC Flight
app.kill()
