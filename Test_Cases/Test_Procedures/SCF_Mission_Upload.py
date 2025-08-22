
from pywinauto import Application, mouse
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

# Connect to the window by title
update_rate_window = Application(backend="win32").connect(title="Select Update Rate And Display Units")

# Get the window
update_rate_dlg = update_rate_window.window(title="Select Update Rate And Display Units")

# Get Fast button label
fast_button = update_rate_dlg.child_window(title="Fast", auto_id="button24", control_type="System.Windows.Forms.Button")

# Left click on Fast button
fast_button.click_input()

# Wait for left click to be performed
time.sleep(1)

# Get Link label
connect_dlg = main_window.child_window(auto_id="lightLink", control_type="Gcs.FormControls.IndicatorLight")

# Double click on Link label
connect_dlg.double_click_input()

# Wait for Double click to be performed
time.sleep(2)

# Connect to the window by title
Connect_win = Application(backend="win32").connect(title="Connect to Aircraft")

# Get the window
Connect_win_dlg = Connect_win.window(title="Connect to Aircraft")

# Find the textbox element
element = Connect_win_dlg.child_window(best_match="Edit").wrapper_object()

# Set the text
element.set_edit_text("172.22.13.39")

# Get Connect button
conn_button = Connect_win_dlg.child_window(title="Connect", auto_id="button_ok", control_type="System.Windows.Forms.Button")

# Click on Connect button
conn_button.click_input()

# Wait for the "Connect to Aircraft" window to be closed
time.sleep(2)

# Get the Upload button label
upload_button = main_window.child_window(title="Upload", auto_id="btnUpLoad", control_type="System.Windows.Forms.Button")

# Click on Upload button
upload_button.click_input()

# Wait for the Mission Upload window to appear
time.sleep(2)

# Get window coordinates
rect = main_window.rectangle()
bbox = (rect.left, rect.top, rect.right, rect.bottom)

# Create folder with file name under Result Image
original_img_folder = os.path.join (result_images_dir, file_name.split('.')[0])
os.makedirs(original_img_folder, exist_ok=True)

# Capture screenshot using PIL
img1 = ImageGrab.grab(bbox)
img1.save(os.path.join(original_img_folder, "SCFlight_Window.png"))

# Extract file path of original and small images
original_img_path = os.path.join(original_img_folder, "SCFlight_Window.png")
small_img_path  = os.path.join(test_images_dir, "Mission_Upload_Window.png")

# Load the original and small images
original_img = cv2.imread(original_img_path)
small_img = cv2.imread(small_img_path)

# Extract file names excluding path
original_img_name = original_img_path.split('\\')[-1]
small_img_name = small_img_path.split('\\')[-1]
name = small_img_name.split('.')[0]

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
        out.write('| *Figure1: Test Image for comparing results* |\n')

        out.write('----------------------------\n')
        out.write(f'**{small_img_name}** is *matched* with part of **{original_img_name}** below: \n\n')
        out.write('----------------------------\n')

        out.write(f'| **Title: {original_img_name}** |\n')
        out.write('| :---------------------------- |\n')
        out.write(f'| ![Result Image captured](../Result_Images/{file_name.split('.')[0]}/{original_img_name}) |\n')
        out.write('| *Figure2: Results Image captured to check Test Image* |\n')

        out.write('----------------------------\n')
        out.write(f'Matched part identical to **{small_img_name}** *highlighted* with polygon in **{'matched_region_highlighted_' + name + '.png'}** below: \n\n')
        out.write('----------------------------\n')

        out.write(f'| **Title: {'matched_region_highlighted_' + name + '.png'}** |\n')
        out.write('| :---------------------------- |\n')
        out.write(f'| ![Captured Image against Test Image](../Result_Images/{file_name.split('.')[0]}/{'matched_region_highlighted_' + name + '.png'}) |\n')
        out.write('| *Figure3: Test Image is identified and marked with polygon* |\n')

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

# Connect to the window by title
mission_upload_window = Application(backend="win32").connect(title="Mission Upload")

# Get the window
mission_upload_dlg = mission_upload_window.window(title="Mission Upload")

# Get the Upload File button
upload_button = mission_upload_dlg.child_window(title="Upload File", auto_id="btnLoadExistingPlan", control_type="System.Windows.Forms.Button")

# Click on Upload File button
upload_button.click_input()

# Wait for the Upload dialog to open
time.sleep(2)

# Connect to the window by title
upload_window = Application(backend="win32").connect(title="Open")

# Get the window
upload_window_dlg = upload_window.window(title="Open")

# Get File name: label
file_name = upload_window_dlg.child_window(class_name="Edit")

# Enter the text
file_name.set_edit_text(r"C:\Users\Narayana.Chatta\Downloads\SC3_Local\Simulation\Mission_Plan_24Jul25.spx")

# Get Open button label
open_button = upload_window_dlg.child_window(title="&Open", class_name="Button")

# click on Open button
open_button.click_input()

# Get the window
nav_map = main_window.child_window(auto_id="navMap", control_type="Gcs.FormControls.MapDisplay")

# Get the rectangle of the navMap control
rect = nav_map.rectangle()

# Calculate center of the control
center_x = 1352
center_y = 484

# Move mouse to center and scroll
mouse.move(coords=(center_x, center_y))

# Click on Map Display to focus
nav_map.click_input()

# Zoom in Map using loop
for _ in range(72):
    mouse.scroll(coords=(center_x, center_y), wheel_dist=1)  # Zoom in
    time.sleep(0.1)

# Wait for the Map to be settled
time.sleep(1)

# Get window coordinates
rect = main_window.rectangle()
bbox = (rect.left, rect.top, rect.right, rect.bottom)

# Capture screenshot using PIL
img1 = ImageGrab.grab(bbox)
img1.save(os.path.join(original_img_folder, "SCFlight_Window_With_MissionPlan.png"))

# Extract file path of original and small images
original_img_path = os.path.join(original_img_folder, "SCFlight_Window_With_MissionPlan.png")
small_img_path  = os.path.join(test_images_dir, "NavMap_With_Mission_Plan.png")

# Load the original and small images
original_img = cv2.imread(original_img_path)
small_img = cv2.imread(small_img_path)

# Extract file names excluding path
original_img_name = original_img_path.split('\\')[-1]
small_img_name = small_img_path.split('\\')[-1]
name = small_img_name.split('.')[0]

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
        out.write('| *Figure1: Test Image for comparing results* |\n')

        out.write('----------------------------\n')
        out.write(f'**{small_img_name}** is *matched* with part of **{original_img_name}** below: \n\n')
        out.write('----------------------------\n')

        out.write(f'| **Title: {original_img_name}** |\n')
        out.write('| :---------------------------- |\n')
        out.write(f'| ![Result Image captured](../Result_Images/{file_name.split('.')[0]}/{original_img_name}) |\n')
        out.write('| *Figure2: Results Image captured to check Test Image* |\n')

        out.write('----------------------------\n')
        out.write(f'Matched part identical to **{small_img_name}** *highlighted* with polygon in **{'matched_region_highlighted_' + name + '.png'}** below: \n\n')
        out.write('----------------------------\n')

        out.write(f'| **Title: {'matched_region_highlighted_' + name + '.png'}** |\n')
        out.write('| :---------------------------- |\n')
        out.write(f'| ![Captured Image against Test Image](../Result_Images/{file_name.split('.')[0]}/{'matched_region_highlighted_' + name + '.png'}) |\n')
        out.write('| *Figure3: Test Image is identified and marked with polygon* |\n')

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
