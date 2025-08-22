
from pywinauto import Application, mouse
from PIL import ImageGrab
import cv2
import numpy as np
import pyautogui
import time
import keyboard
from Utils import move

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

file_name.set_edit_text(r"C:\Users\Narayana.Chatta\Downloads\SC3_Local\Simulation\Mission_Plan_24Jul25.spx")
# Enter the text

# Get Open button label
open_button = upload_window_dlg.child_window(title="&Open", class_name="Button")

# click on Open button
open_button.click_input()

# Wait for the mission plan to load
time.sleep(2)

# Get the window
nav_map = main_window.child_window(auto_id="navMap", control_type="Gcs.FormControls.MapDisplay")

# Cartesian coordinates of Center of Nav Map
center_x = 1352
center_y = 484

# Move mouse to center and scroll
mouse.move(coords=(center_x, center_y))

# Click on Map Display to focus
nav_map.click_input()

# Zoom in Map using loop
for _ in range(75):
    mouse.scroll(coords=(center_x, center_y), wheel_dist=1)  # Zoom in
    time.sleep(0.1)

# wait for map to be settled
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

# Extract file path of original and small images
original_img_path = os.path.join(original_img_folder, "SCFlight_Window.png")
small_img_path  = os.path.join(test_images_dir, "NavMap_HomeUAV_DefaultLocation_1.png")

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

# Move to the ending point
(moved_x, moved_y) = move.move_cursor_to_map_xy(139, -607)

# Ensure Nav Map is in focus
pyautogui.click()

# Wait for cursor to settle at desired location
time.sleep(1)

# Press and hold Ctrl
keyboard.press('ctrl')

# Perform double click at a specific position
pyautogui.click()
time.sleep(0.05)
pyautogui.click()

# Release Ctrl
keyboard.release('ctrl')

# Wait for the Edit Waypoint window to appear and user to observe
time.sleep(2)

# Get window coordinates
rect = main_window.rectangle()
bbox = (rect.left, rect.top, rect.right, rect.bottom)

# Capture screenshot using PIL
img1 = ImageGrab.grab(bbox)
img1.save(os.path.join(original_img_folder, "SCFlight_Window.png"))

# Extract file path of original and small images
original_img_path = os.path.join(original_img_folder, "SCFlight_Window.png")
small_img_path  = os.path.join(test_images_dir, "NavMap_HotKeys_EditWaypoint_window.png")

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
Connect_win = Application(backend="win32").connect(title="Edit Waypoint")

# Get the window
edit_wpt_dlg = Connect_win.window(title="Edit Waypoint")

# Get the panel1 label
panel1_label = edit_wpt_dlg.child_window(auto_id="panel1", control_type="System.Windows.Forms.Panel")

# Get Loiter checkbox
loiter_chkbox = panel1_label.child_window(title="Loiter", auto_id="chkBoxLoiter", control_type="System.Windows.Forms.CheckBox")

# Select the checkbox
loiter_chkbox.click_input()

# Wait for value to be updated
time.sleep(0.5)

# Get X Coordinate textbox
x_text = panel1_label.child_window(auto_id="x", control_type="System.Windows.Forms.TextBox")

# Enter the value
x_text.set_edit_text("-190")

# Wait for value to be updated
time.sleep(0.5)
x_text.click_input()

# Get Y Coordinate textbox
y_text = panel1_label.child_window(auto_id="y", control_type="System.Windows.Forms.TextBox")

# Enter the value
y_text.set_edit_text("-2147")

# Wait for value to be updated
time.sleep(0.5)
y_text.click_input()

# Get Height textbox
height_text = panel1_label.child_window(auto_id="height", control_type="System.Windows.Forms.TextBox")

# Enter the value
height_text.set_edit_text("300")

# Wait for value to be updated
time.sleep(0.5)
height_text.click_input()

# Get the Radius label
radius_label = panel1_label.child_window(auto_id="txtBoxRadius", control_type="System.Windows.Forms.NumericUpDown")

# Get the Radius textbox
radius_txtbox = radius_label.child_window(control_type="System.Windows.Forms.UpDownBase+UpDownEdit")

# Enter the value
radius_txtbox.set_edit_text("270")

# Wait for value to be updated
time.sleep(0.5)
radius_txtbox.click_input()

# Get the Direction label
direction_label = panel1_label.child_window(title="Direction", auto_id="groupBox1", control_type="System.Windows.Forms.GroupBox")

# Get the counter clockwise radio button
counter_clk_btn = direction_label.child_window(title="Counter Clockwise", auto_id="radButtCCW", control_type="System.Windows.Forms.RadioButton")

# click on counter clockwise radio button
counter_clk_btn.click_input()

# Wait for value to be updated
time.sleep(0.5)
counter_clk_btn.click_input()

# Get the Duration label
dur_label = panel1_label.child_window(title="Duration", auto_id="groupBox2", control_type="System.Windows.Forms.GroupBox")

# Get the Minutes radio button
min_btn = dur_label.child_window(title="Minutes", auto_id="radButtSeconds", control_type="System.Windows.Forms.RadioButton")

# Click on Minutes radio button
min_btn.click_input()

# Wait for value to be updated
time.sleep(0.5)

# Get the Duration textbox
dur_txtbox = dur_label.child_window(control_type="System.Windows.Forms.UpDownBase+UpDownEdit")

# Enter the value
dur_txtbox.set_edit_text("1.1")

# Wait for value to be updated
time.sleep(0.5)
dur_txtbox.click_input()

# Get Set button
set_btn = edit_wpt_dlg.child_window(title="Set", auto_id="set_button", control_type="System.Windows.Forms.Button")

# Click on set button
set_btn.click_input()

# Wait for action to complete
time.sleep(2)

# Connect to the window by title
Connect_win = Application(backend="win32").connect(title="Change")

# Get the window
change_dlg = Connect_win.window(title="Change")

# Get the Yes button
yes_btn = change_dlg.child_window(title="&Yes", class_name="Button")

# Click on Yes button
yes_btn.click_input()

# Wait for action to complete
time.sleep(2)

# Get window coordinates
rect = main_window.rectangle()
bbox = (rect.left, rect.top, rect.right, rect.bottom)

# Capture screenshot using PIL
img1 = ImageGrab.grab(bbox)
img1.save(os.path.join(original_img_folder, "SCFlight_Window.png"))

# Extract file path of original and small images
original_img_path = os.path.join(original_img_folder, "SCFlight_Window.png")
small_img_path  = os.path.join(test_images_dir, "NavMap_HotKeys_WaypointEdit_1.png")

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

# Move to the ending point
(moved_x, moved_y) = move.move_cursor_to_map_xy(-1347, -1511)

# Ensure Nav Map is in focus
pyautogui.click()

# Wait for cursor to settle at desired location
time.sleep(1)

# Press and hold Ctrl
keyboard.press('ctrl')

# Perform double click at a specific position
pyautogui.click()
time.sleep(0.05)
pyautogui.click()

# Release Ctrl
keyboard.release('ctrl')

# Wait for the New Waypoint window to appear and user to observe
time.sleep(2)

# Get window coordinates
rect = main_window.rectangle()
bbox = (rect.left, rect.top, rect.right, rect.bottom)

# Capture screenshot using PIL
img1 = ImageGrab.grab(bbox)
img1.save(os.path.join(original_img_folder, "SCFlight_Window.png"))

# Extract file path of original and small images
original_img_path = os.path.join(original_img_folder, "SCFlight_Window.png")
small_img_path  = os.path.join(test_images_dir, "NavMap_HotKeys_NewWaypoint_window.png")

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
Connect_win = Application(backend="win32").connect(title="New Waypoint")

# Get the window
new_wpt_dlg = Connect_win.window(title="New Waypoint")

# Get the panel1 label
panel1_label = new_wpt_dlg.child_window(auto_id="panel1", control_type="System.Windows.Forms.Panel")

# Get Loiter checkbox
loiter_chkbox = panel1_label.child_window(title="Loiter", auto_id="chkBoxLoiter", control_type="System.Windows.Forms.CheckBox")

# Select the checkbox
loiter_chkbox.click_input()

# Wait for value to be updated
time.sleep(0.5)

# Get Insert Waypoint before label
insert_wpt_bfre = panel1_label.child_window(auto_id="targetWaypointIdBox", control_type="System.Windows.Forms.NumericUpDown")

# Get Insert Waypoint before textbox
insert_wpt_txtbox = insert_wpt_bfre. child_window(control_type="System.Windows.Forms.UpDownBase+UpDownEdit")

# Enter the value
insert_wpt_txtbox.set_edit_text("10")

# Wait for value to be updated
time.sleep(0.5)
insert_wpt_txtbox.click_input()

# Get X Coordinate textbox
x_text = panel1_label.child_window(auto_id="x", control_type="System.Windows.Forms.TextBox")

# Enter the value
x_text.set_edit_text("-1347")

# Wait for value to be updated
time.sleep(0.5)
x_text.click_input()

# Get Y Coordinate textbox
y_text = panel1_label.child_window(auto_id="y", control_type="System.Windows.Forms.TextBox")

# Enter the value
y_text.set_edit_text("-1511")

# Wait for value to be updated
time.sleep(0.5)
y_text.click_input()

# Get Height textbox
height_text = panel1_label.child_window(auto_id="height", control_type="System.Windows.Forms.TextBox")

# Enter the value
height_text.set_edit_text("280")

# Wait for value to be updated
time.sleep(0.5)
height_text.click_input()

# Get the Radius label
radius_label = panel1_label.child_window(auto_id="txtBoxRadius", control_type="System.Windows.Forms.NumericUpDown")

# Get the Radius textbox
radius_txtbox = radius_label.child_window(control_type="System.Windows.Forms.UpDownBase+UpDownEdit")

# Enter the value
radius_txtbox.set_edit_text("230")

# Wait for value to be updated
time.sleep(0.5)
radius_txtbox.click_input()

# Get the Direction label
direction_label = panel1_label.child_window(title="Direction", auto_id="groupBox1", control_type="System.Windows.Forms.GroupBox")

# Get the counter clockwise radio button
counter_clk_btn = direction_label.child_window(title="Counter Clockwise", auto_id="radButtCCW", control_type="System.Windows.Forms.RadioButton")

# click on counter clockwise radio button
counter_clk_btn.click_input()

# Wait for value to be updated
time.sleep(0.5)
counter_clk_btn.click_input()

# Get the Duration label
dur_label = panel1_label.child_window(title="Duration", auto_id="groupBox2", control_type="System.Windows.Forms.GroupBox")

# Get the Minutes radio button
laps_btn = dur_label.child_window(title="Laps", auto_id="radButtLaps", control_type="System.Windows.Forms.RadioButton")

# Click on Minutes radio button
laps_btn.click_input()

# Wait for value to be updated
time.sleep(0.5)

# Get the Duration textbox
dur_txtbox = dur_label.child_window(control_type="System.Windows.Forms.UpDownBase+UpDownEdit")

# Enter the value
dur_txtbox.set_edit_text("1.1")

# Wait Wait for value to be updated
time.sleep(0.5)
dur_txtbox.click_input()

# Get Set button
set_btn = new_wpt_dlg.child_window(title="Set", auto_id="set_button", control_type="System.Windows.Forms.Button")

# Click on set button
set_btn.click_input()

# Wait for action to complete
time.sleep(2)

# Connect to the window by title
Connect_win = Application(backend="win32").connect(title="Change")

# Get the window
change_dlg = Connect_win.window(title="Change")

# Get the Yes button
yes_btn = change_dlg.child_window(title="&Yes", class_name="Button")

# Click on Yes button
yes_btn.click_input()

# Wait for action to complete
time.sleep(2)

# Get window coordinates
rect = main_window.rectangle()
bbox = (rect.left, rect.top, rect.right, rect.bottom)

# Capture screenshot using PIL
img1 = ImageGrab.grab(bbox)
img1.save(os.path.join(original_img_folder, "SCFlight_Window.png"))

# Extract file path of original and small images
original_img_path = os.path.join(original_img_folder, "SCFlight_Window.png")
small_img_path  = os.path.join(test_images_dir, "NavMap_HotKeys_WaypointNew_1.png")

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

# Press and hold Ctrl
keyboard.press('ctrl')

# Move to the cursor desired coordinates
(moved_x, moved_y) = move.move_cursor_to_map_xy(-1343, -1009)

# Press and hold the right mouse button at the coordinates
mouse.press(button='right', coords=(moved_x, moved_y))

# Move to the ending point
(moved_x, moved_y) = move.move_cursor_to_map_xy(-834, -1467)

# Release the right mouse button
mouse.release(button='right', coords=(moved_x, moved_y))

# Release Ctrl
keyboard.release('ctrl')

# Wait for the Map to move and user to observe
time.sleep(3)

# Get window coordinates
rect = main_window.rectangle()
bbox = (rect.left, rect.top, rect.right, rect.bottom)

# Capture screenshot using PIL
img1 = ImageGrab.grab(bbox)
img1.save(os.path.join(original_img_folder, "SCFlight_Window.png"))

# Extract file path of original and small images
original_img_path = os.path.join(original_img_folder, "SCFlight_Window.png")
small_img_path  = os.path.join(test_images_dir, "NavMap_HotKeys_WaypointDrag_1.png")

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
