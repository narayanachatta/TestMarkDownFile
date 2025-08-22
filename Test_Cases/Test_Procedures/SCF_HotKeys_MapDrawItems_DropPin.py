
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
small_img_path  = os.path.join(test_images_dir, "NavMap_HomeUAV_DefaultLocation.png")

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
(moved_x, moved_y) = move.move_cursor_to_map_xy(-425, 1292)

# Ensure Nav Map is in focus
pyautogui.click()

# Wait for cursor to settle at desired location
time.sleep(1)

# Press I on keyboard
keyboard.send('i')

# Wait for the PinSpotForm window to appear and user to observe
time.sleep(2)

# Get window coordinates
rect = main_window.rectangle()
bbox = (rect.left, rect.top, rect.right, rect.bottom)

# Capture screenshot using PIL
img1 = ImageGrab.grab(bbox)
img1.save(os.path.join(original_img_folder, "SCFlight_Window.png"))

# Extract file path of original and small images
original_img_path = os.path.join(original_img_folder, "SCFlight_Window.png")
small_img_path  = os.path.join(test_images_dir, "NavMap_HotKeys_PinSpotForm_window.png")

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
Connect_win = Application(backend="win32").connect(title="PinSpotForm")

# Get the window
pinspotform_dlg = Connect_win.window(title="PinSpotForm")

# Get Pin Name textbox
pinname_txtbox = pinspotform_dlg.child_window(auto_id="pinNameBox1", control_type="System.Windows.Forms.TextBox")

# Enter the text
pinname_txtbox.set_edit_text("TestPin1")

# Wait for value to be updated
time.sleep(0.5)

# Get the Pin Color button
pincolor_btn = pinspotform_dlg.child_window(auto_id="btnPinColor", control_type="System.Windows.Forms.Button")

# Click the button
pincolor_btn.click_input()

# Wait for the action to perform
time.sleep(0.5)

# Connect to the window by title
Connect_win = Application(backend="win32").connect(title="Colour")

# Get the window
colour_dlg = Connect_win.window(title="Colour")

# Get the Basic Colous dialog
basic_colours_grid = colour_dlg.child_window(class_name="Static", found_index=1)

""" Below are the coordinates (Column, Row) for each colour box under Basic colours section:
(15, 15) (15, 41) (15, 69) (15, 95) (15, 123) (15, 151)
(46, 15) (46, 41) (46, 69) (46, 95) (46, 123) (46, 151)
(81, 15) (81, 41) (81, 69) (81, 95) (81, 123) (81, 151)
(114, 15) (114, 41) (114, 69) (114, 95) (114, 123) (114, 151)
(146, 15) (146, 41) (146, 69) (146, 95) (146, 123) (146, 151)
(179, 15) (179, 41) (179, 69) (179, 95) (179, 123) (179, 151)
(211, 15) (211, 41) (211, 69) (211, 95) (211, 123) (211, 151)
(245, 15) (245, 41) (245, 69) (245, 95) (245, 123) (245, 151) """

# Select the colour grid
basic_colours_grid.click_input(coords=(146, 95))

# Wait for the action to perform
time.sleep(0.5)

# Get OK button of Colour window
ok_btn = colour_dlg.child_window(title="OK", class_name="Button")

# Click on OK button
ok_btn.click_input()

# Wait for the Colour window to close
time.sleep(1)

# Get OK button of PinSpotForm window
ok_btn1 = pinspotform_dlg.child_window(title="Ok", auto_id="btnOk", control_type="System.Windows.Forms.Button")

# Click on OK button
ok_btn1.click_input()

# Wait for the PinSpotForm window to close
time.sleep(1)

# Get window coordinates
rect = main_window.rectangle()
bbox = (rect.left, rect.top, rect.right, rect.bottom)

# Capture screenshot using PIL
img1 = ImageGrab.grab(bbox)
img1.save(os.path.join(original_img_folder, "SCFlight_Window.png"))

# Extract file path of original and small images
original_img_path = os.path.join(original_img_folder, "SCFlight_Window.png")
small_img_path  = os.path.join(test_images_dir, "NavMap_HotKeys_PinSpotForm_1.png")

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
(moved_x, moved_y) = move.move_cursor_to_map_xy(-1310, 861)

# Ensure Nav Map is in focus
pyautogui.click()

# Wait for cursor to settle at desired location
time.sleep(1)

# Press I on keyboard
keyboard.send('i')

# Wait for the PinSpotForm window to appear and user to observe
time.sleep(2)

# Connect to the window by title
Connect_win = Application(backend="win32").connect(title="PinSpotForm")

# Get the window
pinspotform_dlg = Connect_win.window(title="PinSpotForm")

# Get Pin Name textbox
pinname_txtbox = pinspotform_dlg.child_window(auto_id="pinNameBox1", control_type="System.Windows.Forms.TextBox")

# Enter the text
pinname_txtbox.set_edit_text("TestPin2")

# Wait for value to be updated
time.sleep(0.5)

# Get the Pin Color button
pincolor_btn = pinspotform_dlg.child_window(auto_id="btnPinColor", control_type="System.Windows.Forms.Button")

# Click the button
pincolor_btn.click_input()

# Wait for the action to perform
time.sleep(0.5)

# Connect to the window by title
Connect_win = Application(backend="win32").connect(title="Colour")

# Get the window
colour_dlg = Connect_win.window(title="Colour")

# Get the Custom Colous dialog button
custm_colours_btn = colour_dlg.child_window(title="Define Custo&m Colours >>", class_name="Button")

# Select the colour grid
custm_colours_btn.click_input()

# Wait for the action to perform
time.sleep(0.5)

# Get Hue field and enter value
colour_dlg.child_window(class_name="Edit", found_index=0).set_edit_text("160")

# Get Sat field and enter value
colour_dlg.child_window(class_name="Edit", found_index=1).set_edit_text("114")

# Get Lum field and enter value
colour_dlg.child_window(class_name="Edit", found_index=2).set_edit_text("141")

# Get Red field and enter value
colour_dlg.child_window(class_name="Edit", found_index=3).set_edit_text("100")

# Get Green field and enter value
colour_dlg.child_window(class_name="Edit", found_index=4).set_edit_text("150")

# Get Blue field and enter value
colour_dlg.child_window(class_name="Edit", found_index=5).set_edit_text("200")

# Wait for value to be updated
time.sleep(0.5)

# Get Add to Custom Colours button
custm_btn = colour_dlg.child_window(title="&Add to Custom Colours", class_name="Button")

# Click on button
custm_btn.click_input()

# Wait for the action to perform
time.sleep(0.5)

# Get the Basic Colous dialog
custm_colours_grid = colour_dlg.child_window(class_name="Static", found_index=3)

""" Below are the coordinates (Column, Row) for each colour box under Custom colours section:
(15, 15) (15, 41) (15, 69) (15, 95) (15, 123) (15, 151)
(46, 15) (46, 41) (46, 69) (46, 95) (46, 123) (46, 151) """

# Select the colour grid
custm_colours_grid.click_input(coords=(15, 15))

# Get OK button of Colour window
ok_btn = colour_dlg.child_window(title="OK", class_name="Button")

# Click on OK button
ok_btn.click_input()

# Wait for the Colour window to close
time.sleep(1)

# Get OK button of PinSpotForm window
ok_btn1 = pinspotform_dlg.child_window(title="Ok", auto_id="btnOk", control_type="System.Windows.Forms.Button")

# Click on OK button
ok_btn1.click_input()

# Wait for the PinSpotForm window to close
time.sleep(1)

# Get window coordinates
rect = main_window.rectangle()
bbox = (rect.left, rect.top, rect.right, rect.bottom)

# Capture screenshot using PIL
img1 = ImageGrab.grab(bbox)
img1.save(os.path.join(original_img_folder, "SCFlight_Window.png"))

# Extract file path of original and small images
original_img_path = os.path.join(original_img_folder, "SCFlight_Window.png")
small_img_path  = os.path.join(test_images_dir, "NavMap_HotKeys_PinSpotForm_2.png")

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
