
from pywinauto import Application, mouse
from PIL import ImageGrab
import cv2
import numpy as np
import pyautogui
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
for _ in range(73):
    mouse.scroll(coords=(center_x, center_y), wheel_dist=1)  # Zoom in
    time.sleep(0.1)

# wait for map to be settled
time.sleep(1)

# Right click on Map
mouse.right_click(coords = (center_x, center_y))

# Locate the image on screen
location = pyautogui.locateOnScreen(os.path.join(test_images_dir, "MapMenu_ApplicationSettings.png"), confidence=0.85)

if location:
    center = pyautogui.center(location)
    pyautogui.click(center)
    print(f"Clicked at: {center}")

    # Wait for Application Settings window and user to observe
    time.sleep(3)

else:
    out.write("Menu item 'MapMenu->ApplicationSettings' not found.\n")
    # Close the application and terminate execution
    app.close
    sys.exit()

# Connect to the window by title
app_settings_window = Application(backend="win32").connect(title="Application Settings")

# Get the window
app_settings_dlg = app_settings_window.window(title="Application Settings")

# Get Battery Level checkbox
battery_checkbox = app_settings_dlg.child_window(auto_id="want_battery_as_percent_check", control_type="System.Windows.Forms.CheckBox")

# Uncheck the checkbox
battery_checkbox.click_input()

# Wait for user to observe
time.sleep(1)

# Check the checkbox to bring back to original value
battery_checkbox.click_input()

# Wait before next label
time.sleep(1)

# Get the Tile Cache directory
tile_cache_dir = app_settings_dlg.child_window(auto_id="tile_cache_directory_box", control_type="System.Windows.Forms.TextBox")

# Enter the text
tile_cache_dir.set_edit_text("abcdXXXX1234")

# Wait for user to observe
time.sleep(1)

# Enter the original text
tile_cache_dir.set_edit_text(r"C:\Users\Narayana.Chatta\Downloads\SC3_Local\Simulation\tilecache")

# Wait before next label
time.sleep(1)

# Get the Tile Providers label
tile_provider = app_settings_dlg.child_window(auto_id="groupBox2", control_type="System.Windows.Forms.GroupBox")

# Select OSM Terrain Radio button
tile_provider.child_window(auto_id="useOSMTerrain", control_type="System.Windows.Forms.RadioButton").click_input()

# Wait for user to observe
time.sleep(1)

# Select OSM Satellite Radio button
tile_provider.child_window(auto_id="useOSMSatellite", control_type="System.Windows.Forms.RadioButton").click_input()

# Wait for user to observe
time.sleep(1)

# Select Bing Radio button
tile_provider.child_window(auto_id="use_bing_radio", control_type="System.Windows.Forms.RadioButton").click_input()

# Wait for user to observe
time.sleep(1)

# Select Open Street Map Radio button
tile_provider.child_window(auto_id="use_openstreetmap_radio", control_type="System.Windows.Forms.RadioButton").click_input()

# Wait before next label
time.sleep(1)

# Get General section
general_label = app_settings_dlg.child_window(title="General", auto_id="groupBox1", control_type="System.Windows.Forms.GroupBox")

# Get the Map Opacity combo box
map_opacity = general_label.child_window(auto_id="default_map_image_opacity", control_type="System.Windows.Forms.ComboBox")

# Set opacity level to 80%
map_opacity.select("80%")

# Wait for user to observe
time.sleep(1)

# Set opacity level back to 60%
map_opacity.select("60%")

# Wait before next label
time.sleep(1)

# Get the Polygon Opacity combo box
poly_opacity = general_label.child_window(auto_id="PolygonOpacityPercent", control_type="System.Windows.Forms.NumericUpDown")

# Edit Polygon Opacity value to 30%
edit_poly_opacity = poly_opacity.child_window(control_type="System.Windows.Forms.UpDownBase+UpDownEdit")
edit_poly_opacity.set_edit_text("30")

# Wait for user to observe
time.sleep(1)

# Edit Polygon Opacity value back to 15%
edit_poly_opacity.set_edit_text("15")

# Wait before next label
time.sleep(1)

# Get the Plan Dir label
plan_dir = app_settings_dlg.child_window(title="Plan Directories", auto_id="groupBox3", control_type="System.Windows.Forms.GroupBox")

# Get the Include Sub-dir check box
inc_sub_dir = plan_dir.child_window(title="Include subdirectories", auto_id="subdirs_include", control_type="System.Windows.Forms.CheckBox")

# Unselect Include Sub-dir check box
inc_sub_dir.click_input()

# Wait for user to observe
time.sleep(1)

# Unselect Include Sub-dir check box
inc_sub_dir.click_input()

# Wait before next label
time.sleep(1)

# Get the Dir List
dir_list = plan_dir.child_window(auto_id="dir_list", control_type="System.Windows.Forms.ListBox")

# Get the items in list
items_list = dir_list.item_texts()

# Select one of the item available in list
dir_list.select(items_list[0])

# Get the Remove button
remove_btn = plan_dir.child_window(title="Remove", auto_id="remove_dir_button", control_type="System.Windows.Forms.Button")

# Wait for user to observe
time.sleep(1)

# Remove selected item
remove_btn.click_input()

# Wait for user to observe
time.sleep(1)

# Select one of the item available in list
dir_list.select(items_list[1])

# Wait for user to observe
time.sleep(1)

# Remove selected item
remove_btn.click_input()

# Wait before next label
time.sleep(1)

# Get the Add button
add_btn = plan_dir.child_window(title="Add", auto_id="add_dir_button", control_type="System.Windows.Forms.Button")

# Click add button
add_btn.click_input()

# Wait for user to observe
time.sleep(1)

# Connect to the window by title
inc_dir_window = Application(backend="win32").connect(title="Browse For Folder")

# Get the window
inc_dir_dlg = inc_dir_window.window(title="Browse For Folder")

# Access the TreeView control
dir_tree = inc_dir_dlg.child_window(class_name="SysTreeView32")

# Expand and select the desired folder
dir_tree.get_item(r"\Desktop").expand()
time.sleep(1)
dir_tree.get_item(r"\Desktop\This PC").expand()
time.sleep(1)
dir_tree.get_item(r"\Desktop\This PC\Windows (C:)").expand()
time.sleep(1)
dir_tree.get_item(r"\Desktop\This PC\Windows (C:)\Users").expand()
time.sleep(1)
dir_tree.get_item(r"\Desktop\This PC\Windows (C:)\Users\Narayan.Chatta").expand()
time.sleep(1)
dir_tree.get_item(r"\Desktop\This PC\Windows (C:)\Users\Narayan.Chatta\Downloads").expand()
time.sleep(1)
dir_tree.get_item(r"\Desktop\This PC\Windows (C:)\Users\Narayan.Chatta\Downloads\SC3_Local").expand()
time.sleep(1)
dir_tree.get_item(r"\Desktop\This PC\Windows (C:)\Users\Narayan.Chatta\Downloads\SC3_Local\Simulation").expand()
time.sleep(1)
dir_tree.get_item(r"\Desktop\This PC\Windows (C:)\Users\Narayan.Chatta\Downloads\SC3_Local\Simulation\Template").select()

# Get OK button
ok_btn = inc_dir_dlg.child_window(title="OK", class_name="Button")

# Click OK
ok_btn.click_input()

# Wait before next label
time.sleep(1)

# Click add button
add_btn.click_input()

# Wait for user to observe
time.sleep(1)

# Expand and select the desired folder
dir_tree.get_item(r"\Desktop").expand()
time.sleep(1)
dir_tree.get_item(r"\Desktop\This PC").expand()
time.sleep(1)
dir_tree.get_item(r"\Desktop\This PC\Windows (C:)").expand()
time.sleep(1)
dir_tree.get_item(r"\Desktop\This PC\Windows (C:)\Users").expand()
time.sleep(1)
dir_tree.get_item(r"\Desktop\This PC\Windows (C:)\Users\Narayan.Chatta").expand()
time.sleep(1)
dir_tree.get_item(r"\Desktop\This PC\Windows (C:)\Users\Narayan.Chatta\Downloads").expand()
time.sleep(1)
dir_tree.get_item(r"\Desktop\This PC\Windows (C:)\Users\Narayan.Chatta\Downloads\SC3_Local").expand()
time.sleep(1)
dir_tree.get_item(r"\Desktop\This PC\Windows (C:)\Users\Narayan.Chatta\Downloads\SC3_Local\Simulation").select()

# Click OK
ok_btn.click_input()

# Wait before next label
time.sleep(1)

# Get LRTA label
lrta_label = app_settings_dlg.child_window(title="Long Range Tracking Antenna (LRTA)", auto_id="groupBox4", control_type="System.Windows.Forms.GroupBox")

# Get Antenna Director Select button
ant_dir_slct = lrta_label.child_window(auto_id="enable_antenna_director_passthrough_checkbox", control_type="System.Windows.Forms.CheckBox")

# Check the checkbox
ant_dir_slct.click_input()

# Wait for user to observe
time.sleep(1)

# Get Antenna Director Text box
ant_dir_txt_box = lrta_label.child_window(auto_id="antenna_director_text_box", control_type="System.Windows.Forms.TextBox")

# Edit the text box
ant_dir_txt_box.set_edit_text(r"ABCD1234YYYY")

# Wait for user to observe
time.sleep(1)

# Edit the text box again for original value
ant_dir_txt_box.set_edit_text(r"C:\Users\Narayana.Chatta\AppData\Local\Temp\AntennaDirector\coords.csv")

# Wait for user to observe
time.sleep(1)

# Uncheck the checkbox
ant_dir_slct.click_input()

# Wait before next label
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
small_img_path  = os.path.join(test_images_dir, "NavMap_ApplicationSettings.png")

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


# Test Result
out.write('----------------------------\n')
out.write('**Test Result**: *PASS*\n')
out.write('----------------------------\n')

# Close the SC Flight
app.kill()
