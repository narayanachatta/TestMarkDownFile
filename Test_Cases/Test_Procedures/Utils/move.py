
from pywinauto import mouse
import pyautogui
import time

def move_cursor_to_map_xy(map_x, map_y, center_x=1354, center_y=484, scale_x=5.86, scale_y=5.84):
    """
    Moves the cursor to a position on screen based on map coordinates.

    Parameters:
    - map_x, map_y: Coordinates in meters relative to the map center (0, 0)
    - center_x, center_y: Pixel coordinates of the nav map center
    - scale_x, scale_y: Meters per pixel for x and y directions
    """
    # Convert map coordinates to screen coordinates
    screen_x = center_x + (map_x / scale_x)
    screen_y = center_y - (map_y / scale_y)  # y is inverted on screen
    
    # Move the mouse to (screen_x, screen_y)
    mouse.move(coords=(int(screen_x), int(screen_y)))
    
    return (int(screen_x), int(screen_y))
    
def screen_coords_xy(map_x, map_y, center_x=1354, center_y=484, scale_x=5.86, scale_y=5.84):
    """
    Convert map coordinates to screen coordinates.

    Parameters:
    - map_x, map_y: Coordinates in meters relative to the map center (0, 0)
    - center_x, center_y: Pixel coordinates of the nav map center
    - scale_x, scale_y: Meters per pixel for x and y directions
    """
    # Convert map coordinates to screen coordinates
    screen_x = center_x + (map_x / scale_x)
    screen_y = center_y - (map_y / scale_y)  # y is inverted on screen
    
    return (int(screen_x), int(screen_y))
