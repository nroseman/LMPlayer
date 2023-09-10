# TODO: Clean up imports
from dotenv import load_dotenv
import os
import pyautogui
import win32gui
import win32ui
import win32con
import mss
import mss.tools
from pywinauto.application import Application
import cv2
import PIL
from PIL import Image
import pytesseract
import datetime as dt
import time


load_dotenv()
# path_to_tesseract = os.getenv("PATH_TESSERACT")
# pytesseract.pytesseract.tesseract_cmd = path_to_tesseract

target_window = win32gui.FindWindow(None, ('Lords Mobile: Kingdom Wars'))
app = Application()
app.connect(handle=target_window)


window_rect = app.window().child_window().rectangle()
window_region = (window_rect.left, window_rect.top, window_rect.right -
                 window_rect.left, window_rect.bottom - window_rect.top)
print(window_region)

auto_construct = False

collect_box = {
    "top": 805,
    "left": 1545,
    "width": 85,
    "height": 26
}
help_box = {
    "top": 750,
    "left": 1705,
    "width": 60,
    "height": 50
}
location_box = {
    "top": 875,
    "left": 50,
    "width": 145,
    "height": 85
}

change_screen_loc = (115, 920)

help_loc = (1730, 775)
helpall_loc = (920, 920)

exit_loc = (1715, 60)

quest_loc = (1125, 940)
admin_quest_loc = (885, 210)
guild_quest_loc = (1190, 215)
quest_complete_loc = (1545, 432)
quest_start_loc = (1573, 43)
quest_rgb = (255, 247, 153)

mystery_box_loc = (1603, 730)
mystery_claim_loc = (885, 750)


def send_help():
    # - Pyautogui - #
    print("Checking if help requested... ", end="")
    if pyautogui.locateOnScreen(
            'assets/help_button.PNG', region=window_region, grayscale=True, confidence=.9):
        print("Pressing Help... ", end="")
        app.window().click(
            button='left', coords=help_loc)
        time.sleep(.5)
        app.window().click(
            button='left', coords=helpall_loc)
        time.sleep(1)
        app.window().click(
            button='left', coords=exit_loc)
        time.sleep(1)
        print("DONE")
    else:
        print("No help detected")


# This needs work
def click_queue():
    print("Checking project queue... ", end="")
    if pyautogui.locateOnScreen(
            'assets/free_buton.PNG', region=window_region, grayscale=True, confidence=.9):
        print("Pressing Free... ", end="")
        app.window().click(
            button='left', coords=(275, 125))
        time.sleep(1)
        # TODO: Need to find a way to differentiate between construction, research, etc.
        # This only works for construction - will break eventually
        if auto_construct:
            print("Locating next project...")
            app.window().click(
                button='left', coords=(275, 125))
            time.sleep(5)
            app.window().click(
                button='left', coords=(468, 235))
            time.sleep(1)
            app.window().click(
                button='left', coords=(765, 445))
            time.sleep(1)
            app.window().click(
                button='left', coords=(570, 430))
            time.sleep(1)
            app.window().click(
                button='left', coords=(570, 430))
            time.sleep(1)
            app.window().click(
                button='left', coords=(860, 30))
            time.sleep(1)
            app.window().click(
                button='left', coords=(860, 30))
    else:
        print("Nothing completed.")


def check_quests():
    print("Checking for quests...", end=" ")
    app.window().click(
        button='left', coords=quest_loc)
    time.sleep(1)
    for quest in [admin_quest_loc, guild_quest_loc]:
        app.window().click(
            button='left', coords=quest)
        time.sleep(.5)
        if pyautogui.pixel(quest_complete_loc[0], quest_complete_loc[1]) == quest_rgb:
            print("Completing quest...", end=" ")
            app.window().click(
                button='left', coords=quest_complete_loc)
            time.sleep(2)
        if pyautogui.pixel(quest_start_loc[0], quest_start_loc[1]) == quest_rgb:
            print("Starting quest...", end=" ")
            app.window().click(
                button='left', coords=quest_start_loc)
            time.sleep(1)
    print("DONE")
    app.window().click(
        button='left', coords=exit_loc)


def check_box():
    print("checking if mystery box ready...", end=" ")
    # - screenshot and pytesseract maybe overly complicated... - #
    # take_screenshot(box=collect_box)
    # if read_image() == "Collect":

    # get the blue value from mystery box counter
    px_b = pyautogui.pixel(1627, 857)[2]
    if px_b >= 187 and px_b <= 193:  # range due to the color changing with background
        print("clicking mystery box...", end=" ")
        app.window().click(
            button='left', coords=mystery_box_loc)
        time.sleep(1)
        print("clicking claim...", end=" ")
        app.window().click(
            button='left', coords=mystery_claim_loc)
        time.sleep(3)
        print("DONE")
    else:
        print("Box not ready")


def get_location():
    if pyautogui.locateOnScreen(
            'assets/castle.PNG', region=window_region, grayscale=True, confidence=.9):
        return "Overworld"
    if pyautogui.locateOnScreen(
            'assets/map.PNG', region=window_region, grayscale=True, confidence=.9):
        return "Castle"
    return None


def go_to_location(loc):
    current_loc = get_location()
    if current_loc == None:
        print("Can't Navigate")
        return None
    elif current_loc == loc:
        return current_loc
    else:
        app.window().click(
            button='left', coords=change_screen_loc)
        time.sleep(2)
        return True


def take_screenshot(box=None):
    with mss.mss() as sct:
        # The screenshot of box
        if box != None:
            monitor = {
                "top": window_rect.top + box["top"],
                "left": window_rect.left + box["left"],
                "width": box["width"],
                "height": box["height"]
            }
        # screenshot of window_rect
        else:
            monitor = {
                "top": window_rect.top,
                "left": window_rect.left,
                "width": window_region[2],
                "height": window_region[3]
            }

        output = "./assets/window.PNG".format(
            **monitor)

        # Grab the data
        sct_img = sct.grab(monitor)

        # Save to PIL image
        img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")

        output = "./assets/collect.PNG"
        img.save(output)

        # Save to the picture file
        # mss.tools.to_png(sct_img.rgb, sct_img.size, output=output)
        # print(output)


def read_image():
    img_text = pytesseract.image_to_string(
        Image.open("./assets/collect.PNG")).strip()
    print(f"|{img_text}|")
    return img_text


# send_help()
# click_queue()
# take_screenshot()
# read_image()
# check_quests()
# check_box()
# go_to_location("Castle")
app.top_window().set_focus()
while True:
    print("New cycle begins")
    if go_to_location("Castle") == None:
        print("Waiting and will try again in 1 minute")
        time.sleep(60)
        continue
    check_box()
    send_help()
    check_quests()
    print("Cycle Done")
    time.sleep(300)
