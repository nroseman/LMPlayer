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
from PIL import Image, ImageGrab
import pytesseract
import datetime as dt
import time
from time import strftime

load_dotenv()
# path_to_tesseract = os.getenv("PATH_TESSERACT")
# pytesseract.pytesseract.tesseract_cmd = path_to_tesseract

target_window = win32gui.FindWindow(None, ('Lords Mobile: Kingdom Wars'))
app = Application()
app.connect(handle=target_window)


window_rect = app.window().child_window().rectangle()
window_region = (window_rect.left, window_rect.top, window_rect.right -
                 window_rect.left, window_rect.bottom - window_rect.top)
print(window_region)  # must be (70, 31, 1780, 1001)

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

yellow_rgb = (255, 247, 153)
red_rgb = (214, 48, 49)

change_screen_loc = (115, 920)

help_loc = (1730, 775)
helpall_loc = (920, 920)

exit_loc = (1715, 60)

# TODO: are these necessary?
quest_loc = (1125, 940)
quest_notification_rgb = (214, 48, 49)
admin_quest_notification_loc = ()
admin_quest_loc = (885, 210)
guild_quest_loc = (1190, 215)
guild_quest_notification_loc = ()
quest_complete_loc = (1545, 432)
quest_start_loc = (1568, 426)
quest_rgb = (255, 247, 153)
quest_auto_loc = ()

quests = {
    "admin": {
        "notification_rgb": red_rgb,
        "notification_loc": (1014, 167),
        "tab_loc": (885, 210),
        "collect_loc": (1577, 391),
        "start_loc": (1560, 391),
        "auto_loc": (1625, 400),
        "quest_rgb": yellow_rgb,
        "is_auto": True,
        "total_quests": 7,
        "timer": 0

    },
    "guild": {
        "notification_rgb": red_rgb,
        "notification_loc": (1332, 167),
        "tab_loc": (1190, 215),
        "collect_loc": (1577, 391),
        "start_loc": (1560, 391),
        "auto_loc": (1625, 400),
        "quest_rgb": yellow_rgb,
        "is_auto": False,
        "total_quests": 7,
        "timer": 0
    }
}

mystery_box_loc = (1603, 730)
mystery_claim_loc = (885, 750)

guild_loc = (825, 940)
guild_gift_rgb = (255, 77, 36)
guild_tab_loc = (1650, 260)
guild_gift_loc = (1345, 385)
bonus_chest_tab_loc = (1420, 320)
guild_gift_tab_loc = (940, 340)
open_chest_loc = (1540, 440)
open_chest_px = (1575, 466)
open_chest_rgb = (255, 255, 255)

# Are admin/guild quests auto-complete through VIP points
is_admin_auto = True
is_guild_auto = False

timer_1 = 0
timer_5 = 0


def send_help():
    # - Pyautogui - #
    print("Checking if help requested... ", end="")
    if pyautogui.locateOnScreen(
            'assets/help_button.PNG', region=window_region, grayscale=True, confidence=.9):
        print("Pressing help... ", end="")
        app.window().click(
            button='left', coords=help_loc)
        time.sleep(.5)
        app.window().click(
            button='left', coords=helpall_loc)
        time.sleep(1)
        app.window().click(
            button='left', coords=exit_loc)
    else:
        print("No help detected", end=" ", flush=True)
    print("DONE")


# This needs work
def click_queue():
    print("Checking project queue... ", end="", flush=True)
    if pyautogui.locateOnScreen(
            'assets/free_buton.PNG', region=window_region, grayscale=True, confidence=.9):
        print("Pressing Free... ", end="", flush=True)
        app.window().click(
            button='left', coords=(275, 125))
        time.sleep(1)
        # TODO: Need to find a way to differentiate between construction, research, etc.
        # This only works for construction - will break eventually
        if auto_construct:
            print("Locating next project...", flush=True)
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
    print("Checking for quests...", end=" ", flush=True)
    app.window().click(
        button='left', coords=quest_loc)
    time.sleep(2)
    screen = ImageGrab.grab(bbox=(window_rect.left, window_rect.top,
                                  window_rect.right, window_rect.bottom))
    for quest in quests:
        current = quests[quest]
        if screen.getpixel(current["notification_loc"])[0] > 200:  # if it's red
            app.window().click(button='left', coords=current["tab_loc"])
            time.sleep(1)
            if current["is_auto"] == True:
                print(f"{quest} are auto complete...", end=" ", flush=True)
                for x in range(current["total_quests"]):
                    app.window().click(button='left',
                                       coords=current["auto_loc"])
                    time.sleep(.5)
                time.sleep(1)
                print("spam clicking is done")

            # Manually start/complete quests
            # collect button
            elif screen.getpixel(current["collect_loc"]) == current["quest_rgb"]:
                print("Starting quest...", end=" ", flush=True)
                app.window().click(button='left', coords=current["auto_loc"])
                time.sleep(1)
            else:  # Start button
                print("Completing quest and starting next...", end=" ", flush=True)
                app.window().click(button='left',
                                   coords=current["auto_loc"])  # click complete
                time.sleep(1)
                # click start on next quest
                app.window().click(button='left', coords=current["auto_loc"])
                time.sleep(1)
    app.window().click(button='left', coords=exit_loc)
    print("DONE")
    time.sleep(1)


def check_box():
    print("checking if mystery box ready...", end=" ", flush=True)
    # get the blue value from mystery box counter
    px_b = pyautogui.pixel(1627, 857)[2]
    if px_b >= 187 and px_b <= 193:  # range due to the color changing with background
        print("clicking mystery box...", end=" ", flush=True)
        app.window().click(
            button='left', coords=mystery_box_loc)
        time.sleep(1)
        print("clicking claim...", end=" ", flush=True)
        app.window().click(
            button='left', coords=mystery_claim_loc)
        time.sleep(3)
        print("DONE")
    else:
        print("Box not ready")


# TODO: Get more data points
# Does it default to bonus gifts always?
# multiple gifts?
def check_guild_gift():
    print("Checking for guild gifts...", end=" ", flush=True)
    time.sleep(1)
    px1 = pyautogui.pixel(845, 945)
    if px1 == guild_gift_rgb:
        app.window().click(
            button='left', coords=guild_loc)
        time.sleep(1)
        app.window().click(
            button='left', coords=guild_tab_loc)
        time.sleep(1)
        app.window().click(
            button='left', coords=guild_gift_loc)
        time.sleep(1)
        app.window().click(
            button='left', coords=guild_tab_loc)
        time.sleep(1)
        app.window().click(
            button='left', coords=bonus_chest_tab_loc)
        for loc in [bonus_chest_tab_loc, guild_gift_tab_loc]:
            app.window().click(
                button='left', coords=loc)
            px2 = pyautogui.pixel(open_chest_px[0], open_chest_px[1])
            if px2 == open_chest_rgb:
                print("Received gift...", end=" ", flush=True)
                app.window().click(
                    button='left', coords=open_chest_loc)
                time.sleep(1)
                app.window().click(
                    button='left', coords=open_chest_loc)
                time.sleep(1)
        app.window().click(
            button='left', coords=exit_loc)
        time.sleep(1)
        app.window().click(
            button='left', coords=exit_loc)
    print("DONE")
    time.sleep(1)


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
# check_guild_gift()
# pyautogui.displayMousePosition()

# - Main Loop - #
app.top_window().set_focus()  # Window needs to be visible for pyautogui
while True:
    start = time.perf_counter()

    print(f"New cycle begin - {time.strftime('%H:%M:%S')}:")

    last_location = get_location()

    if go_to_location("Castle") == None:
        print("Waiting and will try again in 1 minute")
        time.sleep(60)
        continue

    if timer_1 - start <= 0:
        check_box()
        timer_1 = start + 60

    if timer_5 - start <= 0:
        send_help()
        check_quests()
        check_guild_gift()
        timer_5 = start + 300

    go_to_location(last_location)

    print(f"Cycle Done - {time.strftime('%H:%M:%S')}")
    print("Next cycle in 1 minutes")

    time.sleep(60)

# img = ImageGrab.grab(bbox=(window_rect.left, window_rect.top,
#                      window_rect.right, window_rect.bottom))
# px = img.getpixel(quests["guild"]["notification_loc"])
# print(px)
