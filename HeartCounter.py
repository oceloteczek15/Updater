import win32api
import win32gui
import win32con
import colorsys
import time
from time import sleep
import win32com.client as comctl
from colorama import init, Fore, Style

init(autoreset=True)  # Initialize colorama to automatically reset the console colors

wsh = comctl.Dispatch("WScript.Shell")

def set_console_always_on_top():
    hwnd = win32gui.GetForegroundWindow()
    screen_width = win32api.GetSystemMetrics(0)
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, screen_width-400, 0, 400, 300, win32con.SWP_SHOWWINDOW)

def print_chromatic_text(text):
    reset_color = Style.RESET_ALL
    chromatic_text = ""

    # Generate colors with hues between pink and purple
    pink_hue = 320 / 360  # Pink hue value (320 degrees)
    purple_hue = 270 / 360  # Purple hue value (270 degrees)

    for i, char in enumerate(text):
        if char in ("╔", "╚", "═", "╝", "╔", "╗", "║"):
            chromatic_text += Fore.BLUE + char
        else:
            hue = pink_hue + (purple_hue - pink_hue) * i / len(text)
            r, g, b = [int(255 * c) for c in colorsys.hsv_to_rgb(hue, 1.0, 1.0)]
            chromatic_text += f"\033[38;2;{r};{g};{b}m{char}"

    print(chromatic_text + reset_color)

def add_one_and_print(number, mode):
    set_console_always_on_top()
    sleep(3)
    print_chromatic_text("[i] Started script")
    count = 0
    enter_timer = 0
    while True:
        number += 1
        wsh.SendKeys(number)
        if mode == 1:
            print_chromatic_text(f"[i] Sending {number} (Mode: Enter)")
            wsh.SendKeys("{ENTER}")
            sleep(0.8)
            count += 1
            if count % 10 == 0:
                print_chromatic_text(f"[!] Waiting for 6 seconds after {count} iterations.")
                sleep(6)
        elif mode == 2:
            print_chromatic_text(f"[i] Sending {number} (Mode: Shift + Enter)")
            wsh.SendKeys("+{ENTER}")
            sleep(0.2)

            # Check if 10 seconds have passed and simulate pressing Enter
            if time.time() - enter_timer >= 10:
                print_chromatic_text(f"[!] Simulating Enter key press after {number} iterations.")
                wsh.SendKeys("{ENTER}")
                enter_timer = time.time()  # Reset the timer

try:
    mode_selection_art = r"""
    
██╗░░██╗███████╗░█████╗░██████╗░████████╗░█████╗░░█████╗░██╗░░░██╗███╗░░██╗████████╗███████╗██████╗░
██║░░██║██╔════╝██╔══██╗██╔══██╗╚══██╔══╝██╔══██╗██╔══██╗██║░░░██║████╗░██║╚══██╔══╝██╔════╝██╔══██╗
███████║█████╗░░███████║██████╔╝░░░██║░░░██║░░╚═╝██║░░██║██║░░░██║██╔██╗██║░░░██║░░░█████╗░░██████╔╝
██╔══██║██╔══╝░░██╔══██║██╔══██╗░░░██║░░░██║░░██╗██║░░██║██║░░░██║██║╚████║░░░██║░░░██╔══╝░░██╔══██╗
██║░░██║███████╗██║░░██║██║░░██║░░░██║░░░╚█████╔╝╚█████╔╝╚██████╔╝██║░╚███║░░░██║░░░███████╗██║░░██║
╚═╝░░╚═╝╚══════╝╚═╝░░╚═╝╚═╝░░╚═╝░░░╚═╝░░░░╚════╝░░╚════╝░░╚═════╝░╚═╝░░╚══╝░░░╚═╝░░░╚══════╝╚═╝░░╚═╝

    by oci (testver)
"""



    print_chromatic_text(mode_selection_art)
    print("Choose the writing mode:")
    print("1. Default Mode (Press Enter)")
    print("2. Special Mode (Press Shift + Enter)")
    mode = int(input("Enter the mode number: "))

    if mode not in (1, 2):
        raise ValueError("Invalid mode number. Please choose either 1 or 2.")

    starting_number = int(input("Enter the starting number: "))
    add_one_and_print(starting_number, mode)
except KeyboardInterrupt:
    print(Fore.WHITE + Style.BRIGHT + "\nPrzerwano przez użytkownika.")
except ValueError as e:
    print(Fore.RED + f"Error: {e}")
