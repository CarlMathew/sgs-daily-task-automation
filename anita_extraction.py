import pyautogui
import subprocess
import time



subprocess.Popen([ "C:\\Program Files (x86)\\AniTa\\anita.exe", "C:\\Program Files (x86)\\AniTa\\anita.wcf"])
time.sleep(3)

# pyautogui.click(x=100, y=200)
pyautogui.write("use-idb062")
pyautogui.press("enter")
time.sleep(4)
pyautogui.write("carlm")
pyautogui.press("enter")
time.sleep(4)
pyautogui.write("Test123")
pyautogui.press("enter")
time.sleep(4)
pyautogui.write("Test123")
pyautogui.press("enter")