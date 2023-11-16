import getpass
from pyautogui import *
import pyautogui
import time
import keyboard
import random
import win32api, win32con

def click(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)

test = "0123456789"
x = int(724)
x_dist = 90
y = int(726)
y_dist= 100

coordonnees = [(0,0),(0,0),(0,0),(0,0),(0,0),(0,0),(0,0),(0,0),(0,0),(0,0)]

# for i in range (5):

#     im_haut=pyautogui.screenshot(region=(x+i*x_dist,y,x_dist,y_dist))
#     im_bas=pyautogui.screenshot(region=(x+i*x_dist,y+y_dist,x_dist,y_dist))

#     im_haut.save(rf"D:\Programmation\projet_Bourso\screenimage{i}.png")
#     im_bas.save(rf"D:\Programmation\projet_Bourso\screenimage{5+i}.png")
code = "11110000"

for i in range (5): # pour opti tej l'image (vraiment plus long a coder mais faisable)
 
    j=0
    
    for j in code:
        reference = pyautogui.locateOnScreen(f"screenpimage{j}.png",region=(0,0,pyautogui.size().width,pyautogui.size().height),confidence=0.8)
        win32api.SetCursorPos((reference[0]+20,reference[1]+20))

    j=0
    reference = pyautogui.locateOnScreen(f"{j}.png",region=(x+i*x_dist,819,x_dist,102),confidence=0.7)

    while reference == None:
        reference = pyautogui.locateOnScreen(f"{j}.png",region=(x+i*x_dist,819,x_dist,102),confidence=0.7)
        j+=1
    coordonnees[j]=(coordonnees[0],coordonnees[1])

# for i in test:

#     reference = pyautogui.locateOnScreen(f"{i}.png",region=(0,0,960,1060),confidence=0.7)
#     while reference == None:
#             reference = pyautogui.locateOnScreen(f"{i}.png",region=(0,0,960,1060),confidence=0.7)

#     iml=pyautogui.screenshot(region=(0,0,960,1060))
#     iml.save(r"D:\Programmation\projet_Bourso\screenimage1.png")
    # time.sleep(1)
    #win32api.SetCursorPos((reference[0],reference[1]))

# print("le code est passe")

# SCREEN_SIZE = pyautogui.size()

# iml=pyautogui.screenshot(region=(0.34*SCREEN_SIZE.width,573,(0.67-0.34)*SCREEN_SIZE.width,400))
# iml.save(r"D:\Programmation\projet_Bourso\screenimage1.png")

# x = int(0.34*SCREEN_SIZE.width)
# x_dist = int((0.67-0.34)*SCREEN_SIZE.width)

# pyautogui.locateOnScreen(f"{3}.png",region=(x,573,x_dist,400),grayscale=True)




