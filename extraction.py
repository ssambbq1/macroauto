import pyautogui
import time


print("마우스 커서를 원하는 위치로 이동시키세요. 3초 후에 좌표를 출력합니다.")
time.sleep(3)

# 현재 마우스 커서의 좌표를 가져옵니다.
current_mouse_x, current_mouse_y = pyautogui.position()
print(f"현재 마우스 위치: ({current_mouse_x}, {current_mouse_y})")
