import pyautogui
import time
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import simpledialog, messagebox
import json
import os

CONFIG_PATH = os.path.join(os.path.expanduser("~"), "Desktop", "tnlcopy_coords.json")

def save_coords(coords):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(coords, f, ensure_ascii=False, indent=2)

def load_coords():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return None

def set_coords_gui():
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("좌표 설정", "각 버튼에 대해 마우스를 원하는 위치에 올리고 \"확인\"을 누르세요.\n3초 후 좌표가 자동으로 저장됩니다.\n\"확인\" 버튼 창이 화면에 안보이면 Ctrl+tab을 눌러서 활성화 시키세요.")
    coord_names = [
        ("date_xy", "일자 입력란"),
        ("serch_xy", "조회 버튼"),
        ("ref_date_xy", "복사 기준일자 입력란"),
        ("select_xy", "사원 선택란"),
        ("ref_copy_xy", "이전실적복사 버튼"),
        ("copy_xy", "복사 버튼")
    ]
    coords = {}
    for key, label in coord_names:
        messagebox.showinfo("좌표 설정", f"{label} 위치에 마우스를 올리고 OK를 누르세요.")
        time.sleep(3)
        x, y = pyautogui.position()
        coords[key] = (x, y)
        messagebox.showinfo("좌표 저장", f"{label} 좌표: {x}, {y} 저장됨")
    save_coords(coords)
    messagebox.showinfo("완료", f"좌표가 바탕화면에 저장되었습니다.\n{CONFIG_PATH}")
    return coords

def parse_dates(date_inputs):
    result = []
    for item in date_inputs:
        if '-' in item and '~' in item:
            start_str, end_str = item.split('~')
            start_date = datetime.strptime(start_str.strip(), "%Y-%m-%d")
            end_date = datetime.strptime(end_str.strip(), "%Y-%m-%d")
            delta = (end_date - start_date).days
            for i in range(delta + 1):
                result.append((start_date + timedelta(days=i)).strftime("%Y-%m-%d"))
        else:
            result.append(item.strip())
    return result

def get_user_input():
    root = tk.Tk()
    root.withdraw()
    date_inputs_str = simpledialog.askstring(
        "날짜 입력",
        "T&L을 작성할 날짜를 입력하세요.\n단일: 2025-07-01\n범위: 2025-07-01~2025-07-04\n여러 개는 , 또는 줄바꿈으로 구분"
    )
    if not date_inputs_str:
        messagebox.showerror("입력 오류", "날짜를 입력해야 합니다.")
        exit()
    ref_date = simpledialog.askstring(
        "복사 기준일자 입력",
        "복사 기준일자를 입력하세요 (예: 2025-06-27)"
    )
    if not ref_date:
        messagebox.showerror("입력 오류", "복사 기준일자를 입력해야 합니다.")
        exit()
    date_inputs = []
    for part in date_inputs_str.replace(',', '\n').split('\n'):
        part = part.strip()
        if part:
            date_inputs.append(part)
    return date_inputs, ref_date

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    if messagebox.askyesno("좌표 설정", "좌표를 새로 지정하시겠습니까? (아니오 선택 시 기존 좌표를 불러옵니다)"):
        coords = set_coords_gui()
    else:
        coords = load_coords()
        if not coords:
            messagebox.showerror("오류", "저장된 좌표가 없습니다. 좌표를 먼저 지정하세요.")
            exit()

    date_inputs, ref_date = get_user_input()
    dates = parse_dates(date_inputs)

    date_xy = tuple(coords["date_xy"])
    serch_xy = tuple(coords["serch_xy"])
    ref_date_xy = tuple(coords["ref_date_xy"])
    select_xy = tuple(coords["select_xy"])
    ref_copy_xy = tuple(coords["ref_copy_xy"])
    copy_xy = tuple(coords["copy_xy"])

    # 복사 기준일자 입력 (한 번만)
    time.sleep(2)
    pyautogui.click(*ref_date_xy)
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(1)
    pyautogui.press('delete')
    time.sleep(1)
    for char in ref_date:
        pyautogui.press(char)
        time.sleep(0.1)
    time.sleep(1)


    # 날짜별 반복 작업
    for date in dates:
        pyautogui.click(*date_xy)
        time.sleep(2)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(1)
        pyautogui.press('delete')
        time.sleep(1)
        for char in date:
            pyautogui.press(char)
            time.sleep(0.1)
        time.sleep(1)

        pyautogui.click(*serch_xy)
        time.sleep(2)
        pyautogui.click(*select_xy)
        time.sleep(1)
        pyautogui.click(*ref_copy_xy)
        time.sleep(1)
        pyautogui.click(*copy_xy)
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)