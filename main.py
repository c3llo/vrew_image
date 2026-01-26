import tkinter as tk
from tkinter import messagebox, scrolledtext
import threading
import time
from pywinauto import Application

def split_paragraphs(text):
    import re
    # 한 줄 이상의 빈 줄(\n\n)을 기준으로 문단을 나눕니다.
    # 이렇게 하면 여러 줄로 된 하나의 문단이 Vrew의 자막 한 칸으로 입력됩니다.
    blocks = re.split(r'\n\s*\n', text.strip())
    return [b.strip() for b in blocks if b.strip()]

def start_input():
    script = text_area.get("1.0", tk.END).strip()
    if not script:
        messagebox.showwarning("경고", "대본을 입력하세요")
        return

    paragraphs = split_paragraphs(script)

    # UI가 멈추지 않도록 별도 스레드에서 실행
    thread = threading.Thread(target=run_automation, args=(paragraphs,))
    thread.daemon = True
    thread.start()

def run_automation(paragraphs):
    from pywinauto import Desktop
    print("\n--- 실제 Vrew 창 탐색 시작 ---")
    all_windows = Desktop(backend="uia").windows()
    
    target_window = None
    
    # 1. 실제 Vrew 프로젝트 창 찾기 (본인 프로그램 제외, 점 표시가 있거나 Vrew 이름 포함)
    for w in all_windows:
        title = w.window_text()
        if title and "Vrew" in title and "대본 자동 입력기" not in title:
            print(f"후보 발견: [{title}] (Handle: {w.handle})")
            if "●" in title: # 수정 중인 창 우선
                target_window = w
                break
            if not target_window:
                target_window = w

    if not target_window:
        print("Vrew 창을 찾지 못했습니다.")
        root.after(0, lambda: messagebox.showerror("에러", "Vrew 창을 찾을 수 없습니다. Vrew가 실행 중인지 확인해 주세요."))
        return

    try:
        print(f"최종 선택된 창: {target_window.window_text()} (Handle: {target_window.handle})")
        
        # 제목이 아닌 고유 '핸들(handle)'을 사용하여 연결 (중복 매칭 방지)
        app = Application(backend="uia").connect(handle=target_window.handle)
        vrew_win = app.window(handle=target_window.handle)
        
        # 창 활성화
        if vrew_win.get_show_state() == 2: # 2: 최소화
            vrew_win.restore()
        vrew_win.set_focus()
        print("성공적으로 연결 및 포커스 완료")

    except Exception as e:
        error_msg = str(e)
        print(f"연결 실패: {error_msg}")
        root.after(0, lambda m=error_msg: messagebox.showerror("에러", f"Vrew 창에 연결할 수 없습니다.\n원인: {m}"))
        return

    # 입력창(Edit) 찾기 - Vrew는 Chromium 기반이므로 구조가 복잡할 수 있음
    # 입력 시작 (수동 포커스 방식)
    print("입력을 시작합니다. Vrew 창에 포커스를 둡니다.")
    
    for i, para in enumerate(paragraphs, 1):
        root.after(0, lambda i=i: status_label.config(text=f"{i}/{len(paragraphs)} 문단 입력 중..."))
        
        try:
            # 창 활성화
            vrew_win.set_focus()
            
            # 1. 클립보드에 현재 문단 복사
            root.clipboard_clear()
            root.clipboard_append(para)
            root.update() 
            
            # 2. Vrew에 입력 (전체선택 -> 삭제 -> 붙여넣기)
            # 아주 짧은 pause와 sleep으로 처리 속도 향상
            vrew_win.type_keys("^a", pause=0.01)
            time.sleep(0.01)
            vrew_win.type_keys("{BACKSPACE}")
            vrew_win.type_keys("^v", pause=0.01)
            time.sleep(0.1) # 붙여넣기 완료를 위한 최소 대기

            # 3. 다음 자막 칸 이동 (마지막 문단이 아닐 때만 Tab 수행)
            if i < len(paragraphs):
                vrew_win.type_keys("{TAB}")
                time.sleep(0.05)
            
        except Exception as e:
            error_msg = str(e)
            print(f"입력 작업 중 오류: {error_msg}")
            root.after(0, lambda m=error_msg: messagebox.showerror("에러", f"입력 중 오류가 발생했습니다.\n{m}"))
            break

    root.after(0, lambda: status_label.config(text="완료"))
    root.after(0, lambda: messagebox.showinfo("완료", "모든 문단의 자동 입력이 완료되었습니다."))

def update_paragraph_count(event=None):
    """실시간으로 문단 개수를 계산하여 라벨에 표시합니다."""
    text = text_area.get("1.0", tk.END).strip()
    if not text:
        count_label.config(text="감지된 문단: 0개", fg="#666")
        return
    
    paragraphs = split_paragraphs(text)
    count = len(paragraphs)
    count_label.config(text=f"감지된 문단: {count}개", fg="#1976D2")

# GUI 설정
root = tk.Tk()
root.title("Vrew 대본 자동 입력기")
root.geometry("1000x900")  
root.minsize(600, 700)     

root.columnconfigure(0, weight=1)
root.rowconfigure(3, weight=1) # 텍스트 영역 가로/세로 확장

# 상단 섹션
header_frame = tk.Frame(root, pady=10)
header_frame.grid(row=0, column=0, sticky="ew")

instruction_label = tk.Label(
    header_frame, 
    text="Vrew의 첫 번째 자막 입력창을 '한 번 클릭'한 후\n[자동 입력 시작] 버튼을 눌러주세요.", 
    font=("Malgun Gothic", 13, "bold"),
    fg="#D32F2F",
    pady=10
)
instruction_label.pack()

# 문단 개수 표시 라벨 (입력창 바로 위)
count_label = tk.Label(root, text="감지된 문단: 0개", font=("Malgun Gothic", 10, "bold"), fg="#666")
count_label.grid(row=1, column=0, sticky="w", padx=25)

# 중간 라벨
info_label = tk.Label(root, text=" 아래에 대본을 입력하세요 (문단 사이를 '빈 줄'로 구분):", font=("Malgun Gothic", 10), fg="#666")
info_label.grid(row=2, column=0, sticky="w", padx=20)

# 대본 입력창
text_area = scrolledtext.ScrolledText(root, font=("Malgun Gothic", 11), undo=True)
text_area.grid(row=3, column=0, sticky="nsew", padx=20, pady=5)

# 실시간 문단 카운트 바인딩
text_area.bind("<KeyRelease>", update_paragraph_count)
text_area.bind("<ButtonRelease>", update_paragraph_count) # 마우스 붙여넣기 대응

# 하단 섹션
footer_frame = tk.Frame(root, pady=15)
footer_frame.grid(row=4, column=0, sticky="ew")

status_label = tk.Label(footer_frame, text="대기 중", fg="blue", font=("Malgun Gothic", 10))
status_label.pack()

start_button = tk.Button(
    footer_frame, 
    text="자동 입력 시작", 
    command=start_input, 
    bg="#4CAF50", 
    fg="white", 
    font=("Malgun Gothic", 13, "bold"), 
    padx=40, 
    pady=15,
    cursor="hand2"
)
start_button.pack(pady=10)

root.mainloop()
