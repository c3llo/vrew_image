import json
import re
import tkinter as tk
from tkinter import messagebox, scrolledtext
import threading
import time
from pywinauto import Application

def try_extract_script_from_json(text):
    text = text.strip()
    if not text.startswith("{") or not text.strip().endswith("}"):
        return False, text
    normalized = re.sub(r'"\s*\n\s*"', '", "', text)
    try:
        data = json.loads(normalized)
        if not isinstance(data, dict):
            return False, text
        def key_order(k):
            nums = [int(m.group()) for m in re.finditer(r'\d+', k)]
            return (nums, k) if nums else ([-1], k)
        keys = sorted(data.keys(), key=key_order)
        parts = []
        for k in keys:
            v = data[k]
            if isinstance(v, str) and v.strip():
                parts.append(v.strip())
        if not parts:
            return False, text
        return True, "\n\n".join(parts)
    except (json.JSONDecodeError, TypeError):
        return False, text

def get_lines_from_text(text):
    return [s.strip() for s in text.split("\n") if s.strip()]

def split_into_n_paragraphs(lines, n):
    if n <= 0 or not lines:
        return []
    if n == 1:
        return ["\n".join(lines)]
    result = []
    baseSize = len(lines) // n
    remainder = len(lines) % n
    idx = 0
    for i in range(n):
        chunkSize = baseSize + (1 if i < remainder else 0)
        chunk = lines[idx:idx + chunkSize]
        idx += chunkSize
        result.append("\n".join(chunk))
    return result

def split_by_sentence_count(lines, n):
    if n <= 0 or not lines:
        return []
    result = []
    for i in range(0, len(lines), n):
        chunk = lines[i:i + n]
        result.append("\n".join(chunk))
    return result

def do_split_paragraphs():
    script = text_area.get("1.0", tk.END).strip()
    if not script:
        messagebox.showwarning("경고", "대본을 입력하세요")
        return
    try:
        n = int(paragraph_count_var.get().strip())
        if n < 1 or n > 999:
            messagebox.showwarning("경고", "문단 갯수는 1~999 사이로 입력하세요")
            return
    except ValueError:
        messagebox.showwarning("경고", "문단 갯수에 올바른 숫자를 입력하세요 (1~999)")
        return
    lines = get_lines_from_text(script)
    if not lines:
        messagebox.showwarning("경고", "유효한 줄이 없습니다")
        return
    if n > len(lines):
        messagebox.showwarning("경고", f"줄 수({len(lines)}개)보다 문단 수({n}개)가 많을 수 없습니다")
        return
    paragraphs = split_into_n_paragraphs(lines, n)
    newText = "\n\n".join(paragraphs)
    text_area.delete("1.0", tk.END)
    text_area.insert("1.0", newText)
    update_paragraph_count()

def do_split_by_sentences():
    script = text_area.get("1.0", tk.END).strip()
    if not script:
        messagebox.showwarning("경고", "대본을 입력하세요")
        return
    try:
        n = int(sentence_count_var.get().strip())
        if n < 1 or n > 999:
            messagebox.showwarning("경고", "문장 갯수는 1~999 사이로 입력하세요")
            return
    except ValueError:
        messagebox.showwarning("경고", "문장 갯수에 올바른 숫자를 입력하세요 (1~999)")
        return
    lines = get_lines_from_text(script)
    if not lines:
        messagebox.showwarning("경고", "유효한 줄이 없습니다")
        return
    paragraphs = split_by_sentence_count(lines, n)
    newText = "\n\n".join(paragraphs)
    text_area.delete("1.0", tk.END)
    text_area.insert("1.0", newText)
    update_paragraph_count()

def split_paragraphs(text, remove_headers=False):
    import re
    
    if remove_headers:
        # '챕터 [숫자]'로 시작하는 행을 행 단위로 먼저 제거합니다.
        lines = text.split('\n')
        # 패턴: 시작부분(공백허용) + 챕터 + 공백(허용) + 숫자
        pattern = re.compile(r'^\s*챕터\s*\d+')
        
        filtered_lines = []
        for line in lines:
            if not pattern.match(line):
                filtered_lines.append(line)
        
        # 다시 텍스트로 합칩니다.
        text = '\n'.join(filtered_lines)
        
    # 한 줄 이상의 빈 줄(\n\n)을 기준으로 문단을 나눕니다.
    blocks = re.split(r'\n\s*\n', text.strip())
    # 각 블록에서 앞뒤 공백을 제거하고 빈 블록은 제외합니다.
    return [b.strip() for b in blocks if b.strip()]

def start_input():
    script = text_area.get("1.0", tk.END).strip()
    if not script:
        messagebox.showwarning("경고", "대본을 입력하세요")
        return

    paragraphs = split_paragraphs(script, remove_headers=remove_headers_var.get())

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

def get_lines_for_count(text, remove_headers=False):
    lines = get_lines_from_text(text)
    if remove_headers:
        pattern = re.compile(r'^\s*챕터\s*\d+')
        lines = [line for line in lines if not pattern.match(line)]
    return lines

def update_paragraph_count(event=None):
    """실시간으로 문단/문장 개수를 계산하여 라벨에 표시합니다."""
    text = text_area.get("1.0", tk.END).strip()
    if not text:
        count_label.config(text="감지된 문단: 0개", fg="#666")
        sentence_count_label.config(text="감지된 문장: 0개", fg="#666")
        return
    
    remove_headers = remove_headers_var.get()
    paragraphs = split_paragraphs(text, remove_headers=remove_headers)
    lines = get_lines_for_count(text, remove_headers=remove_headers)
    count_label.config(text=f"감지된 문단: {len(paragraphs)}개", fg="#1976D2")
    sentence_count_label.config(text=f"감지된 문장: {len(lines)}개", fg="#1976D2")

# GUI 설정
root = tk.Tk()
root.title("Vrew 대본 자동 입력기 by Moon_8800")
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

# 옵션 섹션 (문단 개수 및 챕터 삭제 체크박스)
option_frame = tk.Frame(root)
option_frame.grid(row=1, column=0, sticky="ew", padx=20)

count_label = tk.Label(option_frame, text="감지된 문단: 0개", font=("Malgun Gothic", 10, "bold"), fg="#666")
count_label.pack(side="left", padx=5)

sentence_count_label = tk.Label(option_frame, text="감지된 문장: 0개", font=("Malgun Gothic", 10, "bold"), fg="#666")
sentence_count_label.pack(side="left", padx=5)

remove_headers_var = tk.BooleanVar(value=True)  # 기본적으로 체크됨
remove_headers_checkbox = tk.Checkbutton(
    option_frame, 
    text="챕터 제목 제외하기 ( 예: 챕터 1: ... )", 
    variable=remove_headers_var,
    font=("Malgun Gothic", 10),  # 폰트 크기 원복
    command=update_paragraph_count
)
remove_headers_checkbox.pack(side="right", padx=5)

# 중간 라벨
info_label = tk.Label(root, text=" 아래에 대본을 입력하세요 (문단 사이를 '빈 줄'로 구분):", font=("Malgun Gothic", 10), fg="#666")
info_label.grid(row=2, column=0, sticky="w", padx=20)

# 대본 입력창
text_area = scrolledtext.ScrolledText(root, font=("Malgun Gothic", 11), undo=True)
text_area.grid(row=3, column=0, sticky="nsew", padx=20, pady=5)

def on_paste(event):
    try:
        raw = root.clipboard_get()
    except tk.TclError:
        return
    ok, text = try_extract_script_from_json(raw)
    if ok:
        text_area.insert(tk.INSERT, text)
        update_paragraph_count()
    else:
        text_area.insert(tk.INSERT, raw)
        update_paragraph_count()
    return "break"

text_area.bind("<Control-v>", on_paste)
text_area.bind("<KeyRelease>", update_paragraph_count)
text_area.bind("<ButtonRelease>", update_paragraph_count)

# 하단 섹션
footer_frame = tk.Frame(root, pady=15)
footer_frame.grid(row=4, column=0, sticky="ew")

status_label = tk.Label(footer_frame, text="대기 중", fg="blue", font=("Malgun Gothic", 10))
status_label.pack()

bottom_row = tk.Frame(footer_frame)
bottom_row.pack(pady=10)

left_frame = tk.Frame(bottom_row)
left_frame.pack(side="left", padx=(0, 40))

start_button = tk.Button(
    left_frame,
    text="자동입력시작",
    command=start_input,
    bg="#4CAF50",
    fg="white",
    font=("Malgun Gothic", 13, "bold"),
    padx=40,
    pady=15,
    cursor="hand2"
)
start_button.pack()

right_frame = tk.Frame(bottom_row)
right_frame.pack(side="right")

paragraph_count_var = tk.StringVar(value="5")

def on_paragraph_spin(delta):
    try:
        current = int(paragraph_count_var.get().strip())
    except ValueError:
        current = 5
    current = max(1, min(999, current + delta))
    paragraph_count_var.set(str(current))

row1 = tk.Frame(right_frame)
row1.pack(pady=(0, 8))
tk.Label(row1, text="문단 갯수", font=("Malgun Gothic", 10)).pack(side="left", padx=(0, 5))
paragraph_count_spin = tk.Spinbox(
    row1,
    textvariable=paragraph_count_var,
    from_=1,
    to=999,
    width=5,
    font=("Malgun Gothic", 11),
    justify="center"
)
paragraph_count_spin.pack(side="left", padx=(0, 10))
split_button = tk.Button(
    row1,
    text="문단 나누기",
    command=do_split_paragraphs,
    bg="#2196F3",
    fg="white",
    font=("Malgun Gothic", 11, "bold"),
    padx=20,
    pady=10,
    cursor="hand2"
)
split_button.pack(side="left")
def paragraph_wheel(e):
    d = e.delta if hasattr(e, "delta") else (120 if e.num == 4 else -120)
    on_paragraph_spin(1 if d > 0 else -1)
    return "break"
paragraph_count_spin.bind("<MouseWheel>", paragraph_wheel)
paragraph_count_spin.bind("<Button-4>", lambda e: (on_paragraph_spin(1), "break")[-1])
paragraph_count_spin.bind("<Button-5>", lambda e: (on_paragraph_spin(-1), "break")[-1])

sentence_count_var = tk.StringVar(value="2")

def on_sentence_spin(delta):
    try:
        current = int(sentence_count_var.get().strip())
    except ValueError:
        current = 2
    current = max(1, min(999, current + delta))
    sentence_count_var.set(str(current))

row2 = tk.Frame(right_frame)
row2.pack()
tk.Label(row2, text="문장 갯수", font=("Malgun Gothic", 10)).pack(side="left", padx=(0, 5))
sentence_count_spin = tk.Spinbox(
    row2,
    textvariable=sentence_count_var,
    from_=1,
    to=999,
    width=5,
    font=("Malgun Gothic", 11),
    justify="center"
)
sentence_count_spin.pack(side="left", padx=(0, 10))
def sentence_wheel(e):
    d = e.delta if hasattr(e, "delta") else (120 if e.num == 4 else -120)
    on_sentence_spin(1 if d > 0 else -1)
    return "break"
sentence_count_spin.bind("<MouseWheel>", sentence_wheel)
sentence_count_spin.bind("<Button-4>", lambda e: (on_sentence_spin(1), "break")[-1])
sentence_count_spin.bind("<Button-5>", lambda e: (on_sentence_spin(-1), "break")[-1])

split_sentences_button = tk.Button(
    row2,
    text="문장갯수로 나누기",
    command=do_split_by_sentences,
    bg="#2196F3",
    fg="white",
    font=("Malgun Gothic", 11, "bold"),
    padx=20,
    pady=10,
    cursor="hand2"
)
split_sentences_button.pack(side="left")

root.mainloop()
