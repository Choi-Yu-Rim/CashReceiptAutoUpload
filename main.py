import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import queue
import os
import json
import time
import tempfile
import shutil
from datetime import datetime

import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# ─────────────────────────────────────────
# 설정
# ─────────────────────────────────────────
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
ROWS_PER_FILE = 200
LOGIN_URL = "https://taxadmin.tosspayments.com/"
UPLOAD_URL = "https://taxadmin.tosspayments.com/receipt/receiptAdmitReq.jsp?pageNum=2&subNum=3"

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"user_id": "", "user_pw": ""}


def save_config(user_id, user_pw):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({"user_id": user_id, "user_pw": user_pw}, f, ensure_ascii=False)


# ─────────────────────────────────────────
# 엑셀 분할
# ─────────────────────────────────────────
def split_excel(file_path: str, output_dir: str):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    if len(all_rows) < 2:
        raise ValueError("엑셀 파일에 데이터가 없습니다 (헤더 외 행이 필요합니다).")
    header = all_rows[0]
    data_rows = all_rows[1:]
    chunks = [data_rows[i:i + ROWS_PER_FILE] for i in range(0, len(data_rows), ROWS_PER_FILE)]
    output_files = []
    for idx, chunk in enumerate(chunks, 1):
        new_wb = openpyxl.Workbook()
        new_ws = new_wb.active
        new_ws.append(list(header))
        for row in chunk:
            new_ws.append(list(row))
        out_path = os.path.join(output_dir, f"현금영수증_분할_{idx}.xlsx")
        new_wb.save(out_path)
        output_files.append(out_path)
    return output_files


# ─────────────────────────────────────────
# 메인 앱
# ─────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("현금영수증 자동 업로드")
        self.geometry("600x740")
        self.resizable(False, False)

        self.config_data = load_config()
        self.msg_queue: queue.Queue = queue.Queue()
        self.running = False

        self._build_ui()
        self._poll_queue()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)

        # ── 제목 ──
        title_frame = ctk.CTkFrame(self, fg_color="transparent")
        title_frame.grid(row=0, column=0, padx=24, pady=(24, 0), sticky="ew")

        ctk.CTkLabel(
            title_frame, text="현금영수증 자동 업로드",
            font=ctk.CTkFont(size=22, weight="bold")
        ).pack(anchor="w")
        ctk.CTkLabel(
            title_frame, text="토스페이먼츠 현금영수증 일괄 업로드 도우미",
            font=ctk.CTkFont(size=12), text_color="gray"
        ).pack(anchor="w", pady=(2, 0))

        # ── 1. 파일 선택 ──
        f1 = ctk.CTkFrame(self)
        f1.grid(row=1, column=0, padx=24, pady=(16, 0), sticky="ew")
        f1.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(f1, text="1.  현금영수증 엑셀 파일 선택",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, columnspan=2, padx=16, pady=(14, 8), sticky="w")

        self.file_var = ctk.StringVar()
        ctk.CTkEntry(f1, textvariable=self.file_var, state="readonly",
                     font=ctk.CTkFont(size=11), height=36).grid(
            row=1, column=0, padx=(16, 8), pady=(0, 14), sticky="ew")
        ctk.CTkButton(f1, text="파일 찾기", width=90, height=36,
                      command=self._browse_file).grid(
            row=1, column=1, padx=(0, 16), pady=(0, 14))

        # ── 2. 로그인 정보 ──
        f2 = ctk.CTkFrame(self)
        f2.grid(row=2, column=0, padx=24, pady=(12, 0), sticky="ew")
        f2.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(f2, text="2.  토스페이먼츠 로그인 정보",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, columnspan=2, padx=16, pady=(14, 8), sticky="w")

        ctk.CTkLabel(f2, text="아이디", font=ctk.CTkFont(size=12), width=60).grid(
            row=1, column=0, padx=(16, 8), pady=4, sticky="w")
        self.id_var = ctk.StringVar(value=self.config_data.get("user_id", ""))
        ctk.CTkEntry(f2, textvariable=self.id_var, height=36,
                     font=ctk.CTkFont(size=12)).grid(
            row=1, column=1, padx=(0, 16), pady=4, sticky="ew")

        ctk.CTkLabel(f2, text="비밀번호", font=ctk.CTkFont(size=12), width=60).grid(
            row=2, column=0, padx=(16, 8), pady=4, sticky="w")
        self.pw_var = ctk.StringVar(value=self.config_data.get("user_pw", ""))
        ctk.CTkEntry(f2, textvariable=self.pw_var, show="●", height=36,
                     font=ctk.CTkFont(size=12)).grid(
            row=2, column=1, padx=(0, 16), pady=4, sticky="ew")

        has_saved = bool(self.config_data.get("user_id") or self.config_data.get("user_pw"))
        saved_text = "✅  저장된 설정을 불러왔습니다. 틀리면 수정 후 시작하세요." if has_saved else "⚠️  저장된 설정이 없습니다. 입력 후 시작하세요."
        saved_color = "#2e7d32" if has_saved else "#b45309"
        ctk.CTkLabel(f2, text=saved_text, font=ctk.CTkFont(size=11),
                     text_color=saved_color).grid(
            row=3, column=0, columnspan=2, padx=16, pady=(6, 14), sticky="w")

        # ── 시작 버튼 ──
        self.start_btn = ctk.CTkButton(
            self, text="▶   업로드 시작",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=50, corner_radius=10,
            command=self._start
        )
        self.start_btn.grid(row=3, column=0, padx=24, pady=16, sticky="ew")

        # ── 진행 상황 ──
        f3 = ctk.CTkFrame(self)
        f3.grid(row=4, column=0, padx=24, pady=(0, 0), sticky="ew")
        f3.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(f3, text="진행 상황",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, padx=16, pady=(14, 6), sticky="w")

        self.status_var = ctk.StringVar(value="대기 중...")
        self.status_label = ctk.CTkLabel(
            f3, textvariable=self.status_var,
            font=ctk.CTkFont(size=12), text_color="#555555", anchor="w"
        )
        self.status_label.grid(row=1, column=0, padx=16, pady=(0, 6), sticky="ew")

        self.progress_bar = ctk.CTkProgressBar(f3, height=16, corner_radius=8)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=2, column=0, padx=16, pady=(0, 6), sticky="ew")

        self.count_var = ctk.StringVar(value="")
        ctk.CTkLabel(f3, textvariable=self.count_var,
                     font=ctk.CTkFont(size=11), text_color="gray").grid(
            row=3, column=0, padx=16, pady=(0, 14), sticky="e")

        # ── 로그 ──
        f4 = ctk.CTkFrame(self)
        f4.grid(row=5, column=0, padx=24, pady=(12, 24), sticky="nsew")
        f4.grid_columnconfigure(0, weight=1)
        f4.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(5, weight=1)

        ctk.CTkLabel(f4, text="작업 로그",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, padx=16, pady=(14, 6), sticky="w")

        self.log_text = ctk.CTkTextbox(
            f4, font=ctk.CTkFont(family="Courier New", size=11),
            state="disabled", wrap="word"
        )
        self.log_text.grid(row=1, column=0, padx=16, pady=(0, 16), sticky="nsew")

    # ── 이벤트 ──────────────────────────────────
    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="현금영수증 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")]
        )
        if path:
            self.file_var.set(path)

    def _start(self):
        if self.running:
            return

        file_path = self.file_var.get().strip()
        user_id = self.id_var.get().strip()
        user_pw = self.pw_var.get().strip()

        if not file_path:
            messagebox.showwarning("알림", "엑셀 파일을 선택해주세요.")
            return
        if not os.path.isfile(file_path):
            messagebox.showwarning("알림", "선택한 파일이 존재하지 않습니다.")
            return
        if not user_id or not user_pw:
            messagebox.showwarning("알림", "아이디와 비밀번호를 입력해주세요.")
            return

        save_config(user_id, user_pw)

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self.progress_bar.set(0)
        self.count_var.set("")

        self.running = True
        self.start_btn.configure(state="disabled", text="⏳  업로드 중...")

        t = threading.Thread(target=self._run, args=(file_path, user_id, user_pw), daemon=True)
        t.start()

    # ── 큐 폴링 ─────────────────────────────────
    def _poll_queue(self):
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                kind = msg[0]
                if kind == "log":
                    _, text, color = msg
                    ts = datetime.now().strftime("%H:%M:%S")
                    self.log_text.configure(state="normal")
                    self.log_text.insert("end", f"[{ts}] {text}\n")
                    self.log_text.see("end")
                    self.log_text.configure(state="disabled")
                elif kind == "status":
                    _, text, color = msg
                    self.status_var.set(text)
                    self.status_label.configure(text_color=color)
                elif kind == "progress":
                    _, cur, total = msg
                    self.progress_bar.set(cur / total if total > 0 else 0)
                    self.count_var.set(f"{cur} / {total} 파일 완료")
                elif kind == "done":
                    self.start_btn.configure(state="normal", text="▶   업로드 시작")
                    self.running = False
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _log(self, text, color=None):
        self.msg_queue.put(("log", text, color))

    def _set_status(self, text, color="#555555"):
        self.msg_queue.put(("status", text, color))

    def _set_progress(self, cur, total):
        self.msg_queue.put(("progress", cur, total))

    # ── 메인 작업 스레드 ─────────────────────────
    def _run(self, file_path, user_id, user_pw):
        driver = None
        tmp_dir = None

        try:
            # 1단계: 엑셀 분할
            self._set_status("엑셀 파일 분할 중...", "#1565c0")
            self._log("[ 1단계 ] 엑셀 파일 분할 시작")
            tmp_dir = tempfile.mkdtemp(prefix="cash_receipt_")
            try:
                split_files = split_excel(file_path, tmp_dir)
            except Exception as e:
                self._log(f"❌ 엑셀 분할 실패: {e}")
                self._set_status("❌ 실패: 엑셀 파일을 읽을 수 없습니다", "#c62828")
                self.msg_queue.put(("done",))
                return
            total = len(split_files)
            self._log(f"✅ 분할 완료 → 총 {total}개 파일 생성됨")
            self._set_progress(0, total)

            # 2단계: 크롬 드라이버
            self._set_status("크롬 드라이버 준비 중... (처음 실행 시 다운로드)", "#1565c0")
            self._log("[ 2단계 ] 크롬 드라이버 초기화 중...")
            self._log("  (처음 실행 시 드라이버 다운로드로 시간이 걸릴 수 있습니다)")
            try:
                opts = Options()
                service = Service(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=opts)
                driver.minimize_window()
            except Exception as e:
                self._log(f"❌ 크롬 드라이버 초기화 실패: {e}")
                self._log("  → 크롬 브라우저가 설치되어 있는지 확인해주세요")
                self._set_status("❌ 실패: 크롬 드라이버 오류", "#c62828")
                self.msg_queue.put(("done",))
                return
            self._log("✅ 크롬 드라이버 초기화 완료")

            # 3단계: 로그인
            self._set_status("토스페이먼츠 로그인 중...", "#1565c0")
            self._log("[ 3단계 ] 로그인 시도 중...")
            try:
                driver.get(LOGIN_URL)
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.ID, "user_id"))
                ).send_keys(user_id)
                driver.find_element(By.ID, "user_pw").send_keys(user_pw)
                driver.find_element(By.ID, "login").click()
                WebDriverWait(driver, 15).until(EC.url_contains("main"))
                self._log("✅ 로그인 완료")
            except Exception as e:
                self._log(f"❌ 로그인 실패: {e}")
                self._log("  → 아이디/비밀번호를 다시 확인해주세요")
                self._set_status("❌ 실패: 로그인 오류", "#c62828")
                self.msg_queue.put(("done",))
                return

            # 4단계: 파일 업로드
            self._set_status(f"파일 업로드 중... (0 / {total})", "#1565c0")
            self._log("[ 4단계 ] 파일 업로드 시작")
            for idx, file in enumerate(split_files, 1):
                fname = os.path.basename(file)
                self._set_status(f"업로드 중: {fname}  ({idx} / {total})", "#1565c0")
                self._log(f"  [{idx}/{total}] {fname} 업로드 중...")
                try:
                    driver.get(UPLOAD_URL)
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.ID, "uploadFile"))
                    )
                    driver.find_element(By.ID, "uploadFile").send_keys(os.path.abspath(file))
                    time.sleep(4)
                    self._log(f"  ✅ [{idx}/{total}] {fname} 업로드 완료")
                    self._set_progress(idx, total)
                except Exception as e:
                    self._log(f"  ❌ [{idx}/{total}] {fname} 업로드 실패: {e}")
                    self._log("  → 오류 발생으로 나머지 업로드를 중단합니다")
                    self._set_status(f"❌ {idx}번째 파일 오류로 중단됨 → {fname}", "#c62828")
                    self._set_progress(idx, total)
                    self.msg_queue.put(("done",))
                    return

            # 완료
            self._set_status(f"✅ 완료!  총 {total}개 파일 모두 업로드 성공", "#2e7d32")
            self._log("")
            self._log(f"🎉 모든 작업 완료!  {total}개 파일 업로드 성공")
            self.after(0, lambda: messagebox.showinfo(
                "완료", f"총 {total}개 파일을 모두 업로드했습니다!\n\n창을 닫아도 됩니다."
            ))

        except Exception as e:
            self._log(f"❌ 예기치 못한 오류: {e}")
            self._set_status("❌ 예기치 못한 오류 발생", "#c62828")
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass
            if tmp_dir and os.path.exists(tmp_dir):
                try:
                    shutil.rmtree(tmp_dir)
                except Exception:
                    pass
            self.msg_queue.put(("done",))


if __name__ == "__main__":
    app = App()
    app.mainloop()
