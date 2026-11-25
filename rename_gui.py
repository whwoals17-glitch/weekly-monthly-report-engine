import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from filename_normalizer import batch_rename_in_folder


class RenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("사업소 주간/월간 보고서 파일명 자동 변환기 (01~15번 표준화)")
        self.root.geometry("860x580")
        self.root.configure(bg="#F8FAFC")

        # 상단 안내
        header_frame = tk.Frame(root, bg="#F8FAFC", pady=12, padx=20)
        header_frame.pack(fill="x")

        tk.Label(
            header_frame,
            text="🏷️ 사업소 보고서 파일명 일괄 표준화 프로그램",
            font=("맑은 고딕", 16, "bold"),
            bg="#F8FAFC",
            fg="#1E293B",
        ).pack(anchor="w")

        tk.Label(
            header_frame,
            text="제각각인 파일명이나 잘못된 번호(예: 08_제주, 12_인재개발원, 날짜생략 등)를 자동 분석해 01~15번 표준 파일명으로 일괄 변경합니다.",
            font=("맑은 고딕", 9),
            bg="#F8FAFC",
            fg="#64748B",
        ).pack(anchor="w", pady=(4, 0))

        # 폴더 선택 영역
        ctrl_frame = tk.Frame(root, bg="#FFFFFF", pady=14, padx=16, highlightbackground="#E2E8F0", highlightthickness=1)
        ctrl_frame.pack(fill="x", padx=20, pady=6)

        row1 = tk.Frame(ctrl_frame, bg="#FFFFFF")
        row1.pack(fill="x")

        tk.Label(row1, text="대상 폴더:", font=("맑은 고딕", 10, "bold"), bg="#FFFFFF", width=9, anchor="w").pack(side="left")
        self.folder_var = tk.StringVar(value=os.getcwd())
        self.folder_entry = tk.Entry(row1, textvariable=self.folder_var, font=("맑은 고딕", 10), relief="solid", bd=1)
        self.folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=4)

        tk.Button(
            row1,
            text="📂 폴더 찾아보기",
            font=("맑은 고딕", 9, "bold"),
            bg="#334155",
            fg="white",
            relief="flat",
            padx=12,
            pady=4,
            command=self.browse_folder,
        ).pack(side="left")

        row2 = tk.Frame(ctrl_frame, bg="#FFFFFF", pady=8)
        row2.pack(fill="x")

        tk.Label(row2, text="기준 날짜:", font=("맑은 고딕", 10, "bold"), bg="#FFFFFF", width=9, anchor="w").pack(side="left")
        self.date_var = tk.StringVar(value="")
        tk.Entry(row2, textvariable=self.date_var, font=("맑은 고딕", 10), width=16, relief="solid", bd=1).pack(side="left", ipady=3)
        tk.Label(
            row2,
            text="(비워두면 폴더 내 다수결 날짜로 자동 통일됩니다. 직접 지정 시 예: 20261023)",
            font=("맑은 고딕", 9),
            bg="#FFFFFF",
            fg="#64748B",
        ).pack(side="left", padx=8)

        self.unify_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            row2,
            text="모든 파일 날짜 통일",
            variable=self.unify_var,
            bg="#FFFFFF",
            font=("맑은 고딕", 9),
            command=self.preview_files,
        ).pack(side="right")

        self.convert_doc_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            row2,
            text="구형(.doc) → .docx 자동 변환",
            variable=self.convert_doc_var,
            bg="#FFFFFF",
            font=("맑은 고딕", 9, "bold"),
            fg="#4F46E5",
            command=self.preview_files,
        ).pack(side="right", padx=(0, 12))

        # 버튼 영역
        btn_frame = tk.Frame(root, bg="#F8FAFC", pady=6, padx=20)
        btn_frame.pack(fill="x")

        tk.Button(
            btn_frame,
            text="🔍 변경 미리보기 (새로고침)",
            font=("맑은 고딕", 10, "bold"),
            bg="#64748B",
            fg="white",
            relief="flat",
            padx=16,
            pady=7,
            command=self.preview_files,
        ).pack(side="left", padx=(0, 10))

        tk.Button(
            btn_frame,
            text="⚡ 지금 파일명 일괄 변경 실행",
            font=("맑은 고딕", 10, "bold"),
            bg="#4F46E5",
            fg="white",
            relief="flat",
            padx=20,
            pady=7,
            command=self.execute_rename,
        ).pack(side="left")

        self.status_label = tk.Label(btn_frame, text="", font=("맑은 고딕", 10, "bold"), bg="#F8FAFC", fg="#059669")
        self.status_label.pack(side="right")

        # 결과 테이블 (Treeview)
        table_frame = tk.Frame(root, bg="#FFFFFF", padx=20, pady=8)
        table_frame.pack(fill="both", expand=True)

        columns = ("num", "branch", "old_name", "new_name", "status")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)
        self.tree.heading("num", text="순번")
        self.tree.heading("branch", text="인식된 사업소")
        self.tree.heading("old_name", text="현재 파일명 (변경 전)")
        self.tree.heading("new_name", text="표준 파일명 (변경 후)")
        self.tree.heading("status", text="상태")

        self.tree.column("num", width=55, anchor="center")
        self.tree.column("branch", width=110, anchor="center")
        self.tree.column("old_name", width=280, anchor="w")
        self.tree.column("new_name", width=290, anchor="w")
        self.tree.column("status", width=90, anchor="center")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 초기 미리보기 실행
        self.preview_files()

    def browse_folder(self):
        selected = filedialog.askdirectory(initialdir=self.folder_var.get(), title="보고서 파일들이 있는 폴더 선택")
        if selected:
            self.folder_var.set(selected)
            self.preview_files()

    def preview_files(self):
        folder = self.folder_var.get().strip()
        target_date = self.date_var.get().strip() or None
        unify = self.unify_var.get()

        for row in self.tree.get_children():
            self.tree.delete(row)

        if not folder or not os.path.isdir(folder):
            self.status_label.config(text="폴더 경로를 확인해주세요.", fg="#DC2626")
            return

        try:
            conv_doc = self.convert_doc_var.get()
            results = batch_rename_in_folder(folder, unify_date=unify, target_date=target_date, dry_run=True, convert_doc=conv_doc)
            for r in results:
                self.tree.insert("", "end", values=(r["branch_num"], r["branch_name"], r["old_name"], r["new_name"], r["message"]))
            to_change = sum(1 for r in results if r["status"] == "preview")
            self.status_label.config(
                text=f"총 {len(results)}개 파일 감지 (변경/변환 대상: {to_change}개)",
                fg="#4F46E5" if to_change > 0 else "#059669",
            )
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def execute_rename(self):
        folder = self.folder_var.get().strip()
        target_date = self.date_var.get().strip() or None
        unify = self.unify_var.get()
        conv_doc = self.convert_doc_var.get()

        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("폴더 확인", "유효한 폴더를 먼저 선택해주세요.")
            return

        try:
            self.status_label.config(text="⏳ 파일명 변경 및 .doc → .docx 변환 중...", fg="#D97706")
            self.root.update_idletasks()
            results = batch_rename_in_folder(folder, unify_date=unify, target_date=target_date, dry_run=False, convert_doc=conv_doc)
            for row in self.tree.get_children():
                self.tree.delete(row)
            for r in results:
                self.tree.insert("", "end", values=(r["branch_num"], r["branch_name"], r["old_name"], r["new_name"], r["message"]))

            renamed_cnt = sum(1 for r in results if r["status"] == "renamed")
            unchanged_cnt = sum(1 for r in results if r["status"] == "unchanged")
            err_cnt = sum(1 for r in results if r["status"] == "error")

            self.status_label.config(text=f"✅ 변경 완료: {renamed_cnt}개 (유지: {unchanged_cnt}개)", fg="#059669")
            messagebox.showinfo(
                "일괄 변경 완료",
                f"파일명 표준화가 완료되었습니다!\n\n"
                f"• 새로 변경됨: {renamed_cnt}개\n"
                f"• 이미 표준 규격: {unchanged_cnt}개\n"
                f"• 오류/건너뜀: {err_cnt}개",
            )
        except Exception as e:
            messagebox.showerror("오류 발생", f"파일명 변경 중 오류가 발생했습니다:\n{e}")


def main():
    # 커맨드라인 인자로 폴더가 주어지면 즉시 변환하거나 해당 폴더로 GUI 오픈
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        target_dir = sys.argv[2] if len(sys.argv) > 2 else os.getcwd()
        res = batch_rename_in_folder(target_dir, unify_date=True, dry_run=False)
        for r in res:
            print(f"[{r['branch_num']}] {r['old_name']} -> {r['new_name']} ({r['message']})")
        return

    root = tk.Tk()
    app = RenamerApp(root)
    if len(sys.argv) > 1 and os.path.isdir(sys.argv[1]):
        app.folder_var.set(sys.argv[1])
        app.preview_files()
    root.mainloop()


if __name__ == "__main__":
    main()
