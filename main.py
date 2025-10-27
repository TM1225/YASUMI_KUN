import csv
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pythoncom
    import win32com.client  # type: ignore
except ImportError:  # pragma: no cover - import guard for PyInstaller build machines
    pythoncom = None
    win32com = None


STATUS_MAP = {
    "休み": 3,
    "ooo": 3,
    "不在": 3,
    "outofoffice": 3,
    "外出": 2,
    "忙しい": 2,
    "busy": 2,
    "仮": 1,
    "tentative": 1,
    "在席": 0,
    "空き": 0,
    "free": 0,
    "他所勤務": 4,
    "workingelsewhere": 4,
}

EXPECTED_FIELDS = [
    "Date",
    "Start",
    "End",
    "Subject",
    "Status",
    "Location",
    "Body",
]


@dataclass
class ScheduleRow:
    row_number: int
    date: str
    start: str
    end: str
    subject: str
    status: str
    location: str
    body: str
    busy: int


class SchedulerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Outlook CSV Scheduler")
        self.root.geometry("960x640")
        self.csv_path: Optional[Path] = None
        self.rows: List[ScheduleRow] = []

        self.path_var = tk.StringVar()

        self._build_ui()

    def _build_ui(self) -> None:
        top_frame = ttk.Frame(self.root, padding=12)
        top_frame.pack(fill="x")

        select_button = ttk.Button(top_frame, text="CSVを選択", command=self.select_csv)
        select_button.pack(side="left")

        path_label = ttk.Label(top_frame, textvariable=self.path_var, wraplength=760)
        path_label.pack(side="left", padx=(12, 0))

        tree_frame = ttk.Frame(self.root, padding=(12, 0, 12, 0))
        tree_frame.pack(fill="both", expand=True)

        columns = ("No", "Date", "Start", "End", "Subject", "Status", "Busy", "Location")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=18)
        headings = {
            "No": "No",
            "Date": "Date",
            "Start": "Start",
            "End": "End",
            "Subject": "Subject",
            "Status": "Status",
            "Busy": "Busy",
            "Location": "Location",
        }

        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=100, anchor="center")

        self.tree.column("Subject", width=200, anchor="w")
        self.tree.column("Status", width=120, anchor="center")
        self.tree.column("Location", width=180, anchor="w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        bottom_frame = ttk.Frame(self.root, padding=12)
        bottom_frame.pack(fill="both")

        self.register_button = ttk.Button(bottom_frame, text="予定を登録", command=self.on_register)
        self.register_button.pack(anchor="w")

        log_label = ttk.Label(bottom_frame, text="処理ログ")
        log_label.pack(anchor="w", pady=(12, 4))

        self.log_text = tk.Text(bottom_frame, height=10, state="disabled")
        self.log_text.pack(fill="both", expand=True)

        log_scroll = ttk.Scrollbar(bottom_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side="right", fill="y")

    def select_csv(self) -> None:
        file_path = filedialog.askopenfilename(
            title="予定CSVを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not file_path:
            return

        try:
            rows = self._load_csv(Path(file_path))
        except ValueError as exc:
            messagebox.showerror("CSV読込エラー", str(exc))
            return

        self.csv_path = Path(file_path)
        self.path_var.set(str(self.csv_path))
        self.rows = rows
        self._refresh_tree()
        self._append_log(f"CSV読込: {len(self.rows)}件を取得")

    def _load_csv(self, path: Path) -> List[ScheduleRow]:
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            headers = reader.fieldnames
            if headers is None:
                raise ValueError("CSVヘッダを検出できません。")
            missing = [field for field in EXPECTED_FIELDS if field not in headers]
            if missing:
                raise ValueError(f"CSVヘッダが不正です。欠損: {', '.join(missing)}")

            rows: List[ScheduleRow] = []
            for index, raw in enumerate(reader, start=2):
                if raw is None:
                    continue
                if not any((value or "").strip() for value in raw.values()):
                    continue
                date = (raw.get("Date") or "").strip()
                start = (raw.get("Start") or "").strip()
                end = (raw.get("End") or "").strip()
                subject = (raw.get("Subject") or "").strip()
                status = (raw.get("Status") or "").strip()
                location = (raw.get("Location") or "").strip()
                body = (raw.get("Body") or "").strip()
                busy = map_busy_status(status)

                rows.append(
                    ScheduleRow(
                        row_number=index,
                        date=date,
                        start=start,
                        end=end,
                        subject=subject,
                        status=status,
                        location=location,
                        body=body,
                        busy=busy,
                    )
                )
        return rows

    def _refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for idx, row in enumerate(self.rows, start=1):
            self.tree.insert(
                "",
                "end",
                values=(
                    idx,
                    row.date,
                    row.start,
                    row.end,
                    row.subject,
                    row.status,
                    row.busy,
                    row.location,
                ),
            )

    def on_register(self) -> None:
        if not self.rows:
            messagebox.showwarning("CSV未選択", "先にCSVファイルを読み込んでください。")
            return
        if win32com is None or pythoncom is None:
            messagebox.showerror("依存モジュール不足", "win32com.client と pythoncom が必要です。")
            return

        self.register_button.config(state="disabled")
        self._append_log("予定登録を開始")

        worker = threading.Thread(target=self._register_appointments, daemon=True)
        worker.start()

    def _register_appointments(self) -> None:
        pythoncom.CoInitialize()
        try:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
            except Exception as exc:  # pragma: no cover - COM errors depend on host
                self._append_log(f"[NG] Outlook起動に失敗: {exc}")
                return

            for row in self.rows:
                start_dt, end_dt, error = self._parse_datetimes(row)
                if error:
                    self._append_log(f"[NG] 行{row.row_number} {error}")
                    continue
                try:
                    appointment = outlook.CreateItem(1)  # 1 = olAppointmentItem
                    appointment.Subject = row.subject or ""
                    appointment.Start = start_dt
                    appointment.End = end_dt
                    appointment.BusyStatus = row.busy
                    if row.location:
                        appointment.Location = row.location
                    if row.body:
                        appointment.Body = row.body
                    appointment.Save()
                    self._append_log(
                        f"[OK] {row.date} {row.start}-{row.end} {row.subject or '(件名なし)'} BusyStatus={row.busy}"
                    )
                except Exception as exc:  # pragma: no cover - COM errors depend on host
                    self._append_log(f"[NG] 行{row.row_number} Outlook登録失敗: {exc}")
        finally:
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self.register_button.config(state="normal"))
            self._append_log("予定登録を終了")

    def _parse_datetimes(self, row: ScheduleRow):
        if not row.date:
            return None, None, "Date が空です"
        if not row.start or not row.end:
            return None, None, "Start/End 不正"
        try:
            start_dt = datetime.strptime(f"{row.date} {row.start}", "%Y-%m-%d %H:%M")
            end_dt = datetime.strptime(f"{row.date} {row.end}", "%Y-%m-%d %H:%M")
        except ValueError:
            return None, None, "Start/End 不正"
        if end_dt <= start_dt:
            return None, None, "End は Start より後である必要があります"
        return start_dt, end_dt, None

    def _append_log(self, message: str) -> None:
        def write_line() -> None:
            self.log_text.config(state="normal")
            self.log_text.insert("end", f"{message}\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")

        self.root.after(0, write_line)


def map_busy_status(value: str) -> int:
    key = (value or "").strip().lower()
    if not key:
        return 2
    return STATUS_MAP.get(key, 2)


def main() -> None:
    root = tk.Tk()
    app = SchedulerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
