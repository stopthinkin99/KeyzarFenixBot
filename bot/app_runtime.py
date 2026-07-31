from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

from config import POLL_SECONDS
from email_reader.invoice_sender import send_current_keyzar_invoice
from excel_reports.daily_report import get_current_invoice_path
from email_reader.outlook_reader import (
    OutlookMailbox,
    OutlookReader,
    load_saved_mailbox,
    save_outlook_mailbox,
)
from fenix.login_session import save_fenix_login_session
from processing.workflow import run_once


class KeyzarFenixApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()

        self.title("Keyzar Fenix Bot - v1")
        self.geometry("960x700")
        self.minsize(860, 600)

        self.stop_event = threading.Event()
        self.worker_thread: threading.Thread | None = None
        self.log_queue: queue.Queue[str] = queue.Queue()

        self._build_ui()
        self.after(150, self._drain_log_queue)

        # Start monitoring automatically after the window opens.
        self.after(800, self.start_bot)

    def _build_ui(self) -> None:
        ttk.Label(
            self,
            text="Keyzar Fenix Bot",
            font=("Segoe UI", 19, "bold"),
        ).pack(pady=(18, 5))

        ttk.Label(
            self,
            text=(
                "Reads the latest 30 Keyzar emails, checks Fenix, "
                "blocks available stones, and appends them to the current invoice."
            ),
            wraplength=820,
            justify="center",
        ).pack(pady=(0, 14))

        status_frame = ttk.Frame(self)
        status_frame.pack(fill="x", padx=24)

        ttk.Label(
            status_frame,
            text="Status:",
            font=("Segoe UI", 10, "bold"),
        ).pack(side="left")

        self.status_var = tk.StringVar(value="Starting")
        ttk.Label(
            status_frame,
            textvariable=self.status_var,
        ).pack(side="left", padx=(8, 24))

        ttk.Label(
            status_frame,
            text="Outlook mailbox:",
            font=("Segoe UI", 10, "bold"),
        ).pack(side="left")

        saved_mailbox = load_saved_mailbox()
        self.mailbox_var = tk.StringVar(
            value=saved_mailbox.get("display_name", "") or "Default Outlook Inbox"
        )
        ttk.Label(
            status_frame,
            textvariable=self.mailbox_var,
        ).pack(side="left", padx=(8, 0))

        button_frame = ttk.Frame(self)
        button_frame.pack(fill="x", padx=24, pady=14)

        self.start_button = ttk.Button(
            button_frame,
            text="Start Bot",
            command=self.start_bot,
        )
        self.start_button.pack(side="left", padx=(0, 8))

        self.stop_button = ttk.Button(
            button_frame,
            text="Stop Bot",
            command=self.stop_bot,
        )
        self.stop_button.pack(side="left", padx=(0, 8))

        self.outlook_button = ttk.Button(
            button_frame,
            text="Connect Outlook",
            command=self.connect_outlook,
        )
        self.outlook_button.pack(side="left", padx=(0, 8))

        self.login_button = ttk.Button(
            button_frame,
            text="Login to Fenix",
            command=self.login_to_fenix,
        )
        self.login_button.pack(side="left", padx=(0, 8))

        self.open_excel_button = ttk.Button(
            button_frame,
            text="Open Excel",
            command=self.open_excel,
        )
        self.open_excel_button.pack(side="left", padx=(0, 8))

        self.send_button = ttk.Button(
            button_frame,
            text="Send Now",
            command=self.send_now,
        )
        self.send_button.pack(side="left")

        ttk.Label(
            self,
            text=(
                "Use Connect Outlook once on a new computer to select the mailbox "
                "that contains Keyzar orders. Send Now emails the current Excel file "
                "to Pune and deletes the local file only after Outlook accepts it."
            ),
            wraplength=840,
            justify="left",
        ).pack(fill="x", padx=24, pady=(0, 10))

        self.log_box = scrolledtext.ScrolledText(
            self,
            wrap="word",
            state="disabled",
            font=("Consolas", 9),
        )
        self.log_box.pack(
            fill="both",
            expand=True,
            padx=24,
            pady=(0, 18),
        )

        self.protocol("WM_DELETE_WINDOW", self._close_app)

    def log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put(f"[{timestamp}] {message}")

    def _drain_log_queue(self) -> None:
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_box.configure(state="normal")
                self.log_box.insert("end", message + "\n")
                self.log_box.see("end")
                self.log_box.configure(state="disabled")
        except queue.Empty:
            pass

        self.after(150, self._drain_log_queue)

    def start_bot(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            self.log("The bot is already running.")
            return

        self.stop_event.clear()
        self.status_var.set("Running")
        self.log(
            f"Monitoring started. Outlook will be checked every "
            f"{POLL_SECONDS} seconds."
        )

        self.worker_thread = threading.Thread(
            target=self._monitor_loop,
            daemon=True,
        )
        self.worker_thread.start()

    def stop_bot(self) -> None:
        self.stop_event.set()
        self.status_var.set("Stopping")
        self.log("Stop requested. The current cycle will finish first.")

    def _monitor_loop(self) -> None:
        while not self.stop_event.is_set():
            self.log("Starting Outlook/Fenix processing cycle.")

            try:
                result = run_once(log_callback=self.log)
                self.log(
                    "Cycle finished: "
                    f"processed={result['processed']}, "
                    f"skipped={result['skipped']}, "
                    f"failed={result['failed']}."
                )
            except Exception as exc:
                self.log(
                    f"Cycle failed: {type(exc).__name__}: {exc}"
                )

            if self.stop_event.wait(POLL_SECONDS):
                break

        self.status_var.set("Stopped")
        self.log("Bot stopped.")


    def connect_outlook(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo(
                "Stop Bot First",
                "Please click Stop Bot and wait until the status says Stopped before selecting an Outlook mailbox.",
            )
            return

        self.outlook_button.configure(state="disabled")
        self.status_var.set("Checking Outlook")
        self.log("Reading available Classic Outlook mailboxes...")
        threading.Thread(target=self._discover_outlook_worker, daemon=True).start()

    def _discover_outlook_worker(self) -> None:
        reader = OutlookReader()
        mailboxes = []
        try:
            reader.connect()
            mailboxes = reader.list_mailboxes()
        except Exception as exc:
            self.log(f"Outlook connection failed: {type(exc).__name__}: {exc}")
            self.after(0, lambda error=str(exc): messagebox.showerror("Outlook Connection Failed", error))
        finally:
            reader.disconnect()
            self.after(0, lambda: self.outlook_button.configure(state="normal"))
            self.status_var.set("Stopped")

        if mailboxes:
            self.after(0, lambda: self._show_mailbox_dialog(mailboxes))

    def _show_mailbox_dialog(self, mailboxes: list[OutlookMailbox]) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Select Outlook Mailbox")
        dialog.geometry("650x280")
        dialog.transient(self)
        dialog.grab_set()

        ttk.Label(
            dialog,
            text="Select the mailbox containing Keyzar orders:",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", padx=20, pady=(20, 8))

        choices = [
            f"{m.display_name} — {m.total_items} items, {m.unread_items} unread"
            for m in mailboxes
        ]
        selected_text = tk.StringVar(value=choices[0])
        combo = ttk.Combobox(
            dialog,
            textvariable=selected_text,
            values=choices,
            state="readonly",
            width=75,
        )
        combo.pack(fill="x", padx=20, pady=(0, 14))

        ttk.Label(
            dialog,
            text=(
                "After saving, the app will show the five newest subjects from "
                "that Inbox so you can confirm it is the correct mailbox."
            ),
            wraplength=600,
            justify="left",
        ).pack(anchor="w", padx=20, pady=(0, 18))

        def save_selection() -> None:
            index = combo.current()
            if index < 0:
                return
            mailbox = mailboxes[index]
            save_outlook_mailbox(
                store_id=mailbox.store_id,
                display_name=mailbox.display_name,
            )
            self.mailbox_var.set(mailbox.display_name)
            dialog.destroy()
            self.log(f"Saved Outlook mailbox: {mailbox.display_name}")
            threading.Thread(target=self._test_selected_mailbox_worker, daemon=True).start()

        ttk.Button(dialog, text="Save and Test", command=save_selection).pack(pady=(0, 16))

    def _test_selected_mailbox_worker(self) -> None:
        reader = OutlookReader()
        try:
            reader.connect()
            self.log(f"Connected to Outlook mailbox: {reader.mailbox_display_name}")
            preview = reader.get_recent_email_preview(5)
            self.log(f"Newest messages found: {len(preview)}")
            for email in preview:
                received = (
                    email.received_time.strftime("%m/%d/%Y %I:%M %p")
                    if email.received_time
                    else "Unknown time"
                )
                self.log(f"  {received} | {email.sender_name} | {email.subject}")
            self.after(
                0,
                lambda: messagebox.showinfo(
                    "Outlook Connected",
                    "The Outlook mailbox was saved and tested. Review the app log to confirm the newest subjects.",
                ),
            )
        except Exception as exc:
            self.log(f"Outlook test failed: {type(exc).__name__}: {exc}")
            self.after(0, lambda error=str(exc): messagebox.showerror("Outlook Test Failed", error))
        finally:
            reader.disconnect()

    def login_to_fenix(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo(
                "Stop Bot First",
                (
                    "Please click Stop Bot and wait until the status says "
                    "Stopped before opening the Fenix login window."
                ),
            )
            return

        proceed = messagebox.askyesno(
            "Login to Fenix",
            (
                "A visible Edge window will open.\n\n"
                "Log in to Fenix and open Search Stock. Then return to this "
                "app and confirm that login is complete.\n\n"
                "Continue?"
            ),
        )

        if not proceed:
            return

        self.login_button.configure(state="disabled")
        self.status_var.set("Waiting for Fenix login")

        threading.Thread(
            target=self._login_worker,
            daemon=True,
        ).start()

    def _login_worker(self) -> None:
        try:
            save_fenix_login_session(
                confirmation_callback=self._confirm_login_complete,
                log_callback=self.log,
            )

            self.status_var.set("Stopped")
            self.after(
                0,
                lambda: messagebox.showinfo(
                    "Fenix Login Saved",
                    (
                        "The Fenix login was saved successfully.\n\n"
                        "Click Start Bot to resume monitoring."
                    ),
                ),
            )

        except Exception as exc:
            self.status_var.set("Stopped")
            self.log(
                f"Fenix login failed: {type(exc).__name__}: {exc}"
            )
            self.after(
                0,
                lambda error=str(exc): messagebox.showerror(
                    "Fenix Login Failed",
                    error,
                ),
            )

        finally:
            self.after(
                0,
                lambda: self.login_button.configure(state="normal"),
            )

    def _confirm_login_complete(self) -> bool:
        event = threading.Event()
        answer = {"value": False}

        def ask() -> None:
            answer["value"] = messagebox.askyesno(
                "Save Fenix Login",
                (
                    "Have you finished logging in and opened the "
                    "Search Stock page?"
                ),
            )
            event.set()

        self.after(0, ask)
        event.wait()
        return answer["value"]

    def open_excel(self) -> None:
        workbook_path = get_current_invoice_path()

        if workbook_path is None or not workbook_path.exists():
            messagebox.showinfo(
                "No Excel File",
                "There is no unsent Keyzar invoice available yet.",
            )
            return

        try:
            os.startfile(str(workbook_path))
            self.log(
                f"Opened Excel file: {workbook_path.name}"
            )

        except OSError as exc:
            self.log(
                f"Could not open Excel: "
                f"{type(exc).__name__}: {exc}"
            )

            messagebox.showerror(
                "Open Excel Failed",
                (
                    f"Could not open:\n"
                    f"{workbook_path}\n\n"
                    f"{exc}"
                ),
            )

    def send_now(self) -> None:
        proceed = messagebox.askyesno(
            "Send Current Invoice",
            (
                "Send the current Keyzar invoice now?\n\n"
                "To: salesinvoice@egonservices.com\n"
                "CC: fenixny.bizops@fenixdiamonds.com\n"
                "From: sales@fenixdiamonds.com\n\n"
                "The local Excel file will be deleted only after "
                "Outlook accepts the message."
            ),
        )

        if not proceed:
            return

        self.send_button.configure(state="disabled")
        self.status_var.set("Sending invoice")

        threading.Thread(
            target=self._send_worker,
            daemon=True,
        ).start()

    def _send_worker(self) -> None:
        try:
            sent_path = send_current_keyzar_invoice(
                log_callback=self.log,
            )

            self.status_var.set(
                "Running"
                if self.worker_thread and self.worker_thread.is_alive()
                else "Stopped"
            )

            if sent_path is None:
                self.after(
                    0,
                    lambda: messagebox.showinfo(
                        "No Invoice",
                        "There is no unsent Keyzar invoice to send.",
                    ),
                )
                return

            filename = Path(sent_path).name
            self.after(
                0,
                lambda name=filename: messagebox.showinfo(
                    "Invoice Sent",
                    (
                        f"{name} was submitted successfully.\n\n"
                        "The local invoice was deleted. The next blocked stone "
                        "will start a fresh invoice."
                    ),
                ),
            )

        except Exception as exc:
            self.status_var.set(
                "Running"
                if self.worker_thread and self.worker_thread.is_alive()
                else "Stopped"
            )
            self.log(
                f"Invoice send failed: {type(exc).__name__}: {exc}"
            )
            self.after(
                0,
                lambda error=str(exc): messagebox.showerror(
                    "Invoice Send Failed",
                    (
                        f"{error}\n\n"
                        "The Excel file was not deleted."
                    ),
                ),
            )

        finally:
            self.after(
                0,
                lambda: self.send_button.configure(state="normal"),
            )

    def _close_app(self) -> None:
        self.stop_event.set()
        self.destroy()


if __name__ == "__main__":
    KeyzarFenixApp().mainloop()
