import win32evtlog
import tkinter as tk
from datetime import datetime


LOG_NAME = "Security"
SERVER = "localhost"

def get_last_unlock():
    log = win32evtlog.OpenEventLog(SERVER, LOG_NAME)

    flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ

    found_latest = False

    while True:
        events = win32evtlog.ReadEventLog(log, flags, 0)
        if not events:
            break

        for event in events:
            event_id = event.EventID & 0xFFFF
            time = datetime.fromtimestamp(event.TimeGenerated.timestamp())

            time_str = time.strftime("%d/%m/%Y %H:%M:%S")

            if event_id == 4801:
                if not found_latest:
                    found_latest = True
                    continue

                return {
                    "method": "4801",
                    "time": time_str,
                    "user": event.StringInserts[1] if event.StringInserts else "Desconocido"
                }

    return None

def show_window(unlock):
    root = tk.Tk()
    root.iconbitmap(default="")
    root.resizable(False, False)
    root.title("Información de desbloqueo")
    root.attributes("-topmost", True)

    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()

    win_w = int(screen_w * 0.40)
    win_h = int(screen_h * 0.25)

    x = (screen_w - win_w) // 2
    y = (screen_h - win_h) // 2

    root.geometry(f"{win_w}x{win_h}+{x}+{y}")
    root.configure(bg="#2f2f2f") 

    frame = tk.Frame(root, bg="#2f2f2f", bd=2, relief="flat")
    frame.pack(expand=True, fill="both", padx=10, pady=10)

    header_text = "ÚLTIMO DESBLOQUEO ANTERIOR" if unlock else "No se encontró un desbloqueo anterior"
    header = tk.Label(
        frame,
        text=header_text,
        fg="#FFA500",        # naranja
        bg="#2f2f2f",
        font=("Segoe UI", 20, "bold"),
        justify="center"
    )
    header.pack(pady=(10,5))

    separator = tk.Frame(frame, height=2, bg="#FFA500")
    separator.pack(fill="x", padx=20, pady=(0,10))

    if unlock:
        content_text = f"Usuario: {unlock['user']}\nFecha  : {unlock['time']}"
    else:
        content_text = ""

    content = tk.Label(
        frame,
        text=content_text,
        fg="white",
        bg="#2f2f2f",
        font=("Segoe UI", 16),
        justify="center"
    )
    content.pack()

    root.after(5000, root.destroy)
    root.mainloop()

if __name__ == "__main__":
    unlock = get_last_unlock()
    show_window(unlock)
