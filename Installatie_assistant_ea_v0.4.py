import os
import subprocess
import tkinter as tk
import winreg
import difflib
import threading
import time  # Alleen als je time.sleep gebruikt
import queue
from tkinter import filedialog, messagebox, Toplevel, Label, ttk
from datetime import datetime
from threading import Timer
from tkinter import messagebox
from concurrent.futures import ThreadPoolExecutor

# Pad naar strings.exe (moet in dezelfde map staan als dit script)

import sys

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.getcwd()

strings_exe_path = os.path.join(base_path, "strings.exe")


# Bekende silent install parameters
silent_keywords = ["/S", "/silent", "/quiet", "/verysilent", "/qn"]

# Zorg dat C:\Temp bestaat
log_dir = r"C:\Temp"
os.makedirs(log_dir, exist_ok=True)
log_file_path = os.path.join(log_dir, "install_log.txt")

# Functie om te loggen
def log_installation(file_path, status="Installatie gestart"):
    with open(log_file_path, "a") as log_file:
        log_file.write(f"{datetime.now()}: {status} voor {file_path}\n")

# Functie om strings.exe te gebruiken om silent parameters te detecteren
def detect_silent_parameter(file_path):
    try:
        result = subprocess.run([strings_exe_path, file_path], capture_output=True, text=True, timeout=10)
        output = result.stdout.lower()
        for keyword in silent_keywords:
            if keyword in output:
                return keyword
    except Exception as e:
        print(f"Fout bij strings.exe voor {file_path}: {e}")
    return None

# functie om te controleren of een programma is geïnstalleerd
def is_installed(display_name):
    """
    Controleert of 'display_name' in de Windows Uninstall-registry staat.
    """
    uninstall_paths = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]

    for path in uninstall_paths:
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path)
        except FileNotFoundError:
            continue

        for i in range(winreg.QueryInfoKey(key)[0]):
            subkey_name = winreg.EnumKey(key, i)
            subkey = winreg.OpenKey(key, subkey_name)
            try:
                name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                similarity = difflib.SequenceMatcher(None, display_name.lower(), name.lower()).ratio()
                if similarity > 0.7:
                        return True

            except FileNotFoundError:
                pass
            finally:
                subkey.Close()

        key.Close()

    return False

def update_progress(value):
    progress_bar['value'] = value


# Aantal installaties dat tegelijk mag draaien (pas dit getal aan!)
MAX_PARALLEL_INSTALLS = 2  # Zet op 1 voor sequentieel, 2 of meer voor parallel

# Deze functie voert één installatie uit
def installeer_bestand(file_path, idx, total):
    basename = os.path.basename(file_path)

    # GUI veilig updaten vanuit thread
    root.after(0, lambda: status_label.config(text=f"Bezig met: {basename}"))
    root.after(0, lambda: update_progress(int((idx / total) * 100)))

    log_installation(file_path, "Installatie gestart")

    # Bepaal de juiste parameters en controlleerd of het een .msi is of niet
    param = file_vars[file_path]['param'] or "/S"
    use_silent = file_vars[file_path]['param'] or file_vars[file_path]['override'].get()
    
    if file_path.lower().endswith(".msi"):
        cmd = f'msiexec /i "{file_path}" /quiet /norestart'
    else:
        cmd = f'"{file_path}" {param}' if use_silent else file_path


    # Start installatieproces
    proc = subprocess.Popen(cmd, shell=True)
    proc.wait()

    # Log resultaat
    ret = proc.returncode
    status = "Installatie voltooid" if ret == 0 else f"Fout: exit {ret}"
    log_installation(file_path, status)

# Deze functie start alle installaties via een ThreadPool
def installatie_worker(selected):
    total = len(selected)

    # ThreadPool zorgt voor max X installaties tegelijk
    with ThreadPoolExecutor(max_workers=MAX_PARALLEL_INSTALLS) as executor:
        for idx, file_path in enumerate(selected, start=1):
            executor.submit(installeer_bestand, file_path, idx, total)

    # Als alles klaar is, GUI weer activeren
    root.after(0, lambda: status_label.config(text="Alle installaties voltooid"))
    root.after(0, lambda: start_btn.config(state="normal"))

# Deze functie wordt aangeroepen als je op de knop klikt
def start_installation():
    # Haal geselecteerde bestanden op
    selected = [
        f for f, d in file_vars.items()
        if d['selected'].get() and not d.get('installed', False)
    ]

    if not selected:
        messagebox.showwarning("Geen selectie", "Selecteer ten minste één bestand.")
        return

    # Knop tijdelijk uitschakelen
    start_btn.config(state="disabled")

    # Start installatieproces in aparte thread (zodat GUI niet blokkeert)
    threading.Thread(target=lambda: installatie_worker(selected), daemon=True).start()

# Functie om map te kiezen en bestanden te laden
def select_directory():
    folder_selected = filedialog.askdirectory()
    if not folder_selected:
        return

    # Alles uit de frame verwijderen
    for widget in files_frame.winfo_children():
        widget.destroy()
    file_vars.clear()

    for file_name in os.listdir(folder_selected):
        if not file_name.lower().endswith((".exe", ".msi")):
            continue

        full_path = os.path.join(folder_selected, file_name)
        param = detect_silent_parameter(full_path)

        # --- Nieuw: installatie-check ---
        display_name = os.path.splitext(file_name)[0]
        already = is_installed(display_name)
        # --------------------------------

        selected_var = tk.BooleanVar(value=not already)
        override_var = tk.BooleanVar()

        row = tk.Frame(files_frame)
        row.pack(fill="x", padx=5, pady=2)

        # Bouw de labeltekst
        label_text = file_name
        if param:
            label_text += f" ({param})"
        if already:
            label_text += " — AL GEÏNSTALLEERD"

        # Checkbox voor selectie
        chk = tk.Checkbutton(
            row,
            text=label_text,
            variable=selected_var,
            anchor="w",
            justify="left"
        )
        chk.pack(side="left", fill="x", expand=True)

        # Maak checkbox onklikbaar als al geïnstalleerd
        if already:
            chk.config(state="disabled")

        # Forceer silent-optie
        tk.Checkbutton(row, text="Forceer silent", variable=override_var)\
          .pack(side="right")

        # Sla alle variabelen op, inclusief 'installed'
        file_vars[full_path] = {
            'selected':  selected_var,
            'override':  override_var,
            'param':     param,
            'installed': already         # gebruik later in start_installation()
        }

        log_installation(
            full_path,
            f"Silent scan: {param or 'GEEN'} — {'INSTALLED' if already else 'NOT'}"
        )

# Splashscreen tonen
def show_splash():
    splash = Toplevel()
    splash.overrideredirect(True)
    splash.geometry("300x150+500+300")
    splash.configure(bg="white")
    Label(splash, text="Installatie Assistent", font=("Arial", 16), bg="white").pack(expand=True)
    Timer(2.5, splash.destroy).start()

# GUI opzetten
# --- Root & Scrollable Frame Setup ---
root = tk.Tk()
root.title("Software Installatie Assistent")
root.geometry("800x800")
root.resizable(False, True)

container = tk.Frame(root)
container.pack(fill="both", expand=True, padx=5, pady=5)

canvas   = tk.Canvas(container)
v_scroll = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=v_scroll.set)
v_scroll.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True)

files_frame = tk.Frame(canvas)
canvas.create_window((0, 0), window=files_frame, anchor="nw")

# --- Flipswitches boven de lijst ---
switch_frame = tk.Frame(root)
switch_frame.pack(pady=5)

select_all_var = tk.BooleanVar(value=False)
silent_all_var = tk.BooleanVar(value=False)

def toggle_select_all():
    for data in file_vars.values():
        if not data.get('installed', False):  # Alleen als niet geïnstalleerd
            data['selected'].set(select_all_var.get())

def toggle_silent_all():
    for data in file_vars.values():
        data['override'].set(silent_all_var.get())

tk.Checkbutton(
    switch_frame,
    text="Selecteer alles / Deselecteer alles",
    variable=select_all_var,
    command=toggle_select_all
).pack(side="left", padx=10)

tk.Checkbutton(
    switch_frame,
    text="Silent install all / Silent install none",
    variable=silent_all_var,
    command=toggle_silent_all
).pack(side="left", padx=10)


def on_frame_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))
files_frame.bind("<Configure>", on_frame_configure)

# --- Einde Scrollable Setup ---

# Splashscreen en knoppen
show_splash()

# 1) Knop om map te kiezen
tk.Button(root, text="Selecteer Installatiemap",
          command=select_directory).pack(pady=5)

# 2) Status- en voortgangs‐frame (Stap 3)
status_frame = tk.Frame(root)
status_frame.pack(fill="x", padx=5, pady=5)

status_label = tk.Label(status_frame, text="Klaar", anchor="w")
status_label.pack(fill="x")

progress_bar = ttk.Progressbar(
    status_frame,
    orient="horizontal",
    mode="determinate",
    maximum=100
)
progress_bar.pack(fill="x", pady=2)

# 3) Startknop koppelen aan variabele (Stap 4)
start_btn = tk.Button(
    root,
    text="Start Installatie",
    command=start_installation
)
start_btn.pack(pady=5)

file_vars = {}

# Elegante credit-label onderaan
credit_label = tk.Label(
    root,
    text="Created by Frank van Domselaar  •  BIS|Econocom",
    font=("Segoe UI", 9, "italic"),
    fg="gray30",
    bg=root["bg"],     # Past automatisch bij achtergrondkleur van je GUI
    anchor="center"
)
credit_label.pack(side="bottom", fill="x", pady=(10, 5))
root.mainloop()