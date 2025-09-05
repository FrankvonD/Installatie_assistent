import os
import subprocess
import tkinter as tk
import winreg
import difflib
import threading
import sys
import csv
import ctypes
from tkinter import filedialog, messagebox, Toplevel, Label, ttk
from datetime import datetime
from threading import Timer

# Controleer of het script met administratorrechten draait
# Dit is nodig voor sommige installaties die systeemwijzigingen vereisen
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    messagebox.showwarning("Administratorrechten vereist", "Start dit script als administrator om installaties correct uit te voeren.")


file_vars = {}
# Maximaal 2 installaties tegelijk
installatie_semaphore = threading.Semaphore(2)


# Bepaal het pad naar de juiste versie van strings.exe (x64 of x86)
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS  # Als het script is gebundeld met PyInstaller
else:
    base_path = os.getcwd()   # Normale modus

# Kies automatisch de juiste versie van strings.exe op basis van systeemarchitectuur
arch = 'x64' if sys.maxsize > 2**32 else 'x86'
strings_exe_path = os.path.join(base_path, f"strings_{arch}.exe")



# Bekende silent install parameters
silent_keywords = ["/S", "/silent", "/quiet", "/verysilent", "/qn"]

# Zorg dat C:\Temp bestaat
log_dir = r"C:\Temp"
os.makedirs(log_dir, exist_ok=True)
log_file_path = os.path.join(log_dir, "install_log.txt")

# Functie om te loggen
def log_installation(file_path, status="Installatie gestart"):
    with open(log_file_path, "a", encoding="utf-8") as log_file:
        log_file.write(f"{datetime.now()}: {status} voor {file_path}\n")


# Functie om het logbestand te openen in Kladblok (of standaard teksteditor)

def open_logbestand(file_path=None):
    try:
        os.startfile(log_file_path)
    except Exception as e:
        messagebox.showerror("Fout", f"Kan logbestand niet openen:\n{e}")
        root.after(0, lambda: file_vars[file_path]['status_label'].config(text="❌ FOUT", fg="red"))



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

# Bekende installaties met afwijkende namen in het register
known_display_patterns = {
    "dotnet-sdk": "Microsoft .NET SDK",
    "1password": "1Password",
    "7z": "7-Zip",
    "admin by request": "Admin By Request",
    "agent": "Advanced Monitoring Agent Network Management",
    "crowdstrike": "CrowdStrike Falcon Sensor",
    "dcu": "Dell Command | Update",
    "exclaimer": "Exclaimer Cloud Signature Update Agent",
    "myportal": "myPortal@Work",
    "office": "Microsoft Office",
    "visio": "Microsoft Visio",
    "displaylink": "DisplayLink Graphics",
    "sentinelone": "SentinelOne Agent"
}
# Deze dictionary bevat mappings van bestandsnamen naar bekende DisplayNames
# Deze kunnen worden uitgebreid met meer bekende software die niet automatisch wordt herkend
# Bijvoorbeeld:
# "example_installer": "Example Software",


# functie om te controleren of een programma is geïnstalleerd

def is_installed(file_path):
    """
    Controleert of 'file_path' in de Windows Uninstall-registry staat
    en toont ook de DisplayVersion indien beschikbaar.
    """

    uninstall_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Uninstall")
    ]

    # ✅ Functie om bestandsnaam te matchen met bekende display-namen
    def match_display_name(file_name):
        base_name = os.path.splitext(os.path.basename(file_name))[0].lower()
        for pattern, display in known_display_patterns.items():
            if pattern in base_name:
                return display
        return base_name  # fallback

    zoeknaam = match_display_name(file_path).lower()

    for hive, path in uninstall_paths:
        try:
            key = winreg.OpenKey(hive, path)
        except FileNotFoundError:
            continue

        for i in range(winreg.QueryInfoKey(key)[0]):
            try:
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name)
                name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                similarity = difflib.SequenceMatcher(None, zoeknaam, name.lower()).ratio()
                if similarity > 0.7:
                    try:
                        version, _ = winreg.QueryValueEx(subkey, "DisplayVersion")
                    except FileNotFoundError:
                        version = "onbekend"
                    return True, version
            except Exception:
                continue
            finally:
                subkey.Close()
        key.Close()

    return False, None

# 📤 Export alle geïnstalleerde software naar een CSV-bestand
def export_installed_software_to_csv(csv_path="C:/Temp/geinstalleerde_software.csv"):
    uninstall_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Uninstall")
    ]

    with open(csv_path, mode="w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(["DisplayName", "DisplayVersion", "Publisher", "InstallLocation"])

        for hive, path in uninstall_paths:
            try:
                key = winreg.OpenKey(hive, path)
            except FileNotFoundError:
                continue

            for i in range(winreg.QueryInfoKey(key)[0]):
                try:
                    subkey_name = winreg.EnumKey(key, i)
                    subkey = winreg.OpenKey(key, subkey_name)

                    name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                    try:
                        version, _ = winreg.QueryValueEx(subkey, "DisplayVersion")
                    except FileNotFoundError:
                        version = "onbekend"
                    try:
                        publisher, _ = winreg.QueryValueEx(subkey, "Publisher")
                    except FileNotFoundError:
                        publisher = "onbekend"
                    try:
                        location, _ = winreg.QueryValueEx(subkey, "InstallLocation")
                    except FileNotFoundError:
                        location = "onbekend"

                    writer.writerow([name, version, publisher, location])
                except Exception:
                    continue
                finally:
                    subkey.Close()
            key.Close()

    print(f"✅ CSV opgeslagen op: {csv_path}")


def update_progress(value):
    progress_bar['value'] = value

# Functie om één installatiebestand uit te voeren
def installeer_bestand(file_path, index, total):
    try:
        with installatie_semaphore:
            log_installation(file_path, "Installatie gestart")
            root.after(0, lambda: file_vars[file_path]['status_label'].config(text="Bezig...", fg="blue"))

            param = file_vars[file_path]['param']
            override = file_vars[file_path]['override'].get()

            # Controleer bestandstype
            if file_path.lower().endswith(".msi"):
                command = ["msiexec", "/i", file_path]
                if override:
                    command += ["/qn", "/norestart", "/log", "install.log"]
                elif param:
                    command += param.split()
            elif file_path.lower().endswith(".exe"):
                command = [file_path]
                if override:
                    command += ["/S", "/silent", "/verysilent", "/quiet", "/qn", "/norestart", "/log=install.log"]
                elif param:
                    command += param.split()
            else:
                log_installation(file_path, "FOUT: Ongeldig bestandstype")
                root.after(0, lambda: file_vars[file_path]['status_label'].config(text="❌ Ongeldig bestand", fg="red"))
                return

            try:
                subprocess.run(command, timeout=300, check=True)
                log_installation(file_path, "Installatie voltooid")
                root.after(0, lambda: file_vars[file_path]['status_label'].config(text="✅ VOLTOOID", fg="green"))
            except Exception as e:
                log_installation(file_path, f"FOUT: {e}")
                root.after(0, lambda e=e: messagebox.showerror("Installatiefout", f"Fout bij {file_path}:\n{e}"))
                root.after(0, lambda: file_vars[file_path]['status_label'].config(text="❌ FOUT", fg="red"))

    finally:
        progress = int((index / total) * 100)
        root.after(0, lambda: update_progress(progress))


def installatie_worker(selected):
    total = len(selected)
    threads = []

    # Start voor elk installatiebestand een aparte thread
    for idx, file_path in enumerate(selected, start=1):
        t = threading.Thread(target=installeer_bestand, args=(file_path, idx, total))
        t.start()
        threads.append(t)

    # Wacht in aparte thread tot alles klaar is
    def wacht_op_threads():
        for t in threads:
            t.join()
        root.after(0, lambda: status_label.config(text="Alle installaties voltooid"))
        root.after(0, lambda: start_btn.config(state="normal"))

    threading.Thread(target=wacht_op_threads).start()

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

# 🔄 Automatisch mapping maken van installer-bestanden naar DisplayNames uit CSV
def generate_known_display_names_from_csv(installer_folder, csv_path="C:/Temp/geinstalleerde_software.csv"):
    mapping = {}

    # 1. Lees alle DisplayNames uit de CSV
    display_names = []
    with open(csv_path, mode="r", encoding="utf-8") as csv_file:
        reader = csv.DictReader(csv_file)
        for row in reader:
            display_names.append(row["DisplayName"].lower())

    # 2. Loop door alle installer-bestanden
    for file_name in os.listdir(installer_folder):
        if not file_name.lower().endswith((".exe", ".msi")):
            continue

        base_name = os.path.splitext(file_name)[0].lower()

        # 3. Zoek de beste match met DisplayName
        best_match = difflib.get_close_matches(base_name, display_names, n=1, cutoff=0.6)
        if best_match:
            mapping[base_name] = best_match[0]

    return mapping

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
        full_path = os.path.normpath(full_path)  # Normaliseer het pad

        # Detecteer standaard silent parameter
        param = detect_silent_parameter(full_path)

        # Speciaal geval voor SentinelOne.msi
        if "sentinelone" in file_name.lower():
            param = '/q SITE_TOKEN="eyJ1cmwiOiAiaHR0cHM6Ly9ldWNlMS1zd3ByZDIuc2VudGluZWxvbmUubmV0IiwgInNpdGVfa2V5IjogImYzYjg2NmQwMTU5NzNiY2MifQ==" /NORESTART'


        # --- Nieuw: installatie-check ---
        if not hasattr(root, "dynamic_display_names"):
            root.dynamic_display_names = generate_known_display_names_from_csv(folder_selected)

        display_name = root.dynamic_display_names.get(os.path.splitext(file_name)[0].lower(), os.path.splitext(file_name)[0])

        already, version = is_installed(display_name)

        # --------------------------------

        selected_var = tk.BooleanVar(value=not already)
        override_var = tk.BooleanVar()

        row = tk.Frame(files_frame)
        row.pack(fill="x", padx=5, pady=2)

        # Voeg een statuslabel toe aan de rechterkant van de rij
        status_lbl = tk.Label(row, text="Wachtend", width=15, anchor="w", fg="gray")
        status_lbl.pack(side="right", padx=5)

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
            'installed': already,         # gebruik later in start_installation()
            'version': version,
            'status_label': status_lbl, 
        }


        status = "✅ GEÏNSTALLEERD" if already else "❌ NIET GEÏNSTALLEERD"
        log_installation(full_path, f"Scanresultaat: {status} — Versie: {version or 'onbekend'} — Silent param: {param or 'GEEN'}")


        # Activeer de knoppen als er bestanden zijn gevonden
        select_all_btn.config(state="normal")
        silent_all_btn.config(state="normal")


# Splashscreen tonen
def show_splash():
    if not is_admin():
        messagebox.showwarning(
            "Administratorrechten vereist",
            "Start dit script als administrator om installaties correct uit te voeren."
    )

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

# Nieuw frame voor alle drie de knoppen naast elkaar
top_button_frame = tk.Frame(root)
top_button_frame.pack(pady=5)


# Scrollable frame voor de installatiebestanden
container = tk.Frame(root)
container.pack(fill="both", expand=True, padx=5, pady=5)

canvas   = tk.Canvas(container)
v_scroll = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=v_scroll.set)
v_scroll.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True)

files_frame = tk.Frame(canvas)
canvas.create_window((0, 0), window=files_frame, anchor="nw")


# --- Toggle-knoppen voor selecteren en silent installeren ---
select_all_state = {"selected": False}  # Start met alles uit

def toggle_select_all():
    new_state = not select_all_state["selected"]
    for data in file_vars.values():
        if not data.get('installed', False):
            data['selected'].set(new_state)
    select_all_btn.config(
        text="Deselecteer alles" if new_state else "Selecteer alles"
    )
    select_all_state["selected"] = new_state

silent_all_state = {"override": False}

def toggle_silent_all():
    new_state = not silent_all_state["override"]
    for data in file_vars.values():
        data['override'].set(new_state)
    silent_all_btn.config(
        text="Silent install uit" if new_state else "Silent install aan"
    )
    silent_all_state["override"] = new_state


# Knoppen in top_button_frame plaatsen

# 1. Eerst: Selecteer Installatiemap
tk.Button(
    top_button_frame,
    text="Selecteer Installatiemap",
    command=select_directory
).pack(side="left", padx=5)

# 2. Dan: Selecteer alles
select_all_btn = tk.Button(
    top_button_frame,
    text="Selecteer alles",
    command=toggle_select_all,
    state="disabled"
)
select_all_btn.pack(side="left", padx=5)

# 3. Dan: Silent install aan
silent_all_btn = tk.Button(
    top_button_frame,
    text="Silent install aan",
    command=toggle_silent_all,
    state="disabled"
)
silent_all_btn.pack(side="left", padx=5)

# 4. Start Installatie
start_btn = tk.Button(
    top_button_frame,
    text="Start Installatie",
    command=start_installation,
    width=20
)
start_btn.pack(side="left", padx=5)

# 5. Toon Installatielog
tk.Button(
    top_button_frame,
    text="Toon Installatielog",
    command=open_logbestand,
    width=20
).pack(side="left", padx=5)



# --- Einde Toggle-knoppen ---

# Zorg dat de canvas automatisch de grootte aanpast
def on_frame_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))
files_frame.bind("<Configure>", on_frame_configure)

# --- Einde Scrollable Setup ---

# Splashscreen en knoppen
show_splash()


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

export_installed_software_to_csv()
# 🔄 Automatisch mapping maken van installer-bestanden naar DisplayNames uit CSV
def generate_known_display_names_from_csv(installer_folder, csv_path="C:/Temp/geinstalleerde_software.csv"):
    mapping = {}

    # 1. Lees alle DisplayNames uit de CSV
    display_names = []
    with open(csv_path, mode="r", encoding="utf-8") as csv_file:
        reader = csv.DictReader(csv_file)
        for row in reader:
            display_names.append(row["DisplayName"].lower())

    # 2. Loop door alle installer-bestanden
    for file_name in os.listdir(installer_folder):
        if not file_name.lower().endswith((".exe", ".msi")):
            continue

        base_name = os.path.splitext(file_name)[0].lower()

        # 3. Zoek de beste match met DisplayName
        best_match = difflib.get_close_matches(base_name, display_names, n=1, cutoff=0.6)
        if best_match:
            mapping[base_name] = best_match[0]

    # 4. Print de mapping (optioneel: schrijf naar bestand)
    print("✅ Automatisch gegenereerde known_display_names:")
    for k, v in mapping.items():
        print(f'"{k}": "{v}",')

    return mapping


root.mainloop()