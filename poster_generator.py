import os
import re
import requests
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright
import sys
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import threading # To run blocking tasks in a separate thread

# Set Playwright browsers path for PyInstaller executable
if getattr(sys, 'frozen', False):
    # Running as a bundled executable
    # Assumes browsers are in a 'playwright/browsers' subdirectory relative to the executable
    executable_dir = os.path.dirname(sys.executable)
    playwright_browsers_dir = os.path.join(executable_dir, 'playwright', 'browsers')
    os.environ['PLAYWRIGHT_BROWSERS_PATH'] = playwright_browsers_dir

translations = {
    # Wetter
    'Dry': 'Trocken',
    'Damp': 'Feucht',
    'Wet': 'Nass',
    'Morning' : 'Morgens',
    'Noon': 'Mittag',
    'Evening': 'Abends',
    'Crisp' : 'Heiter',
    'Clear' : 'Klare Sicht',
    'Hazy': 'Dunst',
    'PartCloud': 'Teils bewölkt',
    'HeavyCloud': 'Stark bewölkt',
    'LightCloud': 'Leicht bewölkt',
    'Fog': 'Nebel',
    'NoRain': 'Kein Regen',
    'Rain': 'Regen',
    'Windy': 'Windig',
    'LightRain': 'Regen',
    'HeavyRain': 'Starkregen',
    'HeavyFog': 'Nebel',
    'LightFog': 'Neblig',
    'HeavySnow': 'Starker Schneefall',
    'LightSnow': 'Leichter Schneefall',

    # Zustand / Surface
    'New': 'Guter Zustand',
    'Normal': 'Leicht verschmutzt',
    'Worn': 'Stark abgefahren',
    'Tarmac': 'Asphalt',
    'Gravel': 'Schotter',
    'Snow': 'Schnee',

    # Service
    'Road Side Service': 'Reparaturzeit unterwegs',
    'Service Park': 'Servicepark',
    'minutes': 'Minuten',
    'mechanic': 'Mechaniker',
    'Inexperienced': 'unerfahrene',
    'Proficient': 'versierte',
    'Competent': 'kompetente',
    'Skilled': 'erfahrene',
    'Expert': 'sehr gute',

    # Group label
    'Group': 'Gruppe',

    # Tabellenköpfe
    'Stage name': 'Wertungsprüfung',
    'Distance': 'Distanz',
    'Surface': 'Zustand',
    'Weather': 'Bedingungen',
}

# --- Core Logic Functions (Adapted for GUI) ---

def translate_iteratively(text, dictionary):
    # First, handle specific multi-word phrases if they match exactly (case-insensitive)
    for en_key, de_value in dictionary.items():
         if ' ' in en_key: # Check if it's a multi-word key
             if en_key.lower() == text.lower():
                 return de_value

    translated_text = text
    sorted_translations = sorted(dictionary.items(), key=lambda item: len(item[0]), reverse=True)
    for en_key, de_value in sorted_translations:
        pattern = r'\b' + re.escape(en_key) + r'\b'
        try:
            translated_text = re.sub(pattern, de_value, translated_text, flags=re.IGNORECASE)
        except re.error:
            translated_text = translated_text.replace(en_key, de_value)
    return translated_text.strip()

def fetch_html_content(url, status_callback):
    if not url or not (url.startswith('http://') or url.startswith('https://')):
        return None, "Ungültige URL: Bitte gib eine gültige URL ein."
    try:
        status_callback("Rufe Daten ab...")
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        try:
            content = response.content.decode('latin-1')
        except UnicodeDecodeError:
            response.encoding = response.apparent_encoding or 'utf-8'
            content = response.text
        status_callback("Daten erfolgreich abgerufen.")
        return content, None # Return content and no error
    except requests.exceptions.RequestException as e:
        return None, f"Fehler beim Abrufen: Konnte die URL nicht laden:\n{e}"
    except Exception as e:
        return None, f"Ein unerwarteter Fehler ist aufgetreten:\n{e}"

def generate_poster_data(html_content, translate, status_callback):
    if not html_content:
        return None, "Kein HTML-Inhalt zum Verarbeiten."

    rally_name = "poster" # Default
    total_distance = ""
    car_name = ""
    legs = []
    error_message = None

    try:
        status_callback("Verarbeite HTML...")
        soup = BeautifulSoup(html_content, 'lxml')

        main_content_td = soup.find('td', class_='szdb', style=lambda value: value and 'padding:5px' in value)
        if not main_content_td:
            return None, "Konnte den Hauptinhaltsbereich nicht finden."

        # Extract Rally Name, Distance, Car Group
        rally_info_table = main_content_td.find_all('table', recursive=False)[0]
        rally_name_tag = rally_info_table.find('tr', class_='fejlec').find('td').find('b')
        rally_name = rally_name_tag.get_text(strip=True) if rally_name_tag else 'N/A'

        rally_info_rows = rally_info_table.find_all('tr')
        for row in rally_info_rows:
            cells = row.find_all('td')
            if len(cells) > 1:
                first_cell_text = cells[0].get_text(strip=True)
                if first_cell_text == 'Total Distance Rally:':
                    total_distance = cells[1].get_text(strip=True)
                elif first_cell_text == 'Car Groups:':
                    car_name = cells[1].get_text(strip=True)

        # Extract Stage Data
        all_tables = main_content_td.find_all('table', recursive=False)
        if len(all_tables) < 2:
            return None, "Konnte die Wertungsprüfungstabelle nicht finden."

        stage_table = all_tables[1]
        stage_rows = stage_table.find_all('tr')
        current_leg = None

        for i, stage_row in enumerate(stage_rows):
            if i == 0: continue # Skip header row
            stage_cells = stage_row.find_all('td')
            first_cell = stage_cells[0] if len(stage_cells) > 0 else None

            # Check for Leg Header
            if first_cell and 'lista_kiemelt' in first_cell.get('class', []):
                bold_tag = first_cell.find('b')
                if bold_tag and 'Leg' in bold_tag.get_text():
                    leg_name_raw = bold_tag.get_text(strip=True)
                    leg_name = leg_name_raw.replace('Leg', 'Etappe') if translate else leg_name_raw
                    if len(stage_cells) > 2 and 'lista_kiemelt' in stage_cells[2].get('class', []):
                        distance_bold = stage_cells[2].find('b')
                        if distance_bold:
                            leg_name += f" ({distance_bold.get_text(strip=True)})"
                    current_leg = {"name": leg_name, "items": []}
                    legs.append(current_leg)
                    continue # Move to next row after processing leg header

            # Check for Service Park or Road Side Service
            is_service = 'servicepark' in stage_row.get('class', [])
            is_road_service = current_leg and len(stage_cells) >= 2 and 'Road Side Service' in stage_cells[1].get_text()

            if is_service or is_road_service:
                if current_leg and len(stage_cells) >= 2:
                    full_text_raw = stage_cells[1].get_text(strip=True)
                    full_text_cleaned = re.sub(r'\s*-\s*', ' - ', full_text_raw).strip()
                    service_text = translate_iteratively(full_text_cleaned, translations) if translate else full_text_cleaned
                    current_leg["items"].append({"type": "service", "cleaned_text": service_text})
                continue # Move to next row

            # Process Stage Row
            if current_leg and len(stage_cells) >= 5:
                try:
                    int(stage_cells[0].get_text(strip=True)) # Check if first cell is a stage number
                    stage_name_div = stage_cells[1].find('div')
                    stage_name = re.sub(r'(\r\n|\n|\r)', '', (stage_name_div.get_text(strip=True) if stage_name_div else stage_cells[1].get_text(strip=True)))
                    stage_length = re.sub(r'(\r\n|\n|\r)', '', stage_cells[2].get_text(strip=True))
                    surface_en = re.sub(r'(\r\n|\n|\r)', '', stage_cells[3].get_text(strip=True))
                    weather_en = re.sub(r'(\r\n|\n|\r)', '', stage_cells[4].get_text(strip=True))

                    surface_formatted = re.sub(r'\s*\(([^)]+)\)', r', \1', surface_en)
                    weather_formatted = ', '.join(weather_en.split())

                    surface_de = translate_iteratively(surface_formatted, translations) if translate else surface_formatted
                    weather_de = translate_iteratively(weather_formatted, translations) if translate else weather_formatted

                    current_leg["items"].append({
                        "type": "stage",
                        "name": stage_name,
                        "length": stage_length,
                        "surface": surface_de,
                        "weather": weather_de
                    })
                except (ValueError, IndexError):
                    pass # Ignore rows that don't look like stages

        status_callback("HTML erfolgreich verarbeitet.")
        poster_data = {
            "rally_name": rally_name,
            "total_distance": total_distance,
            "car_name": car_name,
            "legs": legs
        }
        return poster_data, None

    except Exception as e:
        error_message = f"Fehler bei der HTML-Verarbeitung:\n{e}"
        return None, error_message


def create_poster_files(poster_data, save_path, translate, status_callback):
    if not poster_data or not save_path:
        return False, "Fehlende Daten oder Speicherpfad.", None

    rally_name = poster_data['rally_name']
    total_distance = poster_data['total_distance']
    car_name = poster_data['car_name']
    legs = poster_data['legs']

    header_translations = {
        True: {"stage": "Wertungsprüfung", "length": "Länge", "surface": "Zustand", "weather": "Bedingungen", "distance_label": "Distanz"},
        False: {"stage": "Stage name", "length": "Distance", "surface": "Surface", "weather": "Weather", "distance_label": "Distance"}
    }
    headers = header_translations.get(translate, header_translations[True])

    # --- Build HTML ---
    status_callback("Erstelle HTML...")
    poster_html = f"""
<html>
<head>
<title>{rally_name}</title>
<style>
body {{ font-family: Roboto; padding: 25px; background-color: #f4f1e8; color: #333; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
h1 {{ font-family: Impact, 'Arial Black', Gadget, sans-serif; text-align: center; color: #a00000; font-size: 3em; margin-bottom: 10px; text-transform: uppercase; letter-spacing: 1px; }}
h2 {{ font-family: 'Libre Bodoni'; font-weight: bold; text-align: center; font-size: 1.7em; color: #444; margin-bottom: 10px; margin-top: 5px; text-transform: uppercase; }}
h3 {{ font-family: 'Libre Bodoni'; font-weight: bold; text-align: center; font-size: 1.5em; color: #444; margin-bottom: 20px; margin-top: 0; text-transform: uppercase; }}
.leg-header {{ background-color: #4d3d33; color: #f4f1e8; font-weight: bold; text-align: center; padding: 8px; margin-top: 25px; font-size: 1.4em; text-transform: uppercase; border: 1px solid #111; }}
table {{ width: 100%; border-collapse: collapse; margin-top: 0; box-shadow: none; border: 1px solid #444; }}
th, td {{ border: 1px solid #777; padding: 6px 8px; text-align: left; font-size: 0.95em; }}
th {{ background-color: #777; color: #f4f1e8; font-weight: bold; text-transform: uppercase; font-size: 0.95em; }}
.service-row td {{ font-style: italic; background-color: #dddddd; color: #222; padding: 8px 8px; }}
tbody tr:not(.service-row) td:first-child {{ font-weight: bold; }}
.info-box {{ background-color: #e9e5d9; border: 1px solid #555; padding: 15px; border-radius: 0px; margin-top: 25px; line-height: 1.5; font-size: 0.9em; color: #222; box-shadow: none; }}
</style>
</head>
<body>
<h1>{rally_name}</h1>
"""
    car_name_length = len(car_name.replace('<br>', ' '))
    car_font_size = "1.2em" if car_name_length > 150 else ("1.4em" if car_name_length > 80 else "1.7em")
    poster_html += f'<h2 style="font-size: {car_font_size};">{car_name}</h2>'
    poster_html += f"<h3>{headers['distance_label']}: {total_distance}</h3>"

    for leg in legs:
        poster_html += f'<div class="leg-header">{leg["name"]}</div>'
        poster_html += f"""
<table><thead><tr>
<th>{headers['stage']}</th><th>{headers['length']}</th><th>{headers['surface']}</th><th>{headers['weather']}</th>
</tr></thead><tbody>"""
        for item in leg["items"]:
            if item["type"] == 'stage':
                poster_html += f"""
<tr><td>{item['name']}</td><td>{item['length']}</td><td>{item['surface']}</td><td>{item['weather']}</td></tr>"""
            elif item["type"] == 'service':
                service_text = item.get('cleaned_text', '').strip()
                poster_html += f"""<tr class="service-row"><td colspan="4">{service_text}</td></tr>"""
        poster_html += "</tbody></table>"

    if rally_name.startswith("DE-DCR"):
        poster_html += """
<div class="info-box">
Nach jeder Wertungsprüfung ist es Fahrer und Beifahrer gestattet 5 Minuten Reparaturzeit in Anspruch zu nehmen.<br>
Zwischen den Blöcken gibt es 60 Minuten Reparaturzeit mit bis zu 4 Mechanikern.<br>
Kein Wiedereinstieg nach Ausfall. Die Wertung erfolgt rollend über 4 Wochen.<br>
Punktevergabe: Ränge 1–10 erhalten Punkte: Sieger 20, Vize 18, ... 10. Platz 2 Punkte.<br>
Passwort: Willkommen
</div>"""
    poster_html += "</body></html>"

    # --- Save Files ---
    html_path = os.path.splitext(save_path)[0] + ".html"
    try:
        status_callback("Speichere HTML...")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(poster_html)

        status_callback("Erstelle PNG mit Playwright...")
        with sync_playwright() as p:
            # Check if browsers are installed, install if necessary
            try:
                browser = p.chromium.launch()
            except Exception:
                status_callback("Playwright Browser nicht gefunden. Installiere...")
                print("Attempting to install Playwright browsers...")
                os.system(f'"{sys.executable}" -m playwright install chromium')
                status_callback("Browser installiert. Versuche erneut...")
                browser = p.chromium.launch() # Try again

            page = browser.new_page()
            # Use file URI for local HTML
            file_uri = 'file:///' + os.path.abspath(html_path).replace('\\', '/')
            page.goto(file_uri, wait_until='load')
            # page.set_content(poster_html) # Less reliable for complex CSS/fonts
            # page.wait_for_load_state('networkidle') # Wait longer if needed
            page.screenshot(path=save_path, full_page=True)
            browser.close()
        status_callback("PNG erfolgreich erstellt.")
        return True, f"Poster gespeichert:\nPNG: {save_path}\nHTML: {html_path}", html_path

    except Exception as e:
        return False, f"Fehler beim Speichern oder PNG-Erstellung:\n{e}", html_path


# --- Tkinter GUI Application ---

class RallyPosterApp:
    def __init__(self, master):
        self.master = master
        master.title("Rally Poster Generator")
        master.geometry("600x250") # Adjusted size
        master.resizable(False, False)

        # Style
        self.style = ttk.Style()
        self.style.theme_use('clam') # Or 'vista', 'xpnative'

        # Variables
        self.url_var = tk.StringVar()
        self.translate_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Bereit.")
        self.last_save_dir = self._load_last_save_dir()

        # --- Layout ---
        main_frame = ttk.Frame(master, padding="10 10 10 10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # URL Input
        url_label = ttk.Label(main_frame, text="Rallye URL:")
        url_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.url_entry = ttk.Entry(main_frame, textvariable=self.url_var, width=60)
        self.url_entry.grid(row=0, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5)
        self.url_entry.focus_set()
        # Add right-click context menu
        self.create_context_menu(self.url_entry)


        # Translate Checkbox
        translate_check = ttk.Checkbutton(main_frame, text="übersetzen", variable=self.translate_var)
        translate_check.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        # Generate Button
        self.generate_button = ttk.Button(main_frame, text="Poster generieren", command=self.start_generation_thread)
        self.generate_button.grid(row=2, column=1, columnspan=2, pady=10)

        # Status Bar
        status_label = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=5)
        status_label.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))

        # Center window
        master.update_idletasks()
        x = (master.winfo_screenwidth() // 2) - (master.winfo_width() // 2)
        y = (master.winfo_screenheight() // 2) - (master.winfo_height() // 2)
        master.geometry(f'+{x}+{y}')

    def create_context_menu(self, widget):
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label="Cut", command=lambda: widget.event_generate("<<Cut>>"))
        menu.add_command(label="Copy", command=lambda: widget.event_generate("<<Copy>>"))
        menu.add_command(label="Paste", command=lambda: widget.event_generate("<<Paste>>"))
        menu.add_separator()
        menu.add_command(label="Select All", command=lambda: widget.select_range(0, 'end'))

        widget.bind("<Button-3>", lambda event: menu.tk_popup(event.x_root, event.y_root))

    def _load_last_save_dir(self):
        last_path_file = "last_save_path.txt"
        if os.path.exists(last_path_file):
            try:
                with open(last_path_file, "r", encoding="utf-8") as f:
                    saved_path = f.read().strip()
                    if saved_path and os.path.isdir(os.path.dirname(saved_path)):
                        return os.path.dirname(saved_path)
            except Exception:
                pass # Ignore errors reading the file
        return os.getcwd() # Default to current dir

    def _save_last_save_dir(self, file_path):
        last_path_file = "last_save_path.txt"
        if file_path:
            try:
                with open(last_path_file, "w", encoding="utf-8") as f:
                    f.write(file_path) # Save the full path
                self.last_save_dir = os.path.dirname(file_path) # Update internal state
            except Exception:
                pass # Ignore errors writing the file

    def update_status(self, message):
        # Ensure GUI updates happen on the main thread
        self.master.after(0, self.status_var.set, message)

    def show_error(self, message):
        # Ensure GUI updates happen on the main thread
        self.master.after(0, messagebox.showerror, "Fehler", message)
        self.update_status("Fehler aufgetreten. Bereit.")
        self.master.after(0, self.generate_button.config, {"state": "normal"}) # Re-enable button

    def show_success(self, message, png_path, html_path):
         # Ensure GUI updates happen on the main thread
        self.master.after(0, self._show_success_dialog, message, png_path, html_path)
        self.update_status("Erfolgreich abgeschlossen. Bereit.")
        self.master.after(0, self.generate_button.config, {"state": "normal"}) # Re-enable button

    def _show_success_dialog(self, message, png_path, html_path):
        # This runs in the main thread because it's called via master.after
        def open_png():
            try:
                os.startfile(png_path)
            except Exception as e:
                messagebox.showerror("Fehler", f"PNG konnte nicht geöffnet werden:\n{e}")

        def open_html():
            try:
                os.startfile(html_path)
            except Exception as e:
                messagebox.showerror("Fehler", f"HTML konnte nicht geöffnet werden:\n{e}")

        info_dialog = tk.Toplevel(self.master)
        info_dialog.title("Erfolg")
        info_dialog.geometry("500x200")
        info_dialog.resizable(False, False)
        info_dialog.transient(self.master) # Keep on top of main window
        info_dialog.grab_set() # Modal

        label = ttk.Label(info_dialog, text=message, justify="left", wraplength=480)
        label.pack(pady=10, padx=10)

        button_frame = ttk.Frame(info_dialog)
        button_frame.pack(pady=10)

        open_png_button = ttk.Button(button_frame, text="PNG öffnen", command=open_png, width=15)
        open_png_button.pack(side=tk.LEFT, padx=10)

        open_html_button = ttk.Button(button_frame, text="HTML öffnen", command=open_html, width=15)
        open_html_button.pack(side=tk.LEFT, padx=10)

        close_button = ttk.Button(button_frame, text="Schließen", command=info_dialog.destroy, width=15)
        close_button.pack(side=tk.LEFT, padx=10)

        # Center the dialog relative to the main window
        info_dialog.update_idletasks()
        main_x = self.master.winfo_x()
        main_y = self.master.winfo_y()
        main_w = self.master.winfo_width()
        main_h = self.master.winfo_height()
        dialog_w = info_dialog.winfo_width()
        dialog_h = info_dialog.winfo_height()
        x = main_x + (main_w // 2) - (dialog_w // 2)
        y = main_y + (main_h // 2) - (dialog_h // 2)
        info_dialog.geometry(f'+{x}+{y}')

        info_dialog.wait_window() # Wait until dialog is closed


    def start_generation_thread(self):
        # Disable button to prevent multiple clicks
        self.generate_button.config(state="disabled")
        self.update_status("Starte Generierung...")
        # Run the blocking operations in a separate thread
        thread = threading.Thread(target=self.run_generation_process, daemon=True)
        thread.start()

    def run_generation_process(self):
        url = self.url_var.get().strip()
        translate = self.translate_var.get()

        # 1. Fetch HTML
        html_content, error = fetch_html_content(url, self.update_status)
        if error:
            self.show_error(error)
            return

        # 2. Parse HTML to get data (including rally name for save dialog)
        poster_data, error = generate_poster_data(html_content, translate, self.update_status)
        if error:
            self.show_error(error)
            return
        if not poster_data:
             self.show_error("Konnte keine Posterdaten extrahieren.")
             return

        # 3. Ask for Save Path (needs to run in main thread)
        self.master.after(0, self.ask_save_path_and_generate, poster_data, translate)

    def ask_save_path_and_generate(self, poster_data, translate):
        # This method is called via self.master.after, so it runs in the main GUI thread
        rally_name = poster_data.get('rally_name', 'poster')
        safe_name = re.sub(r'[\\/*?:"<>|]', "_", rally_name).strip() + ".png"
        initial_file = safe_name

        save_path = filedialog.asksaveasfilename(
            master=self.master, # Ensure dialog is parented correctly
            title="Speicherort für Poster wählen",
            defaultextension=".png",
            initialdir=self.last_save_dir,
            initialfile=initial_file,
            filetypes=[("PNG Dateien", "*.png")]
        )

        if not save_path:
            self.update_status("Abgebrochen. Bereit.")
            self.generate_button.config(state="normal") # Re-enable button
            return

        self._save_last_save_dir(save_path) # Save the chosen path/directory

        # 4. Generate Poster Files (can run in thread again if needed, but Playwright might prefer main thread)
        # For simplicity, let's run the final step here, but keep UI responsive
        self.update_status("Erstelle Poster-Dateien...")
        # Run the final file creation in a thread to avoid blocking GUI during Playwright
        thread = threading.Thread(target=self.run_file_creation, args=(poster_data, save_path, translate), daemon=True)
        thread.start()

    def run_file_creation(self, poster_data, save_path, translate):
        success, message, html_path = create_poster_files(poster_data, save_path, translate, self.update_status)
        if success:
            self.show_success(message, save_path, html_path)
        else:
            self.show_error(message)
            # Attempt to clean up the potentially created HTML file on error
            if html_path and os.path.exists(html_path):
                try:
                    os.remove(html_path)
                except OSError:
                    pass # Ignore if removal fails


if __name__ == "__main__":
    # Command-line argument handling is removed as per the plan
    # The application will now always start in GUI mode.
    root = tk.Tk()
    app = RallyPosterApp(root)
    root.mainloop()
