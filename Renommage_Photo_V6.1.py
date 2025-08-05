import os
import re
import json
import time
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk, simpledialog
from datetime import datetime
from PIL import Image, ImageTk, ImageEnhance
import requests
import hashlib
import openai
import logging
import subprocess
import time
import subprocess
from tkinter import ttk
# Définition du répertoire utilisateur pour la configuration
user_config_dir = os.path.expanduser("~\\AppData\\Local\\PhotoManagerPro")
os.makedirs(user_config_dir, exist_ok=True)


# Pour CopiePhoto S10 (spécifique à Windows)
import pythoncom
import win32com.client
import win32clipboard
import os
import logging

# Définition du répertoire utilisateur pour la configuration
user_config_dir = os.path.expanduser("~\\AppData\\Local\\PhotoManagerPro")
os.makedirs(user_config_dir, exist_ok=True)


# Création d'un dossier pour stocker le log dans le répertoire local de l'utilisateur
log_dir = os.path.expanduser(r"~\AppData\Local\PhotoManagerPro")
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

log_path = os.path.join(log_dir, "app.log")

logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)


# ------------------------------------------------
# Constantes et fonctions utilitaires
# ------------------------------------------------
CONFIG_FILE = os.path.join(user_config_dir, "config.json")
CORRECTION_FILE = "corrections.json"
BUTTON_PRESETS_FILE = "button_presets.json"
IMGUR_UPLOAD_URL = "https://api.imgur.com/3/image"
HISTORIQUE_PATH = "historique_copie.txt"
EXTRA_BUTTONS_FILE = "extra_buttons.json"

def load_config():
    """
    Charge la configuration JSON. 
    Si le fichier est manquant, vide ou mal formé, on retourne un dict vide.
    """
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError as e:
            logging.warning(f"Config JSON invalide ou vide ({CONFIG_FILE}) : {e}")
            return {}
    return {}

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)

def load_api_key():
    cfg = load_config()
    return (
        cfg.get("openai_api_key", ""),
        cfg.get("imgur_client_id", ""),
        cfg.get("logo_path", ""),
        cfg.get("default_working_directory", "")
    )

def save_api_keys(openai_key, imgur_key):
    cfg = load_config()
    cfg["openai_api_key"] = openai_key
    cfg["imgur_client_id"] = imgur_key
    save_config(cfg)

def set_default_working_directory(directory):
    cfg = load_config()
    cfg["default_working_directory"] = directory
    save_config(cfg)

# Gestion du fichier JSON des presets
def get_preset_file():
    cfg = load_config()
    return cfg.get("button_presets_file", BUTTON_PRESETS_FILE)

def set_preset_file(path):
    cfg = load_config()
    cfg["button_presets_file"] = path
    save_config(cfg)

# Gestion du fichier des boutons supplémentaires
def load_extra_buttons():
    if os.path.exists(EXTRA_BUTTONS_FILE):
        try:
            with open(EXTRA_BUTTONS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
        except Exception as e:
            logging.warning(f"Erreur lecture {EXTRA_BUTTONS_FILE}: {e}")
    return [f"Extra{i+1}" for i in range(30)]


def save_extra_buttons(values):
    with open(EXTRA_BUTTONS_FILE, "w", encoding="utf-8") as f:
        json.dump(values, f, indent=2)

def image_fingerprint(image_path):
    try:
        with open(image_path, "rb") as f:
            return hashlib.md5(f.read()).hexdigest()
    except Exception:
        return None

def save_correction(fp, name):
    data = {}
    if os.path.exists(CORRECTION_FILE):
        with open(CORRECTION_FILE, "r", encoding="utf-8") as f:
            try:
                data = json.load(f)
            except:
                data = {}
    data[fp] = name
    with open(CORRECTION_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

def upload_image_to_imgur(image_path, client_id):
    if not client_id:
        return None, None
    headers = {"Authorization": f"Client-ID {client_id}"}
    try:
        with open(image_path, "rb") as f:
            resp = requests.post(IMGUR_UPLOAD_URL, headers=headers, files={"image": f})
        if resp.status_code == 200:
            data = resp.json()["data"]
            return data["link"], data.get("deletehash")
    except Exception as e:
        logging.error("Erreur upload Imgur: %s", e)
    return None, None

def delete_from_imgur(delete_hash, client_id):
    if not delete_hash or not client_id:
        return
    headers = {"Authorization": f"Client-ID {client_id}"}
    url = f"https://api.imgur.com/3/image/{delete_hash}"
    try:
        requests.delete(url, headers=headers)
    except Exception:
        pass

# ------------------------------------------------
# Classe pour les tooltips (info-bulles)
# ------------------------------------------------
class CreateToolTip:
    def __init__(self, widget, text='widget info'):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        
    def enter(self, event=None):
        self.showtip()
        
    def leave(self, event=None):
        self.hidetip()
        
    def showtip(self):
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert") if self.widget.winfo_exists() else (0, 0, 0, 0)
        x = x + self.widget.winfo_rootx() + 25
        y = y + cy + self.widget.winfo_rooty() + 25
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)
        
    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

# ==========================================================================
#                           MODULE PHOTO RENAME - VERSION 6.1
# ==========================================================================
class PhotoRenameApp:
    def __init__(self, master):
        self.master = master
        master.title("Renommage Photo V6.1")
        right_bg = "#f0f8ff"

        # --- Cadre principal à droite (navigation + onglets) ---
        self.right_frame = tk.Frame(master, bg=right_bg)
        self.right_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # --- Boutons de navigation existants ---
        nav = tk.Frame(self.right_frame, bg=right_bg)
        nav.grid(row=0, column=0, sticky="w", pady=(0,10))
        tk.Button(nav, text="Précédent", command=self.prev_photo).pack(side="left")
        tk.Button(nav, text="Suivant",   command=self.next_photo).pack(side="left")
        tk.Button(nav, text="Renommer",  command=self.rename_photo).pack(side="left")

        # --- Onglets des presets personnalisés ---
        self.tab_control = ttk.Notebook(self.right_frame)
        self.tab_control.grid(row=1, column=0, sticky="nsew")
        self.tab_frames = []
        for i in range(1, 6):
            frame = tk.Frame(self.tab_control, bg=right_bg)
            self.tab_control.add(frame, text=f"Onglet {i}")
            frame.index = i
            self.tab_frames.append(frame)
            # binding du double-clic sur l'étiquette d'onglet
            self.tab_control.tab(i-1, padding=[5, 5, 5, 5])
        self.tab_control.bind("<Double-1>", self.on_tab_double_click)

        # --- Zone d'édition des presets (vos 80 boutons + 45 fixes) ---
        # À adapter : remettez ici votre code existant pour éditer
        # les presets et la zone fixed_frame pour les 45 boutons.
        self.editor_frame = tk.Frame(self.right_frame, bg=right_bg)
        self.editor_frame.grid(row=2, column=0, sticky="ew", pady=(10,0))
        self.build_editor_zone()  # votre méthode existante

        # --- Chargement de la config et initialisation des onglets ---
        for frame in self.tab_frames:
            preset = ConfigManager.Current.TabPresetMapping.get(frame.index)
            if preset:
                self.load_tab_preset(frame, preset)

    def _build_layout(self):
        self.master.columnconfigure(0, weight=2)
        self.master.columnconfigure(1, weight=1)
        self.master.rowconfigure(0, weight=1)

        # Colonne gauche : affichage de l'image dans un Canvas avec scrollbars
        left_bg = "#f9f9f9"
        self.left_frame = tk.Frame(self.master, bg=left_bg)
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.left_frame.columnconfigure(0, weight=1)
        self.left_frame.rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(self.left_frame, bg="black")
        self.canvas.grid(row=0, column=0, sticky="nsew")
        hbar = tk.Scrollbar(self.left_frame, orient="horizontal", command=self.canvas.xview)
        vbar = tk.Scrollbar(self.left_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.config(xscrollcommand=hbar.set, yscrollcommand=vbar.set)
        hbar.grid(row=1, column=0, sticky="ew")
        vbar.grid(row=0, column=1, sticky="ns")
        self.canvas.bind("<ButtonPress-1>", self.start_pan)
        self.canvas.bind("<B1-Motion>", self.do_pan)

        # ✅ Boutons sous l’image : rotation, zoom
        btn_frame = tk.Frame(self.left_frame, bg=left_bg)
        btn_frame.grid(row=2, column=0, pady=5)

        rot_btn = tk.Button(btn_frame, text="↻ Rotation 90°", command=self.rotate_photo, bg="#4CAF50", fg="white")
        rot_btn.pack(side=tk.LEFT, padx=5)
        CreateToolTip(rot_btn, "Faire pivoter l'image de 90° dans le sens horaire")

        zoom_plus = tk.Button(btn_frame, text="Zoom +", command=self.zoom_in, bg="#2196F3", fg="white")
        zoom_plus.pack(side=tk.LEFT, padx=5)
        CreateToolTip(zoom_plus, "Agrandir l'image")

        zoom_moins = tk.Button(btn_frame, text="Zoom -", command=self.zoom_out, bg="#f44336", fg="white")
        zoom_moins.pack(side=tk.LEFT, padx=5)
        CreateToolTip(zoom_moins, "Réduire l'image")

        # Zone log et barre de progression
        self.log = scrolledtext.ScrolledText(self.left_frame, height=8, bg="white")
        self.log.grid(row=3, column=0, sticky="ew", padx=5, pady=5)

        self.progress = ttk.Progressbar(self.left_frame, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=4, column=0, sticky="ew", pady=5)

    # ✅ (le reste de ta colonne droite n’a pas changé, donc ne pas toucher)



        # === Colonne droite ===
        right_bg = "#e8f5e9"
        self.right_frame = tk.Frame(self.master, bg=right_bg)
        self.right_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        self.right_frame.columnconfigure(0, weight=1)

        # === Navigation ===
        self.nav_frame = tk.LabelFrame(self.right_frame, text="Navigation", bg=right_bg)
        self.nav_frame.grid(row=3, column=0, sticky="ew", padx=5, pady=5)

        prev_btn = tk.Button(self.nav_frame, text="⏮️ Précédent", command=self.photo_prec, bg="#8BC34A", fg="white")
        prev_btn.grid(row=0, column=0, padx=5, pady=5)
        CreateToolTip(prev_btn, "Afficher la photo précédente")

        next_btn = tk.Button(self.nav_frame, text="⏭️ Suivant", command=self.photo_suiv, bg="#8BC34A", fg="white")
        next_btn.grid(row=0, column=1, padx=5, pady=5)
        CreateToolTip(next_btn, "Afficher la photo suivante")

        renommer_btn = tk.Button(
            self.nav_frame,
            text="Renommer",
            command=self.renommer_photo,
            bg="#0D47A1",     # Bleu foncé
            fg="white",
            activebackground="#1565C0",  # Bleu un peu plus clair au clic
            font=("Arial", 10, "bold")
)
        renommer_btn.grid(row=0, column=2, padx=5, pady=5)
        CreateToolTip(renommer_btn, "Renommer la photo avec le nom sélectionné")

        # === Numérotation ===
        self.num_frame = tk.LabelFrame(self.right_frame, text="Numérotation Préfixe", bg=right_bg)
        self.num_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        prefix_chk = tk.Checkbutton(self.num_frame, text="Activer la numérotation", variable=self.var_prefix, bg=right_bg)
        prefix_chk.pack(side=tk.LEFT, padx=5)
        CreateToolTip(prefix_chk, "Ajouter un numéro avant le nouveau nom")
        tk.Label(self.num_frame, text="Prochain numéro :", bg=right_bg).pack(side=tk.LEFT)
        self.spin_prefix = tk.Spinbox(self.num_frame, from_=1, to=9999, width=5)
        self.spin_prefix.pack(side=tk.LEFT, padx=5)
        self.spin_prefix.delete(0, "end")
        self.spin_prefix.insert(0, "1")
        CreateToolTip(self.spin_prefix, "Numéro qui sera utilisé pour le prochain renommage")

        # === Actions ===
        self.act_frame = tk.LabelFrame(self.right_frame, text="Actions", bg=right_bg)
        self.act_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        action_btn = tk.Button(self.act_frame, text="Analyser", command=self.analyser_photo, bg="#FFC107", fg="black")
        action_btn.grid(row=0, column=0, padx=5, pady=5)
        renommer_btn = tk.Button(self.act_frame, text="Renommer", command=self.renommer_photo, bg="#4CAF50", fg="white")
        renommer_btn.grid(row=0, column=1, padx=5, pady=5)
        auto_btn = tk.Button(self.act_frame, text="🔁 Tout analyser", command=self.analyser_toutes_photos, bg="#03A9F4", fg="white")
        auto_btn.grid(row=0, column=2, padx=5, pady=5)
        quit_btn = tk.Button(self.act_frame, text="Quitter", command=self.on_quit, bg="#f44336", fg="white")
        quit_btn.grid(row=0, column=3, padx=5, pady=5)
        retour_btn = tk.Button(self.act_frame, text="Retour au menu", command=self.retour_menu, bg="#9E9E9E", fg="white")
        retour_btn.grid(row=1, column=0, columnspan=4, pady=5)

        # === Suggestions IA ===
        self.sug_frame = tk.LabelFrame(self.right_frame, text="Suggestions IA", bg=right_bg)
        self.sug_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        self.suggestions = [tk.StringVar() for _ in range(4)]
        self.radio_sel = tk.StringVar(value="0")
        self.suggestion_entries = []
        for i in range(4):
            rowf = tk.Frame(self.sug_frame, bg=right_bg)
            rowf.pack(anchor="w", pady=2, fill="x")
            radio = tk.Radiobutton(rowf, variable=self.radio_sel, value=i, bg=right_bg)
            radio.pack(side=tk.LEFT, padx=5)
            ent = tk.Entry(rowf, textvariable=self.suggestions[i], width=40)
            ent.pack(side=tk.LEFT, expand=True, fill="x")
            ent.bind("<FocusIn>", lambda e, idx=i: self.radio_sel.set(str(idx)))
            self.suggestion_entries.append(ent)

        # === Onglets des presets personnalisés ===
        
        self.tab_control = ttk.Notebook(self.right_frame)
        self.tab_control.grid(row=4, column=0, sticky="ew", padx=5, pady=5)
        self.tab_frames = []
        for i in range(1, 6):
            frame = tk.Frame(self.tab_control, bg=right_bg)
            self.tab_control.add(frame, text=f"Onglet {i}")
            frame.index = i
            self.tab_frames.append(frame)
        # Charger la config et remplir chaque onglet
        for frame in self.tab_frames:
            preset = ConfigManager.Current.TabPresetMapping.get(frame.index)
            if preset:
                self.load_tab_preset(frame, preset)




        # === Boutons supplémentaires ===
        self.extra_frame = tk.LabelFrame(self.right_frame, text="45 Boutons supplémentaires", bg=right_bg)
        self.extra_frame.grid(row=5, column=0, sticky="ew", padx=5, pady=5)
        self.extra_vars = []
        self.extra_buttons = []
        total_extra = 45
        for i in range(total_extra):
            val = self.extra_values[i] if i < len(self.extra_values) else f"Extra{i+1}"
            var = tk.StringVar(value=val)
            self.extra_vars.append(var)
        rows_e = 5
        cols_e = 9
        for idx, var in enumerate(self.extra_vars):
            r = idx // cols_e
            c = idx % cols_e
            btn = tk.Button(self.extra_frame, textvariable=var, width=12)
            text_val = var.get().strip().lower()
            if text_val.startswith("extra") and text_val[5:].isdigit():
                couleur = "#e0e0e0"
            else:
                couleur = "#d0eaff"
            btn.config(bg=couleur)
            btn.grid(row=r, column=c, padx=5, pady=5)
            btn.bind("<Button-1>", lambda e, i=idx: self.on_extra_left(i))
            btn.bind("<Button-3>", lambda e, i=idx: self.on_extra_right(i))
            self.extra_buttons.append(btn)

        # === Preset des boutons ===
        self.preset_frame = tk.LabelFrame(self.right_frame, text="Paramétrage des boutons", bg=right_bg)
        self.preset_frame.grid(row=6, column=0, sticky="ew", padx=5, pady=5)
        tk.Label(self.preset_frame, text="Preset:", bg=right_bg).pack(side=tk.LEFT, padx=5)
        self.preset_combobox = ttk.Combobox(self.preset_frame, textvariable=self.preset_var)
        self.preset_combobox.pack(side=tk.LEFT, padx=5)
        self.preset_combobox.bind("<<ComboboxSelected>>", self.load_preset)
        btn_save_preset = tk.Button(self.preset_frame, text="Enregistrer paramétrage boutons", command=self.save_current_preset, bg="#3F51B5", fg="white")
        btn_save_preset.pack(side=tk.LEFT, padx=5)
        btn_del_preset = tk.Button(self.preset_frame, text="Supprimer preset", command=self.delete_current_preset, bg="#E91E63", fg="white")
        btn_del_preset.pack(side=tk.LEFT, padx=5)
        self.update_preset_list()

        # === Menu principal ===
        self.menu_bar = tk.Menu(self.master)
        menu_f = tk.Menu(self.menu_bar, tearoff=0)
        menu_f.add_command(label="Ouvrir un dossier", command=self.ouvrir_dossier)
        menu_f.add_separator()
        menu_f.add_command(label="Quitter", command=self.on_quit)
        menu_f.add_command(label="Retour au menu", command=self.retour_menu)
        self.menu_bar.add_cascade(label="Fichier", menu=menu_f)
        menu_p = tk.Menu(self.menu_bar, tearoff=0)
        menu_p.add_command(label="Configurer clés API", command=self.configurer_api)
        menu_p.add_command(label="Configurer logo", command=self.configurer_logo)
        menu_p.add_command(label="Configurer fichier presets", command=self.configurer_preset_file)
        self.menu_bar.add_cascade(label="Paramètres", menu=menu_p)
        menu_o = tk.Menu(self.menu_bar, tearoff=0)
        menu_o.add_command(label="Organiser dossiers", command=self.organiser_dossiers)
        self.menu_bar.add_cascade(label="Organisation dossier", menu=menu_o)
        menu_a = tk.Menu(self.menu_bar, tearoff=0)
        menu_a.add_command(label="Aide détaillée", command=self.afficher_aide)
        menu_a.add_command(label="FAQ", command=self.afficher_faq)
        self.menu_bar.add_cascade(label="Aide", menu=menu_a)
        self.master.config(menu=self.menu_bar)

    def mettre_a_jour_couleurs_boutons(self):
        """
        Parcourt tous les onglets (self.tab_frames) et colore
        chaque bouton selon son texte (btn* = gris, sinon bleu clair).
        """
        # Si les onglets existent, on applique les couleurs
        for frame in getattr(self, "tab_frames", []):
            for widget in frame.winfo_children():
                # Ne traiter que les boutons
                if isinstance(widget, tk.Button):
                    text = widget.cget("text").strip().lower()
                    couleur = "#e0e0e0" if text.startswith("btn") else "#d0eaff"
                    widget.config(bg=couleur)


    def start_pan(self, event):
        self.canvas.scan_mark(event.x, event.y)

    def do_pan(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def configurer_preset_file(self):
        preset_file = filedialog.askopenfilename(
            title="Sélectionner le fichier JSON des presets",
            filetypes=[("Fichiers JSON", "*.json")]
        )
        if preset_file:
            set_preset_file(preset_file)
            self.update_preset_list()
            messagebox.showinfo("OK", f"Fichier presets configuré : {preset_file}")

    def retour_menu(self):
        self.master.destroy()
        launch_main_menu()

    def log_insert(self, msg):
        self.log.insert("end", msg)
        self.log.see("end")

    def on_quit(self):
        rep = messagebox.askokcancel("Quitter", "Voulez-vous vraiment quitter ?")
        if rep:
            if self.imgur_deletions:
                rep2 = messagebox.askyesno("Supprimer sur Imgur?", f"Supprimer {len(self.imgur_deletions)} images sur Imgur?")
                if rep2:
                    for (dh, cid) in self.imgur_deletions:
                        delete_from_imgur(dh, cid)
                    messagebox.showinfo("OK", "Toutes les images ont été supprimées d'Imgur.")
            save_extra_buttons([var.get() for var in self.extra_vars])
            self.master.quit()

    def on_custom_left(self, idx):
        text_label = self.custom_vars[idx].get()
        focused = self.master.focus_get()
        target = focused if focused in self.suggestion_entries else self.suggestion_entries[int(self.radio_sel.get())]
        try:
            sel_start = target.index("sel.first")
            sel_end = target.index("sel.last")
            new_val = target.get()[:int(sel_start)] + text_label + target.get()[int(sel_end):]
        except tk.TclError:
            pos = target.index(tk.INSERT)
            new_val = target.get()[:int(pos)] + text_label + target.get()[int(pos):]
        target.delete(0, tk.END)
        target.insert(0, new_val)

    def on_custom_left_by_label(self, label):
        for i, var in enumerate(self.custom_vars):
            if var.get() == label:
                self.on_custom_left(i)
                break

    def on_custom_right(self, idx):
        new_lbl = simpledialog.askstring("Renommer le bouton", "Nouveau label ?")
        if new_lbl:
            self.custom_vars[idx].set(new_lbl)
            # Mise à jour immédiate de la couleur
            text = new_lbl.strip().lower()
            if text.startswith("btn") and text[3:].isdigit():
                couleur = "#dcdcdc"
            elif text == "":
                couleur = "#f5f5f5"
            else:
                couleur = "#add8e6"
            self.custom_buttons[idx].config(bg=couleur)

    def on_extra_left(self, idx):
        text_label = self.extra_vars[idx].get()
        focused = self.master.focus_get()
        target = focused if focused in self.suggestion_entries else self.suggestion_entries[int(self.radio_sel.get())]
        try:
            sel_start = target.index("sel.first")
            sel_end = target.index("sel.last")
            new_val = target.get()[:int(sel_start)] + text_label + target.get()[int(sel_end):]
        except tk.TclError:
            pos = target.index(tk.INSERT)
            new_val = target.get()[:int(pos)] + text_label + target.get()[int(pos):]
        target.delete(0, tk.END)
        target.insert(0, new_val)

    def on_extra_right(self, idx):
        new_lbl = simpledialog.askstring("Renommer le bouton", "Nouveau label ?")
        if new_lbl:
            self.extra_vars[idx].set(new_lbl)
            text = new_lbl.strip().lower()
            if text.startswith("extra") and text[5:].isdigit():
                couleur = "#dcdcdc"
            elif text == "":
                couleur = "#f5f5f5"
            else:
                couleur = "#add8e6"
            self.extra_buttons[idx].config(bg=couleur)
            save_extra_buttons([var.get() for var in self.extra_vars])

    def rotate_photo(self):
        if not self.photos:
            return
        self.angle = (self.angle - 90) % 360
        self.afficher_photo()

    def zoom_in(self):
        self.user_zoom *= 1.25
        self.afficher_photo()

    def zoom_out(self):
        self.user_zoom /= 1.25
        self.afficher_photo()

    def afficher_photo(self):
        if not self.photos:
            return
        path = self.photos[self.index]
        self.log_insert(f"Photo affichée : {os.path.basename(path)}\n")
        try:
            self.image_original = Image.open(path)
            rotated = self.image_original.rotate(self.angle, expand=True)
            self.canvas.update_idletasks()
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            # Calcul du zoom de "fit-to-window" pour l'image
            zoom_fit = 1.0
            if canvas_width > 0 and canvas_height > 0:
                zoom_fit = min(canvas_width / rotated.width, canvas_height / rotated.height)
            # Le zoom final est le produit du zoom de base par le zoom utilisateur
            final_zoom = zoom_fit * self.user_zoom
            new_width = int(rotated.width * final_zoom)
            new_height = int(rotated.height * final_zoom)
            try:
                resample_filter = Image.Resampling.LANCZOS
            except AttributeError:
                resample_filter = Image.ANTIALIAS
            resized = rotated.resize((new_width, new_height), resample_filter)
            self.photo_image = ImageTk.PhotoImage(resized)
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo_image)
            self.canvas.config(scrollregion=(0, 0, new_width, new_height))
        except Exception as e:
            logging.error("Erreur chargement image: %s", e)
            self.log_insert(f"❌ Erreur chargement image : {e}\n")

    def ouvrir_dossier(self):
        folder = filedialog.askdirectory()
        if folder:
            self.current_folder = folder
            self.photos = sorted([
                os.path.join(folder, f)
                for f in os.listdir(folder)
                if f.lower().endswith((".jpg", ".jpeg", ".png"))
            ])
            self.index = 0
            self.angle = 0
            self.user_zoom = 1.0
            if self.photos:
                self.afficher_photo()

    def photo_prec(self):
        if self.photos:
            self.index = (self.index - 1) % len(self.photos)
            self.angle = 0
            self.user_zoom = 1.0
            self.afficher_photo()

    def photo_suiv(self):
        if self.photos:
            self.index = (self.index + 1) % len(self.photos)
            self.angle = 0
            self.user_zoom = 1.0
            self.afficher_photo()

    # --- Alias pour compatibilité avec les anciens boutons de navigation ---
    def prev_photo(self):
        self.photo_prec()

    def next_photo(self):
        self.photo_suiv()

    def rename_photo(self):
        self.renommer_photo()

    def ask_openai_session(self, path):
        fp = image_fingerprint(path)
        data = {}
        if os.path.exists(CORRECTION_FILE):
            with open(CORRECTION_FILE, "r", encoding="utf-8") as f:
                try:
                    data = json.load(f)
                except:
                    data = {}
        if fp in data:
            self.log_insert("Suggestion issue de l'apprentissage automatique.\n")
            return [data[fp]]
        openai_key, client_id, _, _ = load_api_key()
        openai.api_key = openai_key
        self.log_insert("Chargement de l'image...\n")
        url, d_hash = upload_image_to_imgur(path, client_id)
        if not url:
            self.log_insert("Erreur upload Imgur.\n")
            return ["erreur_nom1.jpg"]
        self.log_insert(f"Image envoyée. URL : {url}\n")
        if d_hash:
            self.imgur_deletions.append((d_hash, client_id))
        try:
            resp = openai.ChatCompletion.create(
                model="gpt-4-turbo",
                messages=[
                    {"role": "system", "content": "Tu es un assistant qui aide à nommer des photos de diagnostics immobiliers."},
                    {"role": "user", "content": [
                        {"type": "text", "text": (
                            "Voici une photo. Propose quatre noms de fichier logiques et structurés, "
                            "séparés par des retours à la ligne, au format categorie_detail.jpg. "
                            "Exemples : electricite_prise.jpg, gaz_chaudiere.jpg, amiante_toiture.jpg. "
                            "Ne commente pas. Réponds uniquement avec les quatre noms."
                        )},
                        {"type": "image_url", "image_url": {"url": url}}
                    ]}
                ],
                max_tokens=100
            )
            txt = resp.choices[0].message.content
            self.log_insert("Réponse reçue de l'IA.\n")
            lines = [l.strip() for l in txt.split("\n") if l.strip()]
            return lines
        except Exception as e:
            logging.error("Erreur OpenAI: %s", e)
            self.log_insert(f"Erreur : {e}\n")
            return ["erreur_nom1.jpg", "erreur_nom2.jpg", "erreur_nom3.jpg", "erreur_nom4.jpg"]

    def analyser_photo(self):
        if not self.photos:
            return
        noms = self.ask_openai_session(self.photos[self.index])
        for i, nm in enumerate(noms[:4]):
            self.suggestions[i].set(nm)
        self.radio_sel.set("0")

    def renommer_photo(self):
        if not self.photos:
            return
        new_name = self.suggestions[int(self.radio_sel.get())].get().strip()
        if not new_name:
            messagebox.showwarning("Nom manquant", "Saisir ou sélectionner un nom")
            return
        if self.var_prefix.get():
            prefix = int(self.spin_prefix.get())
            new_name = f"{prefix}_{new_name}"
            self.spin_prefix.delete(0, "end")
            self.spin_prefix.insert(0, str(prefix + 1))
        old_path = self.photos[self.index]
        folder = os.path.dirname(old_path)
        ext = os.path.splitext(old_path)[1]
        target_path = os.path.join(folder, new_name + ext)
        counter = 1
        while os.path.exists(target_path):
            target_path = os.path.join(folder, f"{new_name}_{counter}{ext}")
            counter += 1
        try:
            os.rename(old_path, target_path)
            self.photos[self.index] = target_path
            fp = image_fingerprint(target_path)
            save_correction(fp, new_name)
            self.log_insert(f"✅ Photo renommée : {os.path.basename(target_path)}\n")
            self.index = (self.index + 1) % len(self.photos)
            self.angle = 0
            self.user_zoom = 1.0
            self.afficher_photo()
        except Exception as e:
            logging.error("Erreur renommage: %s", e)
            self.log_insert(f"❌ Erreur renommage : {e}\n")

    def analyser_toutes_photos(self):
        total = len(self.photos)
        self.progress["maximum"] = total
        for i, old_path in enumerate(self.photos):
            noms = self.ask_openai_session(old_path)
            if noms:
                nm = noms[0].strip()
                if self.var_prefix.get():
                    prefix = int(self.spin_prefix.get())
                    nm = f"{prefix}_{nm}"
                    self.spin_prefix.delete(0, "end")
                    self.spin_prefix.insert(0, str(prefix + 1))
                ext = os.path.splitext(old_path)[1]
                target_path = os.path.join(os.path.dirname(old_path), nm + ext)
                counter = 1
                while os.path.exists(target_path):
                    target_path = os.path.join(os.path.dirname(old_path), f"{nm}_{counter}{ext}")
                    counter += 1
                try:
                    os.rename(old_path, target_path)
                    self.photos[i] = target_path
                    fp = image_fingerprint(target_path)
                    save_correction(fp, nm)
                    self.log_insert(f"✅ {os.path.basename(target_path)} renommée automatiquement\n")
                except Exception as e:
                    logging.error("Erreur renommage auto pour %s: %s", old_path, e)
                    self.log_insert(f"❌ Erreur renommage {old_path}: {e}\n")
            self.progress["value"] = i + 1
            self.master.update_idletasks()
        self.log_insert("\n🔁 Analyse automatique terminée.\n")
        self.index = 0
        self.angle = 0
        self.user_zoom = 1.0
        self.afficher_photo()

    def configurer_api(self):
        cfg_win = tk.Toplevel(self.master)
        cfg_win.title("Configurer clés API")
        op, im, _, _ = load_api_key()
        tk.Label(cfg_win, text="Clé OpenAI :").pack(pady=5)
        var_open = tk.StringVar(value=op)
        tk.Entry(cfg_win, textvariable=var_open, width=60).pack(pady=5)
        tk.Label(cfg_win, text="Client ID Imgur :").pack(pady=5)
        var_img = tk.StringVar(value=im)
        tk.Entry(cfg_win, textvariable=var_img, width=60).pack(pady=5)
        def do_save():
            save_api_keys(var_open.get(), var_img.get())
            messagebox.showinfo("OK", "Clés API enregistrées.")
            cfg_win.destroy()
        tk.Button(cfg_win, text="Enregistrer", command=do_save).pack(pady=10)

    def configurer_logo(self):
        filename = filedialog.askopenfilename(
            title="Sélectionner un logo",
            filetypes=[("Fichiers image", "*.png;*.ico;*.gif;*.jpg;*.jpeg")]
        )
        if filename:
            cfg = load_config()
            cfg["logo_path"] = filename
            save_config(cfg)
            try:
                img = Image.open(filename)
                self.logo_image = ImageTk.PhotoImage(img)
                self.master.iconphoto(True, self.logo_image)
                messagebox.showinfo("OK", "Logo mis à jour.")
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de charger le logo : {e}")

    def configurer_preset_file(self):
        preset_file = filedialog.askopenfilename(
            title="Sélectionner le fichier JSON des presets",
            filetypes=[("Fichiers JSON", "*.json")]
        )
        if preset_file:
            set_preset_file(preset_file)
            self.update_preset_list()
            messagebox.showinfo("OK", f"Fichier presets configuré : {preset_file}")

    def afficher_aide(self):
        aide = tk.Toplevel(self.master)
        aide.title("Aide détaillée - Version 6.1")
        txt = (
            "=== GUIDE D'UTILISATION DE PHOTO RENAME - Version 6.1 ===\n\n"
            "Introduction:\n"
            "    Cette application permet de renommer et d'organiser vos photos de diagnostics immobiliers.\n"
            "    Vous pouvez obtenir des suggestions via l'intelligence artificielle d'OpenAI, appliquer des rotations, zoomer sur l'image, "
            "utiliser des boutons personnalisables (presets) et organiser automatiquement vos photos dans des dossiers.\n\n"
            
            "Étape 1 : Ouvrir un dossier\n"
            "    - Cliquez sur 'Fichier' puis 'Ouvrir un dossier' pour choisir le dossier contenant vos images.\n"
            "    - Si un dossier par défaut a été défini (via CopiePhoto S10), il se charge automatiquement.\n\n"
            
            "Étape 2 : Navigation, rotation et zoom\n"
            "    - Utilisez '⏮️ Précédent' et '⏭️ Suivant' pour naviguer entre les photos.\n"
            "    - Cliquez sur '↻ Rotation 90°' pour faire pivoter l'image de 90° dans le sens horaire.\n"
            "    - 'Zoom +' agrandit l'image et 'Zoom -' la réduit. Une fois agrandie, utilisez les scrollbars ou le clic-glissé pour explorer l'image.\n\n"
            
            "Étape 3 : Renommage et suggestions IA\n"
            "    - Cliquez sur 'Analyser' pour générer jusqu'à 4 suggestions de noms via l'IA.\n"
            "    - Sélectionnez une suggestion (ou saisissez votre propre nom) et cliquez sur 'Renommer'.\n"
            "    - Si le nom existe déjà, un suffixe numérique sera ajouté automatiquement.\n\n"
            
            "Étape 4 : Boutons personnalisables et presets\n"
            "    - Vous disposez de 18 boutons pour insérer rapidement du texte dans la zone de suggestions.\n"
            "    - Le premier bouton affiche toujours le nom du preset en cours (en minuscules suivi d'un underscore, par exemple 'electricite_').\n"
            "    - Dans le menu 'Paramètres', vous pouvez enregistrer, charger ou supprimer des presets et choisir le fichier JSON utilisé pour ces presets.\n\n"
            
            "Étape 5 : Organisation des dossiers\n"
            "    - Utilisez 'Organiser dossiers' pour déplacer automatiquement vos photos renommées dans des sous-dossiers basés sur le premier mot du nouveau nom (par exemple, 'GAZ' pour 'gaz_conducteur.jpg').\n\n"
            
            "Glossaire:\n"
            "    Preset     : Configuration sauvegardée des boutons personnalisables permettant d'insérer rapidement du texte.\n"
            "    DPI        : Dots Per Inch, mesure de la résolution d'une image.\n"
            "    Rotation   : Faire pivoter l'image d'un angle (ex. 90° dans le sens horaire).\n"
            "    Zoom       : Agrandissement ou réduction de la taille de l'image affichée.\n\n"
            
            "FAQ:\n"
            "    Q1 : Comment ouvrir un dossier ?\n"
            "         R1 : Dans le menu 'Fichier', sélectionnez 'Ouvrir un dossier' et choisissez le répertoire contenant vos images.\n\n"
            "    Q2 : Pourquoi le premier bouton affiche-t-il 'preset_' ?\n"
            "         R2 : Il reflète automatiquement le nom du preset en cours afin d'assurer la cohérence de l'insertion.\n\n"
            "    Q3 : Comment ajuster le zoom et déplacer l'image ?\n"
            "         R3 : Utilisez 'Zoom +' et 'Zoom -' pour ajuster l'image. Les scrollbars et le clic-glissé (panning) permettent d'explorer l'image agrandie.\n\n"
            "    Q4 : Où configurer mes presets ?\n"
            "         R4 : Dans le menu 'Paramètres', utilisez 'Configurer fichier presets' pour sélectionner ou modifier le fichier JSON des presets.\n\n"
            
            "Pour toute question supplémentaire, consultez ce guide ou contactez le support technique."
        )
        st = scrolledtext.ScrolledText(aide, width=90, height=30)
        st.pack(padx=10, pady=10)
        st.insert("end", txt)
        st.config(state="disabled")

    def afficher_faq(self):
        faq = tk.Toplevel(self.master)
        faq.title("FAQ - PHOTO RENAME")
        txt = (
            "=== FAQ ===\n\n"
            "Q1 : L'image ne se charge pas ou est noire ?\n"
            "    R1 : Assurez-vous que le dossier contient des fichiers compatibles (.jpg, .jpeg, .png).\n\n"
            "Q2 : Pourquoi le premier bouton affiche-t-il 'preset_' ?\n"
            "    R2 : Il est programmé pour refléter le nom du preset actuel (en minuscules suivi d'un underscore).\n\n"
            "Q3 : Que faire si le renommage échoue ?\n"
            "    R3 : Un message d'erreur apparaît. Consultez le fichier log (app.log) pour plus de détails.\n\n"
            "Q4 : Comment utiliser le zoom ?\n"
            "    R4 : Cliquez sur 'Zoom +' pour agrandir et 'Zoom -' pour réduire l'image. Utilisez les scrollbars ou le clic-glissé pour explorer l'image agrandie.\n\n"
            "Q5 : Où configurer les presets et le fichier JSON associé ?\n"
            "    R5 : Dans le menu 'Paramètres', sélectionnez 'Configurer fichier presets' pour choisir le fichier JSON des presets.\n"
        )
        st = scrolledtext.ScrolledText(faq, width=90, height=20)
        st.pack(padx=10, pady=10)
        st.insert("end", txt)
        st.config(state="disabled")

    # --- Gestion des Presets ---
    def load_presets_data(self):
        """
        Charge :
          - custom_presets: dict { nom_preset: [80 labels] }
          - fixed_buttons : [n labels] (vos 45+ boutons)
        Depuis vos fichiers JSON (button_presets.json & extra-buttons.json).
        """
        base = os.path.dirname(__file__)
        # 1) Custom presets
        fp_presets = os.path.join(base, "button_presets.json")
        with open(fp_presets, "r", encoding="utf-8") as f:
            custom_presets = json.load(f)
        # Si vous avez 81 entrées (1 préfixe + 80 labels), on retire l’élément 0 :
        for k, lst in custom_presets.items():
            if len(lst) >= 81:
                custom_presets[k] = lst[1:81]
            elif len(lst) >= 80:
                custom_presets[k] = lst[:80]
            else:
                # preset invalide => on laisse tel quel
                custom_presets[k] = lst

        # 2) Boutons fixes
        fp_fixed = os.path.join(base, "extra-buttons.json")
        with open(fp_fixed, "r", encoding="utf-8") as f:
            fixed_buttons = json.load(f)

        return custom_presets, fixed_buttons


    def update_preset_list(self):
        presets = self.load_presets_data()
        self.preset_combobox['values'] = list(presets.keys())

    def load_preset(self, event=None):
        presets = self.load_presets_data()
        preset_name = self.preset_var.get().strip()
        if preset_name in presets:
            values = presets[preset_name]
            if len(values) == len(self.custom_vars):
                for i, v in enumerate(values):
                    self.custom_vars[i].set(v)
                preset_clean = preset_name.lower()
                if not preset_clean.endswith("_"):
                    preset_clean += "_"
                self.custom_vars[0].set(preset_clean)
                self.log_insert(f"Preset '{preset_name}' chargé.\n")
            else:
                messagebox.showwarning("Avertissement", "Le preset ne correspond pas au nombre de boutons.")
        else:
            messagebox.showinfo("Info", "Preset introuvable.")
        self.mettre_a_jour_couleurs_boutons()


    def save_current_preset(self):
        preset_name = self.preset_var.get().strip()
        if not preset_name:
            preset_name = simpledialog.askstring("Nom du preset", "Entrez un nom pour le preset :")
            if not preset_name:
                return
            self.preset_var.set(preset_name)
        presets = self.load_presets_data()
        presets[preset_name] = [var.get() for var in self.custom_vars]
        preset_file = get_preset_file()
        with open(preset_file, "w", encoding="utf-8") as f:
            json.dump(presets, f, indent=2)
        messagebox.showinfo("Preset enregistré", f"Le preset '{preset_name}' a été enregistré.")
        self.update_preset_list()

    def delete_current_preset(self):
        preset_name = self.preset_var.get().strip()
        if not preset_name:
            messagebox.showwarning("Supprimer preset", "Aucun preset spécifié.")
            return
        presets = self.load_presets_data()
        if preset_name in presets:
            rep = messagebox.askyesno("Supprimer preset", f"Voulez-vous vraiment supprimer le preset '{preset_name}' ?")
            if rep:
                del presets[preset_name]
                preset_file = get_preset_file()
                with open(preset_file, "w", encoding="utf-8") as f:
                    json.dump(presets, f, indent=2)
                messagebox.showinfo("Preset supprimé", f"Le preset '{preset_name}' a été supprimé.")
                self.update_preset_list()
        else:
            messagebox.showinfo("Supprimer preset", f"Le preset '{preset_name}' n'existe pas.")

    def organiser_dossiers(self):
        if not self.current_folder:
            if self.default_working_directory:
                use_default = messagebox.askyesno(
                    "Dossier par défaut",
                    f"Aucun dossier ouvert.\nUtiliser le dossier par défaut ?\n{self.default_working_directory}"
                )
                if use_default:
                    self.current_folder = self.default_working_directory
                else:
                    chosen = filedialog.askdirectory(title="Choisir un dossier")
                    if chosen:
                        self.current_folder = chosen
            else:
                chosen = filedialog.askdirectory(title="Choisir un dossier")
                if chosen:
                    self.current_folder = chosen
        if not self.current_folder:
            messagebox.showwarning("Attention", "Aucun dossier valide.")
            return
        rep = messagebox.askyesno("Organisation", "Voulez-vous classer les photos renommées dans de nouveaux dossiers ?")
        if not rep:
            return
        if not os.path.exists(CORRECTION_FILE):
            messagebox.showinfo("Aucune correction", "Aucune photo renommée trouvée (corrections.json introuvable).")
            return
        with open(CORRECTION_FILE, "r", encoding="utf-8") as f:
            try:
                corrections = json.load(f)
            except:
                corrections = {}
        moved_count = 0
        for path in self.photos:
            fp = image_fingerprint(path)
            if fp in corrections:
                new_name = corrections[fp]
                name_clean = re.sub(r'^[0-9]+_?', '', new_name)
                folder_name = name_clean.split("_")[0] if "_" in name_clean else name_clean
                folder_name = folder_name.upper()
                target_folder = os.path.join(self.current_folder, folder_name)
                if not os.path.exists(target_folder):
                    try:
                        os.makedirs(target_folder)
                    except Exception as e:
                        self.log_insert(f"❌ Erreur création dossier {target_folder} : {e}\n")
                        continue
                if os.path.dirname(path) != target_folder:
                    try:
                        new_path = os.path.join(target_folder, os.path.basename(path))
                        shutil.move(path, new_path)
                        self.log_insert(f"✅ Déplacé : {os.path.basename(path)} vers {folder_name}\n")
                        moved_count += 1
                    except Exception as e:
                        self.log_insert(f"❌ Erreur déplacement de {path} : {e}\n")
        messagebox.showinfo("Organisation terminée", f"{moved_count} photo(s) déplacée(s).")

    def load_tab_preset(self, frame, preset_name):
        all_presets, fixed_buttons = self.load_presets_data()
        labels80 = all_presets.get(preset_name, [])
        if len(labels80) < 1:
            messagebox.showwarning(
                "Preset introuvable",
                f"Le preset '{preset_name}' n'existe pas ou est vide.")
            return

        # 1) On vide l’onglet
        for w in frame.winfo_children():
            w.destroy()

        # 2) Création des 80 boutons (10×8)
        cols, w_btn, h_btn, m = 10, 90, 25, 4
        for idx, text in enumerate(labels80):
            r, c = divmod(idx, cols)
            btn = tk.Button(frame, text=text, width=int(w_btn/8))
            btn.config(bg=("#e0e0e0" if text.lower().startswith("btn") else "#d0eaff"))
            btn.grid(row=r, column=c, padx=m, pady=m)
            btn.bind("<Button-1>", lambda e, t=text: self.on_custom_left_by_label(t))
            btn.bind("<Button-3>", lambda e, i=idx: self.on_custom_right(frame, i))

        # 3) Ligne fixe (vos 45+ boutons)
        fixed_row = 8
        for idx, text in enumerate(fixed_buttons):
            btn = tk.Button(frame, text=text, width=int(w_btn/8))
            btn.config(bg="#c0c0c0")
            btn.grid(row=fixed_row, column=idx, padx=m, pady=m, sticky="w")
            btn.bind("<Button-1>", lambda e, t=text: self.on_fixed_left(t))
            btn.bind("<Button-3>", lambda e, t=text: self.on_fixed_right(t))

        # 4) On renomme l’onglet
        self.tab_control.tab(frame.index - 1, text=preset_name)

    def on_tab_double_click(self, event):
        """
        Gestionnaire du double-clic sur la barre d'onglets :
        demande à l'utilisateur quel preset affecter.
        """
        # Déterminer quel onglet a été cliqué
        x, y = event.x, event.y
        for idx, frame in enumerate(self.tab_frames):
            bbox = self.tab_control.bbox(idx)
            if bbox and bbox[0] <= x <= bbox[0] + bbox[2]:
                choix = simpledialog.askstring(
                    "Choisir preset",
                    "Nom du preset à charger dans cet onglet :",
                    parent=self.master)
                if choix:
                    # Mémoriser et recharger
                    ConfigManager.Current.TabPresetMapping[frame.index] = choix
                    self.load_tab_preset(frame, choix)
                break

def launch_photo_rename():
    root = tk.Tk()
    app = PhotoRenameApp(root)
    root.mainloop()

# ==========================================================================
#                      MODULE COPIEPHOTO S10 (Simplifié)
# ==========================================================================
fichiers_copies = []
annuler_copie = False
destination_copie = ""
etat_copie_terminee = False

def ouvrir_fenetre_deplacement_ou_suppression():
    global fichiers_copies
    if not fichiers_copies:
        messagebox.showinfo("Aucun fichier", "Aucun fichier à traiter.")
        return
    fenetre = tk.Toplevel()
    fenetre.title("Déplacer ou supprimer les fichiers")
    fenetre.geometry("400x150")
    tk.Label(fenetre, text="Que souhaitez-vous faire des fichiers copiés ?", font=("Arial", 11)).pack(pady=10)
    def supprimer():
        for f in fichiers_copies:
            try:
                os.remove(f)
            except:
                pass
        fichiers_copies.clear()
        messagebox.showinfo("Supprimé", "Tous les fichiers copiés ont été supprimés.")
        fenetre.destroy()
    def deplacer():
        dossier = filedialog.askdirectory(title="Choisir un dossier pour déplacer les fichiers")
        if not dossier:
            return
        for f in fichiers_copies:
            try:
                shutil.move(f, os.path.join(dossier, os.path.basename(f)))
            except Exception as e:
                print("Erreur déplacement :", e)
        fichiers_copies.clear()
        messagebox.showinfo("Déplacement terminé", f"Les fichiers ont été déplacés vers :\n{dossier}")
        fenetre.destroy()
    tk.Button(fenetre, text="🗑️ Supprimer les fichiers", font=("Arial", 11), width=30, command=supprimer).pack(pady=5)
    tk.Button(fenetre, text="📁 Déplacer les fichiers", font=("Arial", 11), width=30, command=deplacer).pack(pady=5)

def lire_chemin_dossier_copie():
    CF_HDROP = 15
    win32clipboard.OpenClipboard()
    try:
        if win32clipboard.IsClipboardFormatAvailable(CF_HDROP):
            data = win32clipboard.GetClipboardData(CF_HDROP)
            if data and os.path.isdir(data[0]):
                return data[0]
        else:
            data = win32clipboard.GetClipboardData()
            if os.path.isdir(data):
                return data
    except:
        return None
    finally:
        win32clipboard.CloseClipboard()

def recuperer_photos_camera():
    try:
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("Shell.Application")
        poste = shell.Namespace(17)

        tel = next((item for item in poste.Items() if "S10+" in item.Name), None)
        if not tel:
            logging.warning("Téléphone non détecté dans Shell.Namespace(17).")
            raise Exception("Téléphone non détecté ou non branché.")

        sd = next((item for item in tel.GetFolder.Items() if "Carte SD" in item.Name), None)
        if not sd:
            logging.warning("Carte SD introuvable dans le téléphone.")
            raise Exception("Carte SD introuvable.")

        dcim = next((item for item in sd.GetFolder.Items() if item.Name == "DCIM"), None)
        if not dcim:
            logging.warning("Dossier DCIM introuvable.")
            raise Exception("Dossier DCIM introuvable.")

        camera = next((item for item in dcim.GetFolder.Items() if item.Name == "Camera"), None)
        if not camera:
            logging.warning("Dossier Camera introuvable.")
            raise Exception("Dossier Camera introuvable.")

        fichiers = [f for f in camera.GetFolder.Items() if not f.IsFolder]
        logging.info(f"{len(fichiers)} fichiers trouvés dans DCIM/Camera.")
        return fichiers

    except Exception as e:
        logging.error("Erreur dans recuperer_photos_camera : %s", e)
        messagebox.showerror("Erreur", f"Impossible d'accéder au téléphone.\nDétail : {e}")
        return []


def extraire_dates_fichiers(fichiers):
    return list(set(f.ModifyDate for f in fichiers if hasattr(f, 'ModifyDate')))

def filtrer_fichiers_par_date(fichiers, date_choisie):
    return [f for f in fichiers if getattr(f, 'ModifyDate', None) == date_choisie]

def enregistrer_historique(destination, nb_photos):
    with open(HISTORIQUE_PATH, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now().strftime('%d/%m/%Y - %H:%M')}]\n")
        f.write(f"Dossier destination : {destination}\n")
        f.write(f"Nombre de photos copiées : {nb_photos}\n---\n")

def copier_via_copyhere(fichiers, destination_path, champ_progression, barre, bouton_annuler_deplacer):
    global annuler_copie, fichiers_copies, destination_copie, etat_copie_terminee
    shell = win32com.client.Dispatch("Shell.Application")
    dossier_dest = shell.NameSpace(destination_path)
    copiees = 0
    total = len(fichiers)
    fichiers_copies.clear()
    etat_copie_terminee = False
    bouton_annuler_deplacer.config(state="normal", text="❌ Annuler la copie")
    for i, f in enumerate(fichiers, start=1):
        if annuler_copie:
            break
        try:
            with open("debug_copie.txt", "a", encoding="utf-8") as logf:
                logf.write(f"Tentative de copie : {getattr(f, 'Name', '???')} vers {destination_path}\n")

            dossier_dest.CopyHere(f, 20)
            
            fichiers_copies.append(os.path.join(destination_path, f.Name))
            copiees += 1
        except Exception as e:
            print("Erreur copie :", e)
            continue
        champ_progression.config(text=f"{i}/{total} fichiers copiés")
        barre['value'] = (i/total) * 100
        barre.update_idletasks()
        time.sleep(0.2)
    if annuler_copie:
        if messagebox.askyesno("Annulation", "Souhaitez-vous supprimer les fichiers déjà copiés ?"):
            for f in fichiers_copies:
                try:
                    os.remove(f)
                except:
                    pass
        fichiers_copies.clear()
        messagebox.showinfo("Copie annulée", "La copie a été annulée.")
    else:
        enregistrer_historique(destination_path, copiees)
        messagebox.showinfo("Copie terminée", f"{copiees} photos copiées.")
        bouton_annuler_deplacer.config(text="📂 Déplacer ou supprimer les fichiers", state="normal")
        etat_copie_terminee = True


        # Proposer d'ouvrir le dossier source (spécifique téléphone MTP)
    rep = messagebox.askyesno(
        "Ouvrir le téléphone",
        "Souhaitez-vous ouvrir l'explorateur Windows pour accéder au dossier source ?"
    )
    if rep:
        try:
            subprocess.Popen("explorer")  # Ouvre une fenêtre "Ce PC"
        except Exception as e:
            logging.error("Erreur ouverture explorateur : %s", e)
            messagebox.showerror("Erreur", f"Impossible d’ouvrir l’explorateur : {e}")




def lancer_copie(destination, fichiers, champ_progression, barre, bouton_annuler):
    from threading import Thread
    import pythoncom
    import win32com.client

    def is_virtual_file(fichier):
        try:
            path = fichier.Path
            return not os.path.exists(path)
        except:
            return True  # Pas d’attribut Path = fichier virtuel

    def worker():
        bouton_annuler.config(state=tk.NORMAL)
        champ_progression.config(text="Copie en cours...")
        barre.config(value=0, maximum=len(fichiers))

        destination_path = destination
        if not os.path.isdir(destination_path):
            os.makedirs(destination_path)

        fichiers_copies = []

        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("Shell.Application")

        destination_path = os.path.normpath(destination_path)
        dossier_shell = shell.NameSpace(destination_path)

        if dossier_shell is None:
            messagebox.showerror("Erreur", f"Impossible d'accéder à {destination_path}")
            return
        
        for idx, f in enumerate(fichiers):
            if bouton_annuler.canceled:
                break
            try:
                dossier_shell.CopyHere(f, 20)  # Forcé, même si non virtuel
                time.sleep(0.2)  # Pause utile pour éviter les ratés MTP
                fichiers_copies.append(os.path.join(destination_path, f.Name))
            except Exception as e:
                print(f"Erreur de copie pour {f.Name} : {e}")
            barre.config(value=idx + 1)
            champ_progression.config(text=f"{idx+1}/{len(fichiers)} fichiers copiés")
            champ_progression.update_idletasks()

        bouton_annuler.config(state=tk.DISABLED)
        champ_progression.config(text="Copie terminée")
        messagebox.showinfo("Copie terminée", f"{len(fichiers_copies)} fichiers copiés.")

    bouton_annuler.canceled = False
    bouton_annuler.config(command=lambda: setattr(bouton_annuler, "canceled", True))
    Thread(target=worker).start()




def action_valider(destination, champ_progression, barre, bouton_annuler_deplacer):
    if not destination or not os.path.isdir(destination):
        messagebox.showwarning("Erreur", "Aucun dossier valide.")
        return
    set_default_working_directory(destination)
    fichiers = recuperer_photos_camera()

    infos = "\n".join([str(getattr(f, 'Path', '???')) for f in fichiers])
    messagebox.showinfo("Fichiers récupérés", f"{len(fichiers)} fichiers trouvés.")

    if not fichiers:
        messagebox.showinfo("Aucune photo", "Aucun fichier trouvé dans DCIM/Camera.")
        return
    dates = extraire_dates_fichiers(fichiers)
    if len(dates) <= 1:
        lancer_copie(destination, fichiers, champ_progression, barre, bouton_annuler_deplacer)
    else:
        choix = tk.Toplevel()
        choix.title("Filtrer les photos par date")
        choix.geometry("400x200")
        tk.Label(choix, text="Des dates différentes ont été détectées.\nSélectionnez celle à copier :", font=("Arial", 11)).pack(pady=10)
        liste = tk.Listbox(choix, selectmode="single", height=6)
        for d in dates:
            liste.insert(tk.END, d)
        liste.pack()
        def valider_selection():
            selected = liste.curselection()
            if selected:
                date_choisie = dates[selected[0]]
                fichiers_filtres = filtrer_fichiers_par_date(fichiers, date_choisie)
                choix.destroy()
                lancer_copie(destination, fichiers_filtres, champ_progression, barre, bouton_annuler_deplacer)
        tk.Button(choix, text="Lancer la copie", command=valider_selection).pack(pady=10)

def choisir_et_lancer(champ_chemin, champ_progression, barre, bouton_annuler_deplacer):
    dossier = filedialog.askdirectory(title="Choisir un dossier de destination")
    if not dossier:
        return

    champ_chemin.config(text=dossier)

    # ✅ Mise à jour du répertoire actif utilisé pour les traitements
    set_default_working_directory(dossier)

    fichiers = recuperer_photos_camera()
    if not fichiers:
        messagebox.showerror("Erreur", "Aucune photo trouvée sur le téléphone.\nVérifiez la connexion ou le dossier DCIM/Camera.")
        return

    dates = extraire_dates_fichiers(fichiers)
    if len(dates) <= 1:
        lancer_copie(dossier, fichiers, champ_progression, barre, bouton_annuler_deplacer)
    else:
        choix = tk.Toplevel()
        choix.title("Filtrer les photos par date")
        choix.geometry("400x200")
        tk.Label(choix, text="Des dates différentes ont été détectées.\nSélectionnez celle à copier :", font=("Arial", 11)).pack(pady=10)
        liste = tk.Listbox(choix, selectmode="single", height=6)
        for d in dates:
            liste.insert(tk.END, d)
        liste.pack()

        def valider_selection():
            selected = liste.curselection()
            if selected:
                date_choisie = dates[selected[0]]
                fichiers_filtres = filtrer_fichiers_par_date(fichiers, date_choisie)
                choix.destroy()
                lancer_copie(dossier, fichiers_filtres, champ_progression, barre, bouton_annuler_deplacer)

        tk.Button(choix, text="Lancer la copie", command=valider_selection).pack(pady=10)



def edit_s10_settings():
    cfg = load_config()
    phone_paths = cfg.get("phone_paths", [])
    default_phone_path = cfg.get("default_phone_path", "")

    win = tk.Toplevel()
    win.title("Paramètres des téléphones portables")
    win.geometry("500x350")

    tk.Label(win, text="Liste des chemins configurés", font=("Arial", 11)).pack(pady=5)

    frame_list = tk.Frame(win)
    frame_list.pack(pady=5, fill="both", expand=True)

    listbox = tk.Listbox(frame_list, height=10)
    listbox.pack(side=tk.LEFT, fill="both", expand=True)

    scrollbar = tk.Scrollbar(frame_list, orient="vertical", command=listbox.yview)
    scrollbar.pack(side=tk.RIGHT, fill="y")
    listbox.config(yscrollcommand=scrollbar.set)

    for path in phone_paths:
        listbox.insert(tk.END, path)

    # ✅ Ajout
    def add_path():
        newp = simpledialog.askstring(
            "Saisir le chemin",
            "Entrez le chemin vers votre téléphone (ex : Ce PC\\S10+\\Carte SD\\DCIM\\Camera)"
        )
        if newp:
            phone_paths.append(newp)
            listbox.insert(tk.END, newp)

    # ❌ Suppression
    def remove_path():
        sel = listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        val = listbox.get(idx)
        if messagebox.askyesno("Supprimer", f"Supprimer le chemin '{val}' ?"):
            phone_paths.remove(val)
            listbox.delete(idx)

    # ⭐ Définir comme chemin par défaut
    def set_default():
        sel = listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        val = listbox.get(idx)
        cfg["default_phone_path"] = val
        messagebox.showinfo("Défaut défini", f"Chemin par défaut :\n{val}")

    # 💾 Enregistrement
    def do_save():
        cfg["phone_paths"] = phone_paths
        save_config(cfg)
        messagebox.showinfo("Enregistré", "Paramètres S10 sauvegardés.")
        win.destroy()

    # 📦 Boutons
    btn_frame = tk.Frame(win)
    btn_frame.pack(pady=10)

    tk.Button(btn_frame, text="Ajouter", width=15, command=add_path).grid(row=0, column=0, padx=5)
    tk.Button(btn_frame, text="Supprimer", width=15, command=remove_path).grid(row=0, column=1, padx=5)
    tk.Button(btn_frame, text="Définir par défaut", width=15, command=set_default).grid(row=0, column=2, padx=5)

    tk.Button(win, text="Enregistrer", command=do_save, font=("Arial", 11)).pack(pady=10)

    win.mainloop()




def edit_usb_settings():
    cfg = load_config()
    usb_paths = cfg.get("usb_paths", [])
    default_usb_path = cfg.get("default_usb_path", "")

    win = tk.Toplevel()
    win.title("Paramètres USB")
    win.geometry("500x350")

    tk.Label(win, text="Liste des chemins configurés (supports USB)", font=("Arial", 11)).pack(pady=5)

    frame_list = tk.Frame(win)
    frame_list.pack(pady=5, fill="both", expand=True)

    listbox = tk.Listbox(frame_list, height=10)
    listbox.pack(side=tk.LEFT, fill="both", expand=True)

    scrollbar = tk.Scrollbar(frame_list, orient="vertical", command=listbox.yview)
    scrollbar.pack(side=tk.RIGHT, fill="y")
    listbox.config(yscrollcommand=scrollbar.set)

    for path in usb_paths:
        listbox.insert(tk.END, path)

    def add_path():
        newp = filedialog.askdirectory(title="Sélectionner un dossier USB")
        if newp:
            usb_paths.append(newp)
            listbox.insert(tk.END, newp)

    def remove_path():
        sel = listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        val = listbox.get(idx)
        if messagebox.askyesno("Supprimer", f"Supprimer le chemin '{val}' ?"):
            usb_paths.remove(val)
            listbox.delete(idx)

    def set_default():
        sel = listbox.curselection()




def launch_copiephoto_s10():
    global annuler_copie, fichiers_copies, destination_copie, etat_copie_terminee
    root = tk.Tk()  # ✅ Création de la fenêtre AVANT tout le reste
    root.title("📸 Copier mes photos S10+")
    root.state('zoomed')
    chemin_copie_var = tk.StringVar(value=lire_chemin_dossier_copie() or "")

    tk.Label(root, text="Assistant de copie depuis le téléphone S10+", font=("Arial", 14, "bold")).pack(pady=10)
    tk.Label(root, text="📋 Dossier détecté dans le presse-papiers :", font=("Arial", 11)).pack()
    champ_chemin = tk.Label(root, textvariable=chemin_copie_var, fg="blue", font=("Arial", 10), wraplength=580)

    def actualiser_chemin():
        nouveau_chemin = lire_chemin_dossier_copie() or ""
        champ_chemin.config(text=nouveau_chemin or "(Aucun dossier détecté)")
        chemin_copie_var.set(nouveau_chemin)

    btn_refresh = tk.Button(root, text="🔄 Actualiser le dossier détecté", font=("Arial", 11), width=40, command=actualiser_chemin)
    btn_refresh.pack(pady=5)

    champ_chemin.pack(pady=5)
    champ_progression = tk.Label(root, text="Aucune copie en cours", font=("Arial", 10))
    champ_progression.pack(pady=5)
    barre = ttk.Progressbar(root, length=500, mode='determinate')
    barre.pack(pady=5)
    bouton_annuler_deplacer = tk.Button(root, text="📂 Déplacer ou supprimer les fichiers", font=("Arial", 11), width=40, state="disabled", command=ouvrir_fenetre_deplacement_ou_suppression)
    bouton_annuler_deplacer.pack(pady=5)
    bouton_copier_auto = tk.Button(root,text="📋 Copier dans ce dossier détecté",font=("Arial", 11),width=40,command=lambda: action_valider(chemin_copie_var.get(), champ_progression, barre, bouton_annuler_deplacer))
    bouton_copier_auto.pack(pady=10)
    bouton_choisir = tk.Button(root, text="📁 Choisir un autre dossier", font=("Arial", 11), width=40, command=lambda: choisir_et_lancer(champ_chemin, champ_progression, barre, bouton_annuler_deplacer))
    bouton_choisir.pack(pady=5)
    tk.Button(root, text="📱 Ouvrir le dossier du téléphone", font=("Arial", 11), width=40, command=ouvrir_dossier_telephone).pack(pady=5)
    tk.Button(root, text="Retour au menu", font=("Arial", 11), command=lambda: [root.destroy(), launch_main_menu()]).pack(pady=5)
    menubar = tk.Menu(root)
    menu_s10 = tk.Menu(menubar, tearoff=0)
    menu_s10.add_command(label="Paramètres S10", command=edit_s10_settings)
    menu_s10.add_command(label="Paramètres USB", command=edit_usb_settings)

    menubar.add_cascade(label="Paramètres S10", menu=menu_s10)
    menubar.add_command(label="Quitter", command=root.destroy)
    root.config(menu=menubar)
    root.mainloop()

def ouvrir_dossier_telephone():
    try:
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("Shell.Application")
        poste = shell.Namespace(17)
        tel = next((item for item in poste.Items() if "S10+" in item.Name), None)
        if not tel:
            raise Exception("Téléphone non détecté")
        sd = next((item for item in tel.GetFolder.Items() if "Carte SD" in item.Name), None)
        if not sd:
            raise Exception("Carte SD non trouvée")
        dcim = next((item for item in sd.GetFolder.Items() if item.Name == "DCIM"), None)
        camera = next((item for item in dcim.GetFolder.Items() if item.Name == "Camera"), None)
        shell.Open(camera)
    except Exception as e:
        messagebox.showerror("Erreur", f"Impossible d’ouvrir le dossier du téléphone.\n\n{e}")














# ==========================================================================
#                     LAUNCHER PRINCIPAL (PHOTO MANAGER PRO)
# ==========================================================================
def launch_main_menu():
    launcher = tk.Tk()
    launcher.title("PHOTO MANAGER PRO")
    launcher.state('zoomed')
    canvas = tk.Canvas(launcher)
    canvas.pack(fill="both", expand=True)
    launcher.update()
    _, _, logo_path, _ = load_api_key()
    if logo_path and os.path.exists(logo_path):
        try:
            img = Image.open(logo_path).convert("RGBA")
            alpha = img.split()[3]
            alpha = alpha.point(lambda p: int(p * 0.3))
            img.putalpha(alpha)
            logo = ImageTk.PhotoImage(img)
            w = launcher.winfo_width()
            h = launcher.winfo_height()
            canvas.create_image(w//2, h//2, anchor=tk.CENTER, image=logo)
            canvas.logo = logo
        except Exception as e:
            logging.error("Erreur affichage logo dans le launcher: %s", e)
    front_frame = tk.Frame(canvas, bg="#ffffff")
    front_frame.place(relx=0.5, rely=0.3, anchor="center")
    tk.Label(front_frame, text="Bienvenue dans PHOTO MANAGER PRO", font=("Arial", 16, "bold"), bg="#ffffff").pack(pady=10)
    tk.Label(front_frame, text="Veuillez choisir une application :", font=("Arial", 12), bg="#ffffff").pack(pady=5)
    def open_copiephoto():
        launcher.destroy()
        launch_copiephoto_s10()
    def open_photo_rename():
        launcher.destroy()
        launch_photo_rename()
    btn1 = tk.Button(front_frame, text="CopiePhoto S10+", font=("Arial", 14), width=20, command=open_copiephoto, bg="#3F51B5", fg="white")
    btn1.pack(pady=10)
    CreateToolTip(btn1, "Lancer l'assistant de copie depuis le téléphone S10+")
    btn2 = tk.Button(front_frame, text="Photo Rename", font=("Arial", 14), width=20, command=open_photo_rename, bg="#4CAF50", fg="white")
    btn2.pack(pady=10)
    CreateToolTip(btn2, "Lancer l'assistant de renommage des photos")
    tk.Button(front_frame, text="Quitter", font=("Arial", 12), width=15, command=launcher.destroy, bg="#f44336", fg="white").pack(pady=20)
    launcher.mainloop()


class ConfigManager:
    """
    Persiste la map onglet→preset dans %APPDATA%/RenommagePhoto/tab_config.json
    """
    path = os.path.join(
        os.getenv("APPDATA") or os.path.expanduser("~/.config"),
        "RenommagePhoto", "tab_config.json")

    @classmethod
    def load(cls):
        try:
            with open(cls.path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except FileNotFoundError:
            data = {}
        cls.Current = cls()
        cls.Current.TabPresetMapping = {
            int(k): v for k, v in data.get("TabPresetMapping", {}).items()
        }

    @classmethod
    def save(cls):
        os.makedirs(os.path.dirname(cls.path), exist_ok=True)
        data = {"TabPresetMapping": cls.Current.TabPresetMapping}
        with open(cls.path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)


if __name__ == "__main__":
    # 1) Avant de créer la fenêtre, on récupère la config (JSON)
    ConfigManager.load()

    # 2) On lance l'application comme avant
    root = tk.Tk()
    app  = PhotoRenameApp(root)
    root.mainloop()

    # 3) Dès que l'utilisateur ferme la fenêtre, on écrit la config à jour
    ConfigManager.save()


