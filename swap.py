import os
import tkinter as tk
import webbrowser
from tkinter import ttk, messagebox, filedialog
import json
import re
import subprocess
import csv
import psutil
import tempfile
import threading
from pathlib import Path
from PIL import Image, ImageDraw
import sys
import configparser


def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


BASE_DIR_SETTINGS = get_base_path()
PROFILE_FILE = os.path.join(BASE_DIR_SETTINGS, "audio_profiles.json")
SETTINGS_FILE = os.path.join(BASE_DIR_SETTINGS, "config.ini")

PROFILE_NAME_REGEX = re.compile(r'^[A-Za-z0-9-]+$')

if hasattr(sys, '_MEIPASS'):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.abspath(".")


class Checker:
    @staticmethod
    def verify_eartrumpet_exe(exe_path):
        try:
            # Try to list devices, should return lines
            cmd = [exe_path, '--list-devices']
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                check=False,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            if result.returncode != 0:
                return False
            # Must contain at least one line matching [Playback] or [Recording]
            for line in result.stdout.splitlines():
                if line.strip().startswith('[Playback]') or line.strip().startswith('[Recording]'):
                    return True
            return False
        except Exception:
            return False


class AutocompleteCombobox(ttk.Combobox):
    def set_completion_list(self, completion_list):
        self._completion_list = sorted(completion_list, key=str.lower)
        self['values'] = self._completion_list
        self['state'] = 'normal'
        self.bind('<KeyRelease>', self.handle_keyrelease)

    def autocomplete(self):
        value = self.get()
        if not value:
            return
        hits = [item for item in self._completion_list if item.lower().startswith(value.lower())]
        if hits:
            self.set(hits[0])
            self.select_range(len(value), tk.END)

    def handle_keyrelease(self, event):
        if event.keysym in ("BackSpace", "Left", "Right", "Up", "Down", "Home", "End", "Tab"):
            return
        self.autocomplete()


class ToolTip:
    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None

    def showtip(self, text, x, y):
        if self.tipwindow:
            self.hidetip()
        if not text:
            return

        x = x + self.widget.winfo_rootx() + 20
        y = y + self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()


class AudioProfileManager:
    def __init__(self):
        self.changes_pending = False
        self.root = tk.Tk()
        self.root.title("SmartWindowsAudioProfiles")
        self.root.iconbitmap(os.path.join(BASE_DIR, "icon.ico"))
        self.root.geometry("1200x600")
        self.root.minsize(800, 600)
        self.config_file = PROFILE_FILE
        self.eartrumpet_path = "EarTrumpet.exe"
        self.profiles = {}
        self.devices = []
        self.ini_path = SETTINGS_FILE
        self.settings = configparser.ConfigParser()
        self.load_ini()
        self.load_config()
        self.create_gui()
        self.center_root()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def center_root(self):
        self.root.update_idletasks()
        w = 1200
        h = 600
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw // 2) - (w // 2)
        y = (sh // 2) - (h // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def load_ini(self):
        self.settings.read(self.ini_path)
        if 'App' not in self.settings:
            self.settings['App'] = {}
        self.eartrumpet_path = self.settings['App'].get('eartrumpet_path', "EarTrumpet.exe")
        self.auto_save_enabled = self.settings['App'].getboolean('auto_save', True)

    def save_ini(self):
        self.settings['App']['eartrumpet_path'] = self.eartrumpet_path
        self.settings['App']['auto_save'] = str(self.auto_save_var.get() if hasattr(self, 'auto_save_var') else True)
        with open(self.ini_path, 'w') as f:
            self.settings.write(f)
        if hasattr(self, "auto_save_var") and self.auto_save_var.get():
            self.changes_pending = False
            self.mark_profiles_tab_unsaved()

    def create_gui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)

        self.profiles_frame = ttk.Frame(notebook)
        notebook.add(self.profiles_frame, text="Profiles")
        self.create_profiles_tab()

        self.devices_frame = ttk.Frame(notebook)
        notebook.add(self.devices_frame, text="Audio Devices")
        self.create_devices_tab()

        self.settings_frame = ttk.Frame(notebook)
        notebook.add(self.settings_frame, text="Settings")
        about_tab = ttk.Frame(notebook)
        notebook.add(about_tab, text="About")
        ttk.Label(about_tab, text="SmartWindowsAudioProfiles\nby dayeggpi\nVersion 1.0.1", font=('Courrier', 9)).pack(pady=50)
        ttk.Button(self.settings_frame, text="Open WindowsVolume Mixer", command=self.open_volume_mixer).pack(pady=5)
        self.create_settings_tab()

    def open_volume_mixer(self):
        try:
            subprocess.run(['start', 'ms-settings:apps-volume'], shell=True)
        except Exception as e:
            try:
                subprocess.run(['start', 'ms-settings:sound'], shell=True)
            except:
                messagebox.showerror("Error", f"Could not open Volume Mixer: {e}", parent=self.root)

    def mark_profiles_tab_unsaved(self):
        notebook = self.profiles_frame.master
        for i in range(notebook.index("end")):
            if "Profiles" in notebook.tab(i, "text"):
                tab_text = "Profiles"
                if self.changes_pending and not (hasattr(self, "auto_save_var") and self.auto_save_var.get()):
                    tab_text += " *"
                notebook.tab(i, text=tab_text)

    def create_profiles_tab(self):
        top_frame = ttk.Frame(self.profiles_frame)
        top_frame.pack(fill='x', padx=5, pady=5)

        ttk.Label(top_frame, text="Profile(s):").pack(side='left')
        self.profile_var = tk.StringVar()
        self.profile_combo = ttk.Combobox(top_frame, textvariable=self.profile_var,
                                          values=list(self.profiles.keys()), state='readonly')
        self.profile_combo.pack(side='left', padx=(5, 10))
        self.profile_combo.bind('<<ComboboxSelected>>', self.on_profile_selected)

        ttk.Button(top_frame, text="New Profile", command=self.new_profile).pack(side='left', padx=2)
        ttk.Button(top_frame, text="Delete Profile", command=self.delete_profile).pack(side='left', padx=2)

        ttk.Button(top_frame, text="Import", command=self.import_profiles).pack(side='right', padx=2)

        rules_frame = ttk.LabelFrame(self.profiles_frame, text="Profile Rules")
        rules_frame.pack(fill='both', expand=True, padx=5, pady=5)

        list_frame = ttk.Frame(rules_frame)
        list_frame.pack(fill='both', expand=True, padx=5, pady=5)

        self.rules_listbox = tk.Listbox(list_frame)
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.rules_listbox.yview)
        self.rules_listbox.configure(yscrollcommand=scrollbar.set)

        self.rules_listbox.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        button_frame = ttk.Frame(rules_frame)
        button_frame.pack(fill='x', padx=5, pady=5)

        ttk.Button(button_frame, text="Edit Rule", command=self.edit_rule).pack(side='left', padx=2)
        ttk.Button(button_frame, text="Delete Rule", command=self.delete_rule).pack(side='left', padx=2)

        self.activate_button = ttk.Button(top_frame, text="Activate Profile", command=self.activate_profile)
        self.activate_button.pack(side='left', padx=2)

        self.export_button = ttk.Button(top_frame, text="Export", command=self.export_profiles)
        self.export_button.pack(side='right', padx=2)

        self.add_rule_button = ttk.Button(button_frame, text="Add Rule", command=self.add_rule)
        self.add_rule_button.pack(side='left', padx=2)

        self.rules_listbox.bind('<Double-1>', lambda e: self.edit_rule())

        self.activate_button.config(state='disabled')
        self.export_button.config(state='disabled')
        self.add_rule_button.config(state='disabled')

        self.auto_save_var = tk.BooleanVar(value=self.auto_save_enabled)
        ttk.Checkbutton(self.profiles_frame, text="Auto-Save Changes", variable=self.auto_save_var,
                        command=self.save_ini).pack(anchor='w', padx=10, pady=(5, 0))

        ttk.Button(self.profiles_frame, text="Save Changes Now", command=self.save_config).pack(anchor='w', padx=10, pady=5)

    def _on_input_listbox_hover(self, event):
        if not hasattr(self, "input_devices"):
            self.input_tip.hidetip()
            return
        index = self.input_devices_listbox.nearest(event.y)
        if index < 0 or index >= len(self.input_devices):
            self.input_tip.hidetip()
            return
        device = self.input_devices[index]
        tooltip = device.get('name', '')
        self.input_tip.showtip(tooltip, event.x, event.y)

    def _on_output_listbox_hover(self, event):
        if not hasattr(self, "output_devices"):
            self.output_tip.hidetip()
            return
        index = self.output_devices_listbox.nearest(event.y)
        if index < 0 or index >= len(self.output_devices):
            self.output_tip.hidetip()
            return
        device = self.output_devices[index]
        tooltip = device.get('name', '')
        self.output_tip.showtip(tooltip, event.x, event.y)

    def create_devices_tab(self):
        self.refresh_button = ttk.Button(self.devices_frame, text="Refresh Device List", command=self.refresh_devices)
        self.refresh_button.pack(pady=10)

        lists_frame = ttk.Frame(self.devices_frame)
        lists_frame.pack(fill='both', expand=True, padx=10, pady=10)

        input_panel = ttk.Frame(lists_frame)
        input_panel.pack(side='left', fill='both', expand=True, padx=10)

        ttk.Label(input_panel, text="INPUT").pack(anchor='center')
        input_listbox_frame = ttk.Frame(input_panel)
        input_listbox_frame.pack(fill='both', expand=True)

        self.input_devices_listbox = tk.Listbox(input_listbox_frame, font=('Courier', 9), selectmode='extended')
        input_scroll_y = ttk.Scrollbar(input_listbox_frame, orient='vertical', command=self.input_devices_listbox.yview)
        input_scroll_x = ttk.Scrollbar(input_panel, orient='horizontal', command=self.input_devices_listbox.xview)
        self.input_devices_listbox.configure(yscrollcommand=input_scroll_y.set, xscrollcommand=input_scroll_x.set)
        self.input_devices_listbox.pack(side='left', fill='both', expand=True)
        input_scroll_y.pack(side='right', fill='y')
        input_scroll_x.pack(fill='x')

        output_panel = ttk.Frame(lists_frame)
        output_panel.pack(side='left', fill='both', expand=True, padx=10)

        ttk.Label(output_panel, text="OUTPUT").pack(anchor='center')
        output_listbox_frame = ttk.Frame(output_panel)
        output_listbox_frame.pack(fill='both', expand=True)

        self.output_devices_listbox = tk.Listbox(output_listbox_frame, font=('Courier', 9), selectmode='extended')
        output_scroll_y = ttk.Scrollbar(output_listbox_frame, orient='vertical', command=self.output_devices_listbox.yview)
        output_scroll_x = ttk.Scrollbar(output_panel, orient='horizontal', command=self.output_devices_listbox.xview)
        self.output_devices_listbox.configure(yscrollcommand=output_scroll_y.set, xscrollcommand=output_scroll_x.set)
        self.output_devices_listbox.pack(side='left', fill='both', expand=True)
        output_scroll_y.pack(side='right', fill='y')
        output_scroll_x.pack(fill='x')

        ttk.Button(self.devices_frame, text="Copy Selected Device Name", command=self.copy_device_id).pack(pady=5)

        self.input_tip = ToolTip(self.input_devices_listbox)
        self.output_tip = ToolTip(self.output_devices_listbox)

        self.input_devices_listbox.bind("<Motion>", self._on_input_listbox_hover)
        self.input_devices_listbox.bind("<Leave>", lambda e: self.input_tip.hidetip())
        self.output_devices_listbox.bind("<Motion>", self._on_output_listbox_hover)
        self.output_devices_listbox.bind("<Leave>", lambda e: self.output_tip.hidetip())

    def update_device_counts(self):
        self.input_devices_listbox.master.master.children['!label'].config(
            text=f"INPUT ({len(self.input_devices)} items)"
        )
        self.output_devices_listbox.master.master.children['!label'].config(
            text=f"OUTPUT ({len(self.output_devices)} items)"
        )

    def open_link(self, event=None):
        webbrowser.open_new(r"https://github.com/File-New-Project/EarTrumpet")

    def create_settings_tab(self):
        path_frame = ttk.LabelFrame(self.settings_frame, text="Configuration")
        path_frame.pack(fill='x', padx=10, pady=10)

        ttk.Label(path_frame, text="Path to EarTrumpet.exe (version with CLI command feature)").pack(anchor='w', padx=5, pady=5)

        path_entry_frame = ttk.Frame(path_frame)
        path_entry_frame.pack(fill='x', padx=5, pady=5)

        self.path_var = tk.StringVar(value=self.eartrumpet_path)
        path_entry = ttk.Entry(path_entry_frame, textvariable=self.path_var, state="readonly")
        path_entry.pack(side='left', fill='x', expand=True)

        ttk.Button(path_entry_frame, text="Browse", command=self.browse_eartrumpet).pack(side='right', padx=(5, 0))

        ttk.Button(path_frame, text="Test EarTrumpet", command=self.test_eartrumpet).pack(pady=5)
        ttk.Label(path_frame, text="EarTrumpet with CLI command line feature is needed for SmartWindowsAudioProfiles to work. \n The 'CLI command' feature is essentiel, I implemented it in EarTrumpet.\nCompile and install that version, if not already, from my github repo:").pack(anchor='w', padx=5, pady=(50, 5))
        link = ttk.Label(path_frame, text="https://github.com/dayeggpi/EarTrumpet", foreground='blue', cursor="hand2", underline=True)
        ttk.Label(path_frame, text="Additionally, you can find a pre-compiled version (decompress the zip, and that's all) from my SWAP repo").pack(anchor='w', padx=5, pady=(50, 5))
        link = ttk.Label(path_frame, text="https://github.com/dayeggpi/SWAP", foreground='blue', cursor="hand2", underline=True)
        link.pack(anchor='w', padx=5, pady=(0, 20))
        link.bind("<Button-1>", self.open_link)

    def new_profile(self):
        dialog = ProfileDialog(self.root, "New Profile")
        result = dialog.result
        if result:
            name = result['name']
            if name in self.profiles:
                messagebox.showerror("Error", "Profile name already exists!", parent=self.root)
                return
            self.profiles[name] = {'rules': []}
            self.update_profile_combo()
            self.profile_var.set(name)
            if "Select a profile..." in self.profile_combo['values']:
                values = list(self.profile_combo['values'])
                values.remove("Select a profile...")
                self.profile_combo['values'] = values
            self.on_profile_selected()
            if getattr(self, 'auto_save_var', True) and self.auto_save_var.get():
                self.save_config()
                self.changes_pending = False
            else:
                self.changes_pending = True
            self.mark_profiles_tab_unsaved()

    def delete_profile(self):
        current = self.profile_var.get()
        print(f"Before deletion: {list(self.profiles.keys())}")
        if current == "Select a profile...":
            messagebox.showwarning("Warning", "Please select a valid profile.", parent=self.root)
            return
        if not current:
            messagebox.showwarning("Warning", "No profile selected!", parent=self.root)
            return

        if messagebox.askyesno("Confirm", f"Delete profile '{current}'?"):
            del self.profiles[current]
            print(f"After deletion: {list(self.profiles.keys())}")
            self.update_profile_combo()
            self.profile_var.set("Select a profile...")
            self.on_profile_selected()

            if getattr(self, 'auto_save_var', True) and self.auto_save_var.get():
                self.save_config()
                self.changes_pending = False
            else:
                self.changes_pending = True
            self.mark_profiles_tab_unsaved()

    def activate_profile(self):
        # Do NOT reload config here; it can overwrite in-memory edits
        current = self.profile_var.get()
        if not current:
            messagebox.showwarning("Warning", "No profile selected!", parent=self.root)
            return

        # Optional: warn if there are unsaved changes
        if self.changes_pending and not (hasattr(self, 'auto_save_var') and self.auto_save_var.get()):
            if not messagebox.askyesno("Unsaved Changes", "You have unsaved changes. Apply with current in-memory rules?", parent=self.root):
                return

        applied = self.apply_profile(current)
        if getattr(self, 'auto_save_var', True) and self.auto_save_var.get():
            self.save_config()

        if applied == 0:
            messagebox.showwarning("Warning", "No rules were applied. Ensure the app has an active audio session, and the EarTrumpet path is correct.", parent=self.root)
        else:
            messagebox.showinfo("Success", f"Profile '{current}' activated with {applied} rule(s).", parent=self.root)
            
    
    def apply_profile(self, profile_name):
        if profile_name not in self.profiles:
            messagebox.showwarning("Warning", f"Profile Name not in profile", parent=self.root) #to delete
            return 0

        rules = self.profiles[profile_name]['rules']
        applied_count = 0
        for rule in rules:
            if self.execute_rule(rule):
                applied_count += 1
            else:
                messagebox.showwarning("Warning", f"{rule} : Rule not applied as app is not running.", parent=self.root) #to delete
        return applied_count

    def execute_rule(self, rule):
        try:
            if rule.get('direction') and rule['direction'] != 'Render':
                return False

            app_target = (rule.get('app_name') or '').strip()
            device_label = (rule.get('device') or rule.get('name') or rule.get('device_name') or '').strip()
            if not app_target or not device_label:
                print("Rule missing app or device. Skipping.")
                return False

            attempts = []
            # Try exactly as selected
            attempts.append([self.eartrumpet_path, '--set', app_target, device_label])

            # If device label ends with " (Default)", also try without it
            suffix = " (Default)"
            if device_label.endswith(suffix):
                attempts.append([self.eartrumpet_path, '--set', app_target, device_label[:-len(suffix)]])

            # If app looks like "Spotify.exe", also try "Spotify"
            if app_target.lower().endswith('.exe'):
                app_no_ext = app_target[:-4]
                attempts.append([self.eartrumpet_path, '--set', app_no_ext, device_label])
                if device_label.endswith(suffix):
                    attempts.append([self.eartrumpet_path, '--set', app_no_ext, device_label[:-len(suffix)]])

            last = None
            for cmd in attempts:
                print(f"Executing: {cmd}")
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    check=False,
                    creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                )
                last = result
                if result.returncode == 0:
                    return True

            if last:
                print(f"EarTrumpet --set failed. rc={last.returncode}\nstdout={last.stdout}\nstderr={last.stderr}")
            return False

        except Exception as e:
            print(f"Error executing rule: {e}")
            return False
            
            def add_unique(seq, value):
                if value and value not in seq:
                    seq.append(value)

            add_unique(app_candidates, app_target_saved)
            if saved_no_ext != app_target_saved:
                add_unique(app_candidates, saved_no_ext)

            for lbl in apps_now:
                lbl_lower = lbl.lower()
                if (lbl_lower == app_target_saved.lower() or
                    lbl_lower == saved_no_ext.lower() or
                    app_target_saved.lower() in lbl_lower or
                    saved_no_ext.lower() in lbl_lower):
                    add_unique(app_candidates, lbl)

            # If no apps reported, still attempt with the saved label(s)
            if not app_candidates:
                app_candidates = [app_target_saved]
                if saved_no_ext != app_target_saved:
                    app_candidates.append(saved_no_ext)

            # Optional: verify device label appears in current devices; if not, still try
            try:
                res_dev = subprocess.run(
                    [self.eartrumpet_path, '--list-devices'],
                    capture_output=True, text=True, check=False,
                    creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                )
                playback_labels = []
                if res_dev.returncode == 0:
                    for ln in res_dev.stdout.splitlines():
                        ln = ln.strip()
                        if ln.startswith('[Playback]'):
                            playback_labels.append(ln[len('[Playback]'):].strip())
                # If the saved device is not seen, we still try, but this helps diagnosis
                if playback_labels and not any(d.lower() == device_label_saved.lower() for d in playback_labels):
                    print(f"Note: device '{device_label_saved}' not currently in --list-devices playback list.")
            except Exception:
                pass

            # Try combinations: app x device, with a couple of retries for timing
            import time
            attempts_log = []
            for _retry in range(2):  # small retry to tolerate session creation timing
                for app_label in app_candidates:
                    for dev_label in device_candidates:
                        cmd = [self.eartrumpet_path, '--set', app_label, dev_label]
                        print(f"Executing: {cmd}")
                        result = subprocess.run(
                            cmd,
                            capture_output=True, text=True, check=False,
                            creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                        )
                        attempts_log.append((cmd, result.returncode, result.stdout, result.stderr))
                        if result.returncode == 0:
                            return True
                time.sleep(0.25)  # short delay then retry

            # Log failures for troubleshooting
            for (cmd, rc, out, err) in attempts_log:
                print(f"Attempt failed rc={rc}\n  cmd={cmd}\n  stdout={out}\n  stderr={err}")
            return False

        except Exception as e:
            print(f"Error executing rule: {e}")
            return False
            
           

    def add_rule(self):
        current_profile = self.profile_var.get()

        if not current_profile or current_profile not in self.profiles:
            messagebox.showwarning("Warning", "Please select a profile before adding a rule.", parent=self.root)
            return

        dialog = RuleDialog(self.root, "Add Rule", self.devices, eartrumpet_path=self.eartrumpet_path)
        result = dialog.result
        if result:
            # result is a single output rule
            self.profiles[current_profile]['rules'].append(result)
            self.update_rules_display()
            if getattr(self, 'auto_save_var', True) and self.auto_save_var.get():
                self.save_config()
                self.changes_pending = False
            else:
                self.changes_pending = True
            self.mark_profiles_tab_unsaved()

    def edit_rule(self):
        selection = self.rules_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "No rule selected!", parent=self.root)
            return

        current_profile = self.profile_var.get()
        display_idx = selection[0]
        rule_idx = self.displayed_rules_indices[display_idx]

        current_rule = self.profiles[current_profile]['rules'][rule_idx]
        dialog = RuleDialog(self.root, "Edit Rule", self.devices, rule_data=current_rule, eartrumpet_path=self.eartrumpet_path)

        result = dialog.result
        if result:
            self.profiles[current_profile]['rules'][rule_idx] = result
            self.update_rules_display()
        if getattr(self, 'auto_save_var', True) and self.auto_save_var.get():
            self.save_config()
            self.changes_pending = False
        else:
            self.changes_pending = True
        self.mark_profiles_tab_unsaved()

    def delete_rule(self):
        selection = self.rules_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "No rule selected!", parent=self.root)
            return

        current_profile = self.profile_var.get()
        display_idx = selection[0]
        rule_idx = self.displayed_rules_indices[display_idx]

        if messagebox.askyesno("Delete", "Delete this Output rule for the app?", parent=self.root):
            del self.profiles[current_profile]['rules'][rule_idx]
            self.update_rules_display()
            if getattr(self, 'auto_save_var', True) and self.auto_save_var.get():
                self.save_config()
                self.changes_pending = False
            else:
                self.changes_pending = True
            self.mark_profiles_tab_unsaved()

    def on_profile_selected(self, event=None):
        selected = self.profile_var.get()
        if selected == "Select a profile..." or selected not in self.profiles:
            self.rules_listbox.delete(0, tk.END)
            self.activate_button.config(state='disabled')
            self.export_button.config(state='disabled')
            self.add_rule_button.config(state='disabled')
            return

        self.activate_button.config(state='normal')
        self.export_button.config(state='normal')
        self.add_rule_button.config(state='normal')
        self.update_rules_display()

    def update_rules_display(self):
        self.rules_listbox.delete(0, tk.END)
        self.displayed_rules_indices = []

        current_profile = self.profile_var.get()
        if current_profile and current_profile in self.profiles:
            for idx, rule in enumerate(self.profiles[current_profile]['rules']):
                # Only display output rules
                if rule.get('direction') and rule['direction'] != 'Render':
                    continue
                app_name = rule.get('app_name', '')
                device_label = rule.get('device') or rule.get('name') or rule.get('device_name') or 'N/A'
                display_text = f"{app_name} -> {device_label}"
                self.rules_listbox.insert(tk.END, display_text)
                self.displayed_rules_indices.append(idx)

    def update_profile_combo(self):
        values = list(self.profiles.keys())
        if values:
            values.insert(0, "Select a profile...")
        self.profile_combo['values'] = values
        self.profile_var.set("Select a profile...")
        self.on_profile_selected()

    def refresh_devices(self):
        self.input_devices = []
        self.output_devices = []
        self.refresh_button.config(text="Loading...", state='disabled')
        threading.Thread(target=self._refresh_devices_thread, daemon=True).start()

    def _refresh_devices_thread(self):
        self.input_devices = []
        self.output_devices = []
        devices = []
        try:
            if not Checker.verify_eartrumpet_exe(self.eartrumpet_path):
                self.root.after(0, lambda: messagebox.showerror("Error", f"EarTrumpet not found or not working.\n\nPlease configure the correct path in Settings tab.", parent=self.root))
                self.root.after(0, lambda: self.refresh_button.config(text="Refresh Device List", state='normal'))
                return

            cmd = [self.eartrumpet_path, '--list-devices']
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                check=False,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            if result.returncode != 0:
                raise RuntimeError(result.stderr.strip() or "EarTrumpet returned an error while listing devices.")

            pattern_head = ('[Playback]', '[Recording]')
            for raw in result.stdout.splitlines():
                line = raw.strip()
                if not line or not (line.startswith('[Playback]') or line.startswith('[Recording]')):
                    continue
                if line.startswith('[Playback]'):
                    direction = 'Render'
                    label = line[len('[Playback]'):].strip()
                else:
                    direction = 'Capture'
                    label = line[len('[Recording]'):].strip()
                # Build normalized device structure
                devices.append({
                    'id': label,                # using label as identifier
                    'name': label,
                    'device_name': label,
                    'item_id': '',
                    'direction': direction,
                    'state': 'Active',
                    'type': 'Device'
                })

            self.root.after(0, self._update_devices_display, devices)

        except Exception as e:
            print(f"General error: {e}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Unexpected error: {e}", parent=self.root))
            self.root.after(0, lambda: self.refresh_button.config(text="Refresh Device List", state='normal'))

    def _update_devices_display(self, devices):
        devices = sorted(devices, key=lambda d: d['name'].lower())
        self.devices = devices

        self.input_devices = []
        self.output_devices = []

        self.input_devices_listbox.delete(0, tk.END)
        self.output_devices_listbox.delete(0, tk.END)

        for device in devices:
            display_text = device['name']
            if device['direction'] == 'Capture':
                self.input_devices.append(device)
                self.input_devices_listbox.insert(tk.END, display_text)
            elif device['direction'] == 'Render':
                self.output_devices.append(device)
                self.output_devices_listbox.insert(tk.END, display_text)

        self.refresh_button.config(text="Refresh Device List", state='normal')
        self.update_device_counts()

        if not self.devices:
            messagebox.showerror("Error", "Failed to refresh device list!\nPlease check that devices are connected or EarTrumpet is working.", parent=self.root)

    def copy_device_id(self):
        selection_input = self.input_devices_listbox.curselection()
        selection_output = self.output_devices_listbox.curselection()
        lines = []
        if selection_input:
            for idx in selection_input:
                device = self.input_devices[idx]
                text = f"{device['name']}"
                lines.append(text)
        if selection_output:
            for idx in selection_output:
                device = self.output_devices[idx]
                text = f"{device['name']}"
                lines.append(text)
        if not lines:
            messagebox.showwarning("Warning", "No device selected!", parent=self.root)
            return
        text_to_copy = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(text_to_copy)
        messagebox.showinfo("Success", f"Copied to clipboard", parent=self.root)

    def browse_eartrumpet(self):
        filename = filedialog.askopenfilename(
            title="Select EarTrumpet.exe",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
        )
        if filename:
            if not Checker.verify_eartrumpet_exe(filename):
                messagebox.showerror("Invalid File", "The selected EarTrumpet.exe could not be verified!\nPlease select the correct file.", parent=self.root)
                return
            self.path_var.set(filename)
            self.eartrumpet_path = filename
            self.refresh_devices()
            self.save_ini()
            messagebox.showinfo("Saved", "EarTrumpet path saved to config.ini", parent=self.root)

    def test_eartrumpet(self):
        try:
            if not Checker.verify_eartrumpet_exe(self.eartrumpet_path):
                messagebox.showerror("Error", "Configured EarTrumpet.exe is invalid or not working!\nPlease fix it in Settings.", parent=self.root)
            else:
                messagebox.showinfo("Success", "EarTrumpet is working correctly!", parent=self.root)
        except FileNotFoundError:
            messagebox.showerror("Error", "EarTrumpet.exe not found!", parent=self.root)
        except Exception as e:
            messagebox.showerror("Error", f"Error testing EarTrumpet: {e}", parent=self.root)

    def import_profiles(self):
        filename = filedialog.askopenfilename(
            title="Import Profiles",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'r') as f:
                    imported_data = json.load(f)

                if not isinstance(imported_data, dict) or 'profiles' not in imported_data or not isinstance(imported_data['profiles'], dict):
                    raise ValueError("This file does not appear to be a valid exported profile file!")

                invalid_profiles = []
                for name, profile in imported_data.get('profiles', {}).items():
                    if not PROFILE_NAME_REGEX.match(name):
                        invalid_profiles.append(name)
                        continue
                    if name in self.profiles:
                        if not messagebox.askyesno("Conflict", f"Profile '{name}' already exists. Overwrite?", parent=self.root):
                            continue
                    self.profiles[name] = profile

                if invalid_profiles:
                    messagebox.showwarning(
                        "Invalid Profile Names",
                        "These profile names were invalid and were NOT imported:\n" +
                        "\n".join(invalid_profiles),
                        parent=self.root
                    )

                self.update_profile_combo()
                if getattr(self, 'auto_save_var', True) and self.auto_save_var.get():
                    self.save_config()
                messagebox.showinfo("Success", "Profiles imported successfully!", parent=self.root)

            except Exception as e:
                messagebox.showerror("Error", f"Error importing profiles: {e}", parent=self.root)

    def export_profiles(self):
        export_window = tk.Toplevel(self.root)
        export_window.title("Export Profiles")
        export_window.geometry("300x400")
        export_window.iconbitmap(os.path.join(BASE_DIR, "icon.ico"))
        export_window.grab_set()
        ttk.Label(export_window, text="Select profiles to export:").pack(pady=5)

        listbox = tk.Listbox(export_window, selectmode=tk.MULTIPLE)
        toggle_var = tk.BooleanVar(value=False)

        def toggle_selection():
            if toggle_var.get():
                listbox.selection_clear(0, tk.END)
                toggle_button.config(text="Select All")
                toggle_var.set(False)
            else:
                listbox.select_set(0, tk.END)
                toggle_button.config(text="Unselect All")
                toggle_var.set(True)

        toggle_button = ttk.Button(export_window, text="Select All", command=toggle_selection)
        toggle_button.pack(pady=5)

        listbox.pack(fill='both', expand=True, padx=10, pady=5)

        for profile in self.profiles.keys():
            listbox.insert(tk.END, profile)

        def export_selected():
            selected_indices = listbox.curselection()
            selected_names = [listbox.get(i) for i in selected_indices]

            if not selected_names:
                messagebox.showwarning("Warning", "No profiles selected!", parent=self.root)
                return

            filename = filedialog.asksaveasfilename(
                title="Export Selected Profiles",
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )

            if filename:
                try:
                    export_data = {
                        'profiles': {k: self.profiles[k] for k in selected_names},
                        'eartrumpet_path': self.eartrumpet_path
                    }
                    with open(filename, 'w') as f:
                        json.dump(export_data, f, indent=2)
                    export_window.destroy()
                    messagebox.showinfo("Success", "Profiles exported successfully!", parent=self.root)
                except Exception as e:
                    messagebox.showerror("Error", f"Error exporting profiles: {e}", parent=self.root)

        ttk.Button(export_window, text="Export Selected", command=export_selected).pack(pady=10)

    def load_config(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    if not isinstance(config, dict) or 'profiles' not in config or not isinstance(config['profiles'], dict):
                        raise ValueError("This file does not appear to be a valid exported profile file!")
                self.profiles = config.get('profiles', {})
                self.eartrumpet_path = config.get('eartrumpet_path', self.eartrumpet_path)
        except Exception as e:
            print(f"Error loading config: {e}")

    def save_config(self):
        try:
            config = {'profiles': self.profiles}
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=2)
            self.changes_pending = False
            self.mark_profiles_tab_unsaved()

        except Exception as e:
            print(f"Error saving config: {e}")

    def show_window(self, icon=None, item=None):
        self.root.deiconify()
        self.root.lift()

    def on_closing(self):
        if self.changes_pending and not self.auto_save_var.get():
            if messagebox.askyesno("Unsaved Changes", "Some changes were not saved. Save now?", parent=self.root):
                self.save_config()
        self.quit_app()

    def quit_app(self, icon=None, item=None):
        self.root.quit()

    def run(self):
        self.update_profile_combo()
        self.update_rules_display()
        self.refresh_devices()
        self.root.mainloop()


class ProfileDialog:
    def __init__(self, parent, title, profile_data=None):
        self.result = None

        self.dialog = tk.Toplevel(parent)
        self.dialog.iconbitmap(os.path.join(BASE_DIR, "icon.ico"))
        self.dialog.title(title)
        self.dialog.geometry("300x150")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()

        ttk.Label(self.dialog, text="Profile Name:").pack(pady=5)
        self.name_var = tk.StringVar(value=profile_data['name'] if profile_data else '')
        name_entry = ttk.Entry(self.dialog, textvariable=self.name_var, width=30)
        name_entry.pack(pady=5)
        name_entry.focus()

        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(pady=20)

        ttk.Button(button_frame, text="OK", command=self.ok_clicked).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.cancel_clicked).pack(side='left', padx=5)

        self.dialog.bind('<Return>', lambda e: self.ok_clicked())
        ProfileDialog.center_window(self.dialog, parent)
        self.dialog.wait_window()

    def ok_clicked(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror("Error", "Profile name cannot be empty!", parent=self.dialog)
            return
        if not PROFILE_NAME_REGEX.match(name):
            messagebox.showerror("Error", "Invalid profile name!\nOnly letters, numbers, and hyphens (-) are allowed.\nNo spaces or special characters.", parent=self.dialog)
            return
        self.result = {'name': name}
        self.dialog.destroy()

    def center_window(dialog, parent):
        dialog.update_idletasks()
        w = dialog.winfo_width()
        h = dialog.winfo_height()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (w // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (h // 2)
        dialog.geometry(f"+{x}+{y}")

    def cancel_clicked(self):
        self.dialog.destroy()


class RuleDialog:
    def __init__(self, parent, title, devices, rule_data=None, eartrumpet_path="EarTrumpet.exe"):
        self.result = None
        self.devices = sorted(devices, key=lambda d: d['name'].lower())
        self.eartrumpet_path = eartrumpet_path

        self.dialog = tk.Toplevel(parent)
        self.dialog.iconbitmap(os.path.join(BASE_DIR, "icon.ico"))
        self.dialog.title(title)
        self.dialog.geometry("600x600")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()

        ttk.Label(self.dialog, text="Application Name (e.g., chrome.exe):").pack(pady=(5, 0))
        ttk.Label(self.dialog, text="(Note: only apps with active audio sessions are listed if available; you can still type an .exe)").pack(pady=5)
        self.app_var = tk.StringVar()
        self.app_combobox = AutocompleteCombobox(self.dialog, textvariable=self.app_var, width=50)
        self.app_combobox.pack(pady=(0, 5))

        # Status + Refresh button
        apps_topbar = ttk.Frame(self.dialog)
        apps_topbar.pack(fill='x', padx=10, pady=(0, 10))
        self.apps_status_var = tk.StringVar(value="")
        ttk.Button(apps_topbar, text="Refresh apps", command=self._refresh_apps).pack(side='left')
        ttk.Label(apps_topbar, textvariable=self.apps_status_var).pack(side='left', padx=8)

        # Initial load
        self._refresh_apps()
        
        ttk.Label(self.dialog, text="Output Device (Render):").pack()
        output_frame = ttk.Frame(self.dialog)
        output_frame.pack(fill='x', padx=10, pady=2)
        self.output_listbox = tk.Listbox(output_frame, height=15, exportselection=False, selectmode="browse")
        output_scroll_y = ttk.Scrollbar(output_frame, orient='vertical', command=self.output_listbox.yview)
        self.output_listbox.configure(yscrollcommand=output_scroll_y.set)
        self.output_listbox.pack(side='left', fill='both', expand=True)
        output_scroll_y.pack(side='right', fill='y')
        output_scroll_x = ttk.Scrollbar(self.dialog, orient='horizontal', command=self.output_listbox.xview)
        self.output_listbox.configure(xscrollcommand=output_scroll_x.set)
        output_scroll_x.pack(fill='x', padx=10, pady=(0, 10))

        self.render_devices = []

        for device in self.devices:
            if device.get("direction") == "Render":
                display_str = device['name']
                self.render_devices.append(device)
                self.output_listbox.insert(tk.END, display_str)

        # Preselect if editing
        if rule_data:
            device_label = (rule_data.get('device') or rule_data.get('name') or rule_data.get('device_name') or '').strip()
            target_labels = [device_label]
            # Also try without ' (Default)' to tolerate list changes
            if device_label.endswith(" (Default)"):
                target_labels.append(device_label[:-len(" (Default)")])

            # Always clear any accidental selection BEFORE searching
            self.output_listbox.selection_clear(0, tk.END)

            # Find the best match (case-insensitive) among devices
            found_index = None
            device_names_lower = [d['name'].strip().lower() for d in self.render_devices]
            for t in target_labels:
                t_lower = t.strip().lower()
                if not t_lower:
                    continue
                try:
                    found_index = device_names_lower.index(t_lower)
                    break
                except ValueError:
                    pass

            if found_index is not None:
                self.output_listbox.selection_set(found_index)
                self.output_listbox.see(found_index)

            # Always prefill the app name if we have it
            if rule_data.get('app_name'):
                self.app_var.set(rule_data['app_name'])
                
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="OK", command=self.ok_clicked).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.cancel_clicked).pack(side='left', padx=5)
        self.app_combobox.focus()
        RuleDialog.center_window(self.dialog, parent)
        self.dialog.wait_window()

         
    def _refresh_apps(self):
        exe_list = self._get_audio_apps()
        self.app_combobox.set_completion_list(exe_list)
        self.apps_status_var.set(f"Found {len(exe_list)} app(s) with audio")

    def _get_audio_apps(self):
        apps = []
        try:
            cmd = [self.eartrumpet_path, '--list-apps']
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                check=False,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            if result.returncode != 0:
                return apps
            for line in result.stdout.splitlines():
                name = line.strip()
                if name:
                    apps.append(name)
            # Deduplicate and sort
            apps = sorted(set(apps), key=str.lower)
        except Exception:
            pass
        return apps
        
        

    def center_window(dialog, parent):
        dialog.update_idletasks()
        w = dialog.winfo_width()
        h = dialog.winfo_height()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (w // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (h // 2)
        dialog.geometry(f"+{x}+{y}")


    def ok_clicked(self):
        app_name = self.app_var.get().strip()
        valid_apps = getattr(self.app_combobox, '_completion_list', [])
        if not app_name:
            messagebox.showerror("Error", "Application name cannot be empty!", parent=self.dialog)
            return

  
        output_sel = self.output_listbox.curselection()
        if not output_sel:
            messagebox.showerror("Error", "You must select an output (Render) device.", parent=self.dialog)
            return

        output_device = self.render_devices[output_sel[0]]
        output_rule = {
            'app_name': app_name,
            'device': output_device['name'],  # the label we pass to EarTrumpet --set
            'name': output_device['name'],
            'direction': 'Render'
        }

        self.result = output_rule
        self.dialog.destroy()
        

    def cancel_clicked(self):
        self.dialog.destroy()


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="SmartWindowsAudioProfiles CLI")
    parser.add_argument("profile_name", nargs="?", help="Profile name to activate (uses audio_profiles.json in app directory)")
    args = parser.parse_args()

    if args.profile_name is not None and not PROFILE_NAME_REGEX.match(args.profile_name):
        print("ERROR: Invalid profile name! Only letters, numbers, and hyphens (-) are allowed. No spaces.")
        sys.exit(1)

    if args.profile_name:
        app = AudioProfileManager()
        try:
            if not os.path.exists(PROFILE_FILE):
                print("ERROR: audio_profiles.json not found in the application directory.")
                sys.exit(1)
            app.load_config()

            if args.profile_name in app.profiles:
                applied = app.apply_profile(args.profile_name)
                if applied > 0:
                    print(f"Profile '{args.profile_name}' activated with {applied} rule(s).")
                    sys.exit(0)
                else:
                    print(f"ERROR: Profile '{args.profile_name}' found but no rules could be applied. Are the target applications running? Is EarTrumpet configured correctly?")
                    sys.exit(2)
            else:
                print(f"ERROR: Profile '{args.profile_name}' not found in audio_profiles.json.")
                sys.exit(1)
        except Exception as e:
            print(f"ERROR: Failed to activate profile '{args.profile_name}': {e}")
            sys.exit(2)
    else:
        app = AudioProfileManager()
        app.run()