import os
import tkinter as tk
from tkinter import ttk, filedialog
import codecs
import struct
import winreg

#
# --- PART 1: Gather Start Menu apps ---
#

def get_start_menu_shortcuts():
    """
    Return a set of "app names" corresponding to .lnk shortcuts in the Start Menu.
    Looks in both system-wide and per-user Start Menu folders.
    
    For simplicity, we:
      - Collect the .lnk filenames (without the .lnk extension).
      - We do NOT parse the .lnk to see its real target.
    """
    start_menu_paths = [
        r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs",
        os.path.join(os.environ['APPDATA'], r"Microsoft\Windows\Start Menu\Programs"),
    ]
    
    app_names = set()
    
    for base_path in start_menu_paths:
        if not os.path.exists(base_path):
            continue
        
        for root, dirs, files in os.walk(base_path):
            for filename in files:
                if filename.lower().endswith(".lnk"):
                    name_no_ext = os.path.splitext(filename)[0]
                    app_names.add(name_no_ext)
    
    return app_names

#
# --- PART 2: Collect usage data from UserAssist ---
#

USERASSIST_GUIDS = [
    '{CEBFF5CD-ACE2-4F4F-9178-9926F41749EA}',
    '{F2A1CB5A-E3CC-4A2E-AF9D-505A7009D442}',
]

def prettify_name(raw_name: str) -> str:
    """
    Simplified function to convert a ROT13-decoded UserAssist name into something friendlier.
    - If it looks like a path, just take the base filename.
    - If it has a '!', typical for UWP apps, take what follows the last '!'.
    - Strip '.exe' if present.
    """
    name = raw_name.strip()
    
    if '\\' in name:
        name = os.path.basename(name)
    if '!' in name:
        parts = name.split('!')
        if parts[-1].strip():
            name = parts[-1].strip()
    if name.lower().endswith('.exe'):
        name = name[:-4]
    
    return name

def get_userassist_usage():
    """
    Reads the UserAssist registry keys and returns a dict:
        { "friendly_app_name": usage_count, ... }
    """
    usage_dict = {}
    
    for guid in USERASSIST_GUIDS:
        key_path = fr"Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist\{guid}\Count"
        
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                i = 0
                while True:
                    try:
                        value_name, value_data, value_type = winreg.EnumValue(key, i)
                        i += 1
                        
                        decoded_name = codecs.decode(value_name, 'rot_13')
                        if len(value_data) >= 8:
                            usage_count = struct.unpack_from("<I", value_data, 4)[0]
                        else:
                            usage_count = 0
                        
                        app_name = prettify_name(decoded_name)
                        
                        # Combine counts if the same app appears in multiple GUIDs
                        usage_dict[app_name] = usage_dict.get(app_name, 0) + usage_count
                    except OSError:
                        break
        except FileNotFoundError:
            pass
    
    return usage_dict

#
# --- PART 3: Combine data & Show in a GUI (Save TXT: only Start Menu items, sorted by name) ---
#

def create_ui(final_list):
    """
    Create a Tkinter TreeView to display the combined results.
    final_list: list of tuples (app_name, usage_count, is_start_menu).
    """

    def on_save_as_txt():
        """Callback: saves only the Start Menu software names to a text file, sorted by name."""
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            # Collect only Start Menu items
            start_menu_only = [app_name for (app_name, _, is_start_menu) in final_list if is_start_menu]
            # Sort alphabetically (case-insensitive)
            start_menu_only.sort(key=str.casefold)
            
            with open(filename, "w", encoding="utf-8") as f:
                for app_name in start_menu_only:
                    f.write(app_name + "\n")
    
    root = tk.Tk()
    root.title("Software List: Start Menu First + Usage Counts")
    root.geometry("700x500")
    
    # -- Top Frame with "Save as TXT" Button
    top_frame = ttk.Frame(root)
    top_frame.pack(fill='x', padx=5, pady=5)
    
    save_button = ttk.Button(top_frame, text="Save as TXT", command=on_save_as_txt)
    save_button.pack(side=tk.RIGHT)
    
    # -- Main Frame with TreeView
    main_frame = ttk.Frame(root)
    main_frame.pack(fill='both', expand=True, padx=5, pady=5)
    
    columns = ('app_name', 'count', 'source')
    tree = ttk.Treeview(main_frame, columns=columns, show='headings', height=20)
    tree.heading('app_name', text='App Name')
    tree.heading('count', text='Usage Count')
    tree.heading('source', text='From Start Menu?')
    
    tree.column('app_name', width=350)
    tree.column('count', width=80, anchor='center')
    tree.column('source', width=120, anchor='center')
    
    # Insert data into the TreeView
    for (app_name, usage_count, is_start_menu) in final_list:
        source_str = "Yes" if is_start_menu else "No"
        tree.insert('', tk.END, values=(app_name, usage_count, source_str))
    
    # Scrollbar
    scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    tree.pack(side=tk.LEFT, fill='both', expand=True)
    
    root.mainloop()

def main():
    # 1) Get the set of Start Menu app names
    start_menu_apps = get_start_menu_shortcuts()
    
    # 2) Get usage data from UserAssist
    usage_dict = get_userassist_usage()
    
    # 3) Build a combined list: (app_name, usage_count, is_start_menu)
    final_list = []
    for sm_app in start_menu_apps:
        usage_count = usage_dict.get(sm_app, 0)
        final_list.append((sm_app, usage_count, True))
    
    for ua_app, ua_count in usage_dict.items():
        if ua_app not in start_menu_apps:
            final_list.append((ua_app, ua_count, False))
    
    # 4) Sort so that Start Menu apps appear first, then by descending usage
    final_list.sort(key=lambda x: (not x[2], x[1]), reverse=True)
    
    # 5) Display in Tkinter
    create_ui(final_list)

if __name__ == "__main__":
    main()
