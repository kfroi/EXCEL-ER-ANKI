import dearpygui.dearpygui as dpg
import pandas as pd
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
import subprocess
import sys
import shutil
import zipfile
import tempfile
import threading

def select_excel_callback(sender, app_data, user_data):
    dpg.set_value("excel_path", app_data[0])
    try:
        file_path = app_data[0]
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.csv':
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
        columns = list(df.columns)
        # Remove old checkboxes if any
        front_children = dpg.get_item_children("front_checkbox_group", 1)
        if front_children:
            for tag in front_children:
                dpg.delete_item(tag)
        back_children = dpg.get_item_children("back_checkbox_group", 1)
        if back_children:
            for tag in back_children:
                dpg.delete_item(tag)
        # Add new checkboxes for each column
        for col in columns:
            dpg.add_checkbox(label=col, tag=f"front_{col}", parent="front_checkbox_group")
            dpg.add_checkbox(label=col, tag=f"back_{col}", parent="back_checkbox_group")
        # Force refresh of child windows
        dpg.configure_item("front_checkbox_group", show=False)
        dpg.configure_item("front_checkbox_group", show=True)
        dpg.configure_item("back_checkbox_group", show=False)
        dpg.configure_item("back_checkbox_group", show=True)
        dpg.set_value("status", "Excel/CSV file loaded.")
        dpg.set_item_user_data("generate_btn", df)
    except Exception as e:
        dpg.set_value("status", f"Error loading file: {e}")

def prepare_media_from_excel(excel_path):
    # 1. Delete everything in media/media
    media_dir = os.path.join("media", "media")
    os.makedirs(media_dir, exist_ok=True)
    if os.path.exists(media_dir):
        for f in os.listdir(media_dir):
            file_path = os.path.join(media_dir, f)
            if os.path.isfile(file_path):
                os.remove(file_path)
    # 2. Duplicate the Excel spreadsheet temporarily
    temp_dir = tempfile.mkdtemp()
    temp_excel = os.path.join(temp_dir, os.path.basename(excel_path))
    shutil.copy2(excel_path, temp_excel)
    # 3. Convert it to a zip file (Excel xlsx is a zip format)
    temp_zip = temp_excel + ".zip"
    shutil.copy2(temp_excel, temp_zip)
    # 4. Extract the zip file
    extract_dir = os.path.join(temp_dir, "extracted")
    os.makedirs(extract_dir, exist_ok=True)
    with zipfile.ZipFile(temp_zip, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)
    # 5. Move all images to media/media, renaming with output filename prefix
    # Get output filename prefix (without extension)
    from dearpygui.dearpygui import get_value
    output_file = get_value("output_file")
    prefix = os.path.splitext(output_file)[0] + "_"
    for root, dirs, files in os.walk(extract_dir):
        for file in files:
            if file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                src = os.path.join(root, file)
                new_name = prefix + file
                dst = os.path.join(media_dir, new_name)
                shutil.copy2(src, dst)
    # 6. Delete the temporary folder
    shutil.rmtree(temp_dir)


# Add a global variable to store the last error message and a log list
last_error_message = ""
log_messages = []

def append_log(msg):
    global log_messages
    log_messages.append(msg)
    # Keep only last 100 messages for performance
    if len(log_messages) > 100:
        log_messages = log_messages[-100:]
    dpg.set_value("log_window", "\n".join(log_messages))
    # Auto-scroll to bottom
    try:
        dpg.set_y_scroll("log_window", dpg.get_y_scroll_max("log_window"))
    except Exception:
        pass

def set_error_message(msg):
    global last_error_message
    last_error_message = msg
    append_log(f"ERROR: {msg}")

def copy_error_callback():
    import pyperclip
    if last_error_message:
        pyperclip.copy(last_error_message)
        # Set cooldown style: darker and no hover highlight
        dpg.configure_item("copy_error_btn", label="Copied!", enabled=False)
        dpg.bind_item_theme("copy_error_btn", "cooldown_theme")
        def reset_btn():
            import time
            time.sleep(3)
            dpg.configure_item("copy_error_btn", label="Copy Error Message", enabled=True)
            dpg.bind_item_theme("copy_error_btn", "default_theme")
        threading.Thread(target=reset_btn, daemon=True).start()

def set_generate_btn_loading():
    dpg.configure_item("generate_btn", enabled=False)
    dpg.bind_item_theme("generate_btn", "loading_theme")
    global _generate_btn_thread, _generate_btn_stop_event
    import time
    frames = ["Generating", "Generating.", "Generating..", "Generating..."]
    _generate_btn_stop_event = threading.Event()
    def animate():
        i = 0
        while not _generate_btn_stop_event.is_set():
            dpg.configure_item("generate_btn", label=frames[i % 4])
            time.sleep(0.4)
            i += 1
    _generate_btn_thread = threading.Thread(target=animate, daemon=True)
    _generate_btn_thread.start()

def stop_generate_btn_loading():
    global _generate_btn_thread, _generate_btn_stop_event
    if '_generate_btn_stop_event' in globals() and _generate_btn_stop_event:
        _generate_btn_stop_event.set()
    if '_generate_btn_thread' in globals() and _generate_btn_thread:
        _generate_btn_thread = None

def set_generate_btn_done():
    dpg.configure_item("generate_btn", label="APKG Generation Complete!", enabled=False)
    dpg.bind_item_theme("generate_btn", "done_theme")
    append_log("APKG generation complete!")
    # Run move_to_anki_media.py after successful generation
    try:
        append_log("Moving images to Anki collection.media directory...")
        result = subprocess.run([sys.executable, "move_to_anki_media.py"], capture_output=True, text=True)
        if result.returncode == 0:
            append_log("Images moved successfully.")
            if result.stdout:
                append_log(result.stdout.strip())
        else:
            append_log(f"Error moving images: {result.stderr.strip()}")
    except Exception as e:
        append_log(f"Exception running move_to_anki_media.py: {e}")
    def reset_btn():
        import time
        time.sleep(3)
        dpg.configure_item("generate_btn", label="Generate Deck", enabled=True)
        dpg.bind_item_theme("generate_btn", "default_theme")
    threading.Thread(target=reset_btn, daemon=True).start()

def generate_deck_callback(sender, app_data, user_data):
    set_generate_btn_loading()
    append_log("Started APKG generation process.")
    df = dpg.get_item_user_data("generate_btn")
    if df is None:
        stop_generate_btn_loading()
        set_error_message("Please select an Excel file.")
        dpg.configure_item("generate_btn", label="Generate Deck", enabled=True)
        dpg.bind_item_theme("generate_btn", "default_theme")
        return
    columns = list(df.columns)
    front_cols = [col for col in columns if dpg.get_value(f"front_{col}")]
    back_cols = [col for col in columns if dpg.get_value(f"back_{col}")]
    append_log(f"Selected front columns: {front_cols}")
    append_log(f"Selected back columns: {back_cols}")
    if not front_cols or not back_cols:
        stop_generate_btn_loading()
        set_error_message("Please select columns for both front and back.")
        dpg.configure_item("generate_btn", label="Generate Deck", enabled=True)
        dpg.bind_item_theme("generate_btn", "default_theme")
        return
    outdir = dpg.get_value("output_dir")
    outfile = dpg.get_value("output_file")
    outpath = os.path.join(outdir, outfile)
    excel_path = dpg.get_value("excel_path")
    append_log(f"Output directory: {outdir}")
    append_log(f"Output file: {outfile}")
    append_log(f"Excel path: {excel_path}")
    # Prepare media files from Excel
    try:
        append_log("Preparing media from Excel...")
        prepare_media_from_excel(excel_path)
        append_log("Media prepared.")
    except Exception as e:
        stop_generate_btn_loading()
        set_error_message(f"Error preparing media: {e}")
        dpg.configure_item("generate_btn", label="Generate Deck", enabled=True)
        dpg.bind_item_theme("generate_btn", "default_theme")
        return
    # Save the selected DataFrame to a temp CSV
    temp_csv = os.path.join(outdir, "_anki_temp_input.csv")
    append_log(f"Saving temp CSV: {temp_csv}")
    df.to_csv(temp_csv, index=False)
    # Call make_apkg.py with the temp CSV and output path as arguments
    try:
        append_log("Calling make_apkg.py...")
        result = subprocess.run([
            sys.executable, "make_apkg.py", temp_csv, outpath
        ], capture_output=True, text=True)
        stop_generate_btn_loading()
        if result.returncode == 0:
            append_log("make_apkg.py completed successfully.")
            set_generate_btn_done()
        else:
            set_error_message(f"Error: {result.stderr}")
            dpg.configure_item("generate_btn", label="Generate Deck", enabled=True)
            dpg.bind_item_theme("generate_btn", "default_theme")
    except Exception as e:
        stop_generate_btn_loading()
        set_error_message(f"Error running make_apkg.py: {e}")
        dpg.configure_item("generate_btn", label="Generate Deck", enabled=True)
        dpg.bind_item_theme("generate_btn", "default_theme")

def browse_file_callback():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Excel or CSV file",
        filetypes=[
            ("Excel/CSV files", "*.xlsx *.xls *.xlsm *.xlsb *.ods *.csv"),
            ("All files", "*.*")
        ]
    )
    if file_path:
        dpg.set_value("excel_path", file_path)
        # Scan and update columns immediately after file selection
        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext == '.csv':
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            columns = list(df.columns)
            # Remove old checkboxes if any
            front_children = dpg.get_item_children("front_checkbox_group", 1)
            if front_children:
                for tag in front_children:
                    dpg.delete_item(tag)
            back_children = dpg.get_item_children("back_checkbox_group", 1)
            if back_children:
                for tag in back_children:
                    dpg.delete_item(tag)
            # Add new checkboxes for each column
            for col in columns:
                dpg.add_checkbox(label=col, tag=f"front_{col}", parent="front_checkbox_group")
                dpg.add_checkbox(label=col, tag=f"back_{col}", parent="back_checkbox_group")
            # Force refresh of child windows
            dpg.configure_item("front_checkbox_group", show=False)
            dpg.configure_item("front_checkbox_group", show=True)
            dpg.configure_item("back_checkbox_group", show=False)
            dpg.configure_item("back_checkbox_group", show=True)
            append_log("Excel/CSV file loaded.")
            dpg.set_item_user_data("generate_btn", df)
        except Exception as e:
            set_error_message(f"Error loading file: {e}")

def main():
    dpg.create_context()
    # Define themes for the button
    with dpg.theme(tag="cooldown_theme"):
        with dpg.theme_component(dpg.mvButton):
            dpg.add_theme_color(dpg.mvThemeCol_Button, (60, 60, 60, 255), category=dpg.mvThemeCat_Core)  # darker
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, (60, 60, 60, 255), category=dpg.mvThemeCat_Core)  # no highlight
            dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, (60, 60, 60, 255), category=dpg.mvThemeCat_Core)
    with dpg.theme(tag="default_theme"):
        with dpg.theme_component(dpg.mvButton):
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, (120, 120, 120, 255), category=dpg.mvThemeCat_Core)  # gray hover
    with dpg.theme(tag="loading_theme"):
        with dpg.theme_component(dpg.mvButton):
            dpg.add_theme_color(dpg.mvThemeCol_Button, (100, 100, 100, 255), category=dpg.mvThemeCat_Core)  # gray
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, (100, 100, 100, 255), category=dpg.mvThemeCat_Core)
            dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, (100, 100, 100, 255), category=dpg.mvThemeCat_Core)
    with dpg.theme(tag="done_theme"):
        with dpg.theme_component(dpg.mvButton):
            dpg.add_theme_color(dpg.mvThemeCol_Button, (60, 60, 60, 255), category=dpg.mvThemeCat_Core)  # darker
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, (60, 60, 60, 255), category=dpg.mvThemeCat_Core)
            dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, (60, 60, 60, 255), category=dpg.mvThemeCat_Core)
    with dpg.theme(tag="browse_theme"):
        with dpg.theme_component(dpg.mvButton):
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, (120, 120, 120, 255), category=dpg.mvThemeCat_Core)
    with dpg.window(label="Anki Deck Generator", width=800, height=600, no_resize=True, tag="primary_window", no_scrollbar=True, no_scroll_with_mouse=True):
        with dpg.group():
            dpg.add_text("Select Excel file:")
            with dpg.group(horizontal=True):
                dpg.add_input_text(tag="excel_path", readonly=True)
                dpg.add_spacer(width=1)
                browse_btn = dpg.add_button(label="Browse", tag="browse_btn", callback=lambda: browse_file_callback())
                dpg.bind_item_theme(browse_btn, "browse_theme")
            dpg.add_spacer(height=5)
            dpg.add_text("Tip: If you don't see your file, type *.* in the file name box and press Enter to show all files.")
            dpg.add_spacer(height=5)
            dpg.add_separator()
            dpg.add_spacer(height=5)
        with dpg.group(horizontal=True):
            dpg.add_text("Select columns for Front of card:")
            dpg.add_spacer(width=90)
            dpg.add_text("Select columns for Back of card:")
        with dpg.group(horizontal=True):
            with dpg.child_window(tag="front_checkbox_group", height=120, width=300):
                pass
            dpg.add_spacer(width=20)
            with dpg.child_window(tag="back_checkbox_group", height=120, width=300):
                pass
        dpg.add_spacer(height=7)
        with dpg.group():
            dpg.add_text("Select output location:")
            dpg.add_input_text(tag="output_dir", default_value=str(Path.home() / "Downloads"))
            dpg.add_spacer(height=5)
            dpg.add_text("Output filename:")
            dpg.add_input_text(tag="output_file", default_value="anki_deck.apkg")
            dpg.add_spacer(height=2)
            dpg.add_button(label="Generate Deck", tag="generate_btn", callback=generate_deck_callback, enabled=True)
            dpg.bind_item_theme("generate_btn", "default_theme")
            dpg.add_spacer(height=7)
            dpg.add_button(label="Copy Error Message", tag="copy_error_btn", callback=lambda: copy_error_callback(), enabled=True)
        dpg.add_spacer(height=10)
        dpg.add_text("Log:")
        dpg.add_input_text(tag="log_window", multiline=True, readonly=True, height=220, width=780, default_value="", tab_input=False)
    import ctypes
    user32 = ctypes.windll.user32
    screen_width = user32.GetSystemMetrics(0)
    screen_height = user32.GetSystemMetrics(1)
    dpg.create_viewport(title='Anki Deck Generator', width=800, height=600, resizable=True, max_width=screen_width, max_height=screen_height)
    dpg.setup_dearpygui()
    dpg.maximize_viewport()
    dpg.set_primary_window("primary_window", True)
    dpg.show_viewport()
    dpg.start_dearpygui()
    dpg.destroy_context()

if __name__ == '__main__':
    main()
