import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import random
import string
import zlib
import requests
import shutil
import tempfile
import cairosvg
import sys
from PIL import Image, ImageGrab
import win32com.client
from tkinterdnd2 import TkinterDnD, DND_FILES

# Determine if the script is running as a bundled executable or from the source directory
if getattr(sys, 'frozen', False):
    # Running as a bundled executable
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Running as a regular script
    BASE_DIR = os.path.dirname(__file__)

# Paths for the source and backup directories
SOURCE_DIR_FILE = os.path.join(BASE_DIR, "_sourcedir.txt")
BACKUP_DIR_FILE = os.path.join(BASE_DIR, "_backupdir.txt")

def convert_svg_to_png(svg_path, output_path):
    try:
        # Convert the SVG to PNG using cairosvg
        cairosvg.svg2png(url=svg_path, write_to=output_path)
    except Exception as e:
        raise ValueError(f"Failed to convert SVG to PNG: {e}")

def read_directory_from_file(file_path):
    try:
        with open(file_path, 'r') as f:
            directory = f.read().strip()
        return directory
    except Exception as e:
        raise FileNotFoundError(f"Failed to read {file_path}: {e}")

def ensure_valid_directory(directory_path):
    if not directory_path:
        raise ValueError("Directory path is empty!")

    if not os.path.isabs(directory_path):
        raise ValueError(f"Invalid directory path: '{directory_path}'. It must be an absolute path!")

    if not os.path.exists(directory_path):
        try:
            os.makedirs(directory_path)
        except Exception as e:
            raise OSError(f"Failed to create directory '{directory_path}': {e}")

def validate_sourcedir():
    # Check if the file exists
    if not os.path.exists(SOURCE_DIR_FILE):
        messagebox.showerror("Error", "_sourcedir.txt file is missing!")
        return False

    try:
        # Read the contents of the file
        with open(SOURCE_DIR_FILE, 'r') as f:
            source_dir = f.read().strip()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read _sourcedir.txt: {e}")
        return False

    # Check if the path is empty or invalid
    if not source_dir:
        messagebox.showerror("Error", "_sourcedir.txt is blank!")
        return False

    # Ensure the path is absolute (avoiding ambiguous or invalid entries like "abc")
    if not os.path.isabs(source_dir):
        messagebox.showerror("Error", f"Invalid directory path in _sourcedir.txt: '{source_dir}' must be an absolute path!")
        return False

    # Attempt to create the directory if it doesn't exist
    if not os.path.exists(source_dir):
        try:
            os.makedirs(source_dir)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create source directory '{source_dir}': {e}")
            return False

    return True

def validate_backupdir():
    try:
        # Read the contents of the file
        with open(BACKUP_DIR_FILE, 'r') as f:
            backup_dir = f.read().strip()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read _backupdir.txt: {e}")
        return None

    # If the backup directory is empty, contains only spaces or tabs, silently skip
    if not backup_dir:
        return None

    # Ensure the path is absolute
    if not os.path.isabs(backup_dir):
        messagebox.showerror("Error", f"Invalid directory path in _backupdir.txt: '{backup_dir}' must be an absolute path!")
        return None

    # Attempt to create the directory if it doesn't exist
    if not os.path.exists(backup_dir):
        try:
            os.makedirs(backup_dir)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create backup directory '{backup_dir}': {e}")
            return None

    return backup_dir

def backup_ico_files():
    # Validate the backup directory
    backup_dir = validate_backupdir()

    # If the backup directory is not set, skip the backup process
    if not backup_dir:
        return

    try:
        source_dir = read_directory_from_file(SOURCE_DIR_FILE)

        if not source_dir or not os.path.exists(source_dir):
            messagebox.showerror("Error", "Invalid source directory!")
            return

        # Copy all .ico files from the source directory to the backup directory
        for file_name in os.listdir(source_dir):
            if file_name.lower().endswith(".ico"):
                source_file_path = os.path.join(source_dir, file_name)
                backup_file_path = os.path.join(backup_dir, file_name)
                shutil.copy(source_file_path, backup_file_path)

    except Exception as e:
        messagebox.showerror("Warning", f"Failed to backup ICO files: {e}")

def generate_crc32_name():
    random_string = ''.join(random.choices(string.ascii_letters, k=50))
    return f"{zlib.crc32(random_string.encode()):08x}"

def create_icon_with_multiple_sizes(image_path, save_directory):
    try:
        icon_name = f"{generate_crc32_name()}.ico"
        icon_save_path = os.path.join(save_directory, icon_name)

        # Open the image using Pillow
        img = Image.open(image_path)

        # Ensure image has an alpha channel for transparency
        img = img.convert("RGBA")

        # Define standard icon sizes Windows expects
        sizes = [16, 32, 48, 64, 128, 256]

        # Save the image as an ICO file with multiple sizes
        img.save(icon_save_path, format='ICO', sizes=[(size, size) for size in sizes])

        return icon_save_path
    except Exception as e:
        raise OSError(f"Failed to create icon: {e}")

def browse_lnk():
    file_path = filedialog.askopenfilename(filetypes=[("Shortcut files", "*.lnk")])
    lnk_entry.delete(0, tk.END)
    lnk_entry.insert(0, file_path)

def browse_image():
    file_path = filedialog.askopenfilename(
        filetypes=[("Supported files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tiff;*.webp;*.ico;*.exe;*.dll;*.svg")]
    )
    png_entry.delete(0, tk.END)
    png_entry.insert(0, file_path)

def apply_icon():
    try:
        lnk_path = lnk_entry.get()
        png_or_url = png_entry.get()

        # Reset progress bar at the start
        progress_var.set(0)
        root.update_idletasks()

        # Read the source directory
        source_dir = read_directory_from_file(SOURCE_DIR_FILE)
        ensure_valid_directory(source_dir)
        progress_var.set(10)  # Progress: Source directory validated
        root.update_idletasks()

        # Handle SVG files
        if png_or_url.lower().endswith(".svg"):
            temp_png_path = os.path.join(source_dir, "converted_image.png")
            try:
                convert_svg_to_png(png_or_url, temp_png_path)
                progress_var.set(50)  # Progress: SVG converted
                root.update_idletasks()

                icon_save_path = create_icon_with_multiple_sizes(temp_png_path, source_dir)
                os.remove(temp_png_path)  # Clean up the temporary PNG file
                progress_var.set(70)  # Progress: Icon generated
                root.update_idletasks()
            except Exception as e:
                progress_var.set(0)  # Reset progress on error
                messagebox.showerror("Error", f"Failed to process SVG file: {e}")
                return

        elif png_or_url.lower().endswith(".ico"):
            icon_save_path = png_or_url
            progress_var.set(30)  # Progress: ICO file selected
            root.update_idletasks()

        elif png_or_url.lower().endswith(".exe") or png_or_url.lower().endswith(".dll"):
            selected_index = tk.simpledialog.askinteger(
                "Icon Index", f"Enter the icon index for '{png_or_url}':"
            )
            if selected_index is None:
                messagebox.showinfo("Action Cancelled", "Icon application was cancelled.")
                progress_var.set(0)  # Reset progress on cancellation
                return

            icon_save_path = f"{png_or_url}, {selected_index}"
            progress_var.set(40)  # Progress: EXE/DLL index selected
            root.update_idletasks()

        elif png_or_url.startswith("http"):
            response = requests.get(png_or_url, stream=True)
            if response.status_code == 200:
                temp_image_path = os.path.join(source_dir, "temp_image")
                file_extension = ".png"  # Default extension
                temp_image_path += file_extension

                with open(temp_image_path, 'wb') as temp_file:
                    temp_file.write(response.content)
                progress_var.set(60)  # Progress: File downloaded
                root.update_idletasks()

                icon_save_path = create_icon_with_multiple_sizes(temp_image_path, source_dir)
                os.remove(temp_image_path)  # Clean up the temporary file
                progress_var.set(80)  # Progress: Icon generated
                root.update_idletasks()
            else:
                progress_var.set(0)  # Reset progress on error
                messagebox.showerror("Error", f"Failed to download image from URL. Status code: {response.status_code}")
                return
        elif os.path.exists(png_or_url) and any(png_or_url.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff', '.webp']):
            icon_save_path = create_icon_with_multiple_sizes(png_or_url, source_dir)
            progress_var.set(80)  # Progress: Icon generated
            root.update_idletasks()
        else:
            progress_var.set(0)  # Reset progress on error
            messagebox.showerror("Error", "Invalid image file, URL, or unsupported format!")
            return

        # Apply the icon to the shortcut
        if os.path.exists(lnk_path) and lnk_path.lower().endswith(".lnk"):
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(lnk_path)
            shortcut.IconLocation = icon_save_path
            shortcut.Save()
            progress_var.set(100)  # Progress: Shortcut updated
            root.update_idletasks()

            messagebox.showinfo("Success", f"Icon applied successfully as '{os.path.basename(icon_save_path)}'!")
            backup_ico_files()
        else:
            progress_var.set(0)  # Reset progress on error
            messagebox.showerror("Error", "Invalid LNK file!")
    except Exception as e:
        progress_var.set(0)  # Reset progress on error
        messagebox.showerror("Error", f"An error occurred: {e}")

def on_drop_lnk(event):
    lnk_path = event.data.strip('"').strip('{}')
    if os.path.exists(lnk_path) and lnk_path.lower().endswith(".lnk"):
        lnk_entry.delete(0, tk.END)
        lnk_entry.insert(0, lnk_path)
    else:
        messagebox.showerror("Error", "Please drop a valid LNK file.")

def on_drop_image(event):
    image_path = event.data.strip('"').strip('{}')
    if os.path.exists(image_path) and any(image_path.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff', '.webp', '.exe', '.dll', '.svg']):
        png_entry.delete(0, tk.END)
        png_entry.insert(0, image_path)
    else:
        messagebox.showerror("Error", "Please drop a valid image file.")

def delete_orphaned_icons():
    try:
        # Read and validate the source directory
        source_dir = read_directory_from_file(SOURCE_DIR_FILE)
        ensure_valid_directory(source_dir)

        # Validate the backup directory
        backup_dir = validate_backupdir()
        
        # Get desktop path for checking shortcuts
        desktop_path = os.path.join(os.environ["USERPROFILE"], "Desktop")
        if not os.path.exists(desktop_path):
            messagebox.showerror("Error", "Desktop path not found!")
            return

        # Get a list of all .lnk files on the desktop
        desktop_shortcuts = [os.path.join(desktop_path, f) for f in os.listdir(desktop_path) if f.lower().endswith(".lnk")]

        # Store all used icons
        used_icons = set()

        # Check each shortcut for its icon
        shell = win32com.client.Dispatch("WScript.Shell")
        for shortcut_path in desktop_shortcuts:
            try:
                shortcut = shell.CreateShortcut(shortcut_path)
                if shortcut.IconLocation:
                    icon_path = shortcut.IconLocation.split(",")[0]  # Extract icon path
                    used_icons.add(os.path.abspath(icon_path))
            except Exception as e:
                # Skip problematic shortcuts
                print(f"Failed to read shortcut '{shortcut_path}': {e}")

        # Get all .ico files in the source directory
        orphaned_icons = []
        for file_name in os.listdir(source_dir):
            if file_name.lower().endswith(".ico"):
                icon_path = os.path.join(source_dir, file_name)
                if os.path.abspath(icon_path) not in used_icons:
                    orphaned_icons.append(icon_path)

        # Delete orphaned icons
        for icon_path in orphaned_icons:
            try:
                os.remove(icon_path)
            except Exception as e:
                messagebox.showwarning("Warning", f"Failed to delete orphaned icon '{icon_path}': {e}")

        # Replace icons in the backup directory
        if backup_dir:
            try:
                # Clear backup directory
                shutil.rmtree(backup_dir)
                os.makedirs(backup_dir)

                # Copy remaining .ico files from source directory to backup directory
                for file_name in os.listdir(source_dir):
                    if file_name.lower().endswith(".ico"):
                        shutil.copy(os.path.join(source_dir, file_name), os.path.join(backup_dir, file_name))
            except Exception as e:
                messagebox.showwarning("Warning", f"Failed to update backup directory: {e}")

        messagebox.showinfo("Success", "Orphaned icons deleted and backup replaced successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during orphaned icon deletion: {e}")

def paste_image_from_clipboard():
    try:
        # Check if the clipboard contains image data
        img = ImageGrab.grabclipboard()

        # Handle the specific case where clipboard returns a list (common with browser images)
        if isinstance(img, list):
            messagebox.showerror(
                "Error",
                "Cannot paste clipboard image this way; chances are you're trying to copy the image from a web browser. Try pasting into the input field or try again."
            )
            return

        # If the clipboard does not contain valid image data or lacks the `save` attribute
        if img is None or not hasattr(img, 'save'):
            messagebox.showerror(
                "Error",
                "Cannot paste clipboard image this way; make sure to copy a valid image and try again."
            )
            return

        # Use the system's temporary directory to store the image temporarily
        temp_dir = tempfile.gettempdir()
        temp_image_path = os.path.join(temp_dir, "clipboard_image.png")

        # Save the image from clipboard as a temporary file
        img.save(temp_image_path, format="PNG")

        # Read the source directory path
        source_dir = read_directory_from_file(SOURCE_DIR_FILE)
        ensure_valid_directory(source_dir)

        # Process the temporary image into an icon
        create_icon_with_multiple_sizes(temp_image_path, source_dir)

        # Update the text entry field to show clipboard input
        png_entry.delete(0, tk.END)
        png_entry.insert(0, "<clipboard input>")

        # Delete the temporary file after processing
        os.remove(temp_image_path)

    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred while processing the clipboard image: {e}")

def revert_shortcut_icon():
    try:
        # Get the shortcut (LNK) path from the entry field
        lnk_path = lnk_entry.get()

        # Ensure the path exists and is a valid LNK file
        if not os.path.exists(lnk_path) or not lnk_path.lower().endswith(".lnk"):
            messagebox.showerror("Error", "Invalid shortcut file! Please enter a valid .lnk file.")
            return

        # Use Windows Shell to access the shortcut's properties
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(lnk_path)

        # Reset the icon location to the shortcut's target path
        target_path = shortcut.TargetPath
        if not os.path.exists(target_path):
            messagebox.showerror("Error", f"Target file '{target_path}' does not exist. Cannot revert icon.")
            return

        try:
            # Set IconLocation to the target file with index 0 (default icon of the target)
            shortcut.IconLocation = f"{target_path}, 0"
            shortcut.Save()
            messagebox.showinfo("Success", f"Shortcut icon successfully reverted to the default icon of '{target_path}'.")
        except Exception as inner_error:
            messagebox.showerror(
                "Error",
                f"Failed to revert the shortcut icon: {inner_error}. This may happen if the shortcut is corrupt or restricted. Please try editing manually."
            )

    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred while reverting the shortcut: {e}")

# Validation before launching GUI
if not validate_sourcedir():
    sys.exit()

# GUI Setup
root = TkinterDnD.Tk()
root.title("Conopida")

# Center the window on the screen
window_width = 580
window_height = 200
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_position = (screen_width // 2) - (window_width // 2)
y_position = (screen_height // 2) - (window_height // 2)
root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
root.resizable(False, False)

# Input Fields
# Shortcut (LNK) Path
tk.Label(root, text="Shortcut (LNK) Path:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
lnk_entry = tk.Entry(root, width=50)
lnk_entry.grid(row=0, column=1, padx=10, pady=10)
lnk_button = tk.Button(root, text="Browse", command=browse_lnk)
lnk_button.grid(row=0, column=2, padx=10, pady=10)

# Image File or URL
tk.Label(root, text="Image File or URL:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
png_entry = tk.Entry(root, width=50)
png_entry.grid(row=1, column=1, padx=10, pady=10)
png_button = tk.Button(root, text="Browse", command=browse_image)
png_button.grid(row=1, column=2, padx=10, pady=10)
paste_button = tk.Button(root, text="Paste", command=paste_image_from_clipboard)
paste_button.grid(row=1, column=3)

# Button Group
button_frame = tk.Frame(root)  # Create a frame for the buttons
button_frame.grid(row=2, column=0, columnspan=3, pady=20)  # Position the frame

# Apply Button
apply_button = tk.Button(button_frame, text="Apply", command=apply_icon, width=15)
apply_button.pack(side="left", padx=5)

# Delete Orphaned Icons Button
delete_orphaned_button = tk.Button(button_frame, text="Delete Orphaned Icons", command=delete_orphaned_icons, width=20)
delete_orphaned_button.pack(side="left", padx=5)

# Revert Shortcut Button
revert_button = tk.Button(button_frame, text="Revert to Default", command=revert_shortcut_icon, width=20)
revert_button.pack(side="left", padx=5)

# Drag-and-Drop Support
lnk_entry.drop_target_register(DND_FILES)
lnk_entry.dnd_bind('<<Drop>>', on_drop_lnk)
png_entry.drop_target_register(DND_FILES)
png_entry.dnd_bind('<<Drop>>', on_drop_image)

# Progress Bar
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.grid(row=5, column=0, columnspan=3, sticky="we", padx=10, pady=5)

# Run the Tkinter Event Loop
root.mainloop()