# **Conopida: Icon Manager for Windows Shortcuts**

![Demo](https://github.com/Tomurisk/Conopida/blob/main/Images/demo.gif)

Conopida is a simple yet powerful application to manage icons for Windows shortcuts. With it, you can apply custom icons to `.lnk` files, convert SVG images, and organize icon files with ease.

### **Features:**
- **Apply custom icons** to Windows shortcuts
- **Drag-and-drop** functionality for easy file selection
- Convert **SVG images** to PNG and create multi-size ICO files
- **Backup** your icon files
- Clean up **orphaned icons** that are not being used
- Paste images directly from the **clipboard**

## **How to Use Conopida**

### **Step 1: Set Up Directories**

Before using Conopida, two important files must be configured to set up your directories:

1. **\_sourcedir.txt**: This text file contains the directory path where your icon files are stored. It tells Conopida where to look for images that you want to apply to your shortcuts. The file should contain the absolute path to the folder (e.g., `C:\Users\YourName\Icons`).

2. **\_backupdir.txt**: This text file contains the directory path where the `.ico` backup files will be saved. When an icon is applied to a shortcut, the corresponding `.ico` file will be stored in this backup directory. The file should also contain the absolute path to the folder where backups should be stored (e.g., `C:\Users\YourName\IconBackups`).

Make sure these directories are valid and accessible by the program. If they are not set up correctly, Conopida will notify you to correct them.

---

## **Getting Started**
You can find binary releases in **Releases** section. If you plan to put your hands on the project yourself, here's what you'll need:

### **Prerequisites**
Before using the app, you need to ensure that the following libraries are installed:
1. `Pillow` for image manipulation
2. `cairosvg` for SVG to PNG conversion
3. `requests` for downloading icons from URLs
4. `pywin32` for managing Windows shortcuts
5. `tkinterdnd2` for drag-and-drop functionality in the GUI

To install the necessary libraries, run the following command:

```bash
pip install Pillow cairosvg requests pywin32 tkinterdnd2
```

