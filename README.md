# Automation GUI

Automation GUI is a powerful and flexible Python application that allows users to create automated workflows for various tasks on their computer. With an intuitive graphical interface, users can easily record and playback a series of actions, making it an ideal tool for task automation, testing, and productivity enhancement.

## Features

- **Element Detection**: Automatically detect and interact with UI elements across different applications.
- **Diverse Actions**: Supports a wide range of actions including clicks, keyboard input, window management, file operations, and more.
- **Flexible Scripting**: Generate Python scripts from recorded actions for further customization and integration.
- **Cross-platform Compatibility**: Works on Windows, macOS, and Linux (with some platform-specific limitations).

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/automation-gui.git
   ```
2. Navigate to the project directory:
   ```
   cd automation-gui
   ```
3. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```
   python automation_gui.py
   ```
2. Use the "Start Detection" button to begin capturing UI elements.
3. Select actions from the dropdown menu and add them to your workflow.
4. Click "Generate Script" to create a Python script of your recorded actions.

## Dependencies

- tkinter
- pyautogui
- pygetwindow
- uiautomation (Windows only)
- Pillow

## Disclaimer

This tool is intended for legitimate use cases such as task automation and testing. Users are responsible for ensuring they have the necessary permissions to automate interactions with software and systems.
