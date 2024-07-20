import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import pyautogui
import time
import keyboard
import pygetwindow as gw
import os
import uiautomation as auto
import subprocess
import pyperclip

class AutomationGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Automation GUI v6")
        self.master.geometry("1000x800")

        self.current_element = None
        self.steps = []
        self.detecting = False

        self.create_widgets()

    def create_widgets(self):
        self.main_frame = ttk.Frame(self.master, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Element detection area
        self.detection_frame = ttk.LabelFrame(self.main_frame, text="Element Detection", padding="10")
        self.detection_frame.pack(fill=tk.X, pady=(0, 10))

        self.element_info = tk.StringVar()
        self.element_label = ttk.Label(self.detection_frame, textvariable=self.element_info, wraplength=950, justify=tk.LEFT)
        self.element_label.pack(fill=tk.X)

        # Function selection
        self.function_frame = ttk.LabelFrame(self.main_frame, text="Function Selection", padding="10")
        self.function_frame.pack(fill=tk.X, pady=(0, 10))

        self.functions = [
            "Click", "Double Click", "Right Click", "Type Text", "Press Key", "Wait",
            "Open Program", "Close Program", "Minimize", "Maximize", "Delete",
            "Copy Text", "Paste Text", "Drag and Drop", "Scroll", "Take Screenshot",
            "Move Mouse", "Move Window", "Resize Window", "Switch Window",
            "Press and Hold Key", "Release Key", "Send Hotkey", "Focus Element",
            "Get Text", "Set Value", "Select Item", "Check/Uncheck", "Open URL",
            "Run Command", "Create Folder", "Delete File/Folder", "Rename File/Folder",
            "Read File", "Write to File", "Append to File", "Find Image on Screen",
            "Wait for Element", "Wait for Image", "Loop Start", "Loop End", "If Condition",
            "Else Condition", "End If", "Break Loop", "Continue Loop"
        ]
        self.function_var = tk.StringVar()
        self.function_dropdown = ttk.Combobox(self.function_frame, textvariable=self.function_var, values=self.functions, width=30)
        self.function_dropdown.pack(side=tk.LEFT, padx=(0, 10))

        self.add_button = ttk.Button(self.function_frame, text="Add Step", command=self.add_step)
        self.add_button.pack(side=tk.LEFT, padx=(0, 10))

        self.remove_button = ttk.Button(self.function_frame, text="Remove Step", command=self.remove_step)
        self.remove_button.pack(side=tk.LEFT)

        # Steps list
        self.steps_frame = ttk.LabelFrame(self.main_frame, text="Steps", padding="10")
        self.steps_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.steps_listbox = tk.Listbox(self.steps_frame, width=120, height=20)
        self.steps_listbox.pack(fill=tk.BOTH, expand=True)

        # Control buttons
        self.control_frame = ttk.Frame(self.main_frame)
        self.control_frame.pack(fill=tk.X)

        self.start_button = ttk.Button(self.control_frame, text="Start Detection", command=self.toggle_detection)
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))

        self.complete_button = ttk.Button(self.control_frame, text="Generate Script", command=self.generate_script)
        self.complete_button.pack(side=tk.LEFT)

    def toggle_detection(self):
        if self.detecting:
            self.detecting = False
            self.start_button.config(text="Start Detection")
            keyboard.unhook_all()
        else:
            self.detecting = True
            self.start_button.config(text="Stop Detection")
            keyboard.on_press_key('ctrl', self.on_ctrl_press)
            self.detect_element()

    def on_ctrl_press(self, e):
        if self.detecting:
            self.add_step(auto=True)

    def get_element_info(self):
        try:
            element = auto.ControlFromCursor()
            return {
                "title": element.Name,
                "class_name": element.ClassName,
                "control_type": element.ControlTypeName,
                "automation_id": element.AutomationId,
                "process_id": element.ProcessId,
                "rect": element.BoundingRectangle,
            }
        except Exception as e:
            print(f"Error getting element info: {str(e)}")
            return None

    def detect_element(self):
        if not self.detecting:
            return

        try:
            element_info = self.get_element_info()
            if element_info:
                self.current_element = element_info
                info = f"Title: {element_info['title']}\n" \
                       f"Class Name: {element_info['class_name']}\n" \
                       f"Control Type: {element_info['control_type']}\n" \
                       f"Automation ID: {element_info['automation_id']}\n" \
                       f"Process ID: {element_info['process_id']}\n" \
                       f"Position: {element_info['rect']}"
                self.element_info.set(info)
            else:
                self.current_element = None
                self.element_info.set("No element detected")

        except Exception as e:
            print(f"Error in detect_element: {str(e)}")
            self.element_info.set("Error detecting element")

        self.master.after(100, self.detect_element)

    def add_step(self, auto=False):
        function = self.function_var.get()
        if not function and not auto:
            messagebox.showerror("Error", "Please select a function")
            return

        step_info = f"{function}: " if function else "Detected: "

        if auto or function in ["Click", "Double Click", "Right Click", "Close Program", "Minimize", "Maximize", "Delete", "Focus Element", "Get Text"]:
            if self.current_element is None:
                messagebox.showerror("Error", "No element detected")
                return
            step_info += self.get_element_description(self.current_element)
        elif function in ["Type Text", "Press Key", "Press and Hold Key", "Release Key", "Send Hotkey", "Set Value"]:
            if self.current_element is None:
                messagebox.showerror("Error", "No element detected")
                return
            text = simpledialog.askstring("Input", f"Enter {'text' if function == 'Type Text' else 'key(s)'} to {function.lower()}:")
            if text:
                step_info += f"'{text}' in {self.get_element_description(self.current_element)}"
            else:
                return
        elif function == "Wait":
            seconds = simpledialog.askinteger("Input", "Enter wait time in seconds:")
            if seconds:
                step_info += f"{seconds} seconds"
            else:
                return
        elif function == "Open Program":
            program = filedialog.askopenfilename(title="Select program to open")
            if program:
                step_info += f"{program}"
            else:
                return
        elif function in ["Drag and Drop", "Move Mouse", "Move Window", "Resize Window"]:
            start_x = simpledialog.askinteger("Input", "Enter start X coordinate:")
            start_y = simpledialog.askinteger("Input", "Enter start Y coordinate:")
            end_x = simpledialog.askinteger("Input", "Enter end X coordinate:")
            end_y = simpledialog.askinteger("Input", "Enter end Y coordinate:")
            if all([start_x, start_y, end_x, end_y]):
                step_info += f"from ({start_x}, {start_y}) to ({end_x}, {end_y})"
            else:
                return
        elif function == "Scroll":
            amount = simpledialog.askinteger("Input", "Enter scroll amount (positive for up, negative for down):")
            if amount is not None:
                step_info += f"{amount}"
            else:
                return
        elif function == "Take Screenshot":
            filename = simpledialog.askstring("Input", "Enter filename for screenshot:")
            if filename:
                step_info += f"{filename}"
            else:
                return
        elif function == "Switch Window":
            window_title = simpledialog.askstring("Input", "Enter window title to switch to:")
            if window_title:
                step_info += f"{window_title}"
            else:
                return
        elif function in ["Select Item", "Check/Uncheck"]:
            item = simpledialog.askstring("Input", "Enter item to select or check/uncheck:")
            if item:
                step_info += f"{item}"
            else:
                return
        elif function == "Open URL":
            url = simpledialog.askstring("Input", "Enter URL to open:")
            if url:
                step_info += f"{url}"
            else:
                return
        elif function == "Run Command":
            command = simpledialog.askstring("Input", "Enter command to run:")
            if command:
                step_info += f"{command}"
            else:
                return
        elif function in ["Create Folder", "Delete File/Folder", "Rename File/Folder"]:
            path = filedialog.askdirectory(title=f"Select path for {function}")
            if path:
                if function == "Rename File/Folder":
                    new_name = simpledialog.askstring("Input", "Enter new name:")
                    if new_name:
                        step_info += f"{path} to {new_name}"
                    else:
                        return
                else:
                    step_info += f"{path}"
            else:
                return
        elif function in ["Read File", "Write to File", "Append to File"]:
            file_path = filedialog.askopenfilename(title=f"Select file to {function}")
            if file_path:
                if function in ["Write to File", "Append to File"]:
                    content = simpledialog.askstring("Input", "Enter content to write/append:")
                    if content:
                        step_info += f"{file_path} with content: {content}"
                    else:
                        return
                else:
                    step_info += f"{file_path}"
            else:
                return
        elif function == "Find Image on Screen":
            image_path = filedialog.askopenfilename(title="Select image to find")
            if image_path:
                step_info += f"{image_path}"
            else:
                return
        elif function in ["Wait for Element", "Wait for Image"]:
            timeout = simpledialog.askinteger("Input", "Enter timeout in seconds:")
            if timeout:
                step_info += f"timeout: {timeout} seconds"
            else:
                return
        elif function in ["Loop Start", "If Condition"]:
            condition = simpledialog.askstring("Input", "Enter condition:")
            if condition:
                step_info += f"{condition}"
            else:
                return
        elif function in ["Loop End", "Else Condition", "End If", "Break Loop", "Continue Loop"]:
            step_info += "No additional info needed"

        self.steps.append((function, step_info, self.current_element))
        self.steps_listbox.insert(tk.END, step_info)
        print(f"Added step: {step_info}")

        if not auto:
            self.function_var.set("")

    def remove_step(self):
        selected_indices = self.steps_listbox.curselection()
        if not selected_indices:
            messagebox.showerror("Error", "No step selected")
            return
        
        index = selected_indices[0]
        self.steps_listbox.delete(index)
        del self.steps[index]
        print(f"Removed step at index {index}")

    def get_element_description(self, element_info):
        description = []
        if element_info['title']:
            description.append(f"Title: '{element_info['title']}'")
        if element_info['class_name']:
            description.append(f"Class: '{element_info['class_name']}'")
        if element_info['control_type']:
            description.append(f"Type: '{element_info['control_type']}'")
        if element_info['automation_id']:
            description.append(f"ID: '{element_info['automation_id']}'")
        if element_info['process_id']:
            description.append(f"Process ID: {element_info['process_id']}")
        return ", ".join(description)

    def generate_script(self):
        if not self.steps:
            messagebox.showerror("Error", "No steps added")
            return

        script = """import pyautogui
import pygetwindow as gw
import time
import uiautomation as auto
import os
import subprocess
import pyperclip
import platform
import webbrowser
from PIL import Image

def find_element(**kwargs):
    print(f"Searching for element: {kwargs}")
    timeout = kwargs.get('timeout', 10)
    start_time = time.time()
    while time.time() - start_time < timeout:
        windows = gw.getAllWindows()
        for window in windows:
            try:
                element = auto.ControlFromHandle(window._hWnd)
                if all(getattr(element, k, None) == v for k, v in kwargs.items() if k != 'timeout'):
                    print("Element found!")
                    return element
            except Exception as e:
                print(f"Error checking window: {str(e)}")
        time.sleep(0.5)
    print(f"Element not found after {timeout} seconds")
    return None

def perform_action(element, action, **kwargs):
    if element:
        print(f"Performing action: {action}")
        element.SetFocus()
        time.sleep(0.5)  # Wait for focus
        if action == "click":
            element.Click()
        elif action == "double_click":
            element.DoubleClick()
        elif action == "right_click":
            element.RightClick()
        elif action == "type":
            element.SendKeys(kwargs.get('text', ''))
        elif action == "press_key":
            pyautogui.press(kwargs.get('key', ''))
        elif action == "close":
            element.Close()
        elif action == "minimize":
            element.Minimize()
        elif action == "maximize":
            element.Maximize()
        elif action == "delete":
            element.Delete()
        elif action == "copy":
            element.SetFocus()
            pyautogui.hotkey('ctrl', 'c')
        elif action == "paste":
            element.SetFocus()
            pyautogui.hotkey('ctrl', 'v')
        elif action == "drag_and_drop":
            start_x, start_y, end_x, end_y = kwargs.get('coordinates', (0, 0, 0, 0))
            pyautogui.moveTo(start_x, start_y)
            pyautogui.dragTo(end_x, end_y, duration=1)
        elif action == "scroll":
            pyautogui.scroll(kwargs.get('amount', 0))
        elif action == "move_mouse":
            x, y = kwargs.get('coordinates', (0, 0))
            pyautogui.moveTo(x, y)
        elif action == "move_window":
            x, y = kwargs.get('coordinates', (0, 0))
            element.MoveTo(x, y)
        elif action == "resize_window":
            width, height = kwargs.get('size', (0, 0))
            element.Resize(width, height)
        elif action == "press_and_hold":
            pyautogui.keyDown(kwargs.get('key', ''))
        elif action == "release_key":
            pyautogui.keyUp(kwargs.get('key', ''))
        elif action == "send_hotkey":
            pyautogui.hotkey(*kwargs.get('keys', []))
        elif action == "focus":
            element.SetFocus()
        elif action == "get_text":
            return element.GetWindowText()
        elif action == "set_value":
            element.SetValue(kwargs.get('value', ''))
        elif action == "select_item":
            element.Select(kwargs.get('item', ''))
        elif action == "check_uncheck":
            if kwargs.get('check', True):
                element.Check()
            else:
                element.Uncheck()
        print(f"Action {action} completed")
    else:
        print(f"Cannot perform action {action}: Element not found")

def get_file_association(file_extension):
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, file_extension) as key:
            prog_id = winreg.QueryValue(key, None)
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{prog_id}\\shell\\open\\command") as key:
            command = winreg.QueryValue(key, None)
        return command
    except WindowsError:
        return None

def open_program(path):
    print(f"Opening file or program: {path}")
    try:
        file_extension = os.path.splitext(path)[1].lower()
        
        if platform.system() == 'Windows':
            if file_extension in ['.exe', '.bat', '.cmd']:
                # For executable files, use subprocess.Popen
                subprocess.Popen(path, shell=True)
            elif file_extension == '.lnk':
                # For shortcuts, resolve the target and open it
                import win32com.client
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(path)
                target_path = shortcut.Targetpath
                subprocess.Popen(target_path, shell=True)
            else:
                # For other file types, use the default application
                os.startfile(path)
        elif platform.system() == 'Darwin':  # macOS
            subprocess.call(('open', path))
        else:  # linux variants
            subprocess.call(('xdg-open', path))
        
        print(f"File or program opened successfully")
    except Exception as e:
        print(f"Error opening file or program: {str(e)}")
        # If there's an error, try to find the associated program and run it
        if platform.system() == 'Windows':
            file_extension = os.path.splitext(path)[1]
            associated_command = get_file_association(file_extension)
            if associated_command:
                try:
                    associated_command = associated_command.replace('%1', f'"{path}"')
                    subprocess.Popen(associated_command, shell=True)
                    print(f"Opened with associated program: {associated_command}")
                except Exception as e:
                    print(f"Error opening with associated program: {str(e)}")

def switch_window(title):
    windows = gw.getWindowsWithTitle(title)
    if windows:
        windows[0].activate()
        print(f"Switched to window: {title}")
    else:
        print(f"Window not found: {title}")

def take_screenshot(filename):
    pyautogui.screenshot(filename)
    print(f"Screenshot saved as {filename}")

def open_url(url):
    webbrowser.open(url)
    print(f"Opened URL: {url}")

def run_command(command):
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        print(f"Command executed successfully. Output: {result.stdout}")
    except subprocess.CalledProcessError as e:
        print(f"Error executing command: {e}")

def create_folder(path):
    os.makedirs(path, exist_ok=True)
    print(f"Folder created: {path}")

def delete_file_folder(path):
    if os.path.isfile(path):
        os.remove(path)
    elif os.path.isdir(path):
        import shutil
        shutil.rmtree(path)
    print(f"Deleted: {path}")

def rename_file_folder(old_path, new_name):
    new_path = os.path.join(os.path.dirname(old_path), new_name)
    os.rename(old_path, new_path)
    print(f"Renamed {old_path} to {new_path}")

def read_file(file_path):
    with open(file_path, 'r') as file:
        content = file.read()
    print(f"File content: {content}")
    return content

def write_to_file(file_path, content):
    with open(file_path, 'w') as file:
        file.write(content)
    print(f"Content written to {file_path}")

def append_to_file(file_path, content):
    with open(file_path, 'a') as file:
        file.write(content)
    print(f"Content appended to {file_path}")

def find_image_on_screen(image_path):
    location = pyautogui.locateOnScreen(image_path)
    if location:
        print(f"Image found at: {location}")
        return location
    else:
        print("Image not found on screen")
        return None

def wait_for_element(**kwargs):
    timeout = kwargs.pop('timeout', 10)
    element = find_element(timeout=timeout, **kwargs)
    if element:
        print(f"Element found within {timeout} seconds")
    else:
        print(f"Element not found within {timeout} seconds")
    return element

def wait_for_image(image_path, timeout):
    start_time = time.time()
    while time.time() - start_time < timeout:
        location = pyautogui.locateOnScreen(image_path)
        if location:
            print(f"Image found at: {location}")
            return location
        time.sleep(0.5)
    print(f"Image not found within {timeout} seconds")
    return None

"""
        
        indent_level = 0
        for function, step_info, element_info in self.steps:
            try:
                if function in ["Click", "Double Click", "Right Click", "Close Program", "Minimize", "Maximize", "Delete", "Focus Element", "Get Text", "Copy Text", "Paste Text"]:
                    script += "    " * indent_level
                    script += f"element = find_element({self.get_element_kwargs(element_info)})\n"
                    action = function.lower().replace(" ", "_")
                    if action == "close_program":
                        action = "close"
                    elif action == "get_text":
                        script += "    " * indent_level
                        script += f"text = perform_action(element, '{action}')\n"
                        script += "    " * indent_level
                        script += f"print(f'Got text: {{text}}')\n"
                    else:
                        script += "    " * indent_level
                        script += f"perform_action(element, '{action}')\n"
                elif function in ["Type Text", "Press Key", "Press and Hold Key", "Release Key", "Send Hotkey", "Set Value"]:
                    script += "    " * indent_level
                    script += f"element = find_element({self.get_element_kwargs(element_info)})\n"
                    action = function.lower().replace(" ", "_")
                    text = step_info.split("'")[1]
                    script += "    " * indent_level
                    if function == "Send Hotkey":
                        script += f"perform_action(element, '{action}', keys={text.split(',')})\n"
                    else:
                        script += f"perform_action(element, '{action}', {'text' if function == 'Type Text' else 'key' if function in ['Press Key', 'Press and Hold Key', 'Release Key'] else 'value'}='{text}')\n"
                elif function == "Wait":
                    seconds = int(step_info.split()[0])
                    script += "    " * indent_level
                    script += f"print(f'Waiting for {seconds} seconds')\n"
                    script += "    " * indent_level
                    script += f"time.sleep({seconds})\n"
                elif function == "Open Program":
                    program = step_info.split(": ", 1)[1]
                    script += "    " * indent_level
                    script += f"open_program(r'{program}')\n"
                    script += "    " * indent_level
                    script += f"time.sleep(2)  # Wait for the program to open\n"
                elif function in ["Drag and Drop", "Move Mouse", "Move Window", "Resize Window"]:
                    coords = step_info.split("(")[1].split(")")[0].split(", ")
                    start_x, start_y, end_x, end_y = map(int, coords)
                    script += "    " * indent_level
                    if function == "Drag and Drop":
                        script += f"perform_action(None, 'drag_and_drop', coordinates=({start_x}, {start_y}, {end_x}, {end_y}))\n"
                    elif function == "Move Mouse":
                        script += f"perform_action(None, 'move_mouse', coordinates=({end_x}, {end_y}))\n"
                    elif function == "Move Window":
                        script += f"perform_action(gw.getActiveWindow(), 'move_window', coordinates=({end_x}, {end_y}))\n"
                    elif function == "Resize Window":
                        script += f"perform_action(gw.getActiveWindow(), 'resize_window', size=({end_x - start_x}, {end_y - start_y}))\n"
                elif function == "Scroll":
                    amount = int(step_info.split(": ")[1])
                    script += "    " * indent_level
                    script += f"perform_action(None, 'scroll', amount={amount})\n"
                elif function == "Take Screenshot":
                    filename = step_info.split(": ")[1]
                    script += "    " * indent_level
                    script += f"take_screenshot(r'{filename}')\n"
                elif function == "Switch Window":
                    window_title = step_info.split(": ")[1]
                    script += "    " * indent_level
                    script += f"switch_window('{window_title}')\n"
                elif function in ["Select Item", "Check/Uncheck"]:
                    item = step_info.split(": ")[1]
                    script += "    " * indent_level
                    script += f"element = find_element({self.get_element_kwargs(element_info)})\n"
                    script += "    " * indent_level
                    if function == "Select Item":
                        script += f"perform_action(element, 'select_item', item='{item}')\n"
                    else:
                        script += f"perform_action(element, 'check_uncheck', check={item.lower() == 'check'})\n"
                elif function == "Open URL":
                    url = step_info.split(": ")[1]
                    script += "    " * indent_level
                    script += f"open_url('{url}')\n"
                elif function == "Run Command":
                    command = step_info.split(": ")[1]
                    script += "    " * indent_level
                    script += f"run_command('{command}')\n"
                elif function == "Create Folder":
                    path = step_info.split(": ")[1]
                    script += "    " * indent_level
                    script += f"create_folder(r'{path}')\n"
                elif function == "Delete File/Folder":
                    path = step_info.split(": ")[1]
                    script += "    " * indent_level
                    script += f"delete_file_folder(r'{path}')\n"
                elif function == "Rename File/Folder":
                    old_path, new_name = step_info.split(" to ")
                    old_path = old_path.split(": ")[1]
                    script += "    " * indent_level
                    script += f"rename_file_folder(r'{old_path}', '{new_name}')\n"
                elif function == "Read File":
                    file_path = step_info.split(": ")[1]
                    script += "    " * indent_level
                    script += f"content = read_file(r'{file_path}')\n"
                elif function in ["Write to File", "Append to File"]:
                    file_path, content = step_info.split(" with content: ")
                    file_path = file_path.split(": ")[1]
                    script += "    " * indent_level
                    if function == "Write to File":
                        script += f"write_to_file(r'{file_path}', '{content}')\n"
                    else:
                        script += f"append_to_file(r'{file_path}', '{content}')\n"
                elif function == "Find Image on Screen":
                    image_path = step_info.split(": ")[1]
                    script += "    " * indent_level
                    script += f"location = find_image_on_screen(r'{image_path}')\n"
                elif function in ["Wait for Element", "Wait for Image"]:
                    timeout = int(step_info.split(": ")[1].split()[0])
                    if function == "Wait for Element":
                        script += "    " * indent_level
                        script += f"element = wait_for_element({self.get_element_kwargs(element_info)}, timeout={timeout})\n"
                    else:
                        image_path = step_info.split(": ")[1]
                        script += "    " * indent_level
                        script += f"location = wait_for_image(r'{image_path}', {timeout})\n"
                elif function == "Loop Start":
                    condition = step_info.split(": ")[1]
                    script += "    " * indent_level
                    script += f"while {condition}:\n"
                    indent_level += 1
                elif function == "Loop End":
                    indent_level = max(0, indent_level - 1)
                elif function == "If Condition":
                    condition = step_info.split(": ")[1]
                    script += "    " * indent_level
                    script += f"if {condition}:\n"
                    indent_level += 1
                elif function == "Else Condition":
                    indent_level = max(0, indent_level - 1)
                    script += "    " * indent_level
                    script += "else:\n"
                    indent_level += 1
                elif function == "End If":
                    indent_level = max(0, indent_level - 1)
                elif function == "Break Loop":
                    script += "    " * indent_level
                    script += "break\n"
                elif function == "Continue Loop":
                    script += "    " * indent_level
                    script += "continue\n"
            except Exception as e:
                script += "    " * indent_level
                script += f"print(f'Error in {function}: {str(e)}')\n"
                script += "    " * indent_level
                script += "continue\n"

        script += "\nprint('Script execution completed')\n"

        with open("generated_script.py", "w") as file:
            file.write(script)

        messagebox.showinfo("Script Generated", "The script has been generated and saved as 'generated_script.py'")

    def get_element_kwargs(self, element_info):
        kwargs = {}
        if element_info['title']:
            kwargs['Name'] = element_info['title']
        if element_info['class_name']:
            kwargs['ClassName'] = element_info['class_name']
        if element_info['control_type']:
            kwargs['ControlTypeName'] = element_info['control_type']
        if element_info['automation_id']:
            kwargs['AutomationId'] = element_info['automation_id']
        if element_info['process_id']:
            kwargs['ProcessId'] = element_info['process_id']
        return ', '.join(f"{k}='{v}'" for k, v in kwargs.items())

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomationGUI(root)
    root.mainloop()
