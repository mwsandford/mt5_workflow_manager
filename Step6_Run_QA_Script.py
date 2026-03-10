#!/usr/bin/env python3
"""
Quant Analyzer Script Runner
============================
Automates running Java scripts in Quant Analyzer Pro using Windows UI automation.

This script:
  1. Launches Quant Analyzer (or connects to running instance)
  2. Navigates to the Scripter panel
  3. Opens and runs the specified .java script
  4. Waits for completion
  5. Optionally closes Quant Analyzer

Usage:
------
    # Run a specific script
    python Step7_Run_QA_Script.py "BatchMC_Analysis_TradePC.java"

    # Run script and keep QA open afterward
    python Step7_Run_QA_Script.py "BatchMC_Analysis_TradePC.java" --keep-open

    # Custom QA path
    python Step7_Run_QA_Script.py "BatchMC_Analysis.java" --qa-path "D:\\QA\\QuantAnalyzer4.exe"

    # Debug mode - inspect UI elements without running
    python Step7_Run_QA_Script.py --inspect

    # Show help
    python Step7_Run_QA_Script.py --help

Requirements:
-------------
    pip install pywinauto

Arguments:
----------
    script_name             Name of the .java script to run (e.g., BatchMC_Analysis_TradePC.java)
    --qa-path PATH          Path to QuantAnalyzer4.exe (default: C:\\QuantAnalyzer4\\QuantAnalyzer4.exe)
    --keep-open             Keep Quant Analyzer open after script completes
    --timeout SECONDS       Max time to wait for script completion (default: 600 = 10 minutes)
    --inspect               Debug mode: launch QA and print UI element tree
"""

import argparse
import ctypes
import sys
import os
import time
from datetime import datetime

# Disable PyAutoGUI failsafe globally - required for unattended RDP automation
# Must be done before any pyautogui operations
try:
    import pyautogui
    pyautogui.FAILSAFE = False
    pyautogui.PAUSE = 0.3
except ImportError:
    pass  # Will be caught later when needed


# ============================================================================
# ANSI COLOURS FOR TERMINAL OUTPUT
# ============================================================================
class Colors:
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    GRAY = '\033[90m'
    RESET = '\033[0m'


def init_colors():
    """Enable ANSI colors on Windows."""
    if sys.platform == 'win32':
        os.system('')


def format_duration(seconds: float) -> str:
    """Format duration as hours, minutes, seconds."""
    total_hours = int(seconds // 3600)
    total_minutes = int((seconds % 3600) // 60)
    total_seconds = int(seconds % 60)
    
    if total_hours > 0:
        return f"{total_hours}h {total_minutes}m {total_seconds}s"
    elif total_minutes > 0:
        return f"{total_minutes}m {total_seconds}s"
    else:
        return f"{total_seconds}s"


# ============================================================================
# QUANT ANALYZER AUTOMATION
# ============================================================================
DEFAULT_QA_PATH = r"C:\QuantAnalyzer4\QuantAnalyzer4.exe"
DEFAULT_TIMEOUT = 3600  # 60 minutes

# Keyboard/mouse automation settings for Java app fallback
# These offsets are relative to the window position
SCRIPTER_CLICK_X_OFFSET = 60   # X offset from left edge of window to Scripter icon
SCRIPTER_CLICK_Y_OFFSET = 1076 # Y offset from top of window to Scripter icon
RUN_BUTTON_X_OFFSET = 200      # X offset to Run button in Scripter panel
RUN_BUTTON_Y_OFFSET = 256      # Y offset to Run button
NAVIGATOR_X_OFFSET = 1430      # X offset to Navigator scripts list (middle of script names)
NAVIGATOR_Y_OFFSET = 385       # Y offset to FIRST script in Navigator list
SCRIPT_ROW_HEIGHT = 30         # Pixels between each script row in the list


def check_pywinauto():
    """Check if pywinauto is installed."""
    try:
        import pywinauto
        return True
    except ImportError:
        return False


def check_pyautogui():
    """Check if pyautogui is installed (for Java app fallback)."""
    try:
        import pyautogui
        return True
    except ImportError:
        return False


def check_opencv():
    """Check if opencv-python is installed (for image recognition)."""
    try:
        import cv2
        return True
    except ImportError:
        return False


def get_templates_folder():
    """Get the folder where template images are stored."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(script_dir, "qa_templates")


def find_and_click(template_name: str, confidence: float = 0.8, timeout: int = 10, double: bool = False):
    """Find a template image on screen and click it.
    
    Args:
        template_name: Name of the template file (without path)
        confidence: Match confidence threshold (0-1)
        timeout: How long to search before giving up
        double: Whether to double-click
        
    Returns:
        True if found and clicked, False otherwise
    """
    import pyautogui
    import time
    
    templates_folder = get_templates_folder()
    template_path = os.path.join(templates_folder, template_name)
    
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template image not found: {template_path}")
    
    print(f"  Searching for: {template_name}")
    
    start_time = time.time()
    location = None
    screenshot_failed = False
    
    while time.time() - start_time < timeout:
        try:
            location = pyautogui.locateOnScreen(template_path, confidence=confidence)
            if location:
                break
        except pyautogui.ImageNotFoundException:
            # Image not found, keep trying
            pass
        except Exception as e:
            # Screenshot or other error - may indicate RDP disconnect
            if not screenshot_failed:
                print(f"  {Colors.YELLOW}Warning: Screenshot issue - {e}{Colors.RESET}")
                screenshot_failed = True
        time.sleep(0.5)
    
    if not location:
        if screenshot_failed:
            print(f"  {Colors.RED}Could not find {template_name} - screenshot failed{Colors.RESET}")
            print(f"  {Colors.YELLOW}TIP: RDP may be disconnected. Keep RDP minimized (not closed) for automation to work.{Colors.RESET}")
        else:
            print(f"  {Colors.RED}Could not find {template_name} on screen{Colors.RESET}")
        return False
    
    # Get center of the found region
    center_x = location.left + location.width // 2
    center_y = location.top + location.height // 2
    
    print(f"  Found at ({center_x}, {center_y})")
    
    if double:
        pyautogui.doubleClick(center_x, center_y)
    else:
        pyautogui.click(center_x, center_y)
    
    return True


def capture_templates(qa_path: str):
    """Interactive mode to capture template images for UI elements."""
    import pyautogui
    from PIL import Image
    
    print(f"\n{Colors.CYAN}{'=' * 60}{Colors.RESET}")
    print(f"{Colors.CYAN}TEMPLATE CAPTURE MODE{Colors.RESET}")
    print(f"{Colors.CYAN}{'=' * 60}{Colors.RESET}")
    print()
    
    # Create templates folder
    templates_folder = get_templates_folder()
    os.makedirs(templates_folder, exist_ok=True)
    print(f"Templates will be saved to: {templates_folder}")
    print()
    
    # Check if QA is running
    from pywinauto import Desktop
    desktop = Desktop(backend="uia")
    main_window = None
    
    for win in desktop.windows():
        try:
            title = win.window_text()
            if title and "Quant Analyzer" in title and "Properties" not in title:
                main_window = win
                break
        except Exception:
            pass
    
    if not main_window:
        print(f"{Colors.YELLOW}Quant Analyzer not running. Launching...{Colors.RESET}")
        # Launch QA
        import subprocess
        subprocess.Popen([qa_path], cwd=os.path.dirname(qa_path))
        print("Waiting for Quant Analyzer to start...")
        time.sleep(10)
        
        # Find window again
        for win in desktop.windows():
            try:
                title = win.window_text()
                if title and "Quant Analyzer" in title and "Properties" not in title:
                    main_window = win
                    break
            except Exception:
                pass
    
    if not main_window:
        print(f"{Colors.RED}Could not find Quant Analyzer window{Colors.RESET}")
        return
    
    # Get window position first
    rect = main_window.rectangle()
    
    # Bring to focus by clicking on the window title bar area, then maximize
    print("Maximizing window...")
    try:
        # Click on the window to ensure it has focus (click near top-center of window)
        click_x = rect.left + (rect.right - rect.left) // 2
        click_y = rect.top + 15  # Title bar area
        pyautogui.click(click_x, click_y)
        time.sleep(0.3)
        
        # Now send Win+Up to maximize
        pyautogui.hotkey('win', 'up')
        time.sleep(1)
    except Exception as e:
        print(f"{Colors.YELLOW}Warning: Could not maximize window: {e}{Colors.RESET}")
    
    print(f"Window: {main_window.window_text()}")
    print()
    
    templates_to_capture = [
        ("scripter_icon.png", "Scripter icon in the LEFT MENU (near bottom)"),
        ("run_button.png", "RUN button (green button in Scripter panel toolbar)"),
        ("script_batchmc_minute_data.png", "Script: BatchMC_Minute_Data.java - position at LEFT EDGE of text"),
        ("script_batchmc_tick_data.png", "Script: BatchMC_Tick_Data.java - position at LEFT EDGE of text"),
    ]
    
    print(f"{Colors.YELLOW}Instructions:{Colors.RESET}")
    print("1. When prompted, you have 5 seconds to position your mouse over the element")
    print("2. The script will capture a region around your cursor automatically")
    print("3. Press Ctrl+C to abort at any time")
    print()
    input(f"Press ENTER to begin capturing (then switch to QA window)...")
    print()
    
    for filename, description in templates_to_capture:
        filepath = os.path.join(templates_folder, filename)
        
        if os.path.exists(filepath):
            print(f"{Colors.GREEN}✓ {filename} already exists{Colors.RESET}")
            response = input(f"  Recapture? (y/N): ").strip().lower()
            if response != 'y':
                continue
        
        print(f"\n{Colors.CYAN}Next: {description}{Colors.RESET}")
        print(f"  Switch to QA and position your mouse over this element...")
        
        # Countdown
        for i in range(5, 0, -1):
            print(f"  Capturing in {i}...", end='\r')
            time.sleep(1)
        print(f"  Capturing now!      ")
        
        # Capture region around cursor
        x, y = pyautogui.position()
        
        # Capture a region (adjust size based on element type)
        if 'icon' in filename:
            width, height = 50, 50
            left = x - width // 2
            top = y - height // 2
        elif 'button' in filename:
            width, height = 80, 35
            left = x - width // 2
            top = y - height // 2
        else:
            # Script names - capture from cursor position to the right
            # This captures the distinctive part (QuantPC vs TradePC)
            width, height = 280, 25
            left = x - 20  # Start slightly before cursor
            top = y - height // 2
        
        # Take screenshot of region
        screenshot = pyautogui.screenshot(region=(left, top, width, height))
        screenshot.save(filepath)
        
        print(f"  {Colors.GREEN}✓ Saved: {filepath}{Colors.RESET}")
        print(f"    Captured at: ({x}, {y}) - {width}x{height}")
        
        time.sleep(1)  # Brief pause before next one
    
    print()
    print(f"{Colors.GREEN}Template capture complete!{Colors.RESET}")
    print(f"Templates saved to: {templates_folder}")
    print()
    print("You can now run the script normally and it will use image recognition.")


def safe_click(x: int, y: int, double: bool = False):
    """Click at coordinates with screen bounds validation.
    
    Raises an error if coordinates are outside screen bounds or near corners
    (which would trigger PyAutoGUI's fail-safe).
    """
    import pyautogui
    
    # Get screen size
    screen_width, screen_height = pyautogui.size()
    
    # Define safe margins (avoid corners which trigger fail-safe)
    margin = 10
    
    # Check if coordinates are within safe bounds
    if x < margin or x > screen_width - margin:
        raise ValueError(
            f"Click X coordinate ({x}) is outside safe screen bounds. "
            f"Screen width: {screen_width}. "
            f"This may indicate window position has changed or screen resolution changed. "
            f"Try running --calibrate to recalibrate click positions."
        )
    
    if y < margin or y > screen_height - margin:
        raise ValueError(
            f"Click Y coordinate ({y}) is outside safe screen bounds. "
            f"Screen height: {screen_height}. "
            f"This may indicate window position has changed or screen resolution changed. "
            f"Try running --calibrate to recalibrate click positions."
        )
    
    # Perform the click
    if double:
        pyautogui.doubleClick(x, y)
    else:
        pyautogui.click(x, y)


def launch_or_connect_qa(qa_path: str, timeout: int = 60):
    """Launch Quant Analyzer or connect to existing instance."""
    from pywinauto import Application, Desktop
    from pywinauto.findwindows import ElementNotFoundError
    
    # Verify path exists first
    if not os.path.exists(qa_path):
        raise FileNotFoundError(f"Quant Analyzer not found at: {qa_path}")
    
    # Get the directory containing the exe (needed as working directory)
    qa_dir = os.path.dirname(qa_path)
    
    print(f"  QA Path: {qa_path}")
    print(f"  Working Dir: {qa_dir}")
    
    # Helper to find QA window from all windows
    def find_qa_window():
        """Search all windows for Quant Analyzer."""
        desktop = Desktop(backend="uia")
        for win in desktop.windows():
            try:
                title = win.window_text()
                if title and "Quant Analyzer" in title and "Properties" not in title:
                    return win, title
            except Exception:
                pass
        return None, None
    
    # First, try to connect to an existing instance
    print(f"  Checking for running instance...")
    win, title = find_qa_window()
    if win:
        print(f"  Found running instance: {title}")
        # Connect to the process
        try:
            app = Application(backend="uia").connect(handle=win.handle)
            return app, False
        except Exception as e:
            print(f"  Could not connect: {e}")
    else:
        print(f"  No running instance found")
    
    # Launch new instance with correct working directory
    print(f"  Launching Quant Analyzer...")
    try:
        app = Application(backend="uia").start(qa_path, work_dir=qa_dir)
    except Exception as e:
        raise RuntimeError(f"Failed to start Quant Analyzer: {type(e).__name__}: {e}")
    
    # Wait for main window to appear
    print(f"  Waiting for application window (timeout: {timeout}s)...")
    start_time = time.time()
    while (time.time() - start_time) < timeout:
        win, title = find_qa_window()
        if win:
            print(f"  Found window: {title}")
            # Give it a moment to fully load
            time.sleep(3)
            # Return the app connected to this window
            try:
                app = Application(backend="uia").connect(handle=win.handle)
                return app, True
            except Exception as e:
                print(f"  Warning: Could not reconnect: {e}")
                return app, True
        time.sleep(1)
    
    # Debug: print all visible windows
    print(f"\n  {Colors.YELLOW}Timeout - listing all visible windows:{Colors.RESET}")
    desktop = Desktop(backend="uia")
    for win in desktop.windows():
        try:
            title = win.window_text()
            if title:
                print(f"    - {title}")
        except Exception:
            pass
    
    raise RuntimeError(f"Quant Analyzer window not found within {timeout} seconds")


def find_scripter_panel(main_window):
    """Navigate to the Scripter panel."""
    # Try different approaches to find and click Scripter
    
    # Approach 1: Look for Scripter text/button in left menu
    try:
        scripter_btn = main_window.child_window(title="Scripter", control_type="Text")
        if scripter_btn.exists():
            scripter_btn.click_input()
            time.sleep(1)
            return True
    except Exception:
        pass
    
    # Approach 2: Try as a button
    try:
        scripter_btn = main_window.child_window(title="Scripter", control_type="Button")
        if scripter_btn.exists():
            scripter_btn.click_input()
            time.sleep(1)
            return True
    except Exception:
        pass
    
    # Approach 3: Try as a menu item or list item
    try:
        scripter_btn = main_window.child_window(title_re=".*Scripter.*", control_type="ListItem")
        if scripter_btn.exists():
            scripter_btn.click_input()
            time.sleep(1)
            return True
    except Exception:
        pass
    
    # Approach 4: Try clicking by image/position (left side menu)
    # The Scripter icon appears to be near the bottom of the left menu
    try:
        # Look for any element containing "Scripter" text
        scripter_elements = main_window.descendants(title_re=".*Scripter.*")
        for elem in scripter_elements:
            try:
                elem.click_input()
                time.sleep(1)
                return True
            except Exception:
                continue
    except Exception:
        pass
    
    return False


def find_and_run_script(main_window, script_name: str):
    """Find the script in the Navigator and run it."""
    
    # Step 1: Find the script in the Navigator tree/list
    print(f"  Looking for script: {script_name}")
    
    script_item = None
    
    # Try different approaches to find the script
    # Approach 1: Direct title match
    try:
        script_item = main_window.child_window(title=script_name, control_type="ListItem")
        if not script_item.exists():
            script_item = None
    except Exception:
        pass
    
    # Approach 2: TreeItem
    if script_item is None:
        try:
            script_item = main_window.child_window(title=script_name, control_type="TreeItem")
            if not script_item.exists():
                script_item = None
        except Exception:
            pass
    
    # Approach 3: Text element
    if script_item is None:
        try:
            script_item = main_window.child_window(title=script_name, control_type="Text")
            if not script_item.exists():
                script_item = None
        except Exception:
            pass
    
    # Approach 4: Partial match
    if script_item is None:
        try:
            # Remove .java extension for partial match
            base_name = script_name.replace(".java", "")
            script_item = main_window.child_window(title_re=f".*{base_name}.*")
            if not script_item.exists():
                script_item = None
        except Exception:
            pass
    
    if script_item is None:
        raise Exception(f"Could not find script '{script_name}' in Navigator panel")
    
    # Double-click to open the script
    print(f"  Opening script...")
    script_item.double_click_input()
    time.sleep(1)
    
    # Step 2: Click the Run button
    print(f"  Clicking Run button...")
    
    run_button = None
    
    # Try to find the Run button
    try:
        run_button = main_window.child_window(title="Run", control_type="Button")
        if not run_button.exists():
            run_button = None
    except Exception:
        pass
    
    # Try with different control types
    if run_button is None:
        try:
            # Look for button with Run text
            buttons = main_window.descendants(control_type="Button")
            for btn in buttons:
                try:
                    if "Run" in btn.window_text():
                        run_button = btn
                        break
                except Exception:
                    continue
        except Exception:
            pass
    
    if run_button is None:
        raise Exception("Could not find Run button")
    
    run_button.click_input()
    print(f"  Script started!")
    return True


def wait_for_completion(main_window, timeout: int = 600):
    """Wait for the script to complete by monitoring the UI."""
    print(f"  Waiting for script completion (timeout: {timeout}s)...")
    
    start_time = time.time()
    last_status = ""
    
    while (time.time() - start_time) < timeout:
        # Check if Run button is enabled again (indicates completion)
        try:
            run_button = main_window.child_window(title="Run", control_type="Button")
            if run_button.exists() and run_button.is_enabled():
                # Also check if Stop button is disabled
                try:
                    stop_button = main_window.child_window(title="Stop", control_type="Button")
                    if stop_button.exists() and not stop_button.is_enabled():
                        # Script has finished
                        elapsed = time.time() - start_time
                        print(f"  Script completed in {format_duration(elapsed)}")
                        return True
                except Exception:
                    pass
        except Exception:
            pass
        
        # Check status bar for "Ready" or completion message
        try:
            status_elements = main_window.descendants(control_type="StatusBar")
            for status in status_elements:
                try:
                    text = status.window_text()
                    if text and text != last_status:
                        last_status = text
                        if "Ready" in text or "finished" in text.lower() or "completed" in text.lower():
                            elapsed = time.time() - start_time
                            print(f"  Script completed in {format_duration(elapsed)}")
                            return True
                except Exception:
                    pass
        except Exception:
            pass
        
        # Progress indicator
        elapsed = int(time.time() - start_time)
        if elapsed > 0 and elapsed % 30 == 0:
            print(f"    ... still running ({format_duration(elapsed)} elapsed)")
        
        time.sleep(2)
    
    raise TimeoutError(f"Script did not complete within {timeout} seconds")


def inspect_ui(qa_path: str):
    """Debug mode: Print UI element tree for inspection."""
    from pywinauto import Application, Desktop
    
    print(f"\n{Colors.CYAN}=== UI INSPECTION MODE ==={Colors.RESET}\n")
    
    app, launched = launch_or_connect_qa(qa_path)
    
    # Find the main window
    desktop = Desktop(backend="uia")
    main_window = None
    for win in desktop.windows():
        try:
            title = win.window_text()
            if title and "Quant Analyzer" in title and "Properties" not in title:
                main_window = win
                break
        except Exception:
            pass
    
    if not main_window:
        print(f"{Colors.RED}Could not find Quant Analyzer window{Colors.RESET}")
        return
    
    print(f"\n{Colors.CYAN}Main Window:{Colors.RESET}")
    print(f"  Title: {main_window.window_text()}")
    
    # Get window class and position
    is_java_app = False
    try:
        class_name = main_window.class_name()
        print(f"  Class: {class_name}")
        is_java_app = "Sun" in class_name or "Java" in class_name
    except:
        print(f"  Class: (unknown)")
    
    # Get window rectangle
    try:
        rect = main_window.rectangle()
        print(f"  Position: Left={rect.left}, Top={rect.top}, Right={rect.right}, Bottom={rect.bottom}")
        print(f"  Size: {rect.right - rect.left} x {rect.bottom - rect.top}")
    except Exception as e:
        print(f"  Position: (could not get: {e})")
    
    if is_java_app:
        print(f"\n{Colors.YELLOW}*** JAVA APPLICATION DETECTED ***{Colors.RESET}")
        print(f"  Java apps don't expose UI controls to Windows automation.")
        print(f"  This script will use mouse/keyboard automation instead.")
        print(f"\n{Colors.CYAN}To use mouse automation, we need pyautogui:{Colors.RESET}")
        if check_pyautogui():
            print(f"  {Colors.GREEN}pyautogui is installed - good!{Colors.RESET}")
        else:
            print(f"  {Colors.RED}pyautogui is NOT installed{Colors.RESET}")
            print(f"  Install it with: pip install pyautogui")
    else:
        # Try to print control tree for non-Java apps
        print(f"\n{Colors.CYAN}Control Tree (first 3 levels):{Colors.RESET}")
        try:
            main_window.print_control_identifiers(depth=3)
        except Exception as e:
            print(f"  Error printing tree: {e}")
        
        print(f"\n{Colors.YELLOW}Looking for key elements...{Colors.RESET}")
        
        # Look for Scripter
        print(f"\n  Scripter-related elements:")
        try:
            all_elements = main_window.descendants()
            scripter_items = [e for e in all_elements if 'scripter' in e.window_text().lower()]
            for item in scripter_items[:10]:
                try:
                    print(f"    - '{item.window_text()}' ({item.element_info.control_type})")
                except Exception:
                    pass
            if not scripter_items:
                print(f"    (none found)")
        except Exception as e:
            print(f"    Error: {e}")
        
        # Look for Run button
        print(f"\n  Run/Stop buttons:")
        try:
            buttons = main_window.descendants(control_type="Button")
            for btn in buttons:
                try:
                    text = btn.window_text()
                    if text and ("Run" in text or "Stop" in text):
                        print(f"    - '{text}' (Button, enabled={btn.is_enabled()})")
                except Exception:
                    pass
        except Exception as e:
            print(f"    Error: {e}")
        
        # Look for Navigator/script list
        print(f"\n  List/Tree items (potential script list):")
        try:
            items = main_window.descendants(control_type="ListItem")
            items += main_window.descendants(control_type="TreeItem")
            for item in items[:20]:
                try:
                    text = item.window_text()
                    if text and ".java" in text:
                        print(f"    - '{text}' ({item.element_info.control_type})")
                except Exception:
                    pass
        except Exception as e:
            print(f"    Error: {e}")
    
    print(f"\n{Colors.GREEN}Inspection complete.{Colors.RESET}")
    print(f"{Colors.YELLOW}Press Enter to close Quant Analyzer (or Ctrl+C to keep it open)...{Colors.RESET}")
    
    try:
        input()
        if launched:
            main_window.close()
    except KeyboardInterrupt:
        print("\nKeeping Quant Analyzer open.")


def calibrate_positions(qa_path: str):
    """Calibration mode: show mouse coordinates relative to QA window."""
    from pywinauto import Desktop
    import pyautogui
    
    print(f"\n{Colors.CYAN}=== CALIBRATION MODE ==={Colors.RESET}\n")
    print("This mode helps you find the correct click positions for UI elements.")
    print()
    
    app, launched = launch_or_connect_qa(qa_path)
    
    # Find the main window
    desktop = Desktop(backend="uia")
    main_window = None
    for win in desktop.windows():
        try:
            title = win.window_text()
            if title and "Quant Analyzer" in title and "Properties" not in title:
                main_window = win
                break
        except Exception:
            pass
    
    if not main_window:
        print(f"{Colors.RED}Could not find Quant Analyzer window{Colors.RESET}")
        return
    
    # Get window position first
    rect = main_window.rectangle()
    win_left = rect.left
    win_top = rect.top
    
    # Maximize the window for consistent calibration (click to focus, then use keyboard)
    print("Maximizing window...")
    try:
        # Click on the window to ensure it has focus (click near top-center of window)
        click_x = rect.left + (rect.right - rect.left) // 2
        click_y = rect.top + 15  # Title bar area
        pyautogui.click(click_x, click_y)
        time.sleep(0.3)
        
        # Now send Win+Up to maximize
        pyautogui.hotkey('win', 'up')
        time.sleep(1)
    except Exception as e:
        print(f"{Colors.YELLOW}Warning: Could not maximize window: {e}{Colors.RESET}")
    
    # Get window position AFTER maximizing
    rect = main_window.rectangle()
    win_left = rect.left
    win_top = rect.top
    
    # Get screen size
    screen_width, screen_height = pyautogui.size()
    
    print(f"Window found: {main_window.window_text()}")
    print(f"Window position: Left={win_left}, Top={win_top}")
    print(f"Window size: {rect.right - rect.left} x {rect.bottom - rect.top}")
    print(f"Screen size: {screen_width} x {screen_height}")
    print()
    print(f"{Colors.YELLOW}NOTE: Window has been maximized. Calibrate with maximized window.{Colors.RESET}")
    print()
    print(f"{Colors.YELLOW}Move your mouse to each UI element and note the OFFSET values:{Colors.RESET}")
    print(f"  1. Scripter icon (left menu, near bottom)")
    print(f"  2. Run button (green button in Scripter panel header)")
    print(f"  3. FIRST script name in Navigator (BatchMC_Analysis_QuantPC.java)")
    print(f"  4. SECOND script name to calculate row height")
    print()
    print(f"Press Ctrl+C to exit calibration mode.")
    print()
    print(f"{'Mouse X':>10} {'Mouse Y':>10} | {'X Offset':>10} {'Y Offset':>10}")
    print("-" * 50)
    
    try:
        while True:
            x, y = pyautogui.position()
            x_offset = x - win_left
            y_offset = y - win_top
            print(f"{x:>10} {y:>10} | {x_offset:>10} {y_offset:>10}", end='\r')
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("\n")
        print(f"{Colors.GREEN}Calibration complete.{Colors.RESET}")
        print()
        print(f"Update these values in the script:")
        print(f"  SCRIPTER_CLICK_X_OFFSET = <X offset for Scripter icon in left menu>")
        print(f"  SCRIPTER_CLICK_Y_OFFSET = <Y offset for Scripter icon>")
        print(f"  RUN_BUTTON_X_OFFSET = <X offset for Run button>")
        print(f"  RUN_BUTTON_Y_OFFSET = <Y offset for Run button>")
        print(f"  NAVIGATOR_X_OFFSET = <X offset for middle of script names in Navigator>")
        print(f"  NAVIGATOR_Y_OFFSET = <Y offset for FIRST script (BatchMC_Analysis_QuantPC.java)>")
        print(f"  SCRIPT_ROW_HEIGHT = <pixels between script rows, usually ~23>")


def run_qa_script(script_name: str, qa_path: str, keep_open: bool, timeout: int, output_folder: str = None, use_images: bool = False):
    """Main function to run a script in Quant Analyzer.
    
    Args:
        script_name: Name of the Java script to run
        qa_path: Path to QuantAnalyzer4.exe
        keep_open: Whether to leave QA open after script completes
        timeout: Maximum time to wait for script completion
        output_folder: Folder where completion marker will be written
        use_images: If True, use image recognition for clicking (more reliable across resolutions)
    """
    from pywinauto import Desktop
    import pyautogui
    
    print(f"Script to run: {script_name}")
    print(f"Quant Analyzer: {qa_path}")
    if output_folder:
        print(f"Output folder: {output_folder}")
    if use_images:
        print(f"Mode: Image recognition")
    else:
        print(f"Mode: Coordinate-based")
    print()
    
    # Configure pyautogui for unattended automation
    # FAILSAFE disabled because RDP disconnect causes false triggers
    # This is safe for unattended automation on your own system
    pyautogui.FAILSAFE = False
    pyautogui.PAUSE = 0.5  # Pause between actions
    
    # Step 1: Launch or connect to QA
    print("Step 1: Launching Quant Analyzer...")
    app, launched = launch_or_connect_qa(qa_path)
    
    # Find the main window
    desktop = Desktop(backend="uia")
    main_window = None
    for win in desktop.windows():
        try:
            title = win.window_text()
            if title and "Quant Analyzer" in title and "Properties" not in title:
                main_window = win
                break
        except Exception:
            pass
    
    if not main_window:
        raise RuntimeError("Could not find Quant Analyzer window")
    
    # Bring QA to front and maximize using window handle (not screen coordinates)
    # This avoids accidentally focusing/maximizing other windows (e.g. browser)
    print("  Bringing Quant Analyzer to front and maximizing...")
    try:
        hwnd = main_window.handle
        # SetForegroundWindow + ShowWindow via Win32 API for reliable focus/maximize
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        time.sleep(0.3)
        ctypes.windll.user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
        time.sleep(1)  # Wait for window to settle
    except Exception as e:
        print(f"  {Colors.YELLOW}Warning: Could not maximize window: {e}{Colors.RESET}")
        print(f"  {Colors.YELLOW}Continuing anyway - image recognition doesn't require maximized window{Colors.RESET}")
    
    # Get window position and size AFTER maximizing (or after failed attempt)
    rect = main_window.rectangle()
    win_left = rect.left
    win_top = rect.top
    win_width = rect.right - rect.left
    win_height = rect.bottom - rect.top
    
    # Get screen size for validation
    screen_width, screen_height = pyautogui.size()
    
    print(f"  Window position: ({win_left}, {win_top})")
    print(f"  Window size: {win_width} x {win_height}")
    print(f"  Screen size: {screen_width} x {screen_height}")
    
    # Use image recognition or coordinate-based clicking
    if use_images:
        _run_script_with_images(script_name, main_window, keep_open, timeout, output_folder, launched)
    else:
        _run_script_with_coordinates(script_name, main_window, keep_open, timeout, output_folder, launched,
                                      win_left, win_top, screen_width, screen_height)


def _run_script_with_images(script_name: str, main_window, keep_open: bool, timeout: int, output_folder: str, launched: bool):
    """Run script using image recognition for clicking."""
    import pyautogui
    
    # Check that templates exist
    templates_folder = get_templates_folder()
    if not os.path.exists(templates_folder):
        raise FileNotFoundError(
            f"Templates folder not found: {templates_folder}\n"
            f"Run with --capture-templates first to create template images."
        )
    
    # Determine which script template to use
    script_lower = script_name.lower()
    if 'minute' in script_lower or 'm1' in script_lower:
        script_template = "script_batchmc_minute_data.png"
    elif 'tick' in script_lower:
        script_template = "script_batchmc_tick_data.png"
    else:
        raise ValueError(f"No template defined for script: {script_name}. Expected 'minute' or 'tick' in name.")
    
    try:
        # Bring window to front
        ctypes.windll.user32.SetForegroundWindow(main_window.handle)
        time.sleep(0.5)
        
        # Step 2: Click on Scripter in the left menu
        print("\nStep 2: Clicking on Scripter panel...")
        if not find_and_click("scripter_icon.png", confidence=0.8, timeout=10):
            raise RuntimeError(
                "Could not find Scripter icon on screen.\n"
                "This usually means RDP is disconnected and Windows can't capture screenshots.\n"
                "Solutions:\n"
                "  1. Keep RDP window MINIMIZED instead of closing it\n"
                "  2. Use 'tscon' to disconnect without locking the session\n"
                "  3. Install TightVNC for persistent desktop access"
            )
        time.sleep(1)
        
        # Step 3: Double-click on the script in the Navigator panel
        print(f"\nStep 3: Selecting script '{script_name}' from Navigator...")
        # Use higher confidence for scripts since they look similar
        if not find_and_click(script_template, confidence=0.9, timeout=10, double=True):
            raise RuntimeError(f"Could not find script {script_name} on screen")
        time.sleep(1)
        
        # Step 4: Click the Run button
        print(f"\nStep 4: Clicking Run button...")
        if not find_and_click("run_button.png", confidence=0.8, timeout=10):
            raise RuntimeError("Could not find Run button on screen")
        time.sleep(1)
        
        # Step 5: Wait for completion
        print(f"\nStep 5: Waiting for script completion (timeout: {timeout}s)...")
        
        # Monitor for completion marker file
        marker_file = os.path.join(output_folder, "BatchMC_Complete.txt") if output_folder else None
        
        if marker_file:
            # Delete marker if it exists (Java script should do this, but safety check)
            if os.path.exists(marker_file):
                try:
                    os.remove(marker_file)
                    print(f"  Removed existing completion marker")
                except Exception:
                    pass
            
            print(f"  Waiting for completion marker: {marker_file}")
            start_time = time.time()
            
            while (time.time() - start_time) < timeout:
                if os.path.exists(marker_file):
                    print(f"  {Colors.GREEN}Completion marker found - script completed!{Colors.RESET}")
                    # Read and display marker contents
                    try:
                        with open(marker_file, 'r') as f:
                            contents = f.read().strip()
                            for line in contents.split('\n'):
                                print(f"    {line}")
                    except Exception:
                        pass
                    break
                
                time.sleep(2)
            else:
                print(f"  {Colors.YELLOW}Warning: Timeout waiting for completion marker{Colors.RESET}")
        else:
            print(f"  No output folder specified - waiting {timeout}s...")
            time.sleep(timeout)
        
        print(f"\n{Colors.GREEN}Script execution complete!{Colors.RESET}")
        
    finally:
        # Step 6: Close QA if requested
        if not keep_open:
            print("\nStep 6: Closing Quant Analyzer...")
            try:
                ctypes.windll.user32.SetForegroundWindow(main_window.handle)
                pyautogui.hotkey('alt', 'F4')
                time.sleep(1)
                pyautogui.press('enter')  # Confirm exit dialog
                time.sleep(2)
                print(f"  {Colors.GREEN}Closed{Colors.RESET}")
            except Exception as e:
                print(f"  {Colors.YELLOW}Warning: Could not close cleanly: {e}{Colors.RESET}")
        else:
            print("\nStep 6: Keeping Quant Analyzer open (--keep-open specified)")


def _run_script_with_coordinates(script_name: str, main_window, keep_open: bool, timeout: int, 
                                   output_folder: str, launched: bool,
                                   win_left: int, win_top: int, screen_width: int, screen_height: int):
    """Run script using coordinate-based clicking (original method)."""
    import pyautogui
    
    # Validate that expected click positions will be within screen bounds
    max_y_needed = win_top + SCRIPTER_CLICK_Y_OFFSET
    max_x_needed = win_left + NAVIGATOR_X_OFFSET
    
    if max_y_needed > screen_height - 10:
        raise RuntimeError(
            f"Window position would cause clicks outside screen bounds.\n"
            f"  Expected Scripter Y: {max_y_needed}, Screen height: {screen_height}\n"
            f"  This may happen if the window was restored at a different position after RDP disconnect.\n"
            f"  Please maximize the Quant Analyzer window or move it to top-left corner and try again."
        )
    
    if max_x_needed > screen_width - 10:
        raise RuntimeError(
            f"Window position would cause clicks outside screen bounds.\n"
            f"  Expected Navigator X: {max_x_needed}, Screen width: {screen_width}\n"
            f"  This may happen if the window was restored at a different position after RDP disconnect.\n"
            f"  Please maximize the Quant Analyzer window or move it to top-left corner and try again."
        )
    
    try:
        # Bring window to front
        ctypes.windll.user32.SetForegroundWindow(main_window.handle)
        time.sleep(0.5)
        
        # Step 2: Click on Scripter in the left menu
        print("\nStep 2: Clicking on Scripter panel...")
        scripter_x = win_left + SCRIPTER_CLICK_X_OFFSET
        scripter_y = win_top + SCRIPTER_CLICK_Y_OFFSET
        print(f"  Clicking at ({scripter_x}, {scripter_y})")
        safe_click(scripter_x, scripter_y)
        time.sleep(1)
        
        # Step 3: Double-click on the script in the Navigator panel
        print(f"\nStep 3: Selecting script '{script_name}' from Navigator...")
        
        # Scripts list with their row positions (0-indexed from first script)
        script_list = [
            "BatchMC_Minute_Data.java",
            "BatchMC_Tick_Data.java",
            "EquityControlSample.java",
            "EquityControlSample2.java",
            "LoadSQReportSample.java",
            "LogConsoleTestScript.java",
            "MoneyManagementSimSample.java",
            "MoneyManagementSimSample2.java",
            "MonteCarloSample.java",
            "MonteCarloSample2.java",
            "SampleScriptMC.java",
        ]
        
        # Find script position (or use partial match)
        script_position = -1
        for i, name in enumerate(script_list):
            if script_name.lower() in name.lower() or name.lower() in script_name.lower():
                script_position = i
                break
        
        if script_position < 0:
            print(f"  {Colors.YELLOW}Warning: Script not in known list, using position 0{Colors.RESET}")
            script_position = 0
        else:
            print(f"  Script found at position {script_position + 1} in list")
        
        # Calculate Y position for the script
        # Each row is approximately SCRIPT_ROW_HEIGHT pixels apart
        script_x = win_left + NAVIGATOR_X_OFFSET
        script_y = win_top + NAVIGATOR_Y_OFFSET + (script_position * SCRIPT_ROW_HEIGHT)
        
        print(f"  Double-clicking at ({script_x}, {script_y})")
        safe_click(script_x, script_y, double=True)
        time.sleep(1)
        
        # Step 4: Click the Run button
        print(f"\nStep 4: Clicking Run button...")
        run_x = win_left + RUN_BUTTON_X_OFFSET
        run_y = win_top + RUN_BUTTON_Y_OFFSET
        print(f"  Clicking Run at ({run_x}, {run_y})")
        safe_click(run_x, run_y)
        time.sleep(1)
        
        # Step 5: Wait for completion
        print(f"\nStep 5: Waiting for script completion (timeout: {timeout}s)...")
        
        # Monitor for completion marker file (created by Java script when done)
        marker_file = os.path.join(output_folder, "BatchMC_Complete.txt") if output_folder else None
        
        if marker_file:
            print(f"  Monitoring for completion marker: {marker_file}")
            
            # Delete marker file if it exists (should have been deleted by Java, but just in case)
            if os.path.exists(marker_file):
                try:
                    os.remove(marker_file)
                    print(f"  Removed existing marker file")
                except Exception as e:
                    print(f"  {Colors.YELLOW}Warning: Could not remove existing marker: {e}{Colors.RESET}")
            
            start_time = time.time()
            last_report = 0
            
            while (time.time() - start_time) < timeout:
                elapsed = int(time.time() - start_time)
                
                # Check if completion marker file was created
                if os.path.exists(marker_file):
                    print(f"\n  {Colors.GREEN}Completion marker found - script completed!{Colors.RESET}")
                    print(f"  Completed in {format_duration(elapsed)}")
                    break
                
                # Progress report every 30 seconds
                if elapsed > 0 and elapsed % 30 == 0 and elapsed != last_report:
                    print(f"    ... still running ({format_duration(elapsed)} elapsed)")
                    last_report = elapsed
                
                time.sleep(2)
            else:
                print(f"\n{Colors.YELLOW}Timeout reached. Script may still be running.{Colors.RESET}")
        else:
            # No output folder specified - just wait with timeout
            print(f"  No output folder specified - using timeout only")
            start_time = time.time()
            while (time.time() - start_time) < timeout:
                elapsed = int(time.time() - start_time)
                if elapsed > 0 and elapsed % 30 == 0:
                    print(f"    ... waiting ({format_duration(elapsed)} elapsed)")
                time.sleep(5)
            print(f"\n{Colors.YELLOW}Timeout reached.{Colors.RESET}")
        
        print(f"\n{Colors.GREEN}Script execution completed successfully!{Colors.RESET}")
        
    except Exception as e:
        print(f"\n{Colors.RED}Error: {e}{Colors.RESET}")
        raise
    
    finally:
        # Step 6: Close QA unless --keep-open was specified
        if not keep_open:
            print(f"\nStep 6: Closing Quant Analyzer...")
            try:
                # Bring window to front first
                ctypes.windll.user32.SetForegroundWindow(main_window.handle)
                time.sleep(0.3)
                
                # Try Alt+F4 to close (works better for Java apps)
                pyautogui.hotkey('alt', 'F4')
                time.sleep(1)
                
                # Handle "Are you sure you want to exit?" confirmation dialog
                # Yes button is already selected, just press Enter
                pyautogui.press('enter')
                time.sleep(2)
                
                # Check if it closed
                still_open = False
                for win in desktop.windows():
                    try:
                        title = win.window_text()
                        if title and "Quant Analyzer" in title and "Properties" not in title:
                            still_open = True
                            break
                    except:
                        pass
                
                if still_open:
                    print(f"  {Colors.YELLOW}Window still open, trying again...{Colors.RESET}")
                    pyautogui.hotkey('alt', 'F4')
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(1)
                
                print(f"  Quant Analyzer closed.")
            except Exception as e:
                print(f"  {Colors.YELLOW}Warning: Could not close cleanly: {e}{Colors.RESET}")
        else:
            print(f"\nKeeping Quant Analyzer open as requested.")


def main():
    init_colors()
    
    parser = argparse.ArgumentParser(
        description='Run Java scripts in Quant Analyzer Pro via UI automation',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Run script with coordinate-based clicking (requires calibration)
  %(prog)s "BatchMC_Analysis_TradePC.java" --output-folder "E:\\Trading\\Analysis_Ouput\\TradePC"
  
  # Run script with image recognition (more reliable, resolution-independent)
  %(prog)s "BatchMC_Analysis_TradePC.java" --use-images --output-folder "E:\\Trading\\Analysis_Ouput\\TradePC"
  
  # Capture template images for image recognition mode (one-time setup)
  %(prog)s --capture-templates
  
  # Calibrate coordinate offsets (for coordinate mode)
  %(prog)s --calibrate
  
  # Inspect UI elements (debugging)
  %(prog)s --inspect
        """
    )
    parser.add_argument(
        'script_name',
        nargs='?',
        help='Name of the .java script to run (e.g., BatchMC_Analysis_TradePC.java)'
    )
    parser.add_argument(
        '--qa-path',
        default=DEFAULT_QA_PATH,
        help=f'Path to QuantAnalyzer4.exe (default: {DEFAULT_QA_PATH})'
    )
    parser.add_argument(
        '--keep-open',
        action='store_true',
        help='Keep Quant Analyzer open after script completes'
    )
    parser.add_argument(
        '--timeout',
        type=int,
        default=DEFAULT_TIMEOUT,
        help=f'Max seconds to wait for script completion (default: {DEFAULT_TIMEOUT})'
    )
    parser.add_argument(
        '--inspect',
        action='store_true',
        help='Debug mode: launch QA and print UI element tree'
    )
    parser.add_argument(
        '--calibrate',
        action='store_true',
        help='Calibration mode: show mouse coordinates to find UI element positions'
    )
    parser.add_argument(
        '--capture-templates',
        action='store_true',
        help='Capture template images for image recognition mode'
    )
    parser.add_argument(
        '--use-images',
        action='store_true',
        help='Use image recognition instead of coordinate-based clicking (more reliable)'
    )
    parser.add_argument(
        '--output-folder',
        help='Folder where BatchMC_Results.csv will be created (for completion detection)'
    )
    
    args = parser.parse_args()
    
    # Check pywinauto is installed
    if not check_pywinauto():
        print(f"{Colors.RED}Error: pywinauto is not installed.{Colors.RESET}")
        print(f"\nInstall it with:")
        print(f"  pip install pywinauto")
        sys.exit(1)
    
    # Check pyautogui if running a script or calibrating (needed for Java app mouse automation)
    if (args.script_name or args.calibrate or args.capture_templates or args.use_images) and not check_pyautogui():
        print(f"{Colors.RED}Error: pyautogui is not installed.{Colors.RESET}")
        print(f"\nQuant Analyzer is a Java app that requires mouse/keyboard automation.")
        print(f"Install pyautogui with:")
        print(f"  pip install pyautogui")
        sys.exit(1)
    
    # Check opencv for image recognition
    if args.use_images and not check_opencv():
        print(f"{Colors.RED}Error: opencv-python is not installed.{Colors.RESET}")
        print(f"\nImage recognition mode requires OpenCV.")
        print(f"Install it with:")
        print(f"  pip install opencv-python")
        sys.exit(1)
    
    # Record start time
    start_time = time.time()
    start_datetime = datetime.now()
    
    print(f"{Colors.CYAN}{'=' * 60}{Colors.RESET}")
    print(f"{Colors.CYAN}Quant Analyzer Script Runner{Colors.RESET}")
    print(f"{Colors.CYAN}{'=' * 60}{Colors.RESET}")
    print()
    print(f"{Colors.CYAN}Started:{Colors.RESET} {Colors.GREEN}{start_datetime.strftime('%Y-%m-%d %H:%M:%S')}{Colors.RESET}")
    print()
    
    try:
        if args.inspect:
            inspect_ui(args.qa_path)
        elif args.calibrate:
            calibrate_positions(args.qa_path)
        elif args.capture_templates:
            capture_templates(args.qa_path)
        elif args.script_name:
            run_qa_script(args.script_name, args.qa_path, args.keep_open, args.timeout, args.output_folder, args.use_images)
        else:
            parser.print_help()
            print(f"\n{Colors.RED}Error: Please specify a script name, --inspect, --calibrate, or --capture-templates{Colors.RESET}")
            sys.exit(1)
        
        exit_code = 0
        
    except Exception as e:
        import traceback
        print(f"\n{Colors.RED}Failed: {type(e).__name__}: {e}{Colors.RESET}")
        print(f"\n{Colors.GRAY}Traceback:{Colors.RESET}")
        traceback.print_exc()
        exit_code = 1
    
    # Record end time
    end_datetime = datetime.now()
    total_duration = time.time() - start_time
    
    print()
    print(f"{Colors.CYAN}{'=' * 60}{Colors.RESET}")
    print(f"{Colors.CYAN}Finished:{Colors.RESET}       {Colors.GREEN}{end_datetime.strftime('%Y-%m-%d %H:%M:%S')}{Colors.RESET}")
    print(f"{Colors.CYAN}Total Duration:{Colors.RESET} {Colors.GREEN}{format_duration(total_duration)}{Colors.RESET}")
    
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
