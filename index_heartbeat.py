"""
Modern Scheduler Application
Built with CustomTkinter + APScheduler
Manages and executes .exe files with dedicated logging
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import json
import os
import sys
import subprocess
import threading
import queue
from datetime import datetime
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger
import psutil
import time



# Set appearance
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class TaskManager:
    """Manages task persistence and operations"""
    
    def __init__(self, filename="tasks.json", config_filename="config.json"):
        # Thread safety lock for file operations
        self._lock = threading.RLock()
        
        # Get the directory where the app is running (support both .py and .exe)
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            app_dir = os.path.dirname(sys.executable)
        else:
            # Running as Python script
            app_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Use absolute paths for data files
        self.filename = os.path.join(app_dir, filename)
        self.config_filename = os.path.join(app_dir, config_filename)
        
        self.tasks = self.load_tasks()
        self.config = self.load_config()
        
        # Cleanup orphaned processes from previous crashed sessions
        self.cleanup_orphaned_processes()
    
    def load_tasks(self):
        """Load tasks from JSON file with error handling"""
        if os.path.exists(self.filename):
            try:
                with open(self.filename, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError) as e:
                print(f"Error loading tasks: {e}")
                # Backup corrupted file
                if os.path.exists(self.filename):
                    backup_name = f"{self.filename}.backup"
                    try:
                        os.rename(self.filename, backup_name)
                    except:
                        pass
                return []
        return []
    
    def load_config(self):
        """Load configuration from JSON file with error handling"""
        if os.path.exists(self.config_filename):
            try:
                with open(self.config_filename, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError) as e:
                print(f"Error loading config: {e}")
                return {"last_exe_path": None}
        return {"last_exe_path": None}
    
    def cleanup_orphaned_processes(self):
        """Check for orphaned processes from previous crashed sessions - returns list if found"""
        print("[CLEANUP] Checking for crashed/force-closed sessions...")
        
        # Create path for crash marker file
        crash_marker_file = f"{self.filename}.running"
        
        orphaned_tasks = []
        
        # If marker file exists, it means the app was force-closed or crashed
        if os.path.exists(crash_marker_file):
            print(f"[CLEANUP] Found crash marker file - app was not properly closed!")
            
            # All current tasks are potentially orphaned
            # (they were being run when app crashed)
            orphaned_tasks = self.tasks.copy()
            
            if orphaned_tasks:
                print(f"[CLEANUP] Found {len(orphaned_tasks)} task(s) from crashed session")
                for task in orphaned_tasks:
                    print(f"  - {task['name']}")
            
            # Remove the crash marker file
            try:
                os.remove(crash_marker_file)
                print("[CLEANUP] Removed crash marker file")
            except:
                pass
        else:
            print("[CLEANUP] No crash detected - normal startup")
        
        return orphaned_tasks
    
    def create_crash_marker(self):
        """Create a marker file to detect crashes/force-closes"""
        crash_marker_file = f"{self.filename}.running"
        try:
            with open(crash_marker_file, 'w') as f:
                f.write(f"App started at {datetime.now().isoformat()}\n")
            print(f"[CRASH_MARKER] Created marker file: {crash_marker_file}")
        except Exception as e:
            print(f"[CRASH_MARKER] Error creating marker: {e}")
    
    def remove_crash_marker(self):
        """Remove the crash marker file on clean shutdown"""
        crash_marker_file = f"{self.filename}.running"
        try:
            if os.path.exists(crash_marker_file):
                os.remove(crash_marker_file)
                print(f"[CRASH_MARKER] Removed marker file on clean shutdown")
        except Exception as e:
            print(f"[CRASH_MARKER] Error removing marker: {e}")
    
    def save_config(self):
        """Save configuration to JSON file with error handling"""
        with self._lock:
            try:
                with open(self.config_filename, 'w', encoding='utf-8') as f:
                    json.dump(self.config, f, indent=4)
            except IOError as e:
                print(f"Error saving config: {e}")
    
    def set_last_exe_path(self, path):
        """Save the last selected exe path"""
        self.config["last_exe_path"] = path
        self.save_config()
    
    def get_last_exe_path(self):
        """Get the last selected exe path"""
        return self.config.get("last_exe_path")
    
    def save_tasks(self):
        """Save tasks to JSON file with error handling and atomic write"""
        with self._lock:
            temp_filename = f"{self.filename}.tmp"
            try:
                # Write to temporary file first
                with open(temp_filename, 'w', encoding='utf-8') as f:
                    json.dump(self.tasks, f, indent=4)
                
                # Atomic rename (Windows safe)
                if os.path.exists(self.filename):
                    backup_filename = f"{self.filename}.bak"
                    try:
                        if os.path.exists(backup_filename):
                            os.remove(backup_filename)
                        os.rename(self.filename, backup_filename)
                    except:
                        pass
                
                os.rename(temp_filename, self.filename)
            except IOError as e:
                print(f"Error saving tasks: {e}")
                # Cleanup temp file if it exists
                if os.path.exists(temp_filename):
                    try:
                        os.remove(temp_filename)
                    except:
                        pass
    
    def add_task(self, name, path, interval):
        """Add a new task with safe ID generation"""
        with self._lock:
            # Generate safe ID (find max existing ID + 1)
            max_id = 0
            for task in self.tasks:
                if task.get('id', 0) > max_id:
                    max_id = task['id']
            
            task = {
                "id": max_id + 1,
                "name": name,
                "path": path,
                "interval": interval,
                "status": "Idle",
                "last_run": None
            }
            self.tasks.append(task)
            self.save_tasks()
            return task
    
    def update_task(self, task_id, name, path, interval):
        """Update existing task"""
        with self._lock:
            for task in self.tasks:
                if task["id"] == task_id:
                    task["name"] = name
                    task["path"] = path
                    task["interval"] = interval
                    self.save_tasks()
                    return task
            return None
    
    def delete_task(self, task_id):
        """Delete a task"""
        with self._lock:
            self.tasks = [t for t in self.tasks if t["id"] != task_id]
            self.save_tasks()
    
    def update_status(self, task_id, status, last_run=None):
        """Update task status"""
        with self._lock:
            for task in self.tasks:
                if task["id"] == task_id:
                    task["status"] = status
                    if last_run:
                        task["last_run"] = last_run
                    self.save_tasks()
                    break


class ProcessExecutor:
    """Handles process execution - lightweight mode"""
    
    def __init__(self):
        self.running_processes = {}  # {exe_path: process_object}
    
    def is_running(self, exe_path):
        """Check if a process is already running"""
        exe_path = os.path.normpath(exe_path)
        
        if exe_path in self.running_processes:
            proc = self.running_processes[exe_path]
            if proc.poll() is None:  # Still running
                return True
            else:
                del self.running_processes[exe_path]
        return False
    
    def is_console_app(self, exe_path):
        """Detect if exe is a console application (needs log capture)"""
        if not os.path.exists(exe_path):
            print(f"Executable not found: {exe_path}")
            return False
            
        try:
            # Check if it's a .bat, .cmd, or Python script
            ext = os.path.splitext(exe_path)[1].lower()
            if ext in ['.bat', '.cmd', '.py']:
                return True
            
            # For .exe files, check PE subsystem (Windows specific)
            if ext == '.exe':
                import struct
                with open(exe_path, 'rb') as f:
                    # Read DOS header
                    dos_header = f.read(64)
                    if len(dos_header) < 64 or dos_header[:2] != b'MZ':
                        return False
                    
                    # Get PE header offset
                    pe_offset = struct.unpack('<I', dos_header[60:64])[0]
                    f.seek(pe_offset)
                    
                    # Read PE signature and skip to subsystem field
                    pe_sig = f.read(4)
                    if pe_sig != b'PE\x00\x00':
                        return False
                    
                    # Skip COFF header (20 bytes) to optional header
                    f.read(20)
                    
                    # Read subsystem (offset 68 in optional header)
                    f.read(68)
                    subsystem = struct.unpack('<H', f.read(2))[0]
                    
                    # CUI (Console) = 3, GUI = 2
                    return subsystem == 3
        except (IOError, OSError, struct.error) as e:
            print(f"Error reading PE header from {exe_path}: {e}")
        
        return False  # Default to GUI (no log capture)
    
    def should_capture_output(self, exe_path):
        """Determine if we should capture output for this exe
        
        Always capture console app output, but use non-blocking approach for heavy apps
        """
        if not os.path.exists(exe_path):
            return False
        
        try:
            # If it's a GUI app, don't capture
            if not self.is_console_app(exe_path):
                return False
            
            # It's a console app - always capture, but use appropriate method
            return True
            
        except (IOError, OSError) as e:
            print(f"Error checking file size for {exe_path}: {e}")
            return False
    
    def is_resource_heavy(self, exe_path):
        """Check if exe is resource-intensive (>5MB) - needs special handling"""
        try:
            file_size = os.path.getsize(exe_path)
            # Large files are likely resource-intensive
            return file_size > 5 * 1024 * 1024
        except (IOError, OSError):
            return False
    
    
    def execute(self, exe_path, log_callback=None, needs_logging=None, completion_callback=None, process_ref_callback=None):
        """Execute an .exe file - captures output without blocking the exe
        
        For resource-heavy apps: uses non-blocking pipes to avoid performance impact
        For light console apps: uses standard buffered pipes
        """
        
        # Validate exe_path
        if not exe_path or not isinstance(exe_path, str):
            if log_callback:
                log_callback(f"[x] Invalid executable path\n")
            return None
            
        exe_path = os.path.normpath(exe_path)
        
        if not os.path.exists(exe_path):
            if log_callback:
                log_callback(f"[x] Executable not found: {exe_path}\n")
            return None
        
        if self.is_running(exe_path):
            if log_callback:
                log_callback(f"[!] Process already running, skipping execution\n")
            return "skipped"
        
        # Auto-detect if logging is needed
        if needs_logging is None:
            needs_logging = self.is_console_app(exe_path)
        
        # Check if this is a resource-heavy app
        is_heavy = self.is_resource_heavy(exe_path)
        
        try:
            if needs_logging:
                # Console app - capture output
                print(f"[EXECUTE] Running console app {os.path.basename(exe_path)}")
                
                # Standard pipe-based execution
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = 0  # SW_HIDE
                
                process = subprocess.Popen(
                    exe_path,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    stdin=subprocess.DEVNULL,
                    startupinfo=startupinfo,
                    text=True,
                    bufsize=1,
                    cwd=os.path.dirname(exe_path)
                )
                
                if log_callback:
                    # Start output reader thread (non-blocking)
                    threading.Thread(
                        target=self._stream_output,
                        args=(process.stdout, log_callback, "stdout"),
                        daemon=False
                    ).start()
                    
                    # Monitor completion
                    threading.Thread(
                        target=self._monitor_completion,
                        args=(process, exe_path, log_callback, completion_callback),
                        daemon=False
                    ).start()
                else:
                    # No logging, just monitor completion
                    threading.Thread(
                        target=self._monitor_completion_simple,
                        args=(process, exe_path, completion_callback),
                        daemon=True
                    ).start()
            else:
                # GUI app - run without capturing output
                print(f"[EXECUTE] Running GUI app {os.path.basename(exe_path)}")
                process = subprocess.Popen(
                    exe_path,
                    cwd=os.path.dirname(exe_path) or None
                )
                
                # Just track completion
                threading.Thread(
                    target=self._monitor_completion_simple,
                    args=(process, exe_path, completion_callback),
                    daemon=True
                ).start()
            
            self.running_processes[exe_path] = process
            
            # Send process reference back if callback provided
            if process_ref_callback:
                process_ref_callback(process)
            
            return process
            
        except (FileNotFoundError, OSError, PermissionError) as e:
            if log_callback:
                log_callback(f"[x] Error executing process: {str(e)}\n")
            print(f"Error executing {exe_path}: {e}")
            return None
    
    def _stream_output(self, pipe, log_callback, stream_name="stream"):
        """Stream output with aggressive reading to combat exe-side buffering
        
        The problem: Many console apps buffer output when stdout is redirected.
        Solution: Read aggressively in small chunks with minimal delay.
        """
        import select
        import sys
        
        try:
            buffer = ""
            
            # Use iterator approach which is more efficient than read()
            for line in iter(pipe.readline, ''):
                if line:
                    log_callback(line)
                else:
                    break
                    
        except Exception as e:
            print(f"Stream error: {e}")
        finally:
            try:
                pipe.close()
            except:
                pass
    
    def _monitor_completion(self, process, exe_path, log_callback, completion_callback=None):
        """Monitor process completion with logging"""
        try:
            process.wait()
            if exe_path in self.running_processes:
                del self.running_processes[exe_path]
            log_callback(f"\n[+] Process completed (Exit code: {process.returncode})\n")
        except Exception as e:
            print(f"Error monitoring completion for {exe_path}: {e}")
        finally:
            if completion_callback:
                try:
                    completion_callback()
                except Exception as e:
                    print(f"Error in completion callback: {e}")
    
    def _monitor_completion_simple(self, process, exe_path, completion_callback=None):
        """Monitor process completion without logging (lightweight)"""
        try:
            print(f"[MONITOR] Starting completion monitor for {os.path.basename(exe_path)}")
            process.wait()
            print(f"[MONITOR] Process {os.path.basename(exe_path)} completed with code {process.returncode}")
            if exe_path in self.running_processes:
                del self.running_processes[exe_path]
        except Exception as e:
            print(f"Error monitoring completion for {exe_path}: {e}")
        finally:
            if completion_callback:
                try:
                    print(f"[MONITOR] Calling completion callback for {os.path.basename(exe_path)}")
                    completion_callback()
                except Exception as e:
                    print(f"Error in completion callback: {e}")
    
    def force_cleanup(self, exe_path):
        """Force cleanup of a process from tracking (e.g., when manually terminated)"""
        exe_path = os.path.normpath(exe_path)
        if exe_path in self.running_processes:
            proc = self.running_processes[exe_path]
            # Ensure process is actually dead - kill entire process tree
            if proc.poll() is None:
                try:
                    import psutil
                    try:
                        parent = psutil.Process(proc.pid)
                        children = parent.children(recursive=True)
                        
                        # Terminate all children
                        for child in children:
                            try:
                                child.terminate()
                            except psutil.NoSuchProcess:
                                pass
                        
                        # Terminate parent
                        parent.terminate()
                        
                        # Wait for processes to die
                        gone, alive = psutil.wait_procs([parent] + children, timeout=1)
                        
                        # Force kill survivors
                        for p in alive:
                            try:
                                p.kill()
                            except psutil.NoSuchProcess:
                                pass
                    except psutil.NoSuchProcess:
                        pass  # Process already dead
                except ImportError:
                    # Fallback without psutil
                    proc.terminate()
                    try:
                        proc.wait(timeout=1)
                    except subprocess.TimeoutExpired:
                        proc.kill()
                except (OSError, PermissionError) as e:
                    print(f"Error terminating process: {e}")
            del self.running_processes[exe_path]


class VerticalLogContainer:
    """Vertical stacking log container - logs stack from top to bottom"""
    
    def __init__(self, parent):
        self.parent = parent
        self.log_panels = {}  # {task_name: log_panel_widget}
        
        # Main scrollable container for all logs
        self.scroll_container = ctk.CTkScrollableFrame(
            parent,
            fg_color="#0d0d0d",
            corner_radius=0
        )
        self.scroll_container.pack(fill="both", expand=True, padx=0, pady=0)
    
    def add(self, name):
        """Add a new log panel at the bottom"""
        if name in self.log_panels:
            return self.log_panels[name]
        
        # Create a panel frame
        panel = ctk.CTkFrame(
            self.scroll_container,
            fg_color="#1a1a1a",
            corner_radius=8
        )
        panel.pack(fill="x", padx=10, pady=5, side="top")
        
        self.log_panels[name] = panel
        return panel
    
    def delete(self, name):
        """Delete a log panel"""
        if name not in self.log_panels:
            return
        
        panel = self.log_panels[name]
        panel.destroy()
        del self.log_panels[name]
    
    def get(self, name):
        """Get a log panel by name"""
        return self.log_panels.get(name)
        if self.active_tab == name:
            if self.tabs:
                self.set(list(self.tabs.keys())[0])
            else:
                self.active_tab = None
    
    def tab(self, name):
        """Get tab content widget"""
        if name in self.tabs:
            return self.tabs[name]["content"]
        return None


class LogTab(ctk.CTkScrollableFrame):
    """Individual log tab for console tasks only (lightweight)"""
    
    def __init__(self, parent, task_name, on_close_callback=None, executor=None, exe_path=None):
        super().__init__(parent, fg_color="#1a1a1a")
        
        self.task_name = task_name
        self.log_buffer = []  # Lightweight buffer
        self.max_lines = 500  # Limit log size for performance
        self.process = None  # Store process reference
        self.on_close_callback = on_close_callback
        self.executor = executor  # ProcessExecutor reference
        self.exe_path = exe_path  # Executable path for cleanup
        self.auto_closing = False  # Flag to indicate if auto-closing without user action
        
        # Header frame with close button
        header = ctk.CTkFrame(self, fg_color="#1a1a1a")
        header.pack(fill="x", padx=5, pady=(5, 0))
        
        title_label = ctk.CTkLabel(
            header,
            text=task_name,
            font=("Segoe UI", 11, "bold"),
            text_color="#ffffff"
        )
        title_label.pack(side="left", padx=5, pady=5)
        
        close_btn = ctk.CTkButton(
            header,
            text="✕",
            width=25,
            height=25,
            corner_radius=4,
            font=("Segoe UI", 10, "bold"),
            fg_color="#ef4444",
            hover_color="#dc2626",
            command=self.close_process
        )
        close_btn.pack(side="right", padx=5, pady=5)
        
        # Log text widget
        self.log_text = ctk.CTkTextbox(
            self,
            wrap="word",
            font=("Consolas", 10),
            fg_color="#0d0d0d",
            border_width=0,
            corner_radius=8
        )
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Add initial message
        self.append_log(f"📋 Console Log: {task_name}\n")
        self.append_log(f"⏰ {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
    
    def append_log(self, text):
        """Append text to log (thread-safe, memory-limited)"""
        def _append():
            try:
                self.log_text.configure(state="normal")
                self.log_text.insert("end", text)
                
                # Keep only last N lines for performance - FIXED VERSION
                # Get current line count directly from text widget
                line_count = int(self.log_text.index('end-1c').split('.')[0])
                if line_count > self.max_lines:
                    # Calculate lines to remove from the beginning
                    lines_to_remove = line_count - self.max_lines
                    # Delete from line 1.0 to the calculated line
                    self.log_text.delete("1.0", f"{lines_to_remove}.0")
                
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
            except Exception as e:
                print(f"Error appending log: {e}")
        
        # Ensure UI update happens on main thread
        if threading.current_thread() is threading.main_thread():
            _append()
        else:
            self.after(0, _append)
    
    def close_process(self):
        """Gracefully terminate the running process (only if explicitly requested)"""
        print(f"[CLOSE_PROCESS] Button clicked. Process={self.process}, auto_closing={self.auto_closing}")
        
        # Only terminate if process exists and we're not auto-closing
        if self.process and not self.auto_closing:
            try:
                # Check if process is still running
                poll_result = self.process.poll()
                print(f"[CLOSE_PROCESS] Process poll result: {poll_result}")
                
                if poll_result is None:  # Still running
                    # Kill the entire process tree (parent + all children)
                    self.append_log(f"\n[!] Terminating process and all child processes...\n")
                    print(f"[CLOSE_PROCESS] Killing process tree for PID {self.process.pid}")
                    
                    try:
                        import psutil
                        parent = psutil.Process(self.process.pid)
                        children = parent.children(recursive=True)
                        
                        # Terminate children first
                        for child in children:
                            print(f"[CLOSE_PROCESS] Terminating child PID {child.pid}")
                            try:
                                child.terminate()
                            except psutil.NoSuchProcess:
                                pass
                        
                        # Terminate parent
                        print(f"[CLOSE_PROCESS] Terminating parent PID {parent.pid}")
                        parent.terminate()
                        
                        # Wait for graceful shutdown
                        gone, alive = psutil.wait_procs([parent] + children, timeout=2)
                        
                        # Force kill any remaining
                        for p in alive:
                            print(f"[CLOSE_PROCESS] Force killing PID {p.pid}")
                            try:
                                p.kill()
                            except psutil.NoSuchProcess:
                                pass
                        
                        self.append_log(f"[+] Process tree terminated\n")
                        print(f"[CLOSE_PROCESS] Process tree terminated")
                        
                    except ImportError:
                        # Fallback if psutil not available
                        print(f"[CLOSE_PROCESS] psutil not available, using basic terminate")
                        self.process.terminate()
                        try:
                            self.process.wait(timeout=2.0)
                        except subprocess.TimeoutExpired:
                            self.process.kill()
                            self.process.wait()
                        self.append_log(f"[+] Process terminated\n")
                else:
                    self.append_log(f"[i] Process already completed\n")
                    print(f"[CLOSE_PROCESS] Process already dead")
                
                # Force cleanup from executor's tracking
                if self.executor and self.exe_path:
                    print(f"[CLOSE_PROCESS] Calling force_cleanup on executor")
                    self.executor.force_cleanup(self.exe_path)
            except (OSError, PermissionError) as e:
                self.append_log(f"\n[x] Error terminating process: {str(e)}\n")
                print(f"[CLOSE_PROCESS] Error: {e}")
            except Exception as e:
                print(f"[CLOSE_PROCESS] Unexpected error terminating process: {e}")
        elif not self.process:
            print(f"[CLOSE_PROCESS] No process reference found!")
            self.append_log(f"[x] No process handle available\n")
        else:
            print(f"[CLOSE_PROCESS] Skipping termination (auto_closing={self.auto_closing})")
        
        # Call the callback to close the tab
        if self.on_close_callback:
            try:
                self.on_close_callback()
            except Exception as e:
                print(f"Error in close callback: {e}")


class AddTaskDialog(ctk.CTkToplevel):
    """Dialog for adding/editing tasks"""
    
    def __init__(self, parent, task=None, task_manager=None):
        super().__init__(parent)
        
        self.task = task
        self.task_manager = task_manager
        self.result = None
        
        # Configure window
        self.title("Edit Task" if task else "Add Task")
        self.geometry("500x350")
        self.resizable(False, False)
        
        # Center window
        self.transient(parent)
        self.grab_set()
        
        # Content frame
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Task Name
        ctk.CTkLabel(content, text="Task Name:", font=("Segoe UI", 12)).pack(anchor="w", pady=(0, 5))
        self.name_entry = ctk.CTkEntry(content, height=35, corner_radius=8)
        self.name_entry.pack(fill="x", pady=(0, 15))
        
        # Executable Path
        ctk.CTkLabel(content, text="Executable Path:", font=("Segoe UI", 12)).pack(anchor="w", pady=(0, 5))
        
        path_frame = ctk.CTkFrame(content, fg_color="transparent")
        path_frame.pack(fill="x", pady=(0, 15))
        
        self.path_entry = ctk.CTkEntry(path_frame, height=35, corner_radius=8)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        browse_btn = ctk.CTkButton(
            path_frame,
            text="Browse",
            width=80,
            height=35,
            corner_radius=8,
            command=self.browse_file
        )
        browse_btn.pack(side="right")
        
        # Interval
        ctk.CTkLabel(content, text="Interval (minutes):", font=("Segoe UI", 12)).pack(anchor="w", pady=(0, 5))
        self.interval_entry = ctk.CTkEntry(content, height=35, corner_radius=8)
        self.interval_entry.pack(fill="x", pady=(0, 25))
        
        # Buttons
        button_frame = ctk.CTkFrame(content, fg_color="transparent")
        button_frame.pack(fill="x")
        
        cancel_btn = ctk.CTkButton(
            button_frame,
            text="Cancel",
            height=40,
            corner_radius=8,
            fg_color="#2d2d2d",
            hover_color="#3d3d3d",
            command=self.cancel
        )
        cancel_btn.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        save_btn = ctk.CTkButton(
            button_frame,
            text="Save",
            height=40,
            corner_radius=8,
            command=self.save
        )
        save_btn.pack(side="right", fill="x", expand=True)
        
        # Populate if editing
        if task:
            self.name_entry.insert(0, task["name"])
            self.path_entry.insert(0, task["path"])
            self.interval_entry.insert(0, str(task["interval"]))
        else:
            # Pre-fill with last used exe path if available
            if task_manager:
                last_path = task_manager.get_last_exe_path()
                if last_path and os.path.exists(last_path):
                    self.path_entry.insert(0, last_path)
    
    def browse_file(self):
        """Browse for executable file"""
        filename = filedialog.askopenfilename(
            title="Select Executable",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
        )
        if filename:
            self.path_entry.delete(0, "end")
            self.path_entry.insert(0, filename)
            # Save this path for next time
            if self.task_manager:
                self.task_manager.set_last_exe_path(filename)
    
    def save(self):
        """Save task with validation"""
        name = self.name_entry.get().strip()
        path = self.path_entry.get().strip()
        interval = self.interval_entry.get().strip()
        
        # Validation
        if not name or not path or not interval:
            messagebox.showerror("Error", "All fields are required!")
            return
        
        if len(name) > 100:
            messagebox.showerror("Error", "Task name is too long (max 100 characters)!")
            return
        
        try:
            interval = int(interval)
            if interval <= 0:
                messagebox.showerror("Error", "Interval must be greater than 0!")
                return
            if interval > 10080:  # 1 week in minutes
                messagebox.showwarning("Warning", "Interval is very long (more than 1 week)")
        except ValueError:
            messagebox.showerror("Error", "Interval must be a valid number!")
            return
        
        if not os.path.exists(path):
            messagebox.showerror("Error", "Executable file not found!")
            return
        
        if not os.path.isfile(path):
            messagebox.showerror("Error", "Path must point to a file!")
            return
        
        # Save the path for next time
        if self.task_manager:
            self.task_manager.set_last_exe_path(path)
        
        self.result = {"name": name, "path": path, "interval": interval}
        self.destroy()
    
    def cancel(self):
        """Cancel dialog"""
        self.destroy()


class SchedulerApp(ctk.CTk):
    """Main application window"""
    
    def __init__(self):
        super().__init__()
        
        # Configure window
        self.title("Scheduler")
        self.geometry("1200x700")
        
        # Initialize managers
        self.task_manager = TaskManager()
        
        # Check for orphaned processes from previous crashed sessions
        orphaned_tasks = self.task_manager.cleanup_orphaned_processes()
        
        # Create crash marker file to detect future crashes
        self.task_manager.create_crash_marker()
        
        if orphaned_tasks:
            # Show warning dialog about orphaned processes
            orphaned_names = "\n".join([f"  • {task['name']}" for task in orphaned_tasks])
            warning_message = (
                "⚠ WARNING: Scheduler was not properly closed!\n\n"
                f"The following {len(orphaned_tasks)} task(s) may still be running:\n\n"
                f"{orphaned_names}\n\n"
                "These processes may still be running in the background.\n"
                "Please close them manually through the app or Task Manager.\n\n"
                "To gracefully close them:\n"
                "1. Select the task from the list\n"
                "2. Click 'Execute' to see/close it, OR\n"
                "3. Use Windows Task Manager to stop them\n\n"
                "Click OK to continue."
            )
            # Show warning immediately (before building UI)
            messagebox.showwarning("⚠ Orphaned Processes Detected", warning_message)
        
        self.executor = ProcessExecutor()
        self.scheduler = BackgroundScheduler()
        
        # Start scheduler with error handling
        try:
            self.scheduler.start()
            print("[SCHEDULER] Started successfully")
        except Exception as e:
            print(f"[SCHEDULER] CRITICAL ERROR starting scheduler: {e}")
            messagebox.showerror(
                "Scheduler Error",
                f"Failed to start task scheduler:\n{str(e)}\n\nThe application will still run but tasks won't execute automatically."
            )
        
        # UI components
        self.task_rows = {}
        self.log_tabs = {}
        self.selected_task_id = None
        self.scheduler_paused = False
        self.control_button = None
        
        # Build UI
        self.build_ui()
        
        # Load existing tasks
        self.load_tasks()
        
        # Handle window close
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def build_ui(self):
        """Build the user interface"""
        # Main container
        main_container = ctk.CTkFrame(self, fg_color="#1a1a1a", corner_radius=0)
        main_container.pack(fill="both", expand=True)
        
        # Left panel (Tasks)
        left_panel = ctk.CTkFrame(main_container, width=450, fg_color="#0d0d0d", corner_radius=0)
        left_panel.pack(side="left", fill="both", padx=0, pady=0)
        left_panel.pack_propagate(False)
        
        # Left panel header
        header = ctk.CTkFrame(left_panel, fg_color="transparent", height=80)
        header.pack(fill="x", padx=20, pady=(20, 10))
        
        ctk.CTkLabel(
            header,
            text="Scheduled Tasks",
            font=("Segoe UI", 20, "bold")
        ).pack(side="left", anchor="w")
        
        add_btn = ctk.CTkButton(
            header,
            text="Add Task",
            width=100,
            height=36,
            corner_radius=8,
            font=("Segoe UI", 12),
            command=self.add_task
        )
        add_btn.pack(side="right", anchor="e")
        
        # Task table
        table_container = ctk.CTkFrame(left_panel, fg_color="transparent")
        table_container.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        
        # Table header
        table_header = ctk.CTkFrame(table_container, fg_color="#1a1a1a", height=40, corner_radius=8)
        table_header.pack(fill="x", pady=(0, 5))
        
        ctk.CTkLabel(
            table_header,
            text="Task Name",
            font=("Segoe UI", 11, "bold"),
            width=150
        ).pack(side="left", padx=(15, 0))
        
        ctk.CTkLabel(
            table_header,
            text="Time",
            font=("Segoe UI", 11, "bold"),
            width=80
        ).pack(side="left", padx=(10, 0))
        
        ctk.CTkLabel(
            table_header,
            text="Status",
            font=("Segoe UI", 11, "bold"),
            width=80
        ).pack(side="left", padx=(10, 0))
        
        # Scrollable task list (with hidden scrollbar initially)
        self.task_list = ctk.CTkScrollableFrame(
            table_container,
            fg_color="#1a1a1a",
            corner_radius=8,
            scrollbar_button_color="#1a1a1a",  # Hide scrollbar button initially
            scrollbar_button_hover_color="#2d2d2d"
        )
        self.task_list.pack(fill="both", expand=True)
        
        # Action buttons
        action_frame = ctk.CTkFrame(left_panel, fg_color="transparent", height=60)
        action_frame.pack(fill="x", padx=20, pady=(10, 20))
        
        edit_btn = ctk.CTkButton(
            action_frame,
            text="Edit",
            height=40,
            corner_radius=8,
            fg_color="#2d2d2d",
            hover_color="#3d3d3d",
            font=("Segoe UI", 12),
            command=self.edit_task
        )
        edit_btn.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        delete_btn = ctk.CTkButton(
            action_frame,
            text="Delete",
            height=40,
            corner_radius=8,
            fg_color="#2d2d2d",
            hover_color="#3d3d3d",
            font=("Segoe UI", 12),
            command=self.delete_task
        )
        delete_btn.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        execute_btn = ctk.CTkButton(
            action_frame,
            text="Execute",
            height=40,
            corner_radius=8,
            font=("Segoe UI", 12),
            command=self.execute_task
        )
        execute_btn.pack(side="right", fill="x", expand=True)
        
        # Right panel (Logs)
        right_panel = ctk.CTkFrame(main_container, fg_color="#0d0d0d", corner_radius=0)
        right_panel.pack(side="right", fill="both", expand=True, padx=0, pady=0)
        
        # Right panel header
        log_header = ctk.CTkFrame(right_panel, fg_color="transparent", height=60)
        log_header.pack(fill="x", padx=20, pady=(20, 10))
        
        ctk.CTkLabel(
            log_header,
            text="Logs",
            font=("Segoe UI", 20, "bold")
        ).pack(side="left", anchor="w")
        
        # Vertical log container (stacks logs vertically)
        self.log_container = VerticalLogContainer(right_panel)
        
        # Status bar
        status_bar = ctk.CTkFrame(right_panel, fg_color="#1a1a1a", height=60, corner_radius=0)
        status_bar.pack(fill="x", padx=0, pady=0, side="bottom")
        
        # Status info
        status_info_frame = ctk.CTkFrame(status_bar, fg_color="transparent")
        status_info_frame.pack(side="left", padx=20, pady=10, fill="both", expand=True)
        
        self.status_label = ctk.CTkLabel(
            status_info_frame,
            text="● Scheduler Running",
            font=("Segoe UI", 11),
            text_color="#4ade80"
        )
        self.status_label.pack(side="left")
        
        # Control button (Start/Pause)
        self.control_button = ctk.CTkButton(
            status_bar,
            text="⏸ Pause Scheduler",
            width=150,
            height=40,
            corner_radius=8,
            font=("Segoe UI", 11, "bold"),
            fg_color="#ef4444",
            hover_color="#dc2626",
            command=self.toggle_scheduler
        )
        self.control_button.pack(side="right", padx=20, pady=10)
    
    def add_task_row(self, task):
        """Add a task row to the table"""
        row = ctk.CTkFrame(
            self.task_list,
            fg_color="#0d0d0d",
            corner_radius=8,
            height=50
        )
        row.pack(fill="x", pady=5, padx=5)
        
        # Make row clickable
        row.bind("<Button-1>", lambda e, t=task: self.select_task(t["id"]))
        
        name_label = ctk.CTkLabel(
            row,
            text=task["name"],
            font=("Segoe UI", 11),
            width=150,
            anchor="w"
        )
        name_label.pack(side="left", padx=(15, 0))
        name_label.bind("<Button-1>", lambda e, t=task: self.select_task(t["id"]))
        
        # Format interval as time
        interval_min = task["interval"]
        time_str = f"{interval_min} min" if interval_min < 60 else f"{interval_min // 60}h {interval_min % 60}m"
        
        time_label = ctk.CTkLabel(
            row,
            text=time_str,
            font=("Segoe UI", 11),
            width=80,
            anchor="w"
        )
        time_label.pack(side="left", padx=(10, 0))
        time_label.bind("<Button-1>", lambda e, t=task: self.select_task(t["id"]))
        
        status_label = ctk.CTkLabel(
            row,
            text=task["status"],
            font=("Segoe UI", 11),
            width=80,
            anchor="w",
            text_color="#4ade80" if task["status"] == "Running" else "#94a3b8"
        )
        status_label.pack(side="left", padx=(10, 0))
        status_label.bind("<Button-1>", lambda e, t=task: self.select_task(t["id"]))
        
        self.task_rows[task["id"]] = {
            "frame": row,
            "status_label": status_label
        }
        
        # Update scrollbar visibility
        self.update_scrollbar_visibility()
    
    def update_scrollbar_visibility(self):
        """Show/hide scrollbar based on content size"""
        try:
            # Get the canvas bounds
            self.task_list.update_idletasks()
            canvas_height = self.task_list.winfo_height()
            content_height = self.task_list._canvas.bbox("all")
            
            if content_height and len(content_height) >= 4:
                total_height = content_height[3] - content_height[1]
                # Show scrollbar if content exceeds visible area
                if total_height > canvas_height - 10:
                    self.task_list._scrollbar.configure(fg_color="#555555")
                else:
                    self.task_list._scrollbar.configure(fg_color="#1a1a1a")
        except (AttributeError, IndexError, TypeError) as e:
            # Ignore errors if canvas/scrollbar not ready
            pass
        except Exception as e:
            print(f"Unexpected error updating scrollbar: {e}")
    
    def select_task(self, task_id):
        """Select a task"""
        # Deselect previous
        if self.selected_task_id and self.selected_task_id in self.task_rows:
            self.task_rows[self.selected_task_id]["frame"].configure(fg_color="#0d0d0d")
        
        # Select new
        self.selected_task_id = task_id
        if task_id in self.task_rows:
            self.task_rows[task_id]["frame"].configure(fg_color="#1f6aa5")
    
    def add_task(self):
        """Add a new task"""
        dialog = AddTaskDialog(self, task_manager=self.task_manager)
        self.wait_window(dialog)
        
        if dialog.result:
            task = self.task_manager.add_task(
                dialog.result["name"],
                dialog.result["path"],
                dialog.result["interval"]
            )
            self.add_task_row(task)
            self.schedule_task(task)
    
    def edit_task(self):
        """Edit selected task"""
        if not self.selected_task_id:
            messagebox.showwarning("Warning", "Please select a task to edit")
            return
        
        # Find task
        task = next((t for t in self.task_manager.tasks if t["id"] == self.selected_task_id), None)
        if not task:
            return
        
        dialog = AddTaskDialog(self, task, task_manager=self.task_manager)
        self.wait_window(dialog)
        
        if dialog.result:
            # Update task
            updated_task = self.task_manager.update_task(
                task["id"],
                dialog.result["name"],
                dialog.result["path"],
                dialog.result["interval"]
            )
            
            # Reschedule
            self.scheduler.remove_job(f"task_{task['id']}")
            self.schedule_task(updated_task)
            
            # Refresh UI
            self.refresh_task_list()
    
    def delete_task(self):
        """Delete selected task and its log panel"""
        if not self.selected_task_id:
            messagebox.showwarning("Warning", "Please select a task to delete")
            return
        
        if messagebox.askyesno("Confirm", "Are you sure you want to delete this task?"):
            task_id = self.selected_task_id
            
            # Remove from scheduler
            try:
                self.scheduler.remove_job(f"task_{task_id}")
            except:
                pass
            
            # Remove log tab if it exists
            if task_id in self.log_tabs:
                # Get task name for tab
                task = next((t for t in self.task_manager.tasks if t["id"] == task_id), None)
                if task:
                    try:
                        self.log_container.delete(task["name"])
                    except:
                        pass
                del self.log_tabs[task_id]
            
            # Remove from manager
            self.task_manager.delete_task(task_id)
            
            # Refresh UI
            self.refresh_task_list()
            self.selected_task_id = None
    
    def execute_task(self):
        """Execute selected task immediately"""
        if not self.selected_task_id:
            messagebox.showwarning("Warning", "Please select a task to execute")
            return
        
        task = next((t for t in self.task_manager.tasks if t["id"] == self.selected_task_id), None)
        if task:
            self.run_task(task)
    
    def refresh_task_list(self):
        """Refresh the task list display"""
        # Clear existing rows
        for widget in self.task_list.winfo_children():
            widget.destroy()
        
        self.task_rows.clear()
        
        # Reload tasks
        for task in self.task_manager.tasks:
            self.add_task_row(task)
    
    def load_tasks(self):
        """Load and schedule all tasks"""
        for task in self.task_manager.tasks:
            self.add_task_row(task)
            self.schedule_task(task)
    
    def create_log_panel(self, task_name, task_id, exe_path=None):
        """Create a vertical log panel"""
        panel = self.log_container.add(task_name)
        
        # Create callback for when close button is clicked
        def on_panel_close():
            self.auto_close_panel(task_id)
        
        log_panel = LogTab(
            panel, 
            task_name, 
            on_close_callback=on_panel_close,
            executor=self.executor,
            exe_path=exe_path
        )
        log_panel.pack(fill="both", expand=True)
        self.log_tabs[task_id] = log_panel
        return log_panel
    
    def schedule_task(self, task):
        """Schedule a task for automatic execution"""
        job_id = f"task_{task['id']}"
        
        self.scheduler.add_job(
            func=lambda: self.run_task(task),
            trigger=IntervalTrigger(minutes=task["interval"]),
            id=job_id,
            replace_existing=True
        )
    
    def run_task(self, task):
        """Run a task (lightweight - only log console apps)"""
        # Skip if scheduler is paused
        if self.scheduler_paused:
            return
        
        exe_path = task["path"]
        task_id = task["id"]
        
        print(f"[RUN_TASK] Executing task {task_id}: {os.path.basename(exe_path)}")
        
        # Check if this is a console app that needs logging
        needs_logging = self.executor.is_console_app(exe_path)
        print(f"[RUN_TASK] Detected as {'CONSOLE' if needs_logging else 'GUI'} app")
        
        log_tab = None
        log_callback = None
        
        if needs_logging:
            # Create log tab only for console apps
            if task_id not in self.log_tabs:
                self.create_log_panel(task["name"], task_id, exe_path=exe_path)
            
            log_tab = self.log_tabs[task_id]
            
            # No need to switch to tab in vertical mode - all visible
            
            # Log start
            timestamp = datetime.now().strftime('%H:%M:%S')
            log_tab.append_log(f"\n{'='*50}\n")
            log_tab.append_log(f"{timestamp}  Process started\n")
            log_tab.append_log(f"{'='*50}\n")
            
            # For tif2pdf and similar apps, add a note about buffering
            if 'tif2pdf' in exe_path.lower() or 'pdf' in exe_path.lower():
                log_tab.append_log(f"\n⚠ NOTE: This application may buffer output.\n")
                log_tab.append_log(f"   Output may appear in bursts or only at completion.\n")
                log_tab.append_log(f"   The process is running - check Task Manager if needed.\n\n")
            
            # Create log callback
            def log_callback_fn(text):
                try:
                    log_tab.append_log(text)
                except:
                    pass
            
            log_callback = log_callback_fn
        
        # Update status IMMEDIATELY before execution
        self.update_task_status(task_id, "Running")
        
        # Status watchdog - prevents stuck "Running" status
        watchdog_active = True
        heartbeat_active = True  # For heartbeat indicator
        
        def status_watchdog():
            # Wait 2x the task interval (max expected runtime)
            timeout_seconds = task["interval"] * 60 * 2
            time.sleep(timeout_seconds)
            
            if watchdog_active and task_id in self.task_rows:
                current_status = self.task_rows[task_id]["status_label"].cget("text")
                if current_status == "Running":
                    print(f"[WATCHDOG] Task {task_id} stuck in Running state - forcing to Idle")
                    self.update_task_status(task_id, "Idle")
                    self.task_manager.update_status(task_id, "Idle")
        
        # Heartbeat indicator for long-running processes with buffered output
        def heartbeat():
            """Show periodic heartbeat for processes that buffer output"""
            if not needs_logging or task_id not in self.log_tabs:
                return
            
            count = 0
            while heartbeat_active:
                time.sleep(10)  # Every 10 seconds
                if heartbeat_active and task_id in self.log_tabs:
                    count += 1
                    # Only show heartbeat for known buffering apps
                    if 'tif2pdf' in exe_path.lower() or 'pdf' in exe_path.lower():
                        timestamp = datetime.now().strftime('%H:%M:%S')
                        self.log_tabs[task_id].append_log(f"[{timestamp}] ⏳ Process running... ({count * 10}s elapsed)\n")
        
        watchdog_thread = threading.Thread(target=status_watchdog, daemon=True)
        watchdog_thread.start()
        
        # Start heartbeat for buffering-prone apps
        if needs_logging and ('tif2pdf' in exe_path.lower() or 'pdf' in exe_path.lower()):
            heartbeat_thread = threading.Thread(target=heartbeat, daemon=True)
            heartbeat_thread.start()
        
        # Create completion callback to auto-close tab
        def on_completion():
            nonlocal watchdog_active
            nonlocal heartbeat_active
            watchdog_active = False  # Disable watchdog since process completed normally
            heartbeat_active = False  # Stop heartbeat
            
            # Update status back to Idle
            print(f"[COMPLETION] Setting task {task_id} to Idle")
            self.update_task_status(task_id, "Idle")
            self.task_manager.update_status(
                task_id,
                "Idle",
                datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            )
            
            # Auto-close the tab after a brief delay (2 seconds)
            if needs_logging and task_id in self.log_tabs:
                self.after(2000, lambda: self.auto_close_panel(task_id))
        
        # Create callback to receive process reference
        def on_process_created(process):
            if task_id in self.log_tabs:
                self.log_tabs[task_id].process = process
        
        # Execute in thread
        def execute_thread():
            result = self.executor.execute(
                exe_path, 
                log_callback, 
                needs_logging, 
                completion_callback=on_completion,
                process_ref_callback=on_process_created
            )
            
            # If process was skipped (already running), handle it
            if result == "skipped":
                # DO NOT change status - keep it as Running since process is still running
                # DO NOT close the log tab - let the original process continue logging
                # Just append a skip message to inform user
                if needs_logging and task_id in self.log_tabs:
                    self.log_tabs[task_id].append_log(f"\n[!] Second execution attempt blocked - process already running\n\n")
        
        threading.Thread(target=execute_thread, daemon=True).start()
    
    def update_task_status(self, task_id, status):
        """Thread-safe update task status in UI"""
        def _update():
            try:
                if task_id in self.task_rows:
                    status_label = self.task_rows[task_id]["status_label"]
                    status_label.configure(
                        text=status,
                        text_color="#4ade80" if status == "Running" else "#94a3b8"
                    )
            except Exception as e:
                print(f"Error updating status for task {task_id}: {e}")
        
        # Ensure update happens on main thread
        if threading.current_thread() is threading.main_thread():
            _update()
        else:
            self.after(0, _update)
    
    def auto_close_panel(self, task_id):
        """Auto-close a log panel after task completion"""
        if task_id in self.log_tabs:
            task = next((t for t in self.task_manager.tasks if t["id"] == task_id), None)
            if task:
                try:
                    # Set flag to indicate auto-closing (won't terminate process)
                    self.log_tabs[task_id].auto_closing = True
                    self.log_container.delete(task["name"])
                    del self.log_tabs[task_id]
                except Exception as e:
                    print(f"Error closing panel for task {task_id}: {e}")
    
    def toggle_scheduler(self):
        """Toggle scheduler pause/resume"""
        self.scheduler_paused = not self.scheduler_paused
        
        if self.scheduler_paused:
            # Pause the scheduler properly
            try:
                self.scheduler.pause()
                print("[SCHEDULER] Paused all jobs")
            except Exception as e:
                print(f"[SCHEDULER] Error pausing: {e}")
            
            self.control_button.configure(
                text="▶ Start Scheduler",
                fg_color="#22c55e",
                hover_color="#16a34a"
            )
            self.status_label.configure(
                text="⏸ Scheduler Paused",
                text_color="#fbbf24"
            )
        else:
            # Resume the scheduler properly
            try:
                self.scheduler.resume()
                print("[SCHEDULER] Resumed all jobs")
            except Exception as e:
                print(f"[SCHEDULER] Error resuming: {e}")
            
            self.control_button.configure(
                text="⏸ Pause Scheduler",
                fg_color="#ef4444",
                hover_color="#dc2626"
            )
            self.status_label.configure(
                text="● Scheduler Running",
                text_color="#4ade80"
            )
    
    def on_closing(self):
        """Handle window close - Check for running tasks, warn user, and terminate processes"""
        # Check if any tasks are currently running (check actual process status)
        running_tasks = []
        running_processes_list = []  # Keep track of actual process objects
        
        # Check processes tracked by executor
        for exe_path, process in self.executor.running_processes.items():
            poll_result = process.poll()
            
            if poll_result is None:  # Process is still running
                running_processes_list.append(process)
                
                # Find the task that corresponds to this process
                with self.task_manager._lock:
                    for task in self.task_manager.tasks:
                        task_path_normalized = os.path.normpath(task.get("path"))
                        exe_path_normalized = os.path.normpath(exe_path)
                        if task_path_normalized == exe_path_normalized:
                            running_tasks.append(task)
                            break
        
        # If there are running tasks, warn user
        if running_tasks:
            running_names = ", ".join([task['name'] for task in running_tasks])
            warning_message = (
                f"{len(running_tasks)} task(s) still running: {running_names}\n\n"
                "Close anyway?"
            )
            
            result = messagebox.askyesno(
                "⚠ Running Tasks",
                warning_message,
                icon=messagebox.WARNING
            )
            print(f"[CLOSE_WARNING] User responded: {result}")
            
            if result is False:
                # User clicked NO - don't close
                print("[CLOSE] User chose not to close - running tasks detected")
                return
            else:
                # User clicked YES - proceed with close and TERMINATE all running processes
                print("[CLOSE] User confirmed close despite running tasks - terminating processes")
        
        try:
            # FORCEFULLY TERMINATE all running processes before closing
            print(f"[CLOSE] Checking {len(running_processes_list)} processes for termination...")
            
            import psutil
            for process in running_processes_list:
                try:
                    # Check if process is actually still running
                    if process.poll() is None:
                        # Kill entire process tree (parent + children)
                        print(f"[CLOSE] Terminating process tree for PID {process.pid}...")
                        
                        try:
                            parent = psutil.Process(process.pid)
                            children = parent.children(recursive=True)
                            
                            # Terminate all children first
                            for child in children:
                                print(f"[CLOSE] Terminating child PID {child.pid}")
                                try:
                                    child.terminate()
                                except psutil.NoSuchProcess:
                                    pass
                            
                            # Terminate parent
                            print(f"[CLOSE] Terminating parent PID {parent.pid}")
                            parent.terminate()
                            
                            # Wait for graceful shutdown
                            gone, alive = psutil.wait_procs([parent] + children, timeout=2)
                            
                            # Force kill survivors
                            for p in alive:
                                print(f"[CLOSE] Force killing PID {p.pid}")
                                try:
                                    p.kill()
                                except psutil.NoSuchProcess:
                                    pass
                            
                            print(f"[CLOSE] Process tree terminated")
                        except psutil.NoSuchProcess:
                            print(f"[CLOSE] Process already dead")
                    else:
                        # Process already terminated
                        print(f"[CLOSE] Process {process.pid} already terminated (exit code: {process.returncode})")
                        
                except Exception as e:
                    print(f"[CLOSE] Error handling process: {e}")
            
            # Clear the running_processes dict to prevent orphaned tracking
            self.executor.running_processes.clear()
            print("[CLOSE] All processes handled and cleared from tracking")
            
        except Exception as e:
            print(f"Warning: Error in process termination: {e}")
        
        try:
            # Set all running tasks to Idle before saving
            with self.task_manager._lock:
                for task in self.task_manager.tasks:
                    if task.get("status") == "Running":
                        task["status"] = "Idle"
            
            # Save all current tasks to tasks.json
            self.task_manager.save_tasks()
            print("✓ Tasks saved to disk")
        except Exception as e:
            print(f"Warning: Error saving tasks on close: {e}")
        
        try:
            # Remove crash marker file on clean shutdown
            self.task_manager.remove_crash_marker()
        except Exception as e:
            print(f"Warning: Error removing crash marker: {e}")
        
        try:
            # Shutdown scheduler gracefully
            self.scheduler.shutdown(wait=False)  # Don't wait - faster shutdown
            print("✓ Scheduler shut down")
        except Exception as e:
            print(f"Warning: Error shutting down scheduler: {e}")
        
        # Force close any remaining threads (output readers, monitors, etc.)
        print("[CLOSE] Forcing application exit...")
        
        # Close application
        self.destroy()
        
        # If destroy() doesn't work, force exit
        import sys
        sys.exit(0)


if __name__ == "__main__":
    app = SchedulerApp()
    app.mainloop()
