# Modern Scheduler Application

An ultra-lightweight, intelligent scheduler for managing and executing .exe files - perfect for resource-intensive PCs.

## Features

- **Ultra-Lightweight**: Minimal memory footprint, optimized for busy systems
- **Smart Detection**: Auto-detects console vs GUI apps (zero configuration)
- **Modern Dark UI**: Built with CustomTkinter for a sleek, modern appearance
- **Automated Scheduling**: Schedule .exe files to run at specified intervals
- **Selective Logging**: Only captures output from console apps (GUI apps run natively)
- **No CMD Popups**: Console apps run hidden, GUI apps show their own windows
- **Process Management**: Prevents overlapping executions of the same .exe
- **Persistent Storage**: Tasks saved to JSON for persistence across sessions
- **Real-time Monitoring**: Live status updates for all tasks

## Installation

1. Install Python 3.8 or higher
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Run the application:

```bash
python index.py
```

### Adding Tasks

1. Click "Add Task" button
2. Enter task name, select .exe file, and set interval (in minutes)
3. Task will be automatically scheduled

### Managing Tasks

- **Edit**: Select a task and click "Edit" to modify settings
- **Delete**: Select a task and click "Delete" to remove it
- **Execute**: Select a task and click "Execute" to run immediately

### Logs

- **Console apps only**: Log tabs appear only for CMD/batch/console executables
- **GUI apps**: Run with their own windows, no log capture (zero overhead)
- **Auto-detection**: Scheduler automatically determines app type
- Logs show real-time stdout/stderr output
- Limited to 500 lines per task (performance optimization)
- Timestamps and process status included

## Performance Optimizations

This scheduler is designed for **resource-intensive PCs**:

- ✅ **Minimal Memory**: 500-line log limit per task
- ✅ **Smart Threading**: Only creates threads when needed
- ✅ **Zero GUI Overhead**: GUI apps run natively with no capture
- ✅ **Auto-Detection**: No manual configuration needed
- ✅ **Lightweight UI**: Optimized CustomTkinter components
- ✅ **Efficient Storage**: Simple JSON (not database overhead)

## Technical Details

### Architecture

- **CustomTkinter**: Modern, themed UI components
- **APScheduler**: Background task scheduling with interval triggers
- **subprocess.Popen**: Process execution with output capture
- **Threading**: Non-blocking UI with concurrent process execution
- **JSON**: Simple, readable task persistence

### Key Features

1. **Thread-Safe Logging**: All log updates are thread-safe
2. **Process Deduplication**: Same .exe path won't run concurrently
3. **Smart Detection**: Auto-detects console vs GUI apps
4. **No CMD Popups**: Console apps hidden, GUI apps show naturally
5. **Memory Efficient**: Limited log buffering, minimal overhead
6. **Adaptive Threading**: Only creates threads when needed
7. **Performance First**: Designed for resource-intensive environments

## File Structure

```
Schedulerv2/
├── index.py           # Main application
├── requirements.txt   # Python dependencies
├── tasks.json        # Task storage (auto-created)
└── README.md         # This file
```

## Design Philosophy

Inspired by modern dashboard tools like Notion and Linear:
- Clean, minimal interface
- Rounded corners and soft shadows
- Blue accent colors
- Scalable architecture
- Focus on user experience

## Future Enhancements

- Task grouping and organization
- Advanced search and filtering
- Export logs to file
- Custom notifications
- Task dependencies
- Conditional execution

## License

Free to use and modify.
