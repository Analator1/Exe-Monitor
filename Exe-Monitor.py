import time
import psutil
import win32gui
import win32process
import win32ui
import win32con
import win32api
import threading
from pypresence import Presence, exceptions
import base64
import io
import os
from PIL import Image
from flask import Flask, jsonify, render_template_string, request
from rich.console import Console
import logging
from rich.logging import RichHandler
import json
from datetime import date
from datetime import datetime, timedelta
CACHE_DIR = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "EXE-Monitor")


# Replace these with the actual Exe name and Discord Applicaton IDs.

MAIN_PROCESS_CLIENT_IDS = {
    "Resolve.exe": "123456789012345678",
    "Code.exe":    "123456789012345678",
}

MAIN_PROCESSES = list(MAIN_PROCESS_CLIENT_IDS.keys())

POLL_INTERVAL = 1.0

WEB_SERVER_PORT = 1111

SAVE_INTERVAL = 5

DISCORD_UPDATE_INTERVAL = 15

logging.basicConfig(
    level="INFO",
    format="%(message)s",
    datefmt="[%X]",
    handlers=[
        RichHandler(
            show_path=False,
            show_level=False,
            show_time=False,
            rich_tracebacks=True,
            markup=True
        )
    ],
)

app_logger = logging.getLogger("rich")
console = Console()
process_stats = {}
stats_lock = threading.Lock()

HTML_TEMPLATE = os.path.join(os.path.dirname(__file__), "Main.html")
with open(HTML_TEMPLATE, 'r', encoding='utf-8') as f:
    html_content = f.read()

def update_discord_presence():
    """
    Updates Discord Rich Presence, dynamically connecting with the
    Client ID of the currently focused main application.
    The presence will "stick" to the last focused main application.
    """
    rpc = None
    current_client_id = None
    last_focused_main_app_stats = None

    while True:
        now = time.time()
        main_procs_lower = [p.lower() for p in MAIN_PROCESSES]
        
        currently_focused_stats = None
        with stats_lock:
            for name_lower, stats in process_stats.items():
                if name_lower in main_procs_lower and stats['last_focused']:
                    currently_focused_stats = stats
                    break
        
        if currently_focused_stats:
            last_focused_main_app_stats = currently_focused_stats

        if not last_focused_main_app_stats:
            time.sleep(DISCORD_UPDATE_INTERVAL)
            continue
            
        app_to_display_name = last_focused_main_app_stats.get('original_name')
        target_client_id = None
        original_name = next((name for name in MAIN_PROCESS_CLIENT_IDS if name.lower() == app_to_display_name.lower()), None)
        if original_name:
            target_client_id = MAIN_PROCESS_CLIENT_IDS[original_name]

        if target_client_id != current_client_id:
            if rpc:
                console.print(f"[bold yellow]Switching Discord presence from Client ID: {current_client_id}[/bold yellow]")
                rpc.close()
            
            rpc = None
            current_client_id = None

            if target_client_id:
                try:
                    console.print(f"[bold blue]Connecting to Discord for '{app_to_display_name}' (ID: {target_client_id})...[/bold blue]")
                    rpc = Presence(target_client_id)
                    rpc.connect()
                    current_client_id = target_client_id
                    console.print("[bold blue]Connection successful with Discord RPC.[/bold blue]")
                except (exceptions.InvalidPipe, exceptions.DiscordNotFound):
                    console.print(f"[bold red]Could not connect to Discord for {app_to_display_name}. Is it running?[/bold red]")
                    rpc = None
                except exceptions.InvalidID:
                    console.print(f"[blink bold red]Connection failed for '{app_to_display_name}'. The Client ID '{target_client_id}' is invalid. RPC is not working.[/blink bold red]")
                    rpc = None
        
        if rpc:
            try:
                focus_total_display = last_focused_main_app_stats['total_focused']
                if last_focused_main_app_stats['last_focused'] and last_focused_main_app_stats['focused_start']:
                    focus_total_display += now - last_focused_main_app_stats['focused_start']
                
                hours, remainder = divmod(focus_total_display, 3600)
                minutes, _ = divmod(remainder, 60)
                state_text = f"Focus Time: {int(hours)}h {int(minutes)}m"

                open_total_display = last_focused_main_app_stats['total_open']
                if last_focused_main_app_stats['last_running'] and last_focused_main_app_stats['running_start']:
                    open_total_display += now - last_focused_main_app_stats['running_start']
                
                start_timestamp = now - open_total_display

                rpc.update(
                    details="Doing work!",
                    state=state_text,
                    start=start_timestamp
                )
            except (exceptions.InvalidPipe, exceptions.DiscordError) as e:
                console.print(f"[bold red]Discord connection lost: {e}. Resetting...[/bold red]")
                if rpc:
                    rpc.close()
                rpc = None
                current_client_id = None
        
        time.sleep(DISCORD_UPDATE_INTERVAL)

def get_icon_as_base64(exe_path):
    """Extracts the icon from an exe and returns it as a base64 data URI."""
    if not exe_path:
        return None
    try:
        ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
        ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)
        large, small = win32gui.ExtractIconEx(exe_path, 0)
        hicon = large[0] if large else small[0] if small else None
        if not hicon:
            return None
        
        hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
        hbmp = win32ui.CreateBitmap()
        hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_y)
        hdc = hdc.CreateCompatibleDC()
        hdc.SelectObject(hbmp)
        hdc.DrawIcon((0, 0), hicon)

        bmp_info = hbmp.GetInfo()
        bmp_str = hbmp.GetBitmapBits(True)
        img = Image.frombuffer(
            'RGBA',
            (bmp_info['bmWidth'], bmp_info['bmHeight']),
            bmp_str, 'raw', 'BGRA', 0, 1
        )

        buffered = io.BytesIO()
        img.save(buffered, format="PNG")
        img_str = base64.b64encode(buffered.getvalue()).decode('utf-8')
        data_uri = f"data:image/png;base64,{img_str}"

        win32gui.DestroyIcon(hicon)
        win32gui.DeleteObject(hbmp.GetHandle())
        hdc.DeleteDC()
        
        return data_uri
    except Exception:
        return None

def get_process_info():
    """Gets running process info and foreground window info."""
    running_procs = {}
    for p in psutil.process_iter(['name', 'exe']):
        try:
            p_info = p.info
            if p_info['name'] and p_info['exe']:
                running_procs[p_info['name'].lower()] = {
                    "original_name": p_info['name'],
                    "exe_path": p_info['exe']
                }
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    try:
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd: return running_procs, None, None
        pid = win32process.GetWindowThreadProcessId(hwnd)[1]
        p_name = psutil.Process(pid).name().lower()
        is_minimized = bool(win32gui.IsIconic(hwnd))
        return running_procs, p_name, is_minimized
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.Error):
        return running_procs, None, None

def monitor_processes_worker():
    """The main worker function that runs in a background thread."""
    global process_stats
    
    with stats_lock:
        for name in MAIN_PROCESSES:
            name_lower = name.lower()
            if name_lower not in process_stats:
                process_stats[name_lower] = {
                    'original_name': name, 'total_open': 0.0, 'total_focused': 0.0,
                    'last_running': False, 'running_start': None,
                    'last_focused': False, 'focused_start': None,
                    'icon_data_uri': None
                }
    
    console.print("[bold green]Background monitoring thread started...[/bold green]")
    while True:
        now = time.time()
        running_procs_info, fg_process_name, fg_is_minimized = get_process_info()
        running_names_lower = set(running_procs_info.keys())

        with stats_lock:
            for name_lower, info in running_procs_info.items():
                if name_lower not in process_stats:
                    process_stats[name_lower] = {
                        'original_name': info['original_name'], 'total_open': 0.0, 'total_focused': 0.0,
                        'last_running': False, 'running_start': None,
                        'last_focused': False, 'focused_start': None,
                        'icon_data_uri': get_icon_as_base64(info['exe_path']) or "failed"
                    }
                
                stats = process_stats[name_lower]
                
                if stats['icon_data_uri'] is None:
                    stats['icon_data_uri'] = get_icon_as_base64(info['exe_path']) or "failed"
                
                if not stats['last_running']:
                    stats['running_start'] = now
                stats['last_running'] = True
                
                focused = (name_lower == fg_process_name) and (not fg_is_minimized)
                if focused and not stats['last_focused']:
                    stats['focused_start'] = now
                elif not focused and stats['last_focused'] and stats['focused_start'] is not None:
                    stats['total_focused'] += now - stats['focused_start']
                    stats['focused_start'] = None
                stats['last_focused'] = focused

            all_tracked_names = list(process_stats.keys())
            for name_lower in all_tracked_names:
                stats = process_stats[name_lower]
                if name_lower not in running_names_lower and stats['last_running']:
                    if stats['running_start'] is not None:
                        stats['total_open'] += now - stats['running_start']
                        stats['running_start'] = None
                    if stats['focused_start'] is not None:
                        stats['total_focused'] += now - stats['focused_start']
                        stats['focused_start'] = None
                    stats['last_running'] = False
                    stats['last_focused'] = False
        
        time.sleep(POLL_INTERVAL)

app = Flask(__name__)

@app.route('/')
def index():
    return render_template_string(html_content)

@app.route('/data')
def get_data():
    main_procs_data = []
    other_procs_data = []
    now = time.time()
    main_procs_lower = [p.lower() for p in MAIN_PROCESSES]

    with stats_lock:
        for name_lower, stats in process_stats.items():
            open_total_display = stats['total_open']
            if stats['last_running'] and stats['running_start']:
                open_total_display += now - stats['running_start']

            focus_total_display = stats['total_focused']
            if stats['last_focused'] and stats['focused_start']:
                focus_total_display += now - stats['focused_start']

            status = "Closed"
            if stats['last_running']: status = "Running"
            if stats['last_focused']: status = "Focused"

            icon_uri = stats.get('icon_data_uri')
            if icon_uri == "failed": icon_uri = None
            
            has_icon = bool(icon_uri)

            proc_data = {
                "name": stats.get('original_name', name_lower),
                "status": status,
                "open_total_display": open_total_display,
                "focus_total_display": focus_total_display,
                "icon": icon_uri,
                "has_icon": has_icon
            }
            
            if open_total_display == 0 and name_lower not in main_procs_lower:
                continue

            if name_lower in main_procs_lower:
                main_procs_data.append(proc_data)
            else:
                other_procs_data.append(proc_data)
                
    return jsonify({
        "main_procs": main_procs_data,
        "other_procs": other_procs_data
    })

@app.route('/data/<date_str>')
def get_historical_data(date_str):
    cache_file = os.path.join(CACHE_DIR, f"{date_str}.json")
    if not os.path.exists(cache_file):
        return jsonify({"error": "No data for this date"}), 404
    try:
        with open(cache_file, 'r') as f:
            data = json.load(f)
    except (json.JSONDecodeError, IOError):
        return jsonify({"error": "Invalid data file"}), 500
    
    main_procs_data = []
    other_procs_data = []
    main_procs_lower = [p.lower() for p in MAIN_PROCESSES]
    
    for name_lower, stats in data.items():
        icon_uri = stats.get('icon_data_uri')
        if icon_uri == "failed":
            icon_uri = None
        proc_data = {
            "name": stats.get('original_name', name_lower),
            "status": "Closed",
            "open_total_display": stats.get('total_open', 0.0),
            "focus_total_display": stats.get('total_focused', 0.0),
            "icon": icon_uri,
            "has_icon": bool(icon_uri)
        }
        if name_lower in main_procs_lower:
            main_procs_data.append(proc_data)
        else:
            other_procs_data.append(proc_data)
    
    return jsonify({
        "main_procs": main_procs_data,
        "other_procs": other_procs_data
    })

@app.route('/available_dates')
def get_available_dates():
    dates = []
    for file in os.listdir(CACHE_DIR):
        if file.endswith('.json'):
            dates.append(file[:-5])
    dates.sort(reverse=True)
    return jsonify({"dates": dates})

@app.route('/terminate', methods=['POST'])
def terminate_process():
    name = request.json.get('name')
    if not name:
        return jsonify({'success': False, 'error': 'No name provided'}), 400
    
    try:
        terminated = False
        for p in psutil.process_iter(['name']):
            if p.info['name'] == name:
                p.terminate()
                terminated = True
        return jsonify({'success': terminated})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500
    
@app.route('/reset_timer', methods=['POST'])
def reset_timer():
    name = request.json.get('name')
    if not name:
        return jsonify({'success': False, 'error': 'No name provided'}), 400
    
    with stats_lock:
        name_lower = name.lower()
        if name_lower in process_stats:
            stats = process_stats[name_lower]
            now = time.time()
            stats['total_open'] = 0.0
            stats['total_focused'] = 0.0
            if stats['last_running']:
                stats['running_start'] = now
            if stats['last_focused']:
                stats['focused_start'] = now
            return jsonify({'success': True})
        
    return jsonify({'success': False, 'error': 'Process not found in stats'}), 404

def load_stats_from_today():
    """Loads process stats from today's cache file if it exists."""
    today_str = date.today().strftime("%Y-%m-%d")
    cache_file = os.path.join(CACHE_DIR, f"{today_str}.json")

    if os.path.exists(cache_file):
        try:
            with open(cache_file, 'r') as f:
                data = json.load(f)
                console.print(f"[bold yellow]Loaded previous stats from {cache_file}[/bold yellow]")
                for stats in data.values():
                    stats.update({
                        'last_running': False, 'running_start': None,
                        'last_focused': False, 'focused_start': None
                    })
                return data
        except (json.JSONDecodeError, IOError) as e:
            console.print(f"[bold red]Error loading cache file: {e}. Starting fresh.[/bold red]")
            return {}
    console.print("[bold yellow]No cache file found for today. Starting fresh.[/bold yellow]")
    return {}

def save_stats_worker():
    """Periodically saves the current process stats to a file."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    today_str = date.today().strftime("%Y-%m-%d")
    cache_file = os.path.join(CACHE_DIR, f"{today_str}.json")

    while True:
        time.sleep(SAVE_INTERVAL)
        
        now = time.time()
        current_day = date.today()
        current_day_str = current_day.strftime("%Y-%m-%d")

        if current_day_str != today_str:
            midnight = datetime.combine(current_day, datetime.min.time()).timestamp()
            old_cache_file = cache_file

            with stats_lock:
                data_to_save = {}
                for name_lower, stats in process_stats.items():
                    open_total = stats['total_open']
                    if stats['last_running'] and stats['running_start'] is not None:
                        if stats['running_start'] < midnight:
                            open_total += midnight - stats['running_start']

                    focus_total = stats['total_focused']
                    if stats['last_focused'] and stats['focused_start'] is not None:
                        if stats['focused_start'] < midnight:
                            focus_total += midnight - stats['focused_start']

                    data_to_save[name_lower] = {
                        'original_name': stats.get('original_name', name_lower),
                        'total_open': open_total,
                        'total_focused': focus_total,
                        'icon_data_uri': stats.get('icon_data_uri')
                    }

            try:
                with open(old_cache_file, 'w') as f:
                    json.dump(data_to_save, f, indent=2)
            except IOError as e:
                app_logger.error(f"Could not save stats to {old_cache_file}: {e}")

            with stats_lock:
                for name_lower, stats in process_stats.items():
                    stats['total_open'] = 0.0
                    stats['total_focused'] = 0.0
                    if stats['last_running'] and stats['running_start'] is not None:
                        stats['running_start'] = max(stats['running_start'], midnight)
                    if stats['last_focused'] and stats['focused_start'] is not None:
                        stats['focused_start'] = max(stats['focused_start'], midnight)

            today_str = current_day_str
            cache_file = os.path.join(CACHE_DIR, f"{today_str}.json")

        with stats_lock:
            data_to_save = {}
            for name_lower, stats in process_stats.items():
                open_total = stats['total_open']
                if stats['last_running'] and stats['running_start'] is not None:
                    open_total += now - stats['running_start']

                focus_total = stats['total_focused']
                if stats['last_focused'] and stats['focused_start'] is not None:
                    focus_total += now - stats['focused_start']

                data_to_save[name_lower] = {
                    'original_name': stats.get('original_name', name_lower),
                    'total_open': open_total,
                    'total_focused': focus_total,
                    'icon_data_uri': stats.get('icon_data_uri')
                }

        try:
            with open(cache_file, 'w') as f:
                json.dump(data_to_save, f, indent=2)
        except IOError as e:
            app_logger.error(f"Could not save stats to {cache_file}: {e}")


if __name__ == "__main__":
    process_stats = load_stats_from_today()

    monitor_thread = threading.Thread(target=monitor_processes_worker, daemon=True)
    monitor_thread.start()
    
    save_thread = threading.Thread(target=save_stats_worker, daemon=True)
    save_thread.start()

    discord_thread = threading.Thread(target=update_discord_presence, daemon=True)
    discord_thread.start()

    console.print(f"[bold cyan]Flask server starting...[/bold cyan] Open your browser to [link=http://127.0.0.1:{WEB_SERVER_PORT}]http://1227.0.0.1:{WEB_SERVER_PORT}[/link]")
    app.run(host='0.0.0.0', port=WEB_SERVER_PORT, debug=True)