# cpu_affinity_service_groups.py
#
# Usage:
#   python cpu_affinity_service_groups.py install
#   python cpu_affinity_service_groups.py start
#   python cpu_affinity_service_groups.py stop
#   python cpu_affinity_service_groups.py remove
#
# Requirements:
#   pip install pywin32 psutil
#
# Notes:
# - Config file cpu_affinity_config.json in same folder.
# - Config uses "group": "P"|"E"|"ALL" (case-insensitive). Legacy list form [0,1] still accepted.
# - Smart Unlock: expands to ALL logical CPUs when assigned group is saturated.
# - Service must be installed/started with Administrator privileges.

import os
import sys
import json
import time
import threading
import traceback
from datetime import datetime
from typing import Dict, List

import psutil
import pythoncom
import win32event
import win32service
import win32serviceutil
import win32com.client

import ctypes
from ctypes import Structure, c_size_t, c_uint, c_int, c_ubyte

# ---------- Config & defaults ----------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "cpu_affinity_config.json")
LOG_PATH = os.path.join(SCRIPT_DIR, "cpu_affinity_service_groups.log")

SERVICE_NAME = "CpuAffinityGroupService"
SERVICE_DISPLAY_NAME = "CPU Affinity Group Auto-Setter"
SERVICE_DESC = "Set CPU affinity by core groups (P/E/ALL) with optional Smart Unlock."

# Smart Unlock defaults
UNLOCK_THRESHOLD = 95.0    # percent to trigger expand
REVERT_THRESHOLD = 70.0    # percent to revert back (you chose higher)
SUSTAIN_SECONDS = 5        # seconds to sustain before switching
SAMPLE_INTERVAL = 1.0      # seconds between checks

# ---------- Logging ----------
MAX_LOG_SIZE = 5 * 1024 * 1024  # 5 MB
MAX_LOG_LINES = 500              # keep last 500 lines when truncating

def log(msg: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        if os.path.exists(LOG_PATH) and os.path.getsize(LOG_PATH) > MAX_LOG_SIZE:
            # Truncate: keep only the last MAX_LOG_LINES lines
            try:
                with open(LOG_PATH, "r", encoding="utf-8") as f:
                    lines = f.readlines()
                with open(LOG_PATH, "w", encoding="utf-8") as f:
                    f.writelines(lines[-MAX_LOG_LINES:])
            except Exception:
                # If reading fails, just overwrite the file
                with open(LOG_PATH, "w", encoding="utf-8") as f:
                    f.write(f"[{ts}] Log truncated due to size limit.\n")
        # Append new message
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        # Silently ignore logging errors to avoid crashing the service
        pass

# ---------- Helper: physical->logical map (GetLogicalProcessorInformation) ----------
class SYSTEM_LOGICAL_PROCESSOR_INFORMATION(Structure):
    _fields_ = [
        ("ProcessorMask", c_size_t),
        ("Relationship", c_int),
        ("Reserved", c_ubyte * 16)
    ]

def get_physical_to_logical_map() -> List[List[int]]:
    """Return list where each element is list of logical CPU indices for that physical core."""
    kernel32 = ctypes.windll.kernel32

    # first call to get buffer size
    buf_size = c_uint(0)
    res = kernel32.GetLogicalProcessorInformation(None, ctypes.byref(buf_size))
    # allocate buffer
    buf = (ctypes.c_byte * buf_size.value)()
    res = kernel32.GetLogicalProcessorInformation(ctypes.byref(buf), ctypes.byref(buf_size))
    if res == 0:
        err = ctypes.GetLastError()
        raise OSError(f"GetLogicalProcessorInformation failed (error {err})")

    entry_size = ctypes.sizeof(SYSTEM_LOGICAL_PROCESSOR_INFORMATION)
    count = buf_size.value // entry_size
    mapping: List[List[int]] = []
    for i in range(count):
        start = i * entry_size
        entry_buf = (ctypes.c_byte * entry_size).from_buffer(buf, start)
        entry = SYSTEM_LOGICAL_PROCESSOR_INFORMATION.from_buffer(entry_buf)
        # Relationship == 0 => RelationProcessorCore
        if entry.Relationship == 0:
            mask = int(entry.ProcessorMask)
            logical_indices = []
            bit = 0
            m = mask
            while m:
                if m & 1:
                    logical_indices.append(bit)
                bit += 1
                m >>= 1
            mapping.append(logical_indices)
    # sort by lowest logical index to make ordering deterministic
    mapping.sort(key=lambda lst: min(lst) if lst else 9999)
    return mapping

# ---------- Build group lists ----------
def build_groups():
    """
    Returns tuple: (P_group_logical_indices, E_group_logical_indices, ALL_logical_indices)
    Heuristic: if a physical core maps to >1 logical CPU -> classify as P (HT), else E.
    """
    try:
        phys_map = get_physical_to_logical_map()
        p_list = []
        e_list = []
        for logical_indices in phys_map:
            if len(logical_indices) > 1:
                p_list.extend(logical_indices)
            else:
                e_list.extend(logical_indices)
        # dedupe & sort
        p_list = sorted(set(p_list))
        e_list = sorted(set(e_list))
        all_list = sorted(set(p_list + e_list))
        return p_list, e_list, all_list
    except Exception as e:
        log(f"Failed to build physical->logical map: {e}\n{traceback.format_exc()}")
        # fallback: treat first half as P and second half as E if possible
        logical_count = psutil.cpu_count(logical=True) or 1
        half = max(1, logical_count // 2)
        p_list = list(range(0, half))
        e_list = list(range(half, logical_count))
        all_list = list(range(logical_count))
        log(f"Fallback groups: P={p_list}, E={e_list}, ALL={all_list}")
        return p_list, e_list, all_list

# ---------- Config loader ----------
def load_config() -> Dict[str, dict]:
    """
    Normalize to:
      exe_name.lower() -> {"group": "P"|"E"|"ALL", "smart_unlock": bool, "legacy_physical": [ints] (optional)}
    Accepts legacy simple list: "vlc.exe": [0,1] (treated as physical indices; converted to logical later)
    """
    if not os.path.exists(CONFIG_PATH):
        return {}
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            raw = json.load(f)
    except Exception as e:
        log(f"Failed to load config: {e}")
        return {}

    out: Dict[str, dict] = {}
    for k, v in raw.items():
        if not isinstance(k, str):
            continue
        key = k.strip().lower()
        entry = {"group": None, "smart_unlock": False, "legacy_physical": None}
        if isinstance(v, list):
            # legacy physical core list
            if all(isinstance(x, int) for x in v):
                entry["legacy_physical"] = v
                # leave group None; will map later
        elif isinstance(v, dict):
            grp = v.get("group") or v.get("group_name") or v.get("grp")
            if isinstance(grp, str):
                entry["group"] = grp.strip().upper()
            su = v.get("smart_unlock")
            if isinstance(su, bool):
                entry["smart_unlock"] = su
            # also accept legacy physical key
            cores = v.get("cores") or v.get("physical_cores") or v.get("physical")
            if isinstance(cores, list) and all(isinstance(x, int) for x in cores):
                entry["legacy_physical"] = cores
        else:
            log(f"Ignoring config entry for {k}: unsupported value {v}")
            continue
        # only accept entries that define something (group or cores)
        if entry["group"] or entry["legacy_physical"]:
            out[key] = entry
        else:
            log(f"Ignoring config {k}: no 'group' or legacy core list provided.")
    return out

# ---------- Affinity setter ----------
def set_affinity(pid: int, logical_indices: List[int]):
    try:
        p = psutil.Process(pid)
        if not logical_indices:
            log(f"set_affinity: empty logical_indices for PID {pid}; skipping.")
            return
        p.cpu_affinity(logical_indices)
        log(f"Set affinity: PID={pid} ({p.name()}) -> logical CPUs {logical_indices}")
    except psutil.NoSuchProcess:
        log(f"Process {pid} disappeared before setting affinity.")
    except psutil.AccessDenied:
        log(f"AccessDenied setting affinity for PID {pid}. Service needs sufficient privileges.")
    except Exception as e:
        log(f"Error setting affinity for PID {pid}: {e}\n{traceback.format_exc()}")

# ---------- Smart Unlock monitor (per-process) ----------
class SmartMonitor(threading.Thread):
    def __init__(self, pid: int, base_logical: List[int], all_logical: List[int]):
        super().__init__(daemon=True)
        self.pid = pid
        self.base_logical = list(base_logical)
        self.all_logical = list(all_logical)
        self.stop_event = threading.Event()
        self.expanded = False
        self.over_count = 0
        self.under_count = 0

    def run(self):
        try:
            proc = psutil.Process(self.pid)
        except psutil.NoSuchProcess:
            return
        log(f"SmartMonitor started for PID {self.pid}, base logical {self.base_logical}")
        # prime psutil counters
        try:
            psutil.cpu_percent(interval=None, percpu=True)
        except Exception:
            pass

        while not self.stop_event.is_set():
            if not proc.is_running():
                break
            # sample twice over SAMPLE_INTERVAL to get meaningful values
            try:
                time.sleep(SAMPLE_INTERVAL)
                per_cpu = psutil.cpu_percent(interval=None, percpu=True)
            except Exception:
                # on error, sleep and continue
                time.sleep(SAMPLE_INTERVAL)
                continue

            vals = [per_cpu[i] for i in self.base_logical if i < len(per_cpu)]
            avg = (sum(vals) / len(vals)) if vals else 0.0

            if not self.expanded:
                if avg >= UNLOCK_THRESHOLD:
                    self.over_count += 1
                else:
                    self.over_count = 0
                if self.over_count >= SUSTAIN_SECONDS:
                    # expand to all logical CPUs
                    set_affinity(self.pid, self.all_logical)
                    self.expanded = True
                    log(f"SmartUnlock: PID {self.pid} expanded to ALL (avg {avg:.1f}%)")
                    self.over_count = 0
                    self.under_count = 0
            else:
                if avg <= REVERT_THRESHOLD:
                    self.under_count += 1
                else:
                    self.under_count = 0
                if self.under_count >= SUSTAIN_SECONDS:
                    set_affinity(self.pid, self.base_logical)
                    self.expanded = False
                    log(f"SmartUnlock: PID {self.pid} reverted to base (avg {avg:.1f}%)")
                    self.over_count = 0
                    self.under_count = 0

        log(f"SmartMonitor stopping for PID {self.pid}")

    def stop(self):
        self.stop_event.set()

# ---------- Service ----------
class GroupAffinityService(win32serviceutil.ServiceFramework):
    _svc_name_ = SERVICE_NAME
    _svc_display_name_ = SERVICE_DISPLAY_NAME
    _svc_description_ = SERVICE_DESC

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.stop_requested = False
        self.thread = None

        # runtime
        self.p_group = []
        self.e_group = []
        self.all_group = []
        self.config = {}
        self.monitors: Dict[int, SmartMonitor] = {}
        self.monitor_lock = threading.Lock()

    def SvcStop(self):
        log("Service stop requested.")
        self.stop_requested = True
        win32event.SetEvent(self.hWaitStop)
        # stop monitors
        with self.monitor_lock:
            for m in list(self.monitors.values()):
                m.stop()
            for m in list(self.monitors.values()):
                m.join(timeout=2)
        if self.thread and self.thread.is_alive():
            self.thread.join(timeout=5)
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        log("Service stopping...")

    def SvcDoRun(self):
        log("Service starting...")
        self.thread = threading.Thread(target=self.main, daemon=True)
        self.thread.start()
        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)

    def main(self):
        pythoncom.CoInitialize()
        try:
            # build groups
            self.p_group, self.e_group, self.all_group = build_groups()
            log(f"Detected groups: P={self.p_group}, E={self.e_group}, ALL={self.all_group}")

            # load config
            raw_cfg = load_config()
            if not raw_cfg:
                log("Config empty or missing. Place cpu_affinity_config.json in the service folder.")
            else:
                log(f"Loaded config: {raw_cfg}")

            # normalize config to resolved logical CPU lists per entry
            normalized = {}
            for exe, entry in raw_cfg.items():
                grp = entry.get("group")
                legacy = entry.get("legacy_physical")
                smart = bool(entry.get("smart_unlock", False))
                logical = []
                if grp:
                    g = grp.upper()
                    if g == "P":
                        logical = list(self.p_group)
                    elif g == "E":
                        logical = list(self.e_group)
                    elif g == "ALL":
                        logical = list(self.all_group)
                    else:
                        log(f"Unknown group '{grp}' for {exe}; ignored.")
                elif legacy:
                    # legacy physical indices -> map via physical->logical (reuse earlier heuristic)
                    phys_map = get_physical_to_logical_map()
                    for pidx in legacy:
                        if 0 <= pidx < len(phys_map):
                            logical.extend(phys_map[pidx])
                        else:
                            log(f"Legacy core index {pidx} out of range; ignored for {exe}.")
                logical = sorted(set(logical))
                if logical:
                    normalized[exe] = {"logical": logical, "smart_unlock": smart}
                else:
                    log(f"No logical CPUs resolved for {exe}; entry ignored.")
            self.config = normalized
            log(f"Normalized config: {self.config}")

            # apply to currently running matching processes
            try:
                for proc in psutil.process_iter(["pid", "name"]):
                    name = proc.info.get("name") or ""
                    pid = proc.info.get("pid")
                    if not name or pid is None:
                        continue
                    key = name.lower()
                    if key in self.config:
                        entry = self.config[key]
                        set_affinity(pid, entry["logical"])
                        if entry["smart_unlock"]:
                            self.spawn_monitor(pid, entry["logical"])
            except Exception as e:
                log(f"Error applying affinities at startup: {e}\n{traceback.format_exc()}")

            # setup WMI watcher for new processes
            try:
                locator = win32com.client.Dispatch("WbemScripting.SWbemLocator")
                service = locator.ConnectServer(".", "root\\cimv2")
                query = "SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'"
                watcher = service.ExecNotificationQuery(query)
                log("WMI watcher started.")
            except Exception as e:
                log(f"Failed to start WMI watcher: {e}\n{traceback.format_exc()}")
                watcher = None

            last_cfg_check = time.time()
            while not self.stop_requested:
                # reload config every 5s
                if time.time() - last_cfg_check > 5:
                    try:
                        raw_new = load_config()
                        # quick normalization same as above (lightweight)
                        new_norm = {}
                        for exe, entry in raw_new.items():
                            grp = entry.get("group")
                            legacy = entry.get("legacy_physical")
                            smart = bool(entry.get("smart_unlock", False))
                            logical = []
                            if grp:
                                g = grp.upper()
                                if g == "P":
                                    logical = list(self.p_group)
                                elif g == "E":
                                    logical = list(self.e_group)
                                elif g == "ALL":
                                    logical = list(self.all_group)
                            elif legacy:
                                phys_map = get_physical_to_logical_map()
                                for pidx in legacy:
                                    if 0 <= pidx < len(phys_map):
                                        logical.extend(phys_map[pidx])
                            logical = sorted(set(logical))
                            if logical:
                                new_norm[exe] = {"logical": logical, "smart_unlock": smart}
                        if new_norm != self.config:
                            log(f"Config change detected. New normalized config: {new_norm}")
                            self.config = new_norm
                        last_cfg_check = time.time()
                    except Exception as e:
                        log(f"Error reloading config: {e}")

                if watcher is None:
                    time.sleep(1)
                    continue

                try:
                    evt = watcher.NextEvent(1000)
                except Exception:
                    continue

                if not evt:
                    continue

                try:
                    tgt = evt.Properties_("TargetInstance").Value
                    proc_name = getattr(tgt, "Name", "") or ""
                    proc_pid = getattr(tgt, "ProcessId", None)
                    if proc_name:
                        key = proc_name.lower()
                        if key in self.config:
                            entry = self.config[key]
                            base_logical = entry["logical"]
                            smart = entry["smart_unlock"]
                            attempts = 0
                            success = False
                            while attempts < 6 and not success and not self.stop_requested:
                                try:
                                    set_affinity(int(proc_pid), base_logical)
                                    success = True
                                    if smart:
                                        self.spawn_monitor(int(proc_pid), base_logical)
                                except Exception:
                                    attempts += 1
                                    time.sleep(0.5)
                            if not success:
                                log(f"Failed to set affinity for {proc_name} PID {proc_pid}")
                except Exception as e:
                    log(f"Error handling WMI event: {e}\n{traceback.format_exc()}")

            # cleanup on exit
            with self.monitor_lock:
                for m in list(self.monitors.values()):
                    m.stop()
                for m in list(self.monitors.values()):
                    m.join(timeout=2)

            log("Service main loop exiting.")
        finally:
            pythoncom.CoUninitialize()

    def spawn_monitor(self, pid: int, base_logical: List[int]):
        with self.monitor_lock:
            if pid in self.monitors:
                return
            if not psutil.pid_exists(pid):
                return
            m = SmartMonitor(pid, base_logical, self.all_group)
            self.monitors[pid] = m
            m.start()
            # cleanup thread to remove monitor after it terminates
            def cleaner():
                m.join()
                with self.monitor_lock:
                    if pid in self.monitors:
                        del self.monitors[pid]
            t = threading.Thread(target=cleaner, daemon=True)
            t.start()

if __name__ == "__main__":
    if len(sys.argv) == 1:
        log("Console mode (debug).")
        svc = GroupAffinityService(sys.argv)
        try:
            svc.main()
        except KeyboardInterrupt:
            log("Interrupted by user.")
    else:
        win32serviceutil.HandleCommandLine(GroupAffinityService)
