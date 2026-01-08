# __init__.py
# Copyright (C) 2026 'Chai Chaimee
# Licensed under GNU General Public License. See COPYING.txt for details.

import globalPluginHandler
import ui
import subprocess
import os
import winsound
import logging
import synthDriverHandler
import config
import ctypes
import time
import threading
from ctypes import wintypes
import addonHandler

# Initialize translation
addonHandler.initTranslation()

# Set up logging
logging.basicConfig(level=logging.DEBUG, filename='nvda_plugin.log')

# Define Windows API structures
class LUID(ctypes.Structure):
    _fields_ = [
        ("LowPart", wintypes.DWORD),
        ("HighPart", wintypes.LONG),
    ]

class LUID_AND_ATTRIBUTES(ctypes.Structure):
    _fields_ = [
        ("Luid", LUID),
        ("Attributes", wintypes.DWORD),
    ]

class TOKEN_PRIVILEGES(ctypes.Structure):
    _fields_ = [
        ("PrivilegeCount", wintypes.DWORD),
        ("Privileges", LUID_AND_ATTRIBUTES * 1),
    ]

SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
TOKEN_ADJUST_PRIVILEGES = 0x0020
TOKEN_QUERY = 0x0008
SE_PRIVILEGE_ENABLED = 0x00000002

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = _("VolumePower")

    def script_vol_up(self, gesture):
        """Increase NVDA speech volume by 5%."""
        try:
            synth = synthDriverHandler.getSynth()
            if not synth:
                ui.message(_("Error: No synthesizer available"))
                logging.error("No synthesizer available for volume adjustment")
                return
            current_volume = synth.volume
            new_volume = min(current_volume + 5, 100)
            synth.volume = new_volume
            config.conf["speech"][synth.name]["volume"] = new_volume
            config.conf.save()
            ui.message(_("{volume}%").format(volume=new_volume))
            logging.debug(f"NVDA volume increased to {new_volume}% and saved")
        except Exception as e:
            ui.message(_("Error: Failed to increase NVDA volume"))
            logging.error(f"Failed to increase NVDA volume: {str(e)}")

    script_vol_up.__doc__ = _("Increase NVDA volume 5%")
    script_vol_up.category = scriptCategory

    def script_vol_down(self, gesture):
        """Decrease NVDA speech volume by 5%."""
        try:
            synth = synthDriverHandler.getSynth()
            if not synth:
                ui.message(_("Error: No synthesizer available"))
                logging.error("No synthesizer available for volume adjustment")
                return
            current_volume = synth.volume
            new_volume = max(current_volume - 5, 0)
            synth.volume = new_volume
            config.conf["speech"][synth.name]["volume"] = new_volume
            config.conf.save()
            ui.message(_("{volume}%").format(volume=new_volume))
            logging.debug(f"NVDA volume decreased to {new_volume}% and saved")
        except Exception as e:
            ui.message(_("Error: Failed to decrease NVDA volume"))
            logging.error(f"Failed to decrease NVDA volume: {str(e)}")

    script_vol_down.__doc__ = _("Decrease NVDA volume 5%")
    script_vol_down.category = scriptCategory

    def _enable_shutdown_privilege(self):
        """Enable system shutdown privilege."""
        try:
            token = wintypes.HANDLE()
            if not ctypes.windll.advapi32.OpenProcessToken(
                ctypes.windll.kernel32.GetCurrentProcess(),
                TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY,
                ctypes.byref(token)
            ):
                return False

            luid = LUID()
            if not ctypes.windll.advapi32.LookupPrivilegeValueW(
                None,
                SE_SHUTDOWN_NAME,
                ctypes.byref(luid)
            ):
                return False

            tp = TOKEN_PRIVILEGES()
            tp.PrivilegeCount = 1
            tp.Privileges[0].Luid = luid
            tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED

            if not ctypes.windll.advapi32.AdjustTokenPrivileges(
                token,
                False,
                ctypes.byref(tp),
                ctypes.sizeof(tp),
                None,
                None
            ):
                return False

            return True
        except Exception as e:
            logging.error(f"Error enabling shutdown privilege: {str(e)}")
            return False

    def _shutdown_system(self, reboot=False):
        """Shutdown or restart system using Windows API."""
        try:
            # Windows API constants
            EWX_SHUTDOWN = 0x00000001
            EWX_REBOOT = 0x00000002
            EWX_FORCEIFHUNG = 0x00000010

            # Request shutdown privilege
            if not self._enable_shutdown_privilege():
                ui.message(_("Cannot get shutdown privilege"))
                winsound.Beep(500, 500)
                return

            # Notify user and wait
            ui.message(_("Restart") if reboot else _("Shutdown"))
            time.sleep(3)

            # Call Windows API
            flags = (EWX_REBOOT if reboot else EWX_SHUTDOWN) | EWX_FORCEIFHUNG
            if not ctypes.windll.user32.ExitWindowsEx(flags, 0):
                error = ctypes.windll.kernel32.GetLastError()
                raise Exception(_("ExitWindowsEx failed with error {error}").format(error=error))

            logging.debug(f"System {'reboot' if reboot else 'shutdown'} initiated via Windows API")

        except Exception as e:
            ui.message(_("Error: {error}").format(error=str(e)))
            winsound.Beep(500, 500)
            logging.error(f"Shutdown failed: {str(e)}")

    def script_restart(self, gesture):
        """Restart the system."""
        threading.Thread(target=self._shutdown_system, args=(True,)).start()

    def script_shutdown(self, gesture):
        """Shutdown the system."""
        threading.Thread(target=self._shutdown_system, args=(False,)).start()

    script_restart.__doc__ = _("Restart the system")
    script_restart.category = scriptCategory
    
    script_shutdown.__doc__ = _("Shutdown the system")
    script_shutdown.category = scriptCategory

    # Define gestures
    __gestures = {
        "kb:nvda+shift+pageup": "vol_up",
        "kb:nvda+shift+pagedown": "vol_down",
        "kb:alt+windows+r": "restart",
        "kb:alt+windows+s": "shutdown",
    }