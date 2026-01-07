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
import gui
import wx

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
    scriptCategory = _("Volume&Power")

    def _get_nircmd_path(self):
        """Get the path to nircmd.exe in the tools folder."""
        try:
            base_path = os.path.dirname(__file__)
            nircmd_path = os.path.join(base_path, "tools", "nircmd.exe")
            if os.path.exists(nircmd_path):
                return nircmd_path
            else:
                logging.error(f"nircmd.exe not found at: {nircmd_path}")
                return None
        except Exception as e:
            logging.error(f"Error getting nircmd path: {str(e)}")
            return None

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

    def _shutdown_with_nircmd(self, reboot=False):
        """Shutdown or restart system using nircmd.exe."""
        try:
            nircmd_path = self._get_nircmd_path()
            if not nircmd_path:
                ui.message(_("Error: nircmd.exe not found"))
                winsound.Beep(500, 500)
                return

            ui.message(_("Restart") if reboot else _("Shutdown"))
            winsound.Beep(100, 100)
            
            # Wait 3 seconds before shutdown/restart
            time.sleep(3)
            
            # Execute nircmd command
            command = "exitwin reboot" if reboot else "exitwin poweroff"
            subprocess.run(
                [nircmd_path, command],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            logging.debug(f"System {'reboot' if reboot else 'shutdown'} initiated via nircmd")
            
        except subprocess.CalledProcessError as e:
            ui.message(_("Error: Failed to execute shutdown command"))
            winsound.Beep(500, 500)
            logging.error(f"nircmd failed: {str(e)}")
        except Exception as e:
            ui.message(_("Unexpected error during shutdown"))
            winsound.Beep(500, 500)
            logging.error(f"Unexpected error in shutdown: {str(e)}")

    def script_restart(self, gesture):
        """Restart the system."""
        threading.Thread(target=self._shutdown_with_nircmd, args=(True,)).start()

    def script_shutdown(self, gesture):
        """Shutdown the system."""
        threading.Thread(target=self._shutdown_with_nircmd, args=(False,)).start()

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