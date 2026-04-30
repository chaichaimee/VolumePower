# __init__.py
# Copyright (C) 2026 Chai Chaimee
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
import sys
import threading  # Added for old method
from ctypes import wintypes
import addonHandler
import winVersion
from core import callLater
import nvwave  # For playing sounds through NVDA wave player
import versionInfo  # Added to check NVDA version

# Initialize translation
addonHandler.initTranslation()

# Safely check NVDA version
try:
	# For NVDA 2025.x
	IS_NVDA2026_OR_NEWER = versionInfo.version_year >= 2026
except AttributeError:
	# For NVDA 2026+ without version_year
	# versionInfo.version is a string e.g., "2026.1"
	year_str = versionInfo.version.split('.')[0]
	IS_NVDA2026_OR_NEWER = int(year_str) >= 2026

# Set up logging
logging.basicConfig(level=logging.DEBUG, filename='nvda_plugin.log')

# -------------------------------------------------------------
# Windows API constants (unchanged)
# -------------------------------------------------------------
SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
TOKEN_ADJUST_PRIVILEGES = 0x0020
TOKEN_QUERY = 0x0008
SE_PRIVILEGE_ENABLED = 0x00000002

EWX_SHUTDOWN = 0x00000001
EWX_REBOOT = 0x00000002
EWX_FORCEIFHUNG = 0x00000010

# -------------------------------------------------------------
# WinAPI structure definitions
# -------------------------------------------------------------
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

# -------------------------------------------------------------
# Set explicit argtypes and restype for Windows API functions
# -------------------------------------------------------------
advapi32 = ctypes.windll.advapi32
kernel32 = ctypes.windll.kernel32
user32 = ctypes.windll.user32

# OpenProcessToken
advapi32.OpenProcessToken.argtypes = [
	wintypes.HANDLE,          # ProcessHandle
	wintypes.DWORD,           # DesiredAccess
	ctypes.POINTER(wintypes.HANDLE)  # TokenHandle
]
advapi32.OpenProcessToken.restype = wintypes.BOOL

# LookupPrivilegeValueW
advapi32.LookupPrivilegeValueW.argtypes = [
	wintypes.LPCWSTR,        # lpSystemName
	wintypes.LPCWSTR,        # lpName
	ctypes.POINTER(LUID)     # lpLuid
]
advapi32.LookupPrivilegeValueW.restype = wintypes.BOOL

# AdjustTokenPrivileges
advapi32.AdjustTokenPrivileges.argtypes = [
	wintypes.HANDLE,                      # TokenHandle
	wintypes.BOOL,                        # DisableAllPrivileges
	ctypes.POINTER(TOKEN_PRIVILEGES),     # NewState
	wintypes.DWORD,                        # BufferLength
	ctypes.POINTER(TOKEN_PRIVILEGES),     # PreviousState (optional)
	ctypes.POINTER(wintypes.DWORD)         # ReturnLength (optional)
]
advapi32.AdjustTokenPrivileges.restype = wintypes.BOOL

# ExitWindowsEx
user32.ExitWindowsEx.argtypes = [
	wintypes.UINT,   # uFlags
	wintypes.DWORD   # dwReason
]
user32.ExitWindowsEx.restype = wintypes.BOOL

# GetCurrentProcess
kernel32.GetCurrentProcess.argtypes = []
kernel32.GetCurrentProcess.restype = wintypes.HANDLE

# GetLastError
kernel32.GetLastError.argtypes = []
kernel32.GetLastError.restype = wintypes.DWORD


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = _("VolumePower")

	# ---------------------------------------------------------
	# Volume control scripts (unchanged)
	# ---------------------------------------------------------
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

	# ---------------------------------------------------------
	# Helper to play exit sound using NVDA's wave player
	# ---------------------------------------------------------
	def _play_exit_sound(self):
		"""Play NVDA's exit.wav file using nvwave to respect user's volume settings."""
		try:
			# List of possible base directories where NVDA might be installed
			possible_base_dirs = []

			# 1. globalVars.appDir (standard NVDA installation)
			try:
				import globalVars
				if hasattr(globalVars, 'appDir') and globalVars.appDir:
					possible_base_dirs.append(globalVars.appDir)
			except ImportError:
				pass

			# 2. Directory of the current executable (nvda.exe)
			if sys.executable:
				exe_dir = os.path.dirname(sys.executable)
				possible_base_dirs.append(exe_dir)
				# If the exe is inside an 'app' subfolder (e.g., portable), try parent
				parent_dir = os.path.dirname(exe_dir)
				if parent_dir != exe_dir:
					possible_base_dirs.append(parent_dir)

			# Remove duplicates
			possible_base_dirs = list(dict.fromkeys(possible_base_dirs))

			# Subfolders to try: 'sounds' (older) and 'waves' (newer/portable)
			subfolders = ['sounds', 'waves']

			for base_dir in possible_base_dirs:
				for sub in subfolders:
					sound_path = os.path.join(base_dir, sub, 'exit.wav')
					if os.path.isfile(sound_path):
						# Play using NVDA's wave player (respects user's volume settings)
						nvwave.playWaveFile(sound_path)
						logging.debug(f"Playing exit sound via nvwave from {sound_path}")
						return True

			logging.warning("exit.wav not found in any searched locations")
		except Exception as e:
			logging.error(f"Failed to play exit sound via nvwave: {e}")
		return False

	# ---------------------------------------------------------
	# System shutdown / restart functions
	# ---------------------------------------------------------
	def _enable_shutdown_privilege(self):
		"""Enable system shutdown privilege using explicit ctypes prototypes."""
		try:
			token = wintypes.HANDLE()
			processHandle = kernel32.GetCurrentProcess()
			if not advapi32.OpenProcessToken(
				processHandle,
				TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY,
				ctypes.byref(token)
			):
				errorCode = kernel32.GetLastError()
				logging.error(f"OpenProcessToken failed, error: {errorCode}")
				return False

			luid = LUID()
			if not advapi32.LookupPrivilegeValueW(
				None,
				SE_SHUTDOWN_NAME,
				ctypes.byref(luid)
			):
				errorCode = kernel32.GetLastError()
				logging.error(f"LookupPrivilegeValueW failed, error: {errorCode}")
				ctypes.windll.kernel32.CloseHandle(token)
				return False

			tp = TOKEN_PRIVILEGES()
			tp.PrivilegeCount = 1
			tp.Privileges[0].Luid = luid
			tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED

			if not advapi32.AdjustTokenPrivileges(
				token,
				False,
				ctypes.byref(tp),
				ctypes.sizeof(tp),
				None,
				None
			):
				errorCode = kernel32.GetLastError()
				logging.error(f"AdjustTokenPrivileges failed, error: {errorCode}")
				ctypes.windll.kernel32.CloseHandle(token)
				return False

			ctypes.windll.kernel32.CloseHandle(token)
			return True

		except Exception as e:
			logging.error(f"Error enabling shutdown privilege: {str(e)}")
			return False

	# --- New method (for NVDA 2026+) ---
	def _perform_shutdown(self, reboot=False):
		"""Shutdown or restart system using Windows API (called on main thread after delay)."""
		try:
			# 1. Play exit sound manually to give user audible feedback (with correct volume)
			self._play_exit_sound()

			# 2. Brief pause to let sound play (approximately duration of exit.wav)
			time.sleep(1.5)

			# 3. Enable required privilege
			if not self._enable_shutdown_privilege():
				ui.message(_("Cannot get shutdown privilege"))
				winsound.Beep(500, 500)
				return

			# 4. Call ExitWindowsEx with correct flags
			flags = (EWX_REBOOT if reboot else EWX_SHUTDOWN) | EWX_FORCEIFHUNG
			if not user32.ExitWindowsEx(flags, 0):
				error = kernel32.GetLastError()
				raise Exception(_("ExitWindowsEx failed with error {error}").format(error=error))

			logging.debug(f"System {'reboot' if reboot else 'shutdown'} initiated via Windows API")

		except Exception as e:
			ui.message(_("Error: {error}").format(error=str(e)))
			winsound.Beep(500, 500)
			logging.error(f"Shutdown failed: {str(e)}")

	# --- Old method (for NVDA 2025.x) copied from old script ---
	def _shutdown_system_old(self, reboot=False):
		"""Shutdown or restart system using Windows API (called in separate thread)."""
		try:
			# Request shutdown privilege
			if not self._enable_shutdown_privilege():
				ui.message(_("Cannot get shutdown privilege"))
				winsound.Beep(500, 500)
				return

			# Notify user and wait 3 seconds
			ui.message(_("Restart") if reboot else _("Shutdown"))
			time.sleep(3)

			# Call Windows API
			flags = (EWX_REBOOT if reboot else EWX_SHUTDOWN) | EWX_FORCEIFHUNG
			if not user32.ExitWindowsEx(flags, 0):
				error = kernel32.GetLastError()
				raise Exception(_("ExitWindowsEx failed with error {error}").format(error=error))

			logging.debug(f"System {'reboot' if reboot else 'shutdown'} initiated via Windows API (old method)")

		except Exception as e:
			ui.message(_("Error: {error}").format(error=str(e)))
			winsound.Beep(500, 500)
			logging.error(f"Shutdown failed (old method): {str(e)}")

	# ---------------------------------------------------------
	# Script methods for restart and shutdown (branch by version)
	# ---------------------------------------------------------
	def script_restart(self, gesture):
		"""Restart the system."""
		if IS_NVDA2026_OR_NEWER:
			# Use new method (play exit.wav then shutdown)
			ui.message(_("Restart"))
			callLater(3000, self._perform_shutdown, reboot=True)
		else:
			# Use old method (no sound, use thread)
			threading.Thread(target=self._shutdown_system_old, args=(True,)).start()

	def script_shutdown(self, gesture):
		"""Shutdown the system."""
		if IS_NVDA2026_OR_NEWER:
			# Use new method
			ui.message(_("Shutdown"))
			callLater(3000, self._perform_shutdown, reboot=False)
		else:
			# Use old method
			threading.Thread(target=self._shutdown_system_old, args=(False,)).start()

	script_restart.__doc__ = _("Restart the system")
	script_restart.category = scriptCategory

	script_shutdown.__doc__ = _("Shutdown the system")
	script_shutdown.category = scriptCategory

	# Gesture bindings (unchanged)
	__gestures = {
		"kb:nvda+shift+pageup": "vol_up",
		"kb:nvda+shift+pagedown": "vol_down",
		"kb:alt+windows+r": "restart",
		"kb:alt+windows+s": "shutdown",
	}