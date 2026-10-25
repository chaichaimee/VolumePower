# addonCore.py

import time
import winsound
from threading import Thread

import config
import core
import globalPluginHandler
import gui
import synthDriverHandler
import tones
import ui
import wx
from logHandler import log

from . import audioGuard, channelBalance, powerControl
from .constants import ADDON_SUMMARY
from .countdownDialog import PowerCountdownDialog
from .settings import AudioGuardSettingsPanel

POWER_TAP_THRESHOLD_SECONDS = 0.3
EAR_BEEP_LENGTH_MS = 80
EAR_BEEP_VOLUME = 60

EAR_LAYER_GESTURES = {
	"kb:upArrow": "earUp",
	"kb:downArrow": "earDown",
	"kb:leftArrow": "earLeft",
	"kb:rightArrow": "earRight",
	"kb:enter": "earEqualize",
	"kb:escape": "earExit",
}


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	"""Implementation of global commands for volume control, power control, and audio protection."""
	scriptCategory: str = ADDON_SUMMARY

	def __init__(self, *args, **kwargs) -> None:
		"""Initialization of the add-on: registers settings, starts audio guard checks."""
		super().__init__(*args, **kwargs)
		self._powerTapCount = 0
		self._powerLastTapTime = 0.0
		self._isActive = True
		self._earLayerActive = False
		self._earSide = channelBalance.LEFT
		self._audioGuardEngine = audioGuard.getActiveEngine()
		gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(AudioGuardSettingsPanel)
		Thread(target=self._audioGuardEngine.runStartupChecks, daemon=True).start()
		self._audioGuardEngine.applySleepPrevention()
		if self._audioGuardEngine.settings["reinit"]:
			self._audioGuardEngine.resetSynth()

	def terminate(self, *args, **kwargs) -> None:
		"""This will be called when NVDA is finished with this global plugin."""
		self._isActive = False
		self._exitEarLayer()
		channelBalance.shutdown()
		super().terminate(*args, **kwargs)
		# Stops the self-rescheduling synth-reset loop before anything else, so it
		# cannot keep calling synthDriverHandler.setSynth() against a plugin instance
		# that no longer owns valid state after an add-on disable/reload.
		self._audioGuardEngine.terminate()
		self._audioGuardEngine.releaseSleepPrevention()
		audioGuard.clearActiveEngine()
		try:
			gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(AudioGuardSettingsPanel)
		except ValueError:
			log.debug("Could not remove %s settings panel from NVDA settings dialogs", ADDON_SUMMARY)

	# ---------------------------------------------------------
	# Volume control scripts
	# ---------------------------------------------------------
	def script_vol_up(self, gesture) -> None:
		"""Increase NVDA speech volume by 5%."""
		self._adjustSynthVolume(step=5)

	script_vol_up.__doc__ = _("Increase NVDA volume 5%")
	script_vol_up.category = scriptCategory

	def script_vol_down(self, gesture) -> None:
		"""Decrease NVDA speech volume by 5%."""
		self._adjustSynthVolume(step=-5)

	script_vol_down.__doc__ = _("Decrease NVDA volume 5%")
	script_vol_down.category = scriptCategory

	def _adjustSynthVolume(self, step: int) -> None:
		"""Adjust the current synthesizer's volume by the given signed step and persist it."""
		try:
			synth = synthDriverHandler.getSynth()
			if not synth:
				ui.message(_("Error: No synthesizer available"))
				log.debug("No synthesizer available for volume adjustment")
				return
			newVolume = min(max(synth.volume + step, 0), 100)
			synth.volume = newVolume
			config.conf["speech"][synth.name]["volume"] = newVolume
			config.conf.save()
			ui.message(_("{volume}%").format(volume=newVolume))
			log.debug("NVDA volume changed to %s%% and saved", newVolume)
		except (AttributeError, KeyError, OSError) as volumeError:
			ui.message(_("Error: Failed to change NVDA volume"))
			log.debug("Failed to change NVDA volume: %s", volumeError)

	# ---------------------------------------------------------
	# Left/right ear volume: Shift+NVDA+NumPad Plus opens a layer that stays open until another key is pressed
	# ---------------------------------------------------------
	def getScript(self, gesture):
		script = super().getScript(gesture)
		if script is None and self._earLayerActive and not getattr(gesture, "isModifier", False):
			self._exitEarLayer()
		return script

	def script_earLayer(self, gesture) -> None:
		self._openEarLayer()

	script_earLayer.__doc__ = _(
		"Open the ear volume layer and announce the selected ear: Left and Right arrows select the ear, "
		"Up and Down arrows change it by 5, Enter makes both ears equal, Escape closes. "
		"Any other key also closes the layer"
	)
	script_earLayer.category = scriptCategory

	def script_earLeft(self, gesture) -> None:
		self._earSide = channelBalance.LEFT
		self._runEarChange(0, False, False)

	def script_earRight(self, gesture) -> None:
		self._earSide = channelBalance.RIGHT
		self._runEarChange(0, False, False)

	def script_earUp(self, gesture) -> None:
		self._runEarChange(channelBalance.STEP_PERCENT, False, False)

	def script_earDown(self, gesture) -> None:
		self._runEarChange(-channelBalance.STEP_PERCENT, False, False)

	def script_earEqualize(self, gesture) -> None:
		self._runEarChange(0, True, False)

	def script_earExit(self, gesture) -> None:
		self._exitEarLayer()
		ui.message(_("Ear volume closed"))

	def _openEarLayer(self) -> None:
		if not self._earLayerActive:
			self.bindGestures(EAR_LAYER_GESTURES)
			self._earLayerActive = True
		self._runEarChange(0, False, True)

	def _exitEarLayer(self) -> None:
		if not self._earLayerActive:
			return
		self._earLayerActive = False
		self.bindGestures({gestureId: None for gestureId in EAR_LAYER_GESTURES})
		channelBalance.releaseEndpoint()

	def _runEarChange(self, delta: int, equalize: bool, refresh: bool) -> None:
		side = self._earSide

		def onDone(result) -> None:
			wx.CallAfter(self._announceEarResult, side, delta, equalize, result)

		channelBalance.submitChange(side, delta, equalize, refresh, onDone)

	def _announceEarResult(self, side: int, delta: int, equalize: bool, result) -> None:
		if not self._isActive:
			return
		if result.status != channelBalance.STATUS_OK:
			self._exitEarLayer()
			if result.status == channelBalance.STATUS_MUTED:
				ui.message(_("Audio is muted"))
			elif result.status == channelBalance.STATUS_NOT_STEREO:
				ui.message(_("The output device is not stereo"))
			else:
				ui.message(_("Failed to change ear volume"))
			return
		if equalize:
			ui.message(_("Both ears {level}").format(level=result.left))
			self._earBeep(result.left, EAR_BEEP_VOLUME, EAR_BEEP_VOLUME)
			return
		isLeft = side == channelBalance.LEFT
		level = result.left if isLeft else result.right
		earLabel = _("Left") if isLeft else _("Right")
		if result.limit == channelBalance.LIMIT_MAXIMUM:
			ui.message(_("{ear} {level}, maximum").format(ear=earLabel, level=level))
		elif result.limit == channelBalance.LIMIT_MINIMUM:
			ui.message(_("{ear} {level}, minimum").format(ear=earLabel, level=level))
		else:
			ui.message(_("{ear} {level}").format(ear=earLabel, level=level))
		self._earBeep(level, EAR_BEEP_VOLUME if isLeft else 0, 0 if isLeft else EAR_BEEP_VOLUME)

	def _earBeep(self, level: int, leftVolume: int, rightVolume: int) -> None:
		tones.beep(400 + level * 5, EAR_BEEP_LENGTH_MS, leftVolume, rightVolume)

	# ---------------------------------------------------------
	# Power control script: single tap shuts down, double tap restarts
	# ---------------------------------------------------------
	def script_power(self, gesture) -> None:
		"""Shut down the system on a single tap, restart it on a double tap."""
		currentTime = time.time()
		if currentTime - self._powerLastTapTime > POWER_TAP_THRESHOLD_SECONDS:
			self._powerTapCount = 0
		self._powerTapCount += 1
		self._powerLastTapTime = currentTime
		wx.CallLater(int(POWER_TAP_THRESHOLD_SECONDS * 1000), self._resolvePowerTap)

	script_power.__doc__ = _("Shut down the system (single tap) or restart it (double tap)")
	script_power.category = scriptCategory

	def _resolvePowerTap(self) -> None:
		"""Decide, once the multi-tap window has elapsed, whether to shut down or restart."""
		tapCount = self._powerTapCount
		if tapCount == 0:
			return
		self._powerTapCount = 0
		reboot = tapCount >= 2
		if self._audioGuardEngine.settings["dialogCounterEnabled"]:
			wx.CallAfter(self._showPowerCountdown, reboot)
		else:
			wx.CallAfter(self._executeShutdown, reboot)

	def _showPowerCountdown(self, reboot: bool) -> None:
		"""Show the cancellable 3-2-1 countdown dialog for the requested power action."""
		actionLabel = _("Restart") if reboot else _("Shutdown")
		PowerCountdownDialog(
			gui.mainFrame,
			actionLabel=actionLabel,
			onConfirm=lambda: self._executeShutdown(reboot),
			onCancel=lambda: ui.message(_("Cancelled")),
		).Show()

	def _executeShutdown(self, reboot: bool) -> None:
		"""Run the version-appropriate shutdown or restart sequence after countdown confirmation."""
		ui.message(_("Restart") if reboot else _("Shutdown"))
		if powerControl.IS_NVDA2026_OR_NEWER:
			core.callLater(50, self._runShutdownNewMethod, reboot)
		else:
			Thread(target=self._runShutdownOldMethod, args=(reboot,), daemon=True).start()

	def _runShutdownNewMethod(self, reboot: bool) -> None:
		"""Play the exit sound then shut down or restart (NVDA 2026 and newer).

		Runs entirely on the main thread via core.callLater scheduling. The wait after
		the exit sound must not use time.sleep, since that would freeze NVDA's
		single-threaded event pump and trigger a Watchdog timeout; the remaining delay
		is instead scheduled as a second, non-blocking core.callLater callback.
		"""
		powerControl.playExitSound()
		core.callLater(1500, self._finishShutdownNewMethod, reboot)

	def _finishShutdownNewMethod(self, reboot: bool) -> None:
		"""Perform the actual shutdown or restart once the exit sound delay has elapsed."""
		try:
			powerControl.performShutdown(reboot)
		except RuntimeError as shutdownError:
			self._reportShutdownFailure(shutdownError, "Shutdown failed: %s")

	def _runShutdownOldMethod(self, reboot: bool) -> None:
		"""Shut down or restart without the exit sound step (NVDA 2025.x).

		Runs on a background thread. ui.message and winsound.Beep must not be called
		directly from here, since both interact with NVDA's main-thread speech queue
		and core UI state; on failure the report is marshaled to the main thread via
		wx.CallAfter instead.
		"""
		try:
			powerControl.performShutdown(reboot)
		except RuntimeError as shutdownError:
			wx.CallAfter(self._reportShutdownFailure, shutdownError, "Shutdown failed (old method): %s")

	def _reportShutdownFailure(self, shutdownError: RuntimeError, logMessage: str) -> None:
		"""Report a failed shutdown or restart attempt. Must run on the main thread."""
		ui.message(str(shutdownError))
		winsound.Beep(500, 500)
		log.debug(logMessage, shutdownError)

	__gestures = {
		"kb:nvda+shift+pageup": "vol_up",
		"kb:nvda+shift+pagedown": "vol_down",
		"kb:alt+windows+s": "power",
		"kb:nvda+shift+numpadPlus": "earLayer",
	}
