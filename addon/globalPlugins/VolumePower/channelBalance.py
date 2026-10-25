# channelBalance.py

import gc
import logging
import queue
import struct
import threading
from ctypes import POINTER, cast
from typing import Callable, NamedTuple, Optional, Tuple

import pycaw
from comtypes import CLSCTX_ALL, CoInitialize, CoUninitialize
from logHandler import log
from pycaw.utils import AudioUtilities, IAudioEndpointVolume

LEFT = 0
RIGHT = 1
STEP_PERCENT = 5

STATUS_OK = "ok"
STATUS_MUTED = "muted"
STATUS_NOT_STEREO = "notStereo"
STATUS_FAILED = "failed"

LIMIT_NONE = ""
LIMIT_MAXIMUM = "maximum"
LIMIT_MINIMUM = "minimum"


class BalanceResult(NamedTuple):
	status: str
	left: int
	right: int
	limit: str = LIMIT_NONE


class _Request(NamedTuple):
	kind: str
	side: int
	delta: int
	equalize: bool
	refresh: bool
	onDone: Optional[Callable[[BalanceResult], None]]


def toPercent(scalar: float) -> int:
	return max(0, min(100, int(round(scalar * 100))))


def steppedLevels(left: int, right: int, side: int, delta: int) -> Tuple[int, int]:
	if side == LEFT:
		return max(0, min(100, left + delta)), right
	return left, max(0, min(100, right + delta))


def equalizedLevels(left: int, right: int) -> Tuple[int, int]:
	louder = max(left, right)
	return louder, louder


def limitReached(delta: int, before: Tuple[int, int], after: Tuple[int, int]) -> str:
	if delta == 0 or after != before:
		return LIMIT_NONE
	return LIMIT_MAXIMUM if delta > 0 else LIMIT_MINIMUM


def _openEndpoint():
	device = AudioUtilities.GetSpeakers()
	# Some pycaw releases return a wrapper object instead of the raw IMMDevice.
	raw = device if hasattr(device, "Activate") else getattr(device, "_dev", None)
	if raw is None:
		raise RuntimeError("Unsupported pycaw device object: %s" % type(device).__name__)
	interface = raw.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
	return cast(interface, POINTER(IAudioEndpointVolume))


def _readLevels(endpoint) -> Tuple[int, int]:
	return (
		toPercent(endpoint.GetChannelVolumeLevelScalar(LEFT)),
		toPercent(endpoint.GetChannelVolumeLevelScalar(RIGHT)),
	)


def _apply(endpoint, side: int, delta: int, equalize: bool) -> BalanceResult:
	if endpoint.GetChannelCount() < 2:
		return BalanceResult(STATUS_NOT_STEREO, 0, 0)
	if endpoint.GetMute():
		return BalanceResult(STATUS_MUTED, 0, 0)
	before = _readLevels(endpoint)
	if equalize:
		after = equalizedLevels(*before)
	else:
		after = steppedLevels(before[0], before[1], side, delta)
	# Each channel scalar is the level of that ear as Windows shows it; the Windows
	# device volume follows the louder channel, so only the channels are written.
	if after[0] != before[0]:
		endpoint.SetChannelVolumeLevelScalar(LEFT, after[0] / 100.0, None)
	if after[1] != before[1]:
		endpoint.SetChannelVolumeLevelScalar(RIGHT, after[1] / 100.0, None)
	# Report what Windows now holds, not what was requested.
	left, right = _readLevels(endpoint)
	if log.isEnabledFor(logging.DEBUG):
		log.debug(
			"Ear levels: device %.3f, left %.3f, right %.3f",
			endpoint.GetMasterVolumeLevelScalar(),
			endpoint.GetChannelVolumeLevelScalar(LEFT),
			endpoint.GetChannelVolumeLevelScalar(RIGHT),
		)
	limit = LIMIT_NONE if equalize else limitReached(delta, before, after)
	return BalanceResult(STATUS_OK, left, right, limit)


class _EarVolumeWorker:
	"""Runs every endpoint call on one long-lived thread that owns its COM apartment.

	COM pointers must be released while the apartment that created them is still
	initialized. Short-lived threads that called CoUninitialize right after each key press
	left pointers behind for the garbage collector, which then released them on an
	unrelated thread and logged access violations. One thread that stays initialized until
	shutdown, and that keeps the endpoint only while the ear layer is open, avoids that.
	"""

	def __init__(self) -> None:
		self._requests: "queue.Queue[Optional[_Request]]" = queue.Queue()
		self._startLock = threading.Lock()
		self._thread: Optional[threading.Thread] = None

	def submit(self, request: _Request) -> None:
		with self._startLock:
			if self._thread is None or not self._thread.is_alive():
				self._thread = threading.Thread(target=self._run, daemon=True, name="VolumePowerEar")
				self._thread.start()
		self._requests.put(request)

	def stop(self) -> None:
		with self._startLock:
			if self._thread is None:
				return
			self._requests.put(None)
			self._thread = None

	def _run(self) -> None:
		CoInitialize()
		endpoint = None
		try:
			while True:
				request = self._requests.get()
				if request is None:
					break
				if request.kind == "release":
					endpoint = None
					continue
				endpoint, result = self._handle(endpoint, request)
				if request.onDone is not None:
					request.onDone(result)
		finally:
			endpoint = None
			# Cyclic garbage can still hold a pointer; collect it here, before this
			# thread's apartment is torn down, instead of on whichever thread runs the collector.
			gc.collect()
			CoUninitialize()

	@staticmethod
	def _handle(endpoint, request: _Request):
		if request.refresh:
			endpoint = None
		lastError = None
		for _attempt in (1, 2):
			try:
				if endpoint is None:
					endpoint = _openEndpoint()
				return endpoint, _apply(endpoint, request.side, request.delta, request.equalize)
			except Exception as endpointError:
				# comtypes raises COMError, which does not derive from OSError.
				lastError = endpointError
				endpoint = None
		log.error(
			"Ear volume change failed (pycaw %s, python %s-bit): %r",
			getattr(pycaw, "__file__", "unknown"),
			struct.calcsize("P") * 8,
			lastError,
			exc_info=lastError,
		)
		return None, BalanceResult(STATUS_FAILED, 0, 0)


_worker = _EarVolumeWorker()


def submitChange(
	side: int,
	delta: int,
	equalize: bool,
	refresh: bool,
	onDone: Callable[[BalanceResult], None],
) -> None:
	"""Queue an ear-volume change; onDone runs on the worker thread, so it must marshal to the UI thread.

	With delta 0 and equalize False nothing is written and the current levels are reported.
	refresh=True makes the worker look up the current default output device again.
	"""
	_worker.submit(_Request("change", side, delta, equalize, refresh, onDone))


def releaseEndpoint() -> None:
	"""Drop the cached endpoint on the worker thread when the ear layer closes."""
	_worker.submit(_Request("release", LEFT, 0, False, False, None))


def shutdown() -> None:
	_worker.stop()
