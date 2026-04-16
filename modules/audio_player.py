from __future__ import annotations

from pathlib import Path
import platform
import subprocess
import time
import wave


class AudioPlayer:
    def __init__(self) -> None:
        self._process: subprocess.Popen | None = None
        self._current_path: Path | None = None
        self._is_windows = platform.system() == "Windows"
        self._started_at: float | None = None
        self._duration_ms: int | None = None
        if self._is_windows:
            import winsound

            self._winsound = winsound
        else:
            self._winsound = None

    @property
    def current_path(self) -> Path | None:
        return self._current_path

    def current_position_ms(self) -> int:
        if self._started_at is None:
            return 0
        elapsed = int((time.monotonic() - self._started_at) * 1000)
        if self._duration_ms is not None:
            return max(0, min(elapsed, self._duration_ms))
        return max(0, elapsed)

    def is_playing(self) -> bool:
        if self._is_windows:
            return self._current_path is not None
        return self._process is not None and self._process.poll() is None

    def play(self, path: Path) -> None:
        self.stop()
        self._current_path = path
        self._started_at = time.monotonic()
        self._duration_ms = self._read_duration_ms(path)

        if self._is_windows and self._winsound is not None:
            self._winsound.PlaySound(
                str(path),
                self._winsound.SND_FILENAME | self._winsound.SND_ASYNC,
            )
            return

        command = ["afplay", str(path)] if platform.system() == "Darwin" else ["ffplay", "-nodisp", "-autoexit", str(path)]
        self._process = subprocess.Popen(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

    def stop(self) -> None:
        if self._is_windows and self._winsound is not None:
            self._winsound.PlaySound(None, self._winsound.SND_PURGE)

        if self._process and self._process.poll() is None:
            self._process.terminate()
            try:
                self._process.wait(timeout=1)
            except subprocess.TimeoutExpired:
                self._process.kill()
        self._process = None
        self._current_path = None
        self._started_at = None
        self._duration_ms = None

    @staticmethod
    def _read_duration_ms(path: Path) -> int | None:
        try:
            with wave.open(str(path), 'rb') as wav_file:
                sample_rate = wav_file.getframerate()
                frame_count = wav_file.getnframes()
            if sample_rate <= 0:
                return None
            return int(frame_count * 1000 / sample_rate)
        except Exception:
            return None
