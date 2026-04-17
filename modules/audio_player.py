from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import threading
import wave

import numpy as np
import sounddevice as sd


@dataclass(slots=True)
class ScheduledChunk:
    frame_start: int
    frame_end: int
    dac_start_time: float
    dac_end_time: float


class AudioPlayer:
    def __init__(self) -> None:
        self._current_path: Path | None = None
        self._duration_ms: int | None = None
        self._stream = None
        self._audio_data: np.ndarray | None = None
        self._sample_rate: int = 0
        self._channels: int = 0
        self._write_frame_position: int = 0
        self._audible_frame_position: int = 0
        self._scheduled_chunks: list[ScheduledChunk] = []
        self._playing = False
        self._lock = threading.Lock()

    @property
    def current_path(self) -> Path | None:
        return self._current_path

    def current_position_ms(self) -> int:
        with self._lock:
            if self._sample_rate <= 0:
                return 0
            sample_rate = self._sample_rate
            duration_ms = self._duration_ms
            audible_frame = self._audible_frame_position
            chunks = list(self._scheduled_chunks)
            stream = self._stream
        stream_time = None
        if stream is not None:
            try:
                stream_time = float(stream.time)
            except Exception:
                stream_time = None
        if stream_time is not None and chunks:
            audible_frame = self._frame_at_stream_time(stream_time, chunks, audible_frame)
            with self._lock:
                self._audible_frame_position = audible_frame
                self._scheduled_chunks = [
                    chunk for chunk in self._scheduled_chunks if chunk.dac_end_time >= stream_time - 0.25
                ]
        position_ms = int(audible_frame * 1000 / sample_rate)
        if duration_ms is not None:
            return max(0, min(position_ms, duration_ms))
        return max(0, position_ms)

    def is_playing(self) -> bool:
        with self._lock:
            if not self._playing:
                return False
            if self._stream is None:
                return False
            return bool(self._stream.active)

    def play(self, path: Path) -> None:
        self.stop()
        self._current_path = path
        self._duration_ms = self._read_duration_ms(path)

        audio_data, sample_rate, channels = self._load_wav(path)
        with self._lock:
            self._audio_data = audio_data
            self._sample_rate = sample_rate
            self._channels = channels
            self._write_frame_position = 0
            self._audible_frame_position = 0
            self._scheduled_chunks = []
            self._playing = True

        def callback(outdata, frames, time_info, status) -> None:
            if status:
                pass
            with self._lock:
                if self._audio_data is None:
                    outdata.fill(0)
                    raise sd.CallbackStop()
                start = self._write_frame_position
                end = min(start + frames, len(self._audio_data))
                chunk = self._audio_data[start:end]
                written = len(chunk)
                if written > 0:
                    outdata[:written] = chunk
                    dac_start = float(getattr(time_info, "outputBufferDacTime", 0.0) or getattr(time_info, "currentTime", 0.0))
                    dac_end = dac_start + (written / self._sample_rate)
                    self._scheduled_chunks.append(
                        ScheduledChunk(
                            frame_start=start,
                            frame_end=end,
                            dac_start_time=dac_start,
                            dac_end_time=dac_end,
                        )
                    )
                if written < frames:
                    outdata[written:] = 0
                    self._write_frame_position = len(self._audio_data)
                    raise sd.CallbackStop()
                self._write_frame_position = end

        self._stream = sd.OutputStream(
            samplerate=sample_rate,
            channels=channels,
            dtype=audio_data.dtype,
            callback=callback,
            finished_callback=self._on_finished,
        )
        self._stream.start()

    def stop(self) -> None:
        with self._lock:
            stream = self._stream
            self._stream = None
            self._playing = False
            self._audio_data = None
            self._sample_rate = 0
            self._channels = 0
            self._write_frame_position = 0
            self._audible_frame_position = 0
            self._scheduled_chunks = []
            self._current_path = None
            self._duration_ms = None
        if stream is not None:
            try:
                stream.stop()
            except Exception:
                pass
            try:
                stream.close()
            except Exception:
                pass

    def _on_finished(self) -> None:
        with self._lock:
            if self._audio_data is not None:
                self._audible_frame_position = len(self._audio_data)
            self._playing = False

    @staticmethod
    def _frame_at_stream_time(
        stream_time: float,
        chunks: list[ScheduledChunk],
        fallback_frame: int,
    ) -> int:
        if not chunks:
            return max(0, fallback_frame)
        if stream_time <= chunks[0].dac_start_time:
            return max(0, chunks[0].frame_start)
        previous_end = chunks[0].frame_start
        for chunk in chunks:
            if stream_time < chunk.dac_start_time:
                return max(0, previous_end)
            if chunk.dac_start_time <= stream_time <= chunk.dac_end_time:
                span = max(chunk.dac_end_time - chunk.dac_start_time, 1e-9)
                progress = (stream_time - chunk.dac_start_time) / span
                frame_span = max(chunk.frame_end - chunk.frame_start, 0)
                return max(0, int(chunk.frame_start + frame_span * progress))
            previous_end = chunk.frame_end
        return max(0, chunks[-1].frame_end)

    @staticmethod
    def _load_wav(path: Path) -> tuple[np.ndarray, int, int]:
        with wave.open(str(path), 'rb') as wav_file:
            sample_rate = wav_file.getframerate()
            channels = wav_file.getnchannels()
            sample_width = wav_file.getsampwidth()
            frame_count = wav_file.getnframes()
            raw = wav_file.readframes(frame_count)

        if sample_width == 1:
            data = np.frombuffer(raw, dtype=np.uint8)
        elif sample_width == 2:
            data = np.frombuffer(raw, dtype=np.int16)
        elif sample_width == 4:
            data = np.frombuffer(raw, dtype=np.int32)
        else:
            raise ValueError('この WAV のビット深度にはまだ対応していません。')

        if channels > 1:
            data = data.reshape(-1, channels)
        else:
            data = data.reshape(-1, 1)
        return data, sample_rate, channels

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
