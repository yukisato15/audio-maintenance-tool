from __future__ import annotations

from array import array
from dataclasses import dataclass
from pathlib import Path
import math
import tempfile
import wave


TRIM_BACKUP_PREFIX = ".trim_backup_"


@dataclass(slots=True)
class AudioMetadata:
    sample_rate: int
    channels: int
    sample_width: int
    frame_count: int

    @property
    def duration_ms(self) -> int:
        if self.sample_rate <= 0:
            return 0
        return int(self.frame_count * 1000 / self.sample_rate)


@dataclass(slots=True)
class AudioLevelStats:
    peak_db: float
    rms_db: float
    clipping_ratio: float
    silent_ratio: float


def _backup_path(path: Path) -> Path:
    return path.with_name(f"{TRIM_BACKUP_PREFIX}{path.name}")


def has_trim_backup(path: Path) -> bool:
    return _backup_path(path).exists()


def get_audio_metadata(path: Path) -> AudioMetadata:
    with wave.open(str(path), "rb") as wav_file:
        return AudioMetadata(
            sample_rate=wav_file.getframerate(),
            channels=wav_file.getnchannels(),
            sample_width=wav_file.getsampwidth(),
            frame_count=wav_file.getnframes(),
        )


def _resolve_source(path: Path) -> Path:
    backup = _backup_path(path)
    return backup if backup.exists() else path


def _read_normalized_samples(path: Path) -> tuple[list[float], int, int]:
    source = _resolve_source(path)
    with wave.open(str(source), "rb") as wav_file:
        channels = wav_file.getnchannels()
        sample_width = wav_file.getsampwidth()
        frame_count = wav_file.getnframes()
        raw_frames = wav_file.readframes(frame_count)

    if frame_count <= 0 or not raw_frames:
        return [], channels, frame_count

    if sample_width == 1:
        samples = array("B", raw_frames)
        values = [(sample - 128) / 128 for sample in samples]
    elif sample_width == 2:
        samples = array("h")
        samples.frombytes(raw_frames)
        values = [sample / 32768 for sample in samples]
    elif sample_width == 4:
        samples = array("i")
        samples.frombytes(raw_frames)
        values = [sample / 2147483648 for sample in samples]
    else:
        return [], channels, frame_count

    return values, channels, frame_count


def get_waveform_peaks(path: Path, bucket_count: int = 180) -> list[float]:
    if bucket_count <= 0:
        return []

    values, channels, frame_count = _read_normalized_samples(path)
    if frame_count <= 0 or not values:
        return [0.0] * bucket_count

    peaks = [0.0] * bucket_count
    for frame_index in range(frame_count):
        bucket_index = min(int(frame_index * bucket_count / frame_count), bucket_count - 1)
        frame_peak = 0.0
        base = frame_index * channels
        for channel in range(channels):
            sample_index = base + channel
            if sample_index >= len(values):
                break
            frame_peak = max(frame_peak, abs(values[sample_index]))
        peaks[bucket_index] = max(peaks[bucket_index], min(frame_peak, 1.0))

    return peaks


def analyze_audio_levels(path: Path) -> AudioLevelStats:
    values, channels, frame_count = _read_normalized_samples(path)
    if frame_count <= 0 or not values:
        return AudioLevelStats(peak_db=-120.0, rms_db=-120.0, clipping_ratio=0.0, silent_ratio=1.0)

    frame_peaks: list[float] = []
    squared_sum = 0.0
    clip_count = 0
    silent_count = 0

    for frame_index in range(frame_count):
        base = frame_index * channels
        frame_peak = 0.0
        for channel in range(channels):
            sample_index = base + channel
            if sample_index >= len(values):
                break
            amplitude = abs(values[sample_index])
            squared_sum += values[sample_index] * values[sample_index]
            frame_peak = max(frame_peak, amplitude)
        frame_peaks.append(frame_peak)
        if frame_peak >= 0.995:
            clip_count += 1
        if frame_peak <= 0.01:
            silent_count += 1

    peak = max(frame_peaks, default=0.0)
    rms = (squared_sum / max(len(values), 1)) ** 0.5
    return AudioLevelStats(
        peak_db=20 * math.log10(max(peak, 1e-6)),
        rms_db=20 * math.log10(max(rms, 1e-6)),
        clipping_ratio=clip_count / max(frame_count, 1),
        silent_ratio=silent_count / max(frame_count, 1),
    )


def _trim_frame_range(frame_count: int, sample_rate: int, trim_start_ms: int, trim_end_ms: int) -> tuple[int, int]:
    start_frame = max(0, int(sample_rate * trim_start_ms / 1000))
    end_frame = max(0, int(sample_rate * trim_end_ms / 1000))
    start_frame = min(start_frame, frame_count)
    end_frame = min(end_frame, max(frame_count - start_frame - 1, 0))
    remaining = max(frame_count - start_frame - end_frame, 1)
    return start_frame, remaining


def _write_trimmed_audio(source: Path, destination: Path, trim_start_ms: int, trim_end_ms: int) -> None:
    with wave.open(str(source), "rb") as src:
        params = src.getparams()
        start_frame, remaining_frames = _trim_frame_range(src.getnframes(), src.getframerate(), trim_start_ms, trim_end_ms)
        src.setpos(start_frame)
        frames = src.readframes(remaining_frames)

    with wave.open(str(destination), "wb") as dst:
        dst.setparams(params)
        dst.writeframes(frames)


def create_trim_preview(path: Path, trim_start_ms: int, trim_end_ms: int) -> Path:
    source = _resolve_source(path)
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".wav")
    temp_file.close()
    destination = Path(temp_file.name)
    _write_trimmed_audio(source, destination, trim_start_ms, trim_end_ms)
    return destination


def attenuate_audio_in_place(path: Path, gain_db: float = -3.0) -> None:
    factor = 10 ** (gain_db / 20)
    source = path
    temp_destination = path.with_name(f".trim_tmp_{path.name}")

    with wave.open(str(source), "rb") as src:
        params = src.getparams()
        channels = src.getnchannels()
        sample_width = src.getsampwidth()
        frame_count = src.getnframes()
        raw_frames = src.readframes(frame_count)

    if sample_width == 1:
        samples = array("B", raw_frames)
        adjusted = bytearray()
        for sample in samples:
            centered = sample - 128
            value = int(centered * factor)
            value = max(-128, min(127, value))
            adjusted.append(value + 128)
        output_bytes = bytes(adjusted)
    elif sample_width == 2:
        samples = array("h")
        samples.frombytes(raw_frames)
        adjusted = array("h", (max(-32768, min(32767, int(sample * factor))) for sample in samples))
        output_bytes = adjusted.tobytes()
    elif sample_width == 4:
        samples = array("i")
        samples.frombytes(raw_frames)
        adjusted = array("i", (max(-2147483648, min(2147483647, int(sample * factor))) for sample in samples))
        output_bytes = adjusted.tobytes()
    else:
        raise ValueError("この WAV のビット深度にはまだ対応していません。")

    with wave.open(str(temp_destination), "wb") as dst:
        dst.setparams(params)
        dst.writeframes(output_bytes)

    if path.exists():
        path.unlink()
    temp_destination.rename(path)


def apply_trim_in_place(path: Path, trim_start_ms: int, trim_end_ms: int) -> Path:
    backup = _backup_path(path)
    if not backup.exists():
        path.rename(backup)
    source = backup
    temp_destination = path.with_name(f".trim_tmp_{path.name}")
    _write_trimmed_audio(source, temp_destination, trim_start_ms, trim_end_ms)
    if path.exists():
        path.unlink()
    temp_destination.rename(path)
    return backup


def restore_trim_backup(path: Path) -> None:
    backup = _backup_path(path)
    if not backup.exists():
        raise FileNotFoundError("元に戻せる余白修正バックアップがありません。")
    if path.exists():
        path.unlink()
    backup.rename(path)
