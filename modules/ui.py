from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
from tkinterdnd2 import DND_FILES, TkinterDnD

from modules.audio_editor import analyze_audio_levels, apply_trim_in_place, attenuate_audio_in_place, create_trim_preview, get_audio_metadata, get_waveform_peaks, has_trim_backup, restore_trim_backup
from modules.audio_player import AudioPlayer
from modules.csv_logger import write_rename_log
from modules.file_parser import ParseResult, ParsedAudioFile, parse_audio_folder
from modules.renamer import (
    RenamePlanEntry,
    RenameSettings,
    build_rename_plan,
    execute_rename_plan,
    has_undo_manifest,
    undo_last_rename,
    write_undo_manifest,
)
from modules.settings_store import clear_workflow_state, load_settings, load_workflow_state, save_settings, save_workflow_state


ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")


TABLE_COLUMN_WIDTHS = {
    0: 52,
    1: 72,
    2: 72,
    3: 72,
    4: 132,
    5: 500,
    6: 90,
    7: 220,
}


@dataclass(slots=True)
class FolderSession:
    folder: Path
    parse_result: ParseResult
    selected_filenames: set[str] | None = None
    manual_order: list[str] = field(default_factory=list)
    ok_flags: dict[str, bool] = field(default_factory=dict)
    missing_indices: set[int] = field(default_factory=set)
    undo_selected_filenames: set[str] | None = None
    undo_manual_order: list[str] = field(default_factory=list)
    reviewed_flags: dict[str, bool] = field(default_factory=dict)


class FileRow(ctk.CTkFrame):
    def __init__(
        self,
        master: tk.Misc,
        file_item: ParsedAudioFile,
        initial_ok: bool,
        trim_modified: bool,
        reviewed: bool,
        on_status_change,
        on_play_toggle,
        on_trim,
        on_restore_trim,
        on_drag_start,
        on_drag_end,
    ) -> None:
        is_ng = reviewed and not initial_ok
        if is_ng:
            row_color = ("#fde8e8", "#442727")
            badge_color = "#dc2626"
        elif not reviewed:
            row_color = ("#fff7d6", "#3f3721")
            badge_color = "#ca8a04"
        elif trim_modified:
            row_color = ("#e8f1ff", "#23344c")
            badge_color = "#2563eb"
        elif file_item.duplicate_index:
            row_color = ("#ffedd5", "#45311f")
            badge_color = "#d97706"
        else:
            row_color = "transparent"
            badge_color = None

        super().__init__(master, fg_color=row_color)
        self.file_item = file_item
        self._on_status_change = on_status_change
        self._on_play_toggle = on_play_toggle
        self._on_trim = on_trim
        self._on_restore_trim = on_restore_trim
        self._on_drag_start = on_drag_start
        self._on_drag_end = on_drag_end
        self.ok_var = tk.BooleanVar(value=initial_ok if reviewed and initial_ok else False)
        self.ng_var = tk.BooleanVar(value=is_ng)

        self._configure_columns()

        drag_handle = ctk.CTkLabel(self, text="↕", width=20, anchor="center")
        drag_handle.grid(row=0, column=0, padx=(10, 6), pady=5, sticky="w")
        drag_handle.bind("<ButtonPress-1>", lambda _event: self._on_drag_start(self.file_item.original_filename))
        drag_handle.bind("<ButtonRelease-1>", lambda event: self._on_drag_end(self.file_item.original_filename, event.y_root))

        ctk.CTkCheckBox(self, text="", width=28, variable=self.ok_var, command=self._toggle_ok).grid(
            row=0, column=1, padx=(10, 6), pady=5, sticky="w"
        )
        ctk.CTkCheckBox(self, text="", width=28, variable=self.ng_var, command=self._toggle_ng).grid(
            row=0, column=2, padx=(10, 6), pady=5, sticky="w"
        )

        self.play_button = ctk.CTkButton(self, text="▶", width=42, command=lambda: self._on_play_toggle(self.file_item.path))
        self.play_button.grid(row=0, column=3, padx=(8, 6), pady=5, sticky="w")

        trim_frame = ctk.CTkFrame(self, fg_color="transparent")
        trim_frame.grid(row=0, column=4, padx=(2, 8), pady=5, sticky="w")
        ctk.CTkButton(trim_frame, text="余白", width=58, command=lambda: self._on_trim(self.file_item.path)).grid(row=0, column=0, padx=(0, 4), sticky="w")
        self.restore_trim_button = ctk.CTkButton(
            trim_frame,
            text="戻す",
            width=58,
            fg_color=("#d5d5d5", "#4a4a4a"),
            hover_color=("#c8c8c8", "#5a5a5a"),
            command=lambda: self._on_restore_trim(self.file_item.path),
        )
        self.restore_trim_button.grid(row=0, column=1, sticky="w")
        if not trim_modified:
            self.restore_trim_button.configure(state="disabled")

        status_suffixes: list[str] = []
        if reviewed and is_ng:
            status_suffixes.append("NG")
        if file_item.duplicate_index:
            status_suffixes.append("重複")
        if trim_modified:
            status_suffixes.append("余白修正済み")
        if not reviewed:
            status_suffixes.append("未確認")
        suffix_text = "" if not status_suffixes else "  [" + "] [".join(status_suffixes) + "]"
        ctk.CTkLabel(
            self,
            text=f"{self.file_item.original_filename}{suffix_text}",
            anchor="w",
            text_color=badge_color,
        ).grid(row=0, column=5, padx=(10, 10), pady=5, sticky="ew")
        ctk.CTkLabel(self, text=str(self.file_item.original_index), anchor="center").grid(
            row=0, column=6, padx=8, pady=5, sticky="nsew"
        )
        preview_text = self.file_item.text_portion if self.file_item.text_portion else "(なし)"
        ctk.CTkLabel(self, text=preview_text, anchor="w").grid(
            row=0, column=7, padx=(10, 12), pady=5, sticky="ew"
        )

    def _configure_columns(self) -> None:
        for column, minsize in TABLE_COLUMN_WIDTHS.items():
            self.grid_columnconfigure(column, minsize=minsize, weight=1 if column == 5 else 0)

    def set_playing(self, is_playing: bool) -> None:
        if is_playing:
            self.play_button.configure(text="■", fg_color="#b45309", hover_color="#92400e")
        else:
            self.play_button.configure(text="▶", fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"], hover_color=ctk.ThemeManager.theme["CTkButton"]["hover_color"])

    def _toggle_ok(self) -> None:
        if self.ok_var.get():
            self.ng_var.set(False)
        elif not self.ng_var.get():
            self.ok_var.set(True)
        self._on_status_change(self.file_item.original_filename, self.ok_var.get())

    def _toggle_ng(self) -> None:
        if self.ng_var.get():
            self.ok_var.set(False)
        elif not self.ok_var.get():
            self.ng_var.set(True)
        self._on_status_change(self.file_item.original_filename, self.ok_var.get())


class BatchRenameApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)
        self.title("音声ファイル一括リネームツール")

        self.app_settings = load_settings()
        self.geometry(str(self.app_settings.get("geometry", "1380x820")))
        self.minsize(1220, 760)

        self.folder_sessions: dict[Path, FolderSession] = {}
        self.folder_order: list[Path] = []
        self.current_folder: Path | None = None
        self.file_rows: list[FileRow] = []
        self.missing_vars: dict[int, tk.BooleanVar] = {}
        self.dragging_filename: str | None = None
        self.audio_player = AudioPlayer()
        self.playing_path: Path | None = None
        self.preview_temp_path: Path | None = None
        self.trim_dialog: ctk.CTkToplevel | None = None
        self.trim_waveform_drag_handle: str | None = None

        self.digits_var = tk.StringVar(value=self.app_settings.get("digits", "3桁"))
        self.keep_text_var = tk.BooleanVar(value=bool(self.app_settings.get("keep_text", True)))
        self.move_ng_var = tk.BooleanVar(value=bool(self.app_settings.get("move_ng", True)))
        self.export_csv_var = tk.BooleanVar(value=bool(self.app_settings.get("export_csv", True)))
        self.show_mode_var = tk.StringVar(value="全件")
        self.status_var = tk.StringVar(value="フォルダまたはファイルを追加してください。")
        self.folder_info_var = tk.StringVar(value="対象: 0 件")
        self.current_path_var = tk.StringVar(value="表示中のパス: -")

        self._build_layout()
        self._bind_setting_persistence()
        self._restore_workflow_state()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=5)
        self.grid_columnconfigure(1, weight=3)
        self.grid_rowconfigure(1, weight=1)

        top_bar = ctk.CTkFrame(self)
        top_bar.grid(row=0, column=0, padx=(16, 8), pady=(16, 8), sticky="ew")
        top_bar.grid_columnconfigure(0, weight=1)

        control_row = ctk.CTkFrame(top_bar, fg_color="transparent")
        control_row.grid(row=0, column=0, padx=12, pady=(12, 8), sticky="ew")
        control_row.grid_columnconfigure(4, weight=1)

        ctk.CTkButton(control_row, text="フォルダを追加", width=150, command=self.select_folders).grid(row=0, column=0, padx=(0, 8), sticky="w")
        ctk.CTkButton(control_row, text="ファイルを追加", width=150, command=self.select_files).grid(row=0, column=1, padx=8, sticky="w")
        ctk.CTkButton(
            control_row,
            text="前回リネーム前に戻す",
            width=150,
            height=30,
            fg_color=("#d5d5d5", "#4a4a4a"),
            hover_color=("#c8c8c8", "#5a5a5a"),
            command=self.undo_current_folder,
        ).grid(row=0, column=2, padx=(20, 8), sticky="w")
        ctk.CTkLabel(control_row, textvariable=self.folder_info_var, anchor="e").grid(row=0, column=4, padx=(16, 0), sticky="ew")

        self.folder_list_frame = ctk.CTkFrame(top_bar)
        self.folder_list_frame.grid(row=1, column=0, padx=12, pady=(0, 6), sticky="ew")
        self.folder_list_frame.grid_columnconfigure(0, weight=1)
        self.folder_list_frame.grid_rowconfigure(0, weight=1)

        self.folder_listbox = tk.Listbox(
            self.folder_list_frame,
            height=2,
            selectmode=tk.EXTENDED,
            exportselection=False,
            activestyle="none",
            font=("Arial", 14),
            relief="flat",
            borderwidth=0,
        )
        self.folder_listbox.grid(row=0, column=0, padx=(8, 0), pady=6, sticky="nsew")
        self.folder_listbox.bind("<<ListboxSelect>>", self.on_folder_list_select)
        folder_scrollbar = ctk.CTkScrollbar(self.folder_list_frame, command=self.folder_listbox.yview)
        folder_scrollbar.grid(row=0, column=1, padx=(0, 8), pady=6, sticky="ns")
        self.folder_listbox.configure(yscrollcommand=folder_scrollbar.set)
        self._setup_drop_targets()

        ctk.CTkLabel(
            top_bar,
            textvariable=self.current_path_var,
            anchor="w",
            text_color=("gray35", "gray70"),
        ).grid(row=2, column=0, padx=16, pady=(0, 10), sticky="ew")

        self.table_frame = ctk.CTkFrame(self)
        self.table_frame.grid(row=1, column=0, padx=(16, 8), pady=8, sticky="nsew")
        self.table_frame.grid_columnconfigure(0, weight=1)
        self.table_frame.grid_rowconfigure(2, weight=1)

        toolbar = ctk.CTkFrame(self.table_frame)
        toolbar.grid(row=0, column=0, padx=12, pady=(12, 6), sticky="ew")
        toolbar.grid_columnconfigure(7, weight=1)

        ctk.CTkLabel(toolbar, text="表示設定", anchor="w").grid(row=0, column=0, padx=(0, 8), sticky="w")
        self.mode_segment = ctk.CTkSegmentedButton(
            toolbar,
            values=["全件", "重複のみ", "NGのみ"],
            variable=self.show_mode_var,
            command=lambda _value: self._render_current_folder(),
        )
        self.mode_segment.grid(row=0, column=1, padx=(0, 8), sticky="w")
        ctk.CTkButton(toolbar, text="欠番候補を反映", width=140, command=self.apply_missing_suggestions).grid(row=0, column=2, padx=6, sticky="w")
        ctk.CTkButton(toolbar, text="一覧をクリア", width=120, fg_color=("#d5d5d5", "#4a4a4a"), hover_color=("#c8c8c8", "#5a5a5a"), command=self.clear_folders).grid(row=0, column=3, padx=6, sticky="w")

        header = ctk.CTkFrame(self.table_frame)
        header.grid(row=1, column=0, padx=12, pady=(0, 4), sticky="ew")
        self._configure_table_columns(header)
        headers = {
            0: ("順", "center", (0, 0)),
            1: ("OK", "center", (0, 0)),
            2: ("NG", "center", (0, 0)),
            3: ("再生", "center", (0, 0)),
            4: ("余白", "center", (0, 0)),
            5: ("ファイル名", "w", (10, 10)),
            6: ("元番号", "center", (0, 0)),
            7: ("テキスト部分", "w", (10, 12)),
        }
        for column, (text, anchor, padx) in headers.items():
            sticky = "nsew" if anchor == "center" else "w"
            ctk.CTkLabel(header, text=text, anchor=anchor).grid(row=0, column=column, padx=padx, pady=10, sticky=sticky)

        self.file_scroll = ctk.CTkScrollableFrame(self.table_frame)
        self.file_scroll.grid(row=2, column=0, padx=12, pady=(4, 6), sticky="nsew")
        self.file_scroll.grid_columnconfigure(0, weight=1)

        batch_row = ctk.CTkFrame(self.table_frame, fg_color="transparent")
        batch_row.grid(row=3, column=0, padx=12, pady=(0, 12), sticky="ew")
        for column, minsize in TABLE_COLUMN_WIDTHS.items():
            batch_row.grid_columnconfigure(column, minsize=minsize, weight=1 if column == 5 else 0)
        ctk.CTkButton(batch_row, text="全件OK", width=64, height=28, command=self.mark_all_ok).grid(row=0, column=1, padx=(8, 6), sticky="w")
        ctk.CTkButton(batch_row, text="全件NG", width=64, height=28, command=self.mark_all_ng).grid(row=0, column=2, padx=(8, 6), sticky="w")

        self.side_panel = ctk.CTkFrame(self)
        self.side_panel.grid(row=0, column=1, rowspan=2, padx=(8, 16), pady=(16, 8), sticky="nsew")
        self.side_panel.grid_columnconfigure(0, weight=1)
        self.side_panel.grid_rowconfigure(2, weight=1)

        self.warning_label = ctk.CTkLabel(
            self.side_panel,
            text="重複番号、除外ファイル、Undo 状態をここに表示します。",
            justify="left",
            anchor="w",
            wraplength=420,
        )
        self.warning_label.grid(row=0, column=0, padx=12, pady=(12, 8), sticky="ew")

        action_panel = ctk.CTkFrame(self.side_panel)
        action_panel.grid(row=1, column=0, padx=12, pady=(0, 8), sticky="ew")
        action_panel.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            action_panel,
            text="補助操作",
            anchor="w",
            text_color=("gray35", "gray70"),
        ).grid(row=0, column=0, padx=10, pady=(8, 2), sticky="ew")

        tabview = ctk.CTkTabview(self.side_panel)
        tabview.grid(row=2, column=0, padx=12, pady=8, sticky="nsew")
        tabview.add("欠番")
        tabview.add("設定")

        missing_tab = tabview.tab("欠番")
        missing_tab.grid_columnconfigure(0, weight=1)
        missing_tab.grid_rowconfigure(2, weight=1)
        missing_tab.grid_rowconfigure(3, weight=0)
        ctk.CTkLabel(missing_tab, text="意図的に欠番として残す番号を選択", anchor="w").grid(
            row=0, column=0, padx=12, pady=(12, 8), sticky="ew"
        )
        missing_buttons = ctk.CTkFrame(missing_tab, fg_color="transparent")
        missing_buttons.grid(row=1, column=0, padx=12, pady=(0, 6), sticky="ew")
        ctk.CTkButton(missing_buttons, text="候補をチェック", width=130, command=self.apply_missing_suggestions).grid(
            row=0, column=0, padx=(0, 8), sticky="w"
        )

        self.missing_scroll = ctk.CTkScrollableFrame(missing_tab, height=320)
        self.missing_scroll.grid(row=2, column=0, padx=12, pady=(0, 4), sticky="nsew")
        self.missing_scroll.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(
            missing_tab,
            text="全部外す",
            width=90,
            height=26,
            fg_color="transparent",
            text_color=("gray35", "gray70"),
            hover_color=("#ececec", "#3a3a3a"),
            command=self.clear_missing_checks,
        ).grid(row=3, column=0, padx=12, pady=(0, 10), sticky="e")

        settings_tab = tabview.tab("設定")
        settings_tab.grid_columnconfigure(0, weight=1)
        settings_tab.grid_rowconfigure(5, weight=1)
        ctk.CTkLabel(settings_tab, text="桁数", anchor="w").grid(row=0, column=0, padx=12, pady=(12, 4), sticky="w")
        ctk.CTkOptionMenu(settings_tab, values=["2桁", "3桁", "4桁"], variable=self.digits_var).grid(
            row=1, column=0, padx=12, pady=(0, 12), sticky="ew"
        )
        ctk.CTkCheckBox(settings_tab, text="番号の後ろの元テキストを残す", variable=self.keep_text_var).grid(row=2, column=0, padx=12, pady=6, sticky="w")
        ctk.CTkCheckBox(settings_tab, text="NG ファイルを _NG フォルダへ移動", variable=self.move_ng_var).grid(row=3, column=0, padx=12, pady=6, sticky="w")
        ctk.CTkCheckBox(settings_tab, text="CSV ログを出力", variable=self.export_csv_var).grid(row=4, column=0, padx=12, pady=6, sticky="w")
        ctk.CTkFrame(settings_tab, fg_color="transparent").grid(row=5, column=0, sticky="nsew")

        bottom = ctk.CTkFrame(self)
        bottom.grid(row=2, column=0, columnspan=2, padx=16, pady=(8, 16), sticky="ew")
        bottom.grid_columnconfigure(0, weight=1)

        self.progress = ctk.CTkProgressBar(bottom)
        self.progress.grid(row=0, column=0, padx=12, pady=(12, 6), sticky="ew")
        self.progress.set(0)
        self.status_label = ctk.CTkLabel(bottom, textvariable=self.status_var, anchor="w")
        self.status_label.grid(row=1, column=0, padx=12, pady=(0, 12), sticky="ew")
        ctk.CTkButton(bottom, text="プレビュー", width=110, command=self.open_preview_dialog).grid(
            row=0, column=1, rowspan=2, padx=(12, 6), pady=12, sticky="e"
        )
        ctk.CTkButton(bottom, text="リネーム実行", command=self.run_rename, height=40).grid(
            row=0, column=2, rowspan=2, padx=(6, 12), pady=12, sticky="e"
        )

    def _bind_setting_persistence(self) -> None:
        for variable in (self.digits_var, self.keep_text_var, self.move_ng_var, self.export_csv_var):
            variable.trace_add("write", lambda *_args: self._persist_settings())

    def _persist_settings(self) -> None:
        save_settings(
            {
                "digits": self.digits_var.get(),
                "keep_text": self.keep_text_var.get(),
                "move_ng": self.move_ng_var.get(),
                "export_csv": self.export_csv_var.get(),
                "geometry": self.geometry(),
            }
        )

    def _serialize_session(self, session: FolderSession) -> dict:
        return {
            "folder": str(session.folder),
            "selected_filenames": None if session.selected_filenames is None else sorted(session.selected_filenames),
            "manual_order": list(session.manual_order),
            "ok_flags": dict(session.ok_flags),
            "reviewed_flags": dict(session.reviewed_flags),
            "missing_indices": sorted(session.missing_indices),
            "undo_selected_filenames": None if session.undo_selected_filenames is None else sorted(session.undo_selected_filenames),
            "undo_manual_order": list(session.undo_manual_order),
        }

    def _persist_workflow_state(self) -> None:
        self._save_current_state()
        if not self.folder_order:
            clear_workflow_state()
            return
        save_workflow_state(
            {
                "current_folder": None if self.current_folder is None else str(self.current_folder),
                "folders": [self._serialize_session(self.folder_sessions[folder]) for folder in self.folder_order],
            }
        )

    def _restore_workflow_state(self) -> None:
        state = load_workflow_state()
        folders_state = state.get("folders")
        if not isinstance(folders_state, list):
            return

        restored_paths: list[Path] = []
        for item in folders_state:
            if not isinstance(item, dict):
                continue
            folder = Path(str(item.get("folder", "")))
            if not folder.exists() or not folder.is_dir():
                continue
            selected = item.get("selected_filenames")
            selected_names = set(selected) if isinstance(selected, list) else None
            parse_result = parse_audio_folder(folder, selected_names)
            session = FolderSession(
                folder=folder,
                parse_result=parse_result,
                selected_filenames=selected_names,
                manual_order=list(item.get("manual_order", [])),
                ok_flags=dict(item.get("ok_flags", {})),
                missing_indices={int(index) for index in item.get("missing_indices", [])},
                undo_selected_filenames=None if item.get("undo_selected_filenames") is None else set(item.get("undo_selected_filenames", [])),
                undo_manual_order=list(item.get("undo_manual_order", [])),
                reviewed_flags=dict(item.get("reviewed_flags", {})),
            )
            self._refresh_session(session)
            self.folder_sessions[folder] = session
            self.folder_order.append(folder)
            restored_paths.append(folder)

        if not restored_paths:
            clear_workflow_state()
            return

        desired_current = Path(str(state.get("current_folder", ""))) if state.get("current_folder") else restored_paths[0]
        self.current_folder = desired_current if desired_current in self.folder_sessions else restored_paths[0]
        self._refresh_folder_list()
        self._render_current_folder()
        self.status_var.set("前回の作業状態を復元しました。")

    def on_close(self) -> None:
        self._persist_settings()
        self._persist_workflow_state()
        self._cleanup_preview_temp()
        self.audio_player.stop()
        self.destroy()

    def _configure_table_columns(self, frame: ctk.CTkFrame) -> None:
        for column, minsize in TABLE_COLUMN_WIDTHS.items():
            frame.grid_columnconfigure(column, minsize=minsize, weight=1 if column == 5 else 0)

    def _setup_drop_targets(self) -> None:
        for widget in (self.folder_list_frame, self.folder_listbox):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self.on_folder_drop)

    def _session_label(self, session: FolderSession) -> str:
        count = len(session.parse_result.files)
        suffix = f" ({count}件)" if session.selected_filenames is not None else ""
        marker = "  ↩" if has_undo_manifest(session.folder) else ""
        return f"{session.folder.name}{suffix}{marker}"

    def _default_ok_flags(self, parse_result: ParseResult) -> dict[str, bool]:
        return {file_item.original_filename: True for file_item in parse_result.files}

    def _default_reviewed_flags(self, parse_result: ParseResult) -> dict[str, bool]:
        return {file_item.original_filename: False for file_item in parse_result.files}

    def _refresh_session(self, session: FolderSession) -> None:
        parse_result = parse_audio_folder(session.folder, session.selected_filenames)
        session.parse_result = parse_result
        existing_ok = dict(session.ok_flags)
        existing_reviewed = dict(session.reviewed_flags)
        session.ok_flags = {file_item.original_filename: existing_ok.get(file_item.original_filename, True) for file_item in parse_result.files}
        session.reviewed_flags = {file_item.original_filename: existing_reviewed.get(file_item.original_filename, False) for file_item in parse_result.files}
        current_names = [file_item.original_filename for file_item in parse_result.files]
        ordered = [name for name in session.manual_order if name in current_names]
        for name in current_names:
            if name not in ordered:
                ordered.append(name)
        session.manual_order = ordered

    def _ordered_files(self, session: FolderSession) -> list[ParsedAudioFile]:
        mapping = {file_item.original_filename: file_item for file_item in session.parse_result.files}
        ordered = [mapping[name] for name in session.manual_order if name in mapping]
        for file_item in session.parse_result.files:
            if file_item.original_filename not in session.manual_order:
                ordered.append(file_item)
        return ordered

    def _ask_directories(self) -> list[Path]:
        try:
            raw_result = self.tk.call("tk_chooseDirectory", "-title", "音声ファイルフォルダを選択", "-mustexist", "1", "-multiple", "1")
            selected = list(self.tk.splitlist(raw_result))
        except tk.TclError:
            selected_one = filedialog.askdirectory(title="音声ファイルフォルダを選択", mustexist=True)
            selected = [selected_one] if selected_one else []
        return [Path(item) for item in selected if item and Path(item).is_dir()]

    def select_folders(self) -> None:
        folders = self._ask_directories()
        if folders:
            self.add_folder_sessions(folders)

    def select_files(self) -> None:
        filenames = filedialog.askopenfilenames(
            title="音声ファイルを選択",
            filetypes=[("WAV files", "*.wav"), ("All files", "*")],
        )
        if filenames:
            self.add_file_sessions([Path(name) for name in filenames])

    def add_folder_sessions(self, folders: list[Path]) -> None:
        self._save_current_state()
        added = 0
        excluded_messages: list[str] = []
        for folder in folders:
            if folder in self.folder_sessions:
                continue
            parse_result = parse_audio_folder(folder)
            session = FolderSession(
                folder=folder,
                parse_result=parse_result,
                selected_filenames=None,
                manual_order=[file_item.original_filename for file_item in parse_result.files],
                ok_flags=self._default_ok_flags(parse_result),
                reviewed_flags=self._default_reviewed_flags(parse_result),
            )
            self.folder_sessions[folder] = session
            self.folder_order.append(folder)
            added += 1
            if parse_result.excluded_files:
                excluded_messages.append(f"[{folder.name}]\n" + "\n".join(parse_result.excluded_files))
        self._after_sessions_added(added, excluded_messages)

    def add_file_sessions(self, files: list[Path]) -> None:
        self._save_current_state()
        grouped: dict[Path, set[str]] = {}
        for path in files:
            if path.is_file() and path.suffix.lower() == ".wav":
                grouped.setdefault(path.parent, set()).add(path.name)

        added = 0
        excluded_messages: list[str] = []
        for folder, selected_names in grouped.items():
            session = self.folder_sessions.get(folder)
            if session is None:
                parse_result = parse_audio_folder(folder, selected_names)
                session = FolderSession(
                    folder=folder,
                    parse_result=parse_result,
                    selected_filenames=set(selected_names),
                    manual_order=[file_item.original_filename for file_item in parse_result.files],
                    ok_flags=self._default_ok_flags(parse_result),
                    reviewed_flags=self._default_reviewed_flags(parse_result),
                )
                self.folder_sessions[folder] = session
                self.folder_order.append(folder)
                added += 1
            elif session.selected_filenames is not None:
                session.selected_filenames.update(selected_names)
                self._refresh_session(session)
            if session.parse_result.excluded_files:
                excluded_messages.append(f"[{folder.name}]\n" + "\n".join(session.parse_result.excluded_files))
        self._after_sessions_added(added, excluded_messages)

    def _after_sessions_added(self, added: int, excluded_messages: list[str]) -> None:
        self._refresh_folder_list()
        if self.folder_order and self.current_folder is None:
            self._set_current_folder(self.folder_order[0])
        else:
            self._render_current_folder()
        if excluded_messages:
            messagebox.showwarning(
                "解析できないファイル",
                "先頭番号を解析できないため、以下のファイルを除外しました。\n\n" + "\n\n".join(excluded_messages),
            )
        self.status_var.set(f"{added} 件の対象を追加しました。" if added else "追加できる新規対象はありませんでした。")
        self._persist_workflow_state()

    def _normalize_dropped_folders(self, raw_data: str) -> tuple[list[Path], list[Path]]:
        folders: list[Path] = []
        files: list[Path] = []
        seen_folders: set[Path] = set()
        seen_files: set[Path] = set()
        for item in self.tk.splitlist(raw_data):
            if not item:
                continue
            path = Path(item)
            if path.is_dir() and path not in seen_folders:
                seen_folders.add(path)
                folders.append(path)
            elif path.is_file() and path not in seen_files:
                seen_files.add(path)
                files.append(path)
        return folders, files

    def on_folder_drop(self, event: tk.Event) -> str:
        folders, files = self._normalize_dropped_folders(event.data)
        if folders:
            self.add_folder_sessions(folders)
        if files:
            self.add_file_sessions(files)
        if not folders and not files:
            self.status_var.set("フォルダまたは既存 wav ファイルをドロップしてください。")
        return "break"

    def remove_selected_folders(self) -> None:
        selected_indices = list(self.folder_listbox.curselection())
        if not selected_indices:
            messagebox.showwarning("未選択", "外す対象を一覧から選択してください。")
            return
        self._save_current_state()
        removing_paths = [self.folder_order[index] for index in selected_indices]
        for path in removing_paths:
            self.folder_sessions.pop(path, None)
            if path in self.folder_order:
                self.folder_order.remove(path)
        if self.current_folder in removing_paths:
            self.current_folder = self.folder_order[0] if self.folder_order else None
        self._refresh_folder_list()
        self._render_current_folder()
        self.status_var.set(f"{len(removing_paths)} 件の対象を一覧から外しました。")
        self._persist_workflow_state()

    def clear_folders(self) -> None:
        if not self.folder_order:
            return
        self.folder_sessions.clear()
        self.folder_order.clear()
        self.current_folder = None
        self.file_rows.clear()
        self.missing_vars.clear()
        self.folder_listbox.delete(0, tk.END)
        self.playing_path = None
        self._render_current_folder()
        self._update_folder_info()
        self.status_var.set("対象一覧をクリアしました。")
        clear_workflow_state()

    def on_folder_list_select(self, _event: tk.Event) -> None:
        selection = self.folder_listbox.curselection()
        if selection:
            selected_folder = self.folder_order[selection[0]]
            if selected_folder != self.current_folder:
                self._set_current_folder(selected_folder)

    def _refresh_folder_list(self) -> None:
        selected_paths = {self.folder_order[index] for index in self.folder_listbox.curselection() if index < len(self.folder_order)}
        self.folder_listbox.delete(0, tk.END)
        for folder in self.folder_order:
            self.folder_listbox.insert(tk.END, self._session_label(self.folder_sessions[folder]))
        for index, folder in enumerate(self.folder_order):
            if folder in selected_paths or folder == self.current_folder:
                self.folder_listbox.selection_set(index)
        self._update_folder_info()

    def _update_folder_info(self) -> None:
        count = len(self.folder_order)
        if self.current_folder:
            session = self.folder_sessions[self.current_folder]
            self.folder_info_var.set(f"対象: {count} 件 / 表示中: {self._session_label(session)}")
            self.current_path_var.set(f"表示中のパス: {self.current_folder}")
        else:
            self.folder_info_var.set(f"対象: {count} 件")
            self.current_path_var.set("表示中のパス: -")

    def _save_current_state(self) -> None:
        if not self.current_folder:
            return
        session = self.folder_sessions.get(self.current_folder)
        if session is None:
            return
        session.missing_indices = {index for index, var in self.missing_vars.items() if var.get()}

    def _set_current_folder(self, folder: Path) -> None:
        self._save_current_state()
        self.current_folder = folder
        self._refresh_folder_list()
        self._render_current_folder()

    def _filtered_files(self, session: FolderSession) -> list[ParsedAudioFile]:
        ordered = self._ordered_files(session)
        mode = self.show_mode_var.get()
        if mode == "重複のみ":
            return [file_item for file_item in ordered if file_item.duplicate_index]
        if mode == "NGのみ":
            return [file_item for file_item in ordered if not session.ok_flags.get(file_item.original_filename, True)]
        return ordered

    def _render_current_folder(self) -> None:
        session = self.folder_sessions.get(self.current_folder) if self.current_folder else None
        self._render_file_rows(session)
        self._render_missing_checkboxes(session)
        self._update_warnings(session)
        self._update_folder_info()
        self._sync_play_buttons()
        if session is None:
            self.status_var.set("フォルダまたはファイルを追加してください。")

    def _render_file_rows(self, session: FolderSession | None) -> None:
        for widget in self.file_scroll.winfo_children():
            widget.destroy()
        self.file_rows.clear()

        if session is None:
            ctk.CTkLabel(self.file_scroll, text="上の一覧からフォルダまたはファイルを追加してください。", anchor="w").grid(row=0, column=0, padx=12, pady=12, sticky="w")
            return

        visible_files = self._filtered_files(session)
        if not visible_files:
            ctk.CTkLabel(self.file_scroll, text="現在の表示条件に一致するファイルがありません。", anchor="w").grid(row=0, column=0, padx=12, pady=12, sticky="w")
            return

        for row_number, file_item in enumerate(visible_files, start=1):
            initial_ok = session.ok_flags.get(file_item.original_filename, True)
            row = FileRow(
                self.file_scroll,
                file_item,
                initial_ok,
                has_trim_backup(file_item.path),
                session.reviewed_flags.get(file_item.original_filename, False),
                self._on_file_status_change,
                self.toggle_play_audio,
                self.open_trim_dialog,
                self.restore_trim_file,
                self.start_row_drag,
                self.finish_row_drag,
            )
            row.grid(row=row_number - 1, column=0, padx=4, pady=1, sticky="ew")
            self.file_rows.append(row)

    def _sync_play_buttons(self) -> None:
        for row in self.file_rows:
            row.set_playing(self.playing_path == row.file_item.path)

    def _render_missing_checkboxes(self, session: FolderSession | None) -> None:
        for widget in self.missing_scroll.winfo_children():
            widget.destroy()
        self.missing_vars.clear()
        if session is None:
            ctk.CTkLabel(self.missing_scroll, text="対象を追加すると番号一覧が表示されます。", anchor="w").grid(row=0, column=0, padx=8, pady=8, sticky="w")
            return
        for row_number, index in enumerate(session.parse_result.detected_indices):
            var = tk.BooleanVar(value=index in session.missing_indices)
            self.missing_vars[index] = var
            text_color = ("#2563eb", "#93c5fd") if index in session.missing_indices else None
            ctk.CTkCheckBox(self.missing_scroll, text=f"{index:03d}", variable=var, text_color=text_color, command=lambda idx=index: self._on_missing_toggle(idx)).grid(
                row=row_number, column=0, padx=8, pady=4, sticky="w"
            )

    def _on_missing_toggle(self, index: int) -> None:
        session = self._current_session_or_warn()
        if session is None:
            return
        if self.missing_vars.get(index) and self.missing_vars[index].get():
            session.missing_indices.add(index)
        else:
            session.missing_indices.discard(index)
        self._persist_workflow_state()
        self._render_missing_checkboxes(session)
        self._update_warnings(session)

    def _update_warnings(self, session: FolderSession | None) -> None:
        if session is None:
            self.warning_label.configure(text="重複番号、除外ファイル、Undo 状態をここに表示します。")
            return
        warnings: list[str] = [f"表示中: {self._session_label(session)}"]
        if session.selected_filenames is not None:
            warnings.append("ファイル単体選択モード: 左端の ↕ で行順を並べ替えできます。")
        unreviewed = [name for name, reviewed in session.reviewed_flags.items() if not reviewed]
        trim_modified = [file_item.original_filename for file_item in session.parse_result.files if has_trim_backup(file_item.path)]
        if session.parse_result.duplicate_indices:
            warnings.append("重複番号: " + ", ".join(f"{index:03d}" for index in session.parse_result.duplicate_indices))
        if session.missing_indices:
            warnings.append("欠番指定: " + ", ".join(f"{index:03d}" for index in sorted(session.missing_indices)))
        if trim_modified:
            warnings.append(f"余白修正済み: {len(trim_modified)} 件")
        if unreviewed:
            warnings.append(f"未確認: {len(unreviewed)} 件")
        if session.parse_result.excluded_files:
            warnings.append(f"除外ファイル: {len(session.parse_result.excluded_files)} 件")
        suggestions = self._suggest_missing_indices(session)
        if suggestions:
            warnings.append("欠番候補: " + ", ".join(f"{index:03d}" for index in sorted(suggestions)))
        if has_undo_manifest(session.folder):
            warnings.append("前回のリネームを取り消せます。")
        self.warning_label.configure(text="\n\n".join(warnings))

    def _on_file_status_change(self, filename: str, is_ok: bool) -> None:
        if not self.current_folder:
            return
        session = self.folder_sessions[self.current_folder]
        session.ok_flags[filename] = is_ok
        session.reviewed_flags[filename] = True
        self._persist_workflow_state()
        self._render_current_folder()

    def _monitor_playback(self) -> None:
        if self.playing_path is None:
            return
        if self.audio_player.is_playing():
            self.after(300, self._monitor_playback)
            return
        self.playing_path = None
        self._sync_play_buttons()
        self.status_var.set("再生が終了しました。")

    def toggle_play_audio(self, path: Path) -> None:
        if self.playing_path == path:
            self.stop_audio()
            return
        try:
            self.audio_player.play(path)
            self.playing_path = path
            self._sync_play_buttons()
            self.status_var.set(f"再生中: {path.name}")
            self.after(300, self._monitor_playback)
        except FileNotFoundError:
            messagebox.showerror("再生エラー", f"音声ファイルが見つかりません。\n\n{path.name}")
        except Exception as exc:
            messagebox.showerror("再生エラー", str(exc))

    def stop_audio(self) -> None:
        self.audio_player.stop()
        self.playing_path = None
        self._sync_play_buttons()
        self.status_var.set("再生を停止しました。")

    def _cleanup_preview_temp(self) -> None:
        if self.preview_temp_path and self.preview_temp_path.exists():
            try:
                self.preview_temp_path.unlink()
            except OSError:
                pass
        self.preview_temp_path = None

    @staticmethod
    def _format_duration_ms(duration_ms: int) -> str:
        seconds, milliseconds = divmod(max(duration_ms, 0), 1000)
        minutes, seconds = divmod(seconds, 60)
        return f"{minutes:02d}:{seconds:02d}.{milliseconds:03d}"

    def _format_trim_value(self, duration_ms: int) -> str:
        return f"{duration_ms} ms ({duration_ms / 1000:.2f} 秒)"

    def _draw_trim_waveform(
        self,
        canvas: tk.Canvas,
        peaks: list[float],
        duration_ms: int,
        trim_start_ms: int,
        trim_end_ms: int,
        playhead_x: int | None = None,
    ) -> None:
        canvas.delete("all")
        width = int(canvas.winfo_width() or canvas.cget("width") or 460)
        height = int(canvas.winfo_height() or canvas.cget("height") or 120)
        if width <= 2 or height <= 2:
            return

        canvas.create_rectangle(0, 0, width, height, fill="#1f1f1f", outline="#3b3b3b")
        if not peaks:
            return

        safe_duration = max(duration_ms, 1)
        start_x = max(0, min(width, int(width * trim_start_ms / safe_duration)))
        end_x = max(0, min(width, width - int(width * trim_end_ms / safe_duration)))

        if start_x > 0:
            canvas.create_rectangle(0, 0, start_x, height, fill="#3b3b3b", outline="")
        if end_x < width:
            canvas.create_rectangle(end_x, 0, width, height, fill="#3b3b3b", outline="")

        mid_y = height / 2
        for index, peak in enumerate(peaks):
            x = int(index * width / max(len(peaks) - 1, 1))
            bar_height = max(2, int((height - 16) * peak))
            color = "#2f7fd1" if start_x <= x <= end_x else "#737373"
            canvas.create_line(x, mid_y - bar_height / 2, x, mid_y + bar_height / 2, fill=color, width=1)

        canvas.create_line(start_x, 0, start_x, height, fill="#f5f5f5", dash=(3, 3), width=2)
        canvas.create_line(end_x, 0, end_x, height, fill="#f5f5f5", dash=(3, 3), width=2)
        if playhead_x is not None:
            canvas.create_line(playhead_x, 0, playhead_x, height, fill="#ef4444", width=2)

    def _refresh_current_session_after_audio_edit(self, path: Path) -> None:
        session = self.folder_sessions.get(path.parent)
        if session is None:
            return
        self._refresh_session(session)
        if self.current_folder == session.folder:
            self._render_current_folder()
        else:
            self._refresh_folder_list()
        self._persist_workflow_state()

    def restore_trim_file(self, path: Path) -> None:
        try:
            self.stop_audio()
            self._cleanup_preview_temp()
            restore_trim_backup(path)
            self._refresh_current_session_after_audio_edit(path)
            session = self.folder_sessions.get(path.parent)
            if session is not None:
                session.reviewed_flags[path.name] = True
            self._persist_workflow_state()
            self.status_var.set(f"余白修正前に戻しました: {path.name}")
        except FileNotFoundError:
            messagebox.showwarning("未修正", "このファイルには戻せる余白修正がありません。")
        except Exception as exc:
            messagebox.showerror("余白修正エラー", str(exc))

    @staticmethod
    def _level_improvement_text(level_stats) -> str | None:
        if level_stats.clipping_ratio >= 0.02 or level_stats.peak_db >= -0.1:
            return "音量を少し下げる (-3 dB)"
        return None

    def _format_level_summary(self, level_stats) -> str:
        notes: list[str] = []
        if level_stats.clipping_ratio >= 0.02 or level_stats.peak_db >= -0.1:
            notes.append("音量が少し大きめかも")
        if level_stats.rms_db <= -30:
            notes.append("音量かなり小さめ")
        elif level_stats.rms_db <= -22:
            notes.append("音量やや小さめ")
        if level_stats.silent_ratio >= 0.8:
            notes.append("ほぼ無音")
        summary = " / ".join(notes) if notes else "大きな問題なし"
        return f"音量チェック: peak {level_stats.peak_db:.1f} dB / rms {level_stats.rms_db:.1f} dB / {summary}"

    def _play_trim_preview(self, path: Path, trim_start_ms: int, trim_end_ms: int) -> None:
        try:
            self.stop_audio()
            self._cleanup_preview_temp()
            self.preview_temp_path = create_trim_preview(path, trim_start_ms, trim_end_ms)
            self.audio_player.play(self.preview_temp_path)
            self.status_var.set(f"余白修正プレビュー再生中: {path.name}")
        except Exception as exc:
            self._cleanup_preview_temp()
            messagebox.showerror("余白修正エラー", str(exc))

    def open_trim_dialog(self, path: Path) -> None:
        if not path.exists():
            messagebox.showerror("余白修正エラー", f"音声ファイルが見つかりません。\n\n{path.name}")
            return

        try:
            metadata = get_audio_metadata(path)
            waveform_peaks = get_waveform_peaks(path)
            level_stats = analyze_audio_levels(path)
        except Exception as exc:
            messagebox.showerror("余白修正エラー", str(exc))
            return

        self.stop_audio()
        self._cleanup_preview_temp()

        if self.trim_dialog is not None and self.trim_dialog.winfo_exists():
            self.trim_dialog.destroy()

        dialog = ctk.CTkToplevel(self)
        self.trim_dialog = dialog
        self.trim_waveform_drag_handle = None
        dialog.title("余白修正")
        dialog.geometry("700x700")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        start_var = tk.IntVar(value=0)
        end_var = tk.IntVar(value=0)
        max_trim = max(metadata.duration_ms - 1, 0)

        filename_var = tk.StringVar(value=path.name)
        duration_var = tk.StringVar()
        start_label_var = tk.StringVar()
        end_label_var = tk.StringVar()
        remaining_var = tk.StringVar()
        level_var = tk.StringVar(value=self._format_level_summary(level_stats))
        improvement_var = tk.StringVar(value=self._level_improvement_text(level_stats) or "")

        def current_remaining_ms() -> int:
            return max(metadata.duration_ms - start_var.get() - end_var.get(), 1)

        def current_playhead_x(width: int) -> int | None:
            if self.preview_temp_path is None or self.audio_player.current_path != self.preview_temp_path or not self.audio_player.is_playing():
                return None
            keep_start_x = max(0, min(width, int(width * start_var.get() / max(metadata.duration_ms, 1))))
            keep_end_x = max(0, min(width, width - int(width * end_var.get() / max(metadata.duration_ms, 1))))
            if keep_end_x <= keep_start_x:
                return keep_start_x
            position_ratio = self.audio_player.current_position_ms() / max(current_remaining_ms(), 1)
            position_ratio = max(0.0, min(position_ratio, 1.0))
            return int(keep_start_x + (keep_end_x - keep_start_x) * position_ratio)

        def redraw_waveform() -> None:
            self._draw_trim_waveform(
                waveform_canvas,
                waveform_peaks,
                metadata.duration_ms,
                start_var.get(),
                end_var.get(),
                current_playhead_x(waveform_canvas.winfo_width() or 600),
            )

        def update_labels() -> None:
            trim_start = start_var.get()
            trim_end = end_var.get()
            remaining = current_remaining_ms()
            duration_var.set(f"元の長さ: {self._format_duration_ms(metadata.duration_ms)} / {metadata.duration_ms / 1000:.2f} 秒")
            start_label_var.set(f"先頭を削る: {self._format_trim_value(trim_start)}")
            end_label_var.set(f"末尾を削る: {self._format_trim_value(trim_end)}")
            remaining_var.set(f"適用後の長さ: {self._format_duration_ms(remaining)} / {remaining / 1000:.2f} 秒")
            redraw_waveform()

        def set_start_ms(value_ms: int) -> None:
            allowed = max(metadata.duration_ms - end_var.get() - 1, 0)
            start_var.set(max(0, min(value_ms, allowed)))
            update_labels()

        def set_end_ms(value_ms: int) -> None:
            allowed = max(metadata.duration_ms - start_var.get() - 1, 0)
            end_var.set(max(0, min(value_ms, allowed)))
            update_labels()

        def sync_start(value: str) -> None:
            set_start_ms(int(float(value)))

        def sync_end(value: str) -> None:
            set_end_ms(int(float(value)))

        def reload_audio_state(reset_sliders: bool) -> None:
            nonlocal metadata, waveform_peaks, level_stats, max_trim
            metadata = get_audio_metadata(path)
            waveform_peaks = get_waveform_peaks(path)
            level_stats = analyze_audio_levels(path)
            max_trim = max(metadata.duration_ms - 1, 0)
            start_slider.configure(to=max_trim, number_of_steps=max_trim if max_trim > 0 else 1)
            end_slider.configure(to=max_trim, number_of_steps=max_trim if max_trim > 0 else 1)
            if reset_sliders:
                start_var.set(0)
                end_var.set(0)
            level_var.set(self._format_level_summary(level_stats))
            improvement_var.set(self._level_improvement_text(level_stats) or "")
            improve_button.configure(state="normal" if improvement_var.get() else "disabled")
            refresh_backup_state()

        def refresh_backup_state() -> None:
            restore_button.configure(state="normal" if has_trim_backup(path) else "disabled")
            update_labels()

        def preview_audio() -> None:
            self._play_trim_preview(path, start_var.get(), end_var.get())
            schedule_playhead_refresh()

        def apply_trim() -> None:
            try:
                self.stop_audio()
                self._cleanup_preview_temp()
                apply_trim_in_place(path, start_var.get(), end_var.get())
                self._refresh_current_session_after_audio_edit(path)
                reload_audio_state(reset_sliders=False)
                self.status_var.set(f"余白修正を適用しました: {path.name}")
                messagebox.showinfo("完了", f"余白修正を適用しました。\n\n{path.name}", parent=dialog)
            except Exception as exc:
                messagebox.showerror("余白修正エラー", str(exc), parent=dialog)

        def improve_audio_level() -> None:
            try:
                self.stop_audio()
                self._cleanup_preview_temp()
                attenuate_audio_in_place(path, -3.0)
                self._refresh_current_session_after_audio_edit(path)
                reload_audio_state(reset_sliders=False)
                self.status_var.set(f"音量を少し下げました: {path.name}")
                messagebox.showinfo("完了", f"音量を少し下げました (-3 dB)。\n\n{path.name}", parent=dialog)
            except Exception as exc:
                messagebox.showerror("音量改善エラー", str(exc), parent=dialog)

        def restore_trim() -> None:
            try:
                self.stop_audio()
                self._cleanup_preview_temp()
                restore_trim_backup(path)
                self._refresh_current_session_after_audio_edit(path)
                reload_audio_state(reset_sliders=True)
                self.status_var.set(f"余白修正前に戻しました: {path.name}")
                messagebox.showinfo("完了", f"余白修正前の長さへ戻しました。\n\n{path.name}", parent=dialog)
            except Exception as exc:
                messagebox.showerror("余白修正エラー", str(exc), parent=dialog)

        def ms_from_x(x_position: int) -> int:
            width = max(waveform_canvas.winfo_width(), 1)
            ratio = max(0.0, min(x_position / width, 1.0))
            return int(metadata.duration_ms * ratio)

        def on_waveform_press(event: tk.Event) -> None:
            width = max(waveform_canvas.winfo_width(), 1)
            start_x = int(width * start_var.get() / max(metadata.duration_ms, 1))
            end_x = int(width - width * end_var.get() / max(metadata.duration_ms, 1))
            self.trim_waveform_drag_handle = "start" if abs(event.x - start_x) <= abs(event.x - end_x) else "end"
            on_waveform_drag(event)

        def on_waveform_drag(event: tk.Event) -> None:
            if self.trim_waveform_drag_handle == "start":
                set_start_ms(ms_from_x(event.x))
            elif self.trim_waveform_drag_handle == "end":
                set_end_ms(metadata.duration_ms - ms_from_x(event.x))

        def on_waveform_release(_event: tk.Event) -> None:
            self.trim_waveform_drag_handle = None

        def schedule_playhead_refresh() -> None:
            if not dialog.winfo_exists():
                return
            redraw_waveform()
            if self.preview_temp_path is not None and self.audio_player.current_path == self.preview_temp_path and self.audio_player.is_playing():
                dialog.after(80, schedule_playhead_refresh)

        def close_dialog() -> None:
            self.stop_audio()
            self._cleanup_preview_temp()
            if dialog.winfo_exists():
                dialog.grab_release()
                dialog.destroy()
            self.trim_dialog = None
            self.trim_waveform_drag_handle = None

        dialog.protocol("WM_DELETE_WINDOW", close_dialog)
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(dialog, text="余白修正", font=ctk.CTkFont(size=20, weight="bold"), anchor="w").grid(row=0, column=0, padx=20, pady=(18, 6), sticky="ew")
        ctk.CTkLabel(dialog, textvariable=filename_var, anchor="w").grid(row=1, column=0, padx=20, pady=(0, 4), sticky="ew")
        ctk.CTkLabel(dialog, textvariable=duration_var, anchor="w", text_color=("gray35", "gray70")).grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")

        body = ctk.CTkFrame(dialog)
        body.grid(row=3, column=0, padx=20, pady=8, sticky="nsew")
        body.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(body, text="波形イメージ", anchor="w").grid(row=0, column=0, padx=16, pady=(14, 4), sticky="ew")
        waveform_canvas = tk.Canvas(body, width=620, height=160, highlightthickness=0, bg="#1f1f1f", cursor="sb_h_double_arrow")
        waveform_canvas.grid(row=1, column=0, padx=16, pady=(0, 8), sticky="ew")
        waveform_canvas.bind("<ButtonPress-1>", on_waveform_press)
        waveform_canvas.bind("<B1-Motion>", on_waveform_drag)
        waveform_canvas.bind("<ButtonRelease-1>", on_waveform_release)
        ctk.CTkLabel(body, text="白い破線をドラッグして、先頭と末尾のカット位置を直接調整できます。", anchor="w", text_color=("gray35", "gray70")).grid(row=2, column=0, padx=16, pady=(0, 6), sticky="ew")
        ctk.CTkLabel(body, textvariable=level_var, anchor="w", text_color=("gray35", "gray70")).grid(row=3, column=0, padx=16, pady=(0, 4), sticky="ew")
        ctk.CTkLabel(body, textvariable=improvement_var, anchor="w", text_color=("#2563eb", "#93c5fd")).grid(row=4, column=0, padx=16, pady=(0, 12), sticky="ew")

        ctk.CTkLabel(body, text="先頭の余白", anchor="w").grid(row=5, column=0, padx=16, pady=(0, 4), sticky="ew")
        ctk.CTkLabel(body, textvariable=start_label_var, anchor="w", text_color=("gray35", "gray70")).grid(row=6, column=0, padx=16, pady=(0, 4), sticky="ew")
        start_slider = ctk.CTkSlider(body, from_=0, to=max_trim, number_of_steps=max_trim if max_trim > 0 else 1, variable=start_var, command=sync_start)
        start_slider.grid(row=7, column=0, padx=16, pady=(0, 12), sticky="ew")

        ctk.CTkLabel(body, text="末尾の余白", anchor="w").grid(row=8, column=0, padx=16, pady=(0, 4), sticky="ew")
        ctk.CTkLabel(body, textvariable=end_label_var, anchor="w", text_color=("gray35", "gray70")).grid(row=9, column=0, padx=16, pady=(0, 4), sticky="ew")
        end_slider = ctk.CTkSlider(body, from_=0, to=max_trim, number_of_steps=max_trim if max_trim > 0 else 1, variable=end_var, command=sync_end)
        end_slider.grid(row=10, column=0, padx=16, pady=(0, 12), sticky="ew")

        ctk.CTkLabel(body, textvariable=remaining_var, anchor="w").grid(row=11, column=0, padx=16, pady=(0, 16), sticky="ew")

        button_row = ctk.CTkFrame(dialog, fg_color="transparent")
        button_row.grid(row=4, column=0, padx=20, pady=(8, 20), sticky="ew")
        button_row.grid_columnconfigure(5, weight=1)

        ctk.CTkButton(button_row, text="試聴", width=90, command=preview_audio).grid(row=0, column=0, padx=(0, 8), sticky="w")
        improve_button = ctk.CTkButton(button_row, text="改善", width=90, command=improve_audio_level)
        improve_button.grid(row=0, column=1, padx=8, sticky="w")
        ctk.CTkButton(button_row, text="適用", width=90, command=apply_trim).grid(row=0, column=2, padx=8, sticky="w")
        restore_button = ctk.CTkButton(button_row, text="元の長さに戻す", width=130, fg_color=("#d5d5d5", "#4a4a4a"), hover_color=("#c8c8c8", "#5a5a5a"), command=restore_trim)
        restore_button.grid(row=0, column=3, padx=8, sticky="w")
        ctk.CTkButton(button_row, text="閉じる", width=90, fg_color=("#d5d5d5", "#4a4a4a"), hover_color=("#c8c8c8", "#5a5a5a"), command=close_dialog).grid(row=0, column=4, padx=(8, 0), sticky="w")

        waveform_canvas.bind("<Configure>", lambda _event: update_labels())
        refresh_backup_state()

    def start_row_drag(self, filename: str) -> None:
        if self.show_mode_var.get() != "全件":
            self.status_var.set("並べ替えは『全件』表示のときに使えます。")
            return
        self.dragging_filename = filename
        self.status_var.set(f"並べ替え中: {filename}")

    def finish_row_drag(self, filename: str, y_root: int) -> None:
        if self.dragging_filename != filename or self.show_mode_var.get() != "全件" or not self.current_folder:
            return
        session = self.folder_sessions[self.current_folder]
        target_row = min(
            self.file_rows,
            key=lambda row: abs((row.winfo_rooty() + row.winfo_height() / 2) - y_root),
            default=None,
        )
        self.dragging_filename = None
        if target_row is None:
            return
        self._move_manual_order(session, filename, target_row.file_item.original_filename)
        self._render_current_folder()
        self.status_var.set("並べ替えを更新しました。")

    def _move_manual_order(self, session: FolderSession, source_name: str, target_name: str) -> None:
        if source_name == target_name:
            return
        names = [file_item.original_filename for file_item in self._ordered_files(session)]
        if source_name not in names or target_name not in names:
            return
        names.remove(source_name)
        target_index = names.index(target_name)
        names.insert(target_index, source_name)
        session.manual_order = names

    def mark_all_ok(self) -> None:
        session = self._current_session_or_warn()
        if session is None:
            return
        for file_item in session.parse_result.files:
            session.ok_flags[file_item.original_filename] = True
            session.reviewed_flags[file_item.original_filename] = True
        self._render_current_folder()
        self._persist_workflow_state()

    def mark_all_ng(self) -> None:
        session = self._current_session_or_warn()
        if session is None:
            return
        for file_item in session.parse_result.files:
            session.ok_flags[file_item.original_filename] = False
            session.reviewed_flags[file_item.original_filename] = True
        self._render_current_folder()
        self._persist_workflow_state()

    def _suggest_missing_indices(self, session: FolderSession) -> set[int]:
        candidates: set[int] = set()
        detected = session.parse_result.detected_indices
        if detected:
            for index in range(min(detected), max(detected) + 1):
                if index not in detected:
                    candidates.add(index)
        for index in detected:
            has_ok = any(
                session.ok_flags.get(file_item.original_filename, True)
                for file_item in session.parse_result.files
                if file_item.original_index == index
            )
            if not has_ok:
                candidates.add(index)
        return candidates

    def apply_missing_suggestions(self) -> None:
        session = self._current_session_or_warn()
        if session is None:
            return
        suggestions = self._suggest_missing_indices(session)
        session.missing_indices.update(suggestions)
        self._render_missing_checkboxes(session)
        self._update_warnings(session)
        self.status_var.set(f"{len(suggestions)} 件の欠番候補をチェックしました。")
        self._persist_workflow_state()

    def clear_missing_checks(self) -> None:
        session = self._current_session_or_warn()
        if session is None:
            return
        session.missing_indices.clear()
        self._render_missing_checkboxes(session)
        self._update_warnings(session)
        self._persist_workflow_state()

    def _current_session_or_warn(self) -> FolderSession | None:
        session = self.folder_sessions.get(self.current_folder) if self.current_folder else None
        if session is None:
            messagebox.showwarning("未選択", "表示中の対象がありません。")
        return session

    def _prepare_folder_plans(self) -> tuple[list[tuple[Path, FolderSession, list[RenamePlanEntry]]], list[str]]:
        self._save_current_state()
        settings = RenameSettings(
            digits=int(self.digits_var.get()[0]),
            keep_text=self.keep_text_var.get(),
            move_ng_files=self.move_ng_var.get(),
            export_csv=self.export_csv_var.get(),
        )
        folder_plans: list[tuple[Path, FolderSession, list[RenamePlanEntry]]] = []
        skipped_folders: list[str] = []
        for folder in self.folder_order:
            session = self.folder_sessions[folder]
            ordered_files = self._ordered_files(session)
            if not ordered_files:
                skipped_folders.append(f"{folder.name}: 対象 wav ファイルなし")
                continue
            rename_plan = build_rename_plan(ordered_files, session.ok_flags, session.missing_indices, settings)
            ok_count = sum(1 for entry in rename_plan if entry.status == "OK")
            if ok_count == 0:
                skipped_folders.append(f"{folder.name}: OK ファイルなし")
                continue
            folder_plans.append((folder, session, rename_plan))
        return folder_plans, skipped_folders

    def _plan_display_rows(
        self,
        folder_plans: list[tuple[Path, FolderSession, list[RenamePlanEntry]]],
        skipped_folders: list[str],
    ) -> list[tuple[str, ...]]:
        rows: list[tuple[str, ...]] = []
        for folder, session, plan in folder_plans:
            label = self._session_label(session)
            duplicate_set = set(session.parse_result.duplicate_indices)
            for entry in plan:
                trim_info = ""
                peak_info = ""
                note_parts: list[str] = []
                if entry.source_path and entry.source_path.exists():
                    trim_info = "あり" if has_trim_backup(entry.source_path) else ""
                    peak_info = f"{analyze_audio_levels(entry.source_path).peak_db:.1f}"
                    if not session.reviewed_flags.get(entry.original_filename, False):
                        note_parts.append("未確認")
                if entry.original_index in duplicate_set:
                    note_parts.append("重複番号")
                if entry.status == "MISSING":
                    note_parts.append("意図的欠番")
                rows.append((
                    label,
                    entry.status,
                    str(entry.original_index),
                    "" if entry.new_index is None else str(entry.new_index),
                    entry.original_filename,
                    entry.new_filename,
                    trim_info,
                    peak_info,
                    " / ".join(note_parts),
                ))
        if skipped_folders:
            rows.append(("-", "未処理", "", "", "", "", "", "", " / ".join(skipped_folders)))
        return rows

    def _show_plan_dialog(
        self,
        title: str,
        rows: list[tuple[str, ...]],
        confirm_mode: bool = False,
    ) -> bool:
        dialog = ctk.CTkToplevel(self)
        dialog.title(title)
        dialog.geometry("1320x760")
        dialog.transient(self)
        dialog.grab_set()

        decision = {"confirmed": not confirm_mode}
        ctk.CTkLabel(dialog, text=title, anchor="w").pack(fill="x", padx=16, pady=(16, 8))

        container = ttk.Frame(dialog)
        container.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        columns = ("folder", "status", "old_index", "new_index", "old_name", "new_name", "trim", "peak", "note")
        tree = ttk.Treeview(container, columns=columns, show="headings")
        labels = {
            "folder": "対象",
            "status": "状態",
            "old_index": "元番号",
            "new_index": "新番号",
            "old_name": "元ファイル名",
            "new_name": "新ファイル名",
            "trim": "余白修正",
            "peak": "peak(dB)",
            "note": "備考",
        }
        widths = {
            "folder": 180,
            "status": 80,
            "old_index": 70,
            "new_index": 70,
            "old_name": 240,
            "new_name": 240,
            "trim": 90,
            "peak": 80,
            "note": 220,
        }
        for column in columns:
            tree.heading(column, text=labels[column])
            tree.column(column, width=widths[column], anchor="w")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        for row in rows:
            tree.insert("", "end", values=row)

        def close_with(value: bool) -> None:
            decision["confirmed"] = value
            dialog.grab_release()
            dialog.destroy()

        button_row = ctk.CTkFrame(dialog, fg_color="transparent")
        button_row.pack(fill="x", padx=16, pady=(0, 16))
        if confirm_mode:
            ctk.CTkButton(button_row, text="キャンセル", width=110, fg_color=("#d5d5d5", "#4a4a4a"), hover_color=("#c8c8c8", "#5a5a5a"), command=lambda: close_with(False)).pack(side="right", padx=(8, 0))
            ctk.CTkButton(button_row, text="この内容で実行", width=140, command=lambda: close_with(True)).pack(side="right")
        else:
            ctk.CTkButton(button_row, text="閉じる", width=110, command=lambda: close_with(True)).pack(side="right")

        dialog.wait_window()
        return bool(decision["confirmed"])

    def _report_root_path(self, folders: list[Path]) -> Path:
        common = Path(os.path.commonpath([str(folder) for folder in folders]))
        return common if common.is_dir() else folders[0].parent

    def _build_project_report_rows(
        self,
        folder_plans: list[tuple[Path, FolderSession, list[RenamePlanEntry]]],
    ) -> list[dict[str, str]]:
        rows: list[dict[str, str]] = []
        for folder, session, plan in folder_plans:
            source_entries = [entry for entry in plan if entry.source_path and entry.source_path.exists()]
            peaks = [analyze_audio_levels(entry.source_path).peak_db for entry in source_entries if entry.source_path is not None]
            rows.append({
                "folder_name": folder.name,
                "folder_path": str(folder),
                "ok_count": str(sum(1 for entry in plan if entry.status == "OK")),
                "ng_count": str(sum(1 for entry in plan if entry.status == "NG")),
                "missing_count": str(sum(1 for entry in plan if entry.status == "MISSING")),
                "trim_modified_count": str(sum(1 for entry in source_entries if entry.source_path is not None and has_trim_backup(entry.source_path))),
                "duplicate_indices": ",".join(f"{index:03d}" for index in session.parse_result.duplicate_indices),
                "missing_indices": ",".join(f"{index:03d}" for index in sorted(session.missing_indices)),
                "unreviewed_count": str(sum(1 for reviewed in session.reviewed_flags.values() if not reviewed)),
                "peak_max_db": "" if not peaks else f"{max(peaks):.1f}",
            })
        return rows

    def _write_project_report(
        self,
        folder_plans: list[tuple[Path, FolderSession, list[RenamePlanEntry]]],
        processed_at: str,
    ) -> tuple[Path, Path]:
        folders = [folder for folder, _session, _plan in folder_plans]
        report_root = self._report_root_path(folders)
        safe_stamp = processed_at.replace(":", "").replace("-", "").replace(" ", "_")
        csv_path = report_root / f"batch_report_{safe_stamp}.csv"
        txt_path = report_root / f"batch_report_{safe_stamp}.txt"
        rows = self._build_project_report_rows(folder_plans)
        write_rename_log(rows, csv_path)

        summary_lines = [f"案件レポート {processed_at}"]
        for row in rows:
            summary_lines.extend([
                "",
                f"対象: {row['folder_name']}",
                f"OK: {row['ok_count']} / NG: {row['ng_count']} / 欠番: {row['missing_count']}",
                f"余白修正済み: {row['trim_modified_count']} / 未確認: {row['unreviewed_count']}",
                f"重複番号: {row['duplicate_indices'] or '-'}",
                f"欠番指定: {row['missing_indices'] or '-'}",
                f"peak 最大値: {row['peak_max_db'] or '-'} dB",
                f"パス: {row['folder_path']}",
            ])
        txt_path.write_text("\n".join(summary_lines), encoding="utf-8")
        return csv_path, txt_path

    def open_preview_dialog(self) -> None:
        if not self.folder_order:
            messagebox.showwarning("未選択", "先にフォルダまたはファイルを追加してください。")
            return
        folder_plans, skipped_folders = self._prepare_folder_plans()
        if not folder_plans:
            messagebox.showwarning("処理対象なし", "プレビューできる対象がありません。")
            return
        rows = self._plan_display_rows(folder_plans, skipped_folders)
        self._show_plan_dialog("最終プレビュー", rows, confirm_mode=False)

    def _build_log_rows(
        self,
        folder: Path,
        session: FolderSession,
        plan: list[RenamePlanEntry],
        processed_at: str,
    ) -> list[dict[str, str]]:
        duplicate_set = set(session.parse_result.duplicate_indices)
        rows: list[dict[str, str]] = []
        for entry in plan:
            note_parts: list[str] = []
            peak_db = ""
            trim_modified = ""
            reviewed = ""
            if entry.original_index in duplicate_set:
                note_parts.append("duplicate-index")
            if entry.status == "MISSING":
                note_parts.append("manual-missing")
            if entry.source_path and entry.source_path.exists():
                peak_db = f"{analyze_audio_levels(entry.source_path).peak_db:.1f}"
                trim_modified = "yes" if has_trim_backup(entry.source_path) else ""
                reviewed = "yes" if session.reviewed_flags.get(entry.original_filename, False) else "no"
            rows.append(
                {
                    "original_filename": entry.original_filename,
                    "new_filename": entry.new_filename,
                    "original_index": str(entry.original_index),
                    "new_index": "" if entry.new_index is None else str(entry.new_index),
                    "status": entry.status,
                    "folder_path": str(folder),
                    "processed_at": processed_at,
                    "note": ";".join(note_parts),
                    "peak_db": peak_db,
                    "trim_modified": trim_modified,
                    "reviewed": reviewed,
                }
            )
        return rows

    def _apply_post_rename_session_state(
        self,
        session: FolderSession,
        plan: list[RenamePlanEntry],
        settings: RenameSettings,
    ) -> None:
        session.undo_selected_filenames = None if session.selected_filenames is None else set(session.selected_filenames)
        session.undo_manual_order = list(session.manual_order)

        old_reviewed = dict(session.reviewed_flags)
        new_reviewed: dict[str, bool] = {}
        new_order: list[str] = []
        new_selected: set[str] = set()

        for entry in plan:
            if entry.status == "OK" and entry.new_filename:
                new_name = entry.new_filename
                new_reviewed[new_name] = old_reviewed.get(entry.original_filename, False)
                new_order.append(new_name)
                if session.selected_filenames is not None:
                    new_selected.add(new_name)
            elif entry.status == "NG" and not settings.move_ng_files:
                new_reviewed[entry.original_filename] = old_reviewed.get(entry.original_filename, False)
                new_order.append(entry.original_filename)
                if session.selected_filenames is not None:
                    new_selected.add(entry.original_filename)

        if session.selected_filenames is not None:
            session.selected_filenames = new_selected
            session.manual_order = new_order
        else:
            session.manual_order = new_order
        session.reviewed_flags = new_reviewed
        self._refresh_session(session)

    def undo_current_folder(self) -> None:
        session = self._current_session_or_warn()
        if session is None:
            return
        if not has_undo_manifest(session.folder):
            messagebox.showwarning("履歴なし", "この対象には戻せるリネーム履歴がありません。")
            return
        try:
            self.stop_audio()
            self.progress.set(0)
            undo_last_rename(session.folder, progress_callback=self.progress.set)
            session.selected_filenames = None if session.undo_selected_filenames is None else set(session.undo_selected_filenames)
            session.manual_order = list(session.undo_manual_order)
            self._refresh_session(session)
            self._render_current_folder()
            self._refresh_folder_list()
            self._persist_workflow_state()
            self.progress.set(1)
            self.status_var.set(f"リネーム前に戻しました: {session.folder.name}")
            messagebox.showinfo("完了", f"{session.folder.name} をリネーム前の状態へ戻しました。")
        except Exception as exc:
            self.progress.set(0)
            messagebox.showerror("取り消しエラー", str(exc))

    def run_rename(self) -> None:
        if not self.folder_order:
            messagebox.showwarning("未選択", "先にフォルダまたはファイルを追加してください。")
            return
        folder_plans, skipped_folders = self._prepare_folder_plans()
        if not folder_plans:
            messagebox.showwarning("処理対象なし", "リネームできる対象がありません。")
            return

        rows = self._plan_display_rows(folder_plans, skipped_folders)
        if not self._show_plan_dialog("最終確認", rows, confirm_mode=True):
            self.status_var.set("リネーム実行をキャンセルしました。")
            return

        settings = RenameSettings(
            digits=int(self.digits_var.get()[0]),
            keep_text=self.keep_text_var.get(),
            move_ng_files=self.move_ng_var.get(),
            export_csv=self.export_csv_var.get(),
        )

        try:
            self.stop_audio()
            self.progress.set(0)
            total_folders = len(folder_plans)
            completed_folders = 0
            processed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            for folder, session, rename_plan in folder_plans:
                self.status_var.set(f"処理中: {self._session_label(session)}")
                self.update_idletasks()
                write_undo_manifest(rename_plan, folder, settings)
                log_rows = self._build_log_rows(folder, session, rename_plan, processed_at) if settings.export_csv else []

                def progress_callback(local_progress: float, base: int = completed_folders) -> None:
                    self.progress.set((base + local_progress) / total_folders)

                execute_rename_plan(rename_plan, folder, settings, progress_callback=progress_callback)

                if settings.export_csv:
                    csv_path = folder / "rename_log.csv"
                    write_rename_log(log_rows, csv_path)

                self._apply_post_rename_session_state(session, rename_plan, settings)
                completed_folders += 1
                self.progress.set(completed_folders / total_folders)

            report_csv_path, report_txt_path = self._write_project_report(folder_plans, processed_at)
            self.status_var.set(f"リネームが完了しました。{completed_folders} 件の対象を処理しました。")
            self._render_current_folder()
            self._refresh_folder_list()
            self._persist_workflow_state()

            completion_lines = ["リネームが完了しました。", f"処理対象数: {completed_folders}", f"案件レポートCSV: {report_csv_path}", f"案件レポートTXT: {report_txt_path}"]
            if settings.export_csv:
                completion_lines.append("各フォルダに rename_log.csv を保存しました。")
            if skipped_folders:
                completion_lines.extend(["", "未処理:", *skipped_folders])
            messagebox.showinfo("完了", "\n".join(completion_lines))
        except Exception as exc:
            self.progress.set(0)
            self.status_var.set("エラーが発生しました。")
            messagebox.showerror("エラー", f"処理に失敗しました。\n\n{exc}")
