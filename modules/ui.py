from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
from tkinterdnd2 import DND_FILES, TkinterDnD

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
from modules.settings_store import load_settings, save_settings


ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")


TABLE_COLUMN_WIDTHS = {
    0: 52,
    1: 72,
    2: 72,
    3: 72,
    4: 580,
    5: 90,
    6: 240,
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


class FileRow(ctk.CTkFrame):
    def __init__(
        self,
        master: tk.Misc,
        file_item: ParsedAudioFile,
        initial_ok: bool,
        on_status_change,
        on_play_toggle,
        on_drag_start,
        on_drag_end,
    ) -> None:
        super().__init__(master, fg_color=("#f4f4f4", "#2f2f2f") if file_item.duplicate_index else "transparent")
        self.file_item = file_item
        self._on_status_change = on_status_change
        self._on_play_toggle = on_play_toggle
        self._on_drag_start = on_drag_start
        self._on_drag_end = on_drag_end
        self.ok_var = tk.BooleanVar(value=initial_ok)
        self.ng_var = tk.BooleanVar(value=not initial_ok)

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
        self.play_button.grid(row=0, column=3, padx=(8, 8), pady=5, sticky="w")

        duplicate_suffix = "  [重複]" if self.file_item.duplicate_index else ""
        ctk.CTkLabel(
            self,
            text=f"{self.file_item.original_filename}{duplicate_suffix}",
            anchor="w",
            text_color="#d97706" if self.file_item.duplicate_index else None,
        ).grid(row=0, column=4, padx=(10, 10), pady=5, sticky="ew")
        ctk.CTkLabel(self, text=str(self.file_item.original_index), anchor="center").grid(
            row=0, column=5, padx=8, pady=5, sticky="nsew"
        )
        preview_text = self.file_item.text_portion if self.file_item.text_portion else "(なし)"
        ctk.CTkLabel(self, text=preview_text, anchor="w").grid(
            row=0, column=6, padx=(10, 12), pady=5, sticky="ew"
        )

    def _configure_columns(self) -> None:
        for column, minsize in TABLE_COLUMN_WIDTHS.items():
            self.grid_columnconfigure(column, minsize=minsize, weight=1 if column == 4 else 0)

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
        toolbar.grid_columnconfigure(6, weight=1)

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
            4: ("ファイル名", "w", (10, 10)),
            5: ("元番号", "center", (0, 0)),
            6: ("テキスト部分", "w", (10, 12)),
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
            batch_row.grid_columnconfigure(column, minsize=minsize, weight=1 if column == 4 else 0)
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

    def on_close(self) -> None:
        self._persist_settings()
        self.audio_player.stop()
        self.destroy()

    def _configure_table_columns(self, frame: ctk.CTkFrame) -> None:
        for column, minsize in TABLE_COLUMN_WIDTHS.items():
            frame.grid_columnconfigure(column, minsize=minsize, weight=1 if column == 4 else 0)

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

    def _refresh_session(self, session: FolderSession) -> None:
        parse_result = parse_audio_folder(session.folder, session.selected_filenames)
        session.parse_result = parse_result
        existing_ok = dict(session.ok_flags)
        session.ok_flags = {file_item.original_filename: existing_ok.get(file_item.original_filename, True) for file_item in parse_result.files}
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
                self._on_file_status_change,
                self.toggle_play_audio,
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
            ctk.CTkCheckBox(self.missing_scroll, text=f"{index:03d}", variable=var, command=self._save_current_state).grid(
                row=row_number, column=0, padx=8, pady=4, sticky="w"
            )

    def _update_warnings(self, session: FolderSession | None) -> None:
        if session is None:
            self.warning_label.configure(text="重複番号、除外ファイル、Undo 状態をここに表示します。")
            return
        warnings: list[str] = [f"表示中: {self._session_label(session)}"]
        if session.selected_filenames is not None:
            warnings.append("ファイル単体選択モード: 左端の ↕ で行順を並べ替えできます。")
        if session.parse_result.duplicate_indices:
            warnings.append("重複番号: " + ", ".join(f"{index:03d}" for index in session.parse_result.duplicate_indices))
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
        self._update_warnings(session)

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
        self._render_current_folder()

    def mark_all_ng(self) -> None:
        session = self._current_session_or_warn()
        if session is None:
            return
        for file_item in session.parse_result.files:
            session.ok_flags[file_item.original_filename] = False
        self._render_current_folder()

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

    def clear_missing_checks(self) -> None:
        session = self._current_session_or_warn()
        if session is None:
            return
        session.missing_indices.clear()
        self._render_missing_checkboxes(session)
        self._update_warnings(session)

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

    def open_preview_dialog(self) -> None:
        if not self.folder_order:
            messagebox.showwarning("未選択", "先にフォルダまたはファイルを追加してください。")
            return
        folder_plans, skipped_folders = self._prepare_folder_plans()
        if not folder_plans:
            messagebox.showwarning("処理対象なし", "プレビューできる対象がありません。")
            return

        preview = ctk.CTkToplevel(self)
        preview.title("試行プレビュー")
        preview.geometry("1200x700")
        preview.transient(self)
        preview.grab_set()
        ctk.CTkLabel(preview, text="実際の変更予定一覧", anchor="w").pack(fill="x", padx=16, pady=(16, 8))

        container = ttk.Frame(preview)
        container.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        columns = ("folder", "status", "old_index", "new_index", "old_name", "new_name", "note")
        tree = ttk.Treeview(container, columns=columns, show="headings")
        labels = {
            "folder": "対象",
            "status": "状態",
            "old_index": "元番号",
            "new_index": "新番号",
            "old_name": "元ファイル名",
            "new_name": "新ファイル名",
            "note": "備考",
        }
        widths = {"folder": 200, "status": 80, "old_index": 80, "new_index": 80, "old_name": 260, "new_name": 260, "note": 160}
        for column in columns:
            tree.heading(column, text=labels[column])
            tree.column(column, width=widths[column], anchor="w")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        for folder, session, plan in folder_plans:
            label = self._session_label(session)
            duplicate_set = set(session.parse_result.duplicate_indices)
            for entry in plan:
                note = ""
                if entry.original_index in duplicate_set:
                    note = "重複番号あり"
                elif entry.status == "MISSING":
                    note = "意図的欠番"
                tree.insert(
                    "",
                    "end",
                    values=(
                        label,
                        entry.status,
                        entry.original_index,
                        "" if entry.new_index is None else entry.new_index,
                        entry.original_filename,
                        entry.new_filename,
                        note,
                    ),
                )
        if skipped_folders:
            tree.insert("", "end", values=("-", "未処理", "", "", "", "", " / ".join(skipped_folders)))

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
            note = ""
            if entry.original_index in duplicate_set:
                note = "duplicate-index"
            elif entry.status == "MISSING":
                note = "manual-missing"
            rows.append(
                {
                    "original_filename": entry.original_filename,
                    "new_filename": entry.new_filename,
                    "original_index": str(entry.original_index),
                    "new_index": "" if entry.new_index is None else str(entry.new_index),
                    "status": entry.status,
                    "folder_path": str(folder),
                    "processed_at": processed_at,
                    "note": note,
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

        if session.selected_filenames is not None:
            updated_names: set[str] = set()
            updated_order: list[str] = []
            for entry in plan:
                if entry.status == "OK" and entry.new_filename:
                    updated_names.add(entry.new_filename)
                    updated_order.append(entry.new_filename)
                elif entry.status == "NG" and not settings.move_ng_files:
                    updated_names.add(entry.original_filename)
                    updated_order.append(entry.original_filename)
            session.selected_filenames = updated_names
            session.manual_order = updated_order
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

            for folder, session, rename_plan in folder_plans:
                self.status_var.set(f"処理中: {self._session_label(session)}")
                self.update_idletasks()
                write_undo_manifest(rename_plan, folder, settings)

                def progress_callback(local_progress: float, base: int = completed_folders) -> None:
                    self.progress.set((base + local_progress) / total_folders)

                execute_rename_plan(rename_plan, folder, settings, progress_callback=progress_callback)

                processed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                if settings.export_csv:
                    csv_path = folder / "rename_log.csv"
                    write_rename_log(self._build_log_rows(folder, session, rename_plan, processed_at), csv_path)

                self._apply_post_rename_session_state(session, rename_plan, settings)
                completed_folders += 1
                self.progress.set(completed_folders / total_folders)

            self.status_var.set(f"リネームが完了しました。{completed_folders} 件の対象を処理しました。")
            self._render_current_folder()
            self._refresh_folder_list()

            completion_lines = ["リネームが完了しました。", f"処理対象数: {completed_folders}"]
            if settings.export_csv:
                completion_lines.append("各フォルダに rename_log.csv を保存しました。")
            if skipped_folders:
                completion_lines.extend(["", "未処理:", *skipped_folders])
            messagebox.showinfo("完了", "\n".join(completion_lines))
        except Exception as exc:
            self.progress.set(0)
            self.status_var.set("エラーが発生しました。")
            messagebox.showerror("エラー", f"処理に失敗しました。\n\n{exc}")
