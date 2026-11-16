"""Excel VBAエクスポートツールのGUI。"""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tkinterdnd2 import COPY, DND_FILES, REFUSE_DROP, TkinterDnD

from excel_vba_exporter.exporter import (
    SUPPORTED_EXTENSIONS,
    ExportResult,
    export_workbook,
)


class ExporterApp:
    """ファイル選択とVBAエクスポートを提供するGUI。"""

    def __init__(self) -> None:
        """ウィンドウと入力フォームを初期化する。"""
        self.root = TkinterDnD.Tk()
        self.root.title("Excel VBA エクスポーター")
        self.root.minsize(620, 280)
        self.root.protocol("WM_DELETE_WINDOW", self._close_window)

        self.workbook_path = tk.StringVar()
        self.output_directory = tk.StringVar()
        self.status = tk.StringVar(value="Excelファイルと出力先を指定してください。")
        self._busy = False

        self._build_widgets()
        self._register_drop_targets(self.root)

    def _register_drop_targets(self, widget: tk.Misc) -> None:
        """ウィンドウとすべての子ウィジェットをドロップ先に登録する。

        Args:
            widget: 登録するウィジェットの親。

        """
        widget.drop_target_register(DND_FILES)
        widget.dnd_bind("<<Drop>>", self._handle_drop)
        for child in widget.winfo_children():
            self._register_drop_targets(child)

    def _build_widgets(self) -> None:
        """入力欄、参照ボタン、実行ボタンを配置する。"""
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="対象Excelファイル").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 12), pady=(0, 12)
        )
        self.workbook_entry = ttk.Entry(frame, textvariable=self.workbook_path)
        self.workbook_entry.grid(row=0, column=1, sticky=tk.EW, pady=(0, 12))
        ttk.Button(frame, text="参照...", command=self._choose_workbook).grid(
            row=0, column=2, sticky=tk.E, padx=(8, 0), pady=(0, 12)
        )
        ttk.Label(
            frame,
            text="ウィンドウ内のどこへでもExcelファイルをドロップできます。",
        ).grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=(0, 16))

        ttk.Label(frame, text="出力先フォルダー").grid(
            row=2, column=0, sticky=tk.W, padx=(0, 12), pady=(0, 12)
        )
        ttk.Entry(frame, textvariable=self.output_directory).grid(
            row=2, column=1, sticky=tk.EW, pady=(0, 12)
        )
        ttk.Button(frame, text="参照...", command=self._choose_output_directory).grid(
            row=2, column=2, sticky=tk.E, padx=(8, 0), pady=(0, 12)
        )

        ttk.Label(
            frame,
            text="標準モジュール・クラスモジュール・フォームを出力します。"
            "ブック／シートのコードは対象外です。",
            wraplength=560,
        ).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(0, 16))
        ttk.Label(frame, textvariable=self.status, wraplength=560).grid(
            row=4, column=0, columnspan=3, sticky=tk.W, pady=(0, 12)
        )
        self.export_button = ttk.Button(
            frame, text="エクスポート", command=self._start_export
        )
        self.export_button.grid(row=5, column=2, sticky=tk.E)

    def _choose_workbook(self) -> None:
        """エクスポート元のExcelファイルを選択する。"""
        filetypes = [
            ("Excel マクロ対応ファイル", "*.xls *.xlsm *.xlsb *.xlam"),
            ("すべてのファイル", "*.*"),
        ]
        selected = filedialog.askopenfilename(
            parent=self.root,
            title="対象Excelファイルを選択",
            filetypes=filetypes,
        )
        if selected:
            self.workbook_path.set(selected)

    def _choose_output_directory(self) -> None:
        """エクスポート先フォルダーを選択する。"""
        selected = filedialog.askdirectory(
            parent=self.root,
            title="出力先フォルダーを選択",
            mustexist=True,
        )
        if selected:
            self.output_directory.set(selected)

    def _handle_drop(self, event: tk.Event) -> str:
        """ドロップされたExcelファイルのパスを対象欄に反映する。

        Args:
            event: TkinterDnDが通知するドロップイベント。

        Returns:
            TkinterDnD用のイベント処理結果。

        """
        if self._busy:
            return REFUSE_DROP
        paths = self.root.tk.splitlist(event.data)
        if len(paths) != 1:
            messagebox.showwarning(
                "ファイルを選択してください",
                "一度にドロップできるファイルは1つです。",
                parent=self.root,
            )
            return REFUSE_DROP

        dropped_path = Path(paths[0])
        if dropped_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            messagebox.showerror(
                "対応していないファイル形式",
                "対応しているファイル形式は .xls、.xlsm、.xlsb、.xlam です。",
                parent=self.root,
            )
            return REFUSE_DROP

        self.workbook_path.set(str(dropped_path))
        return COPY

    def _start_export(self) -> None:
        """入力を検証し、バックグラウンドでエクスポートを開始する。"""
        workbook = self.workbook_path.get().strip()
        output = self.output_directory.get().strip()
        if not workbook or not output:
            messagebox.showerror(
                "入力が不足しています",
                "対象Excelファイルと出力先フォルダーを指定してください。",
                parent=self.root,
            )
            return

        source_path = Path(workbook)
        if source_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            messagebox.showerror(
                "対応していないファイル形式",
                "対応しているファイル形式は .xls、.xlsm、.xlsb、.xlam です。",
                parent=self.root,
            )
            return
        if not source_path.is_file():
            messagebox.showerror(
                "対象ファイルが見つかりません",
                f"指定されたファイルを確認してください:\n{source_path}",
                parent=self.root,
            )
            return
        if not Path(output).is_dir():
            messagebox.showerror(
                "出力先が見つかりません",
                f"指定されたフォルダーを確認してください:\n{output}",
                parent=self.root,
            )
            return

        self._busy = True
        self.export_button.configure(state=tk.DISABLED)
        self.status.set("エクスポートしています...")
        threading.Thread(
            target=self._export_in_background,
            args=(source_path, Path(output)),
            daemon=True,
        ).start()

    def _export_in_background(self, source: Path, output: Path) -> None:
        """GUIを停止させずにエクスポートし、結果をUIスレッドへ渡す。

        Args:
            source: エクスポート元Excelファイル。
            output: 選択された出力先フォルダー。

        """
        try:
            result = export_workbook(source, output)
        except Exception as exc:
            self.root.after(0, self._finish_export, None, exc)
        else:
            self.root.after(0, self._finish_export, result, None)

    def _finish_export(
        self, result: ExportResult | None, error: Exception | None
    ) -> None:
        """エクスポート完了を表示し、操作を再び有効にする。

        Args:
            result: 正常終了時のエクスポート結果。
            error: 失敗時の例外。

        """
        self._busy = False
        self.export_button.configure(state=tk.NORMAL)
        if error is not None:
            self.status.set("エクスポートに失敗しました。")
            messagebox.showerror(
                "エクスポートに失敗しました",
                str(error),
                parent=self.root,
            )
            return

        assert result is not None
        count = len(result.files)
        self.status.set(f"エクスポートが完了しました（{count}件）。")
        if count:
            file_list = "\n".join(path.name for path in result.files)
            message = f"{count}件のファイルを出力しました:\n\n{file_list}"
        else:
            message = (
                "エクスポート対象の標準モジュール、クラスモジュール、"
                "フォームは見つかりませんでした。"
            )
        messagebox.showinfo("エクスポート完了", message, parent=self.root)

    def _close_window(self) -> None:
        """エクスポート処理中の終了を防ぎ、処理後に終了できるようにする。"""
        if self._busy:
            messagebox.showwarning(
                "処理中です",
                "エクスポートが完了するまでウィンドウを閉じられません。",
                parent=self.root,
            )
            return
        self.root.destroy()

    def run(self) -> None:
        """Tkinterのイベントループを開始する。"""
        self.root.mainloop()


def main() -> None:
    """GUIアプリケーションを起動する。"""
    ExporterApp().run()
