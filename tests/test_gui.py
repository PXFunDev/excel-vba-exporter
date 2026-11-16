"""GUI全体のドロップ受付と対象パス反映のテスト。"""

from collections.abc import Iterator
from pathlib import Path
from tkinter import Misc
from types import SimpleNamespace

import pytest
from tkinterdnd2 import COPY, REFUSE_DROP

from excel_vba_exporter.gui import ExporterApp


@pytest.fixture(scope="module")
def app() -> Iterator[ExporterApp]:
    """実際のTkDnDウィンドウを非表示で作成し、必ず終了する。

    Yields:
        テスト用GUIアプリケーション。

    """
    application = ExporterApp()
    application.root.withdraw()
    try:
        yield application
    finally:
        application.root.destroy()


def test_drop_on_every_widget_updates_workbook_path(
    app: ExporterApp, tmp_path: Path
) -> None:
    """ボタン・ラベル・余白を含む全ウィジェットでパスを反映する。"""
    source = tmp_path / "日本語 空白" / "sample.xlsm"
    widgets: list[Misc] = [app.root]
    for widget in widgets:
        widgets.extend(widget.winfo_children())
        binding = app.root.tk.splitlist(widget.dnd_bind("<<Drop>>"))
        assert binding, str(widget)
        substitutions = [""] * 18
        substitutions[7] = app.root.tk.call("list", str(source))
        substitutions[15] = str(widget)
        app.workbook_path.set("")

        action = app.root.tk.call(binding[0], *substitutions)

        assert action == COPY
        assert app.workbook_path.get() == str(source)


def test_drop_during_export_does_not_change_workbook(app: ExporterApp) -> None:
    """処理中のドロップを拒否して表示中の対象を維持する。"""
    app.workbook_path.set("original.xlsm")
    app._busy = True

    action = app._handle_drop(SimpleNamespace(data="other.xlsm"))
    app._busy = False

    assert action == REFUSE_DROP
    assert app.workbook_path.get() == "original.xlsm"


@pytest.mark.parametrize("data", ["one.xlsm two.xlsm", "unsupported.txt"])
def test_invalid_drop_is_rejected(
    app: ExporterApp, monkeypatch: pytest.MonkeyPatch, data: str
) -> None:
    """複数ファイルと非対応形式を拒否して対象パスを変更しない。"""
    notifications: list[str] = []
    monkeypatch.setattr(
        "excel_vba_exporter.gui.messagebox.showwarning",
        lambda *args, **kwargs: notifications.append("warning"),
    )
    monkeypatch.setattr(
        "excel_vba_exporter.gui.messagebox.showerror",
        lambda *args, **kwargs: notifications.append("error"),
    )
    app.workbook_path.set("original.xlsm")

    action = app._handle_drop(SimpleNamespace(data=data))

    assert action == REFUSE_DROP
    assert app.workbook_path.get() == "original.xlsm"
    assert len(notifications) == 1
