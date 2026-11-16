"""VBAコンポーネントの出力計画とExcel連携のテスト。"""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from excel_vba_exporter import exporter
from excel_vba_exporter.exporter import ExportError, build_export_plan, export_workbook


class FakeComponent:
    """COMコンポーネントの最小テストダブル。"""

    def __init__(self, name: str, component_type: int) -> None:
        """テスト用コンポーネントを初期化する。

        Args:
            name: コンポーネント名。
            component_type: VBAコンポーネントの種別番号。

        """
        self.Name = name
        self.Type = component_type

    def Export(self, destination: str) -> None:
        """出力先にテスト用ファイルを作成する.

        Args:
            destination: Excel COMから渡される出力先パス。

        """
        Path(destination).write_text("test", encoding="utf-8")


class FakeWorkbook:
    """COMブックの最小テストダブル。"""

    def __init__(self, components: list[FakeComponent]) -> None:
        """テスト用VBAコンポーネント一覧を登録する。

        Args:
            components: テスト用コンポーネント。

        """
        self.VBProject = SimpleNamespace(VBComponents=components, Protection=0)
        self.closed = False

    def Close(self, SaveChanges: bool) -> None:
        """ブックが閉じられたことを記録する.

        Args:
            SaveChanges: Excel COM互換の保存指定。

        """
        self.closed = True


class FakeExcel:
    """Excelアプリケーションの最小テストダブル。"""

    def __init__(self, workbook: Any) -> None:
        """テスト用ブックとExcel状態を初期化する。

        Args:
            workbook: Workbooks.Openが返すテスト用ブック。

        """
        self.Workbooks = SimpleNamespace(Open=lambda *args, **kwargs: workbook)
        self.quit = False

    def Quit(self) -> None:
        """Excelが終了されたことを記録する."""
        self.quit = True


class FakeComRuntime:
    """COM初期化と解除の呼び出しを記録する。"""

    def __init__(self) -> None:
        """初期化／解除の呼び出しフラグを初期化する。"""
        self.initialized = False
        self.uninitialized = False

    def CoInitialize(self) -> None:
        """COM初期化を記録する."""
        self.initialized = True

    def CoUninitialize(self) -> None:
        """COM解除を記録する."""
        self.uninitialized = True


class BlockedWorkbook:
    """VBAプロジェクトへのアクセス失敗を模擬するブック。"""

    @property
    def VBProject(self) -> None:
        """VBAプロジェクトへのアクセス時にエラーを発生させる."""
        raise RuntimeError("access denied")

    def Close(self, SaveChanges: bool) -> None:
        """テストブックのクローズ要求を受け取る."""
        return None


def test_build_export_plan_includes_only_exportable_components(
    tmp_path: Path,
) -> None:
    """標準、クラス、フォームだけを対応する拡張子で計画する."""
    components = [
        FakeComponent("Module1", 1),
        FakeComponent("Class1", 2),
        FakeComponent("UserForm1", 3),
        FakeComponent("ThisWorkbook", 100),
        FakeComponent("Sheet1", 100),
    ]

    plan = build_export_plan(components, tmp_path)

    assert [item.destination.name for item in plan] == [
        "Module1.bas",
        "Class1.cls",
        "UserForm1.frm",
    ]


def test_build_export_plan_refuses_existing_file(tmp_path: Path) -> None:
    """既存ファイルを上書きしない."""
    existing = tmp_path / "Module1.bas"
    existing.write_text("keep", encoding="utf-8")

    with pytest.raises(ExportError, match="既に存在"):
        build_export_plan([FakeComponent("Module1", 1)], tmp_path)

    assert existing.read_text(encoding="utf-8") == "keep"


def test_build_export_plan_refuses_existing_form_resource(
    tmp_path: Path,
) -> None:
    """フォームのリソースファイルを既存ファイルへ上書きしない."""
    resource = tmp_path / "UserForm1.frx"
    resource.write_text("keep", encoding="utf-8")

    with pytest.raises(ExportError, match="既に存在"):
        build_export_plan([FakeComponent("UserForm1", 3)], tmp_path)

    assert resource.read_text(encoding="utf-8") == "keep"


def test_export_workbook_exports_components_and_closes_excel(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """対象コンポーネントを出力し、ブックとExcelおよびCOMを終了する."""
    source = tmp_path / "source.xlsm"
    source.touch()
    output = tmp_path / "output"
    output.mkdir()

    workbook = FakeWorkbook([FakeComponent("Module1", 1), FakeComponent("Sheet1", 100)])
    excel = FakeExcel(workbook)
    com_runtime = FakeComRuntime()
    monkeypatch.setattr(
        exporter,
        "_load_com_runtime",
        lambda: (com_runtime, lambda _: excel),
    )

    result = export_workbook(source, output)

    assert result.files == (output / "Module1.bas",)
    assert result.files[0].is_file()
    assert workbook.closed
    assert excel.quit
    assert excel.AutomationSecurity == exporter.AUTOMATION_SECURITY_FORCE_DISABLE
    assert com_runtime.initialized
    assert com_runtime.uninitialized


def test_export_workbook_reports_vba_access_error_and_cleans_up(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """VBAプロジェクトにアクセスできない場合もExcelを終了する."""
    source = tmp_path / "source.xlsm"
    source.touch()
    output = tmp_path / "output"
    output.mkdir()
    workbook = BlockedWorkbook()
    excel = FakeExcel(workbook)
    com_runtime = FakeComRuntime()
    monkeypatch.setattr(
        exporter,
        "_load_com_runtime",
        lambda: (com_runtime, lambda _: excel),
    )

    with pytest.raises(ExportError, match="オブジェクト モデルへのアクセス") as caught:
        export_workbook(source, output)

    assert "access denied" in str(caught.value)
    assert "Excelをすべて終了" in str(caught.value)
    assert excel.quit
    assert com_runtime.uninitialized


def test_export_workbook_reports_locked_project(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """パスワード保護を信頼設定のエラーとは区別して表示する。"""
    source = tmp_path / "source.xlsm"
    source.touch()
    workbook = FakeWorkbook([FakeComponent("Module1", 1)])
    workbook.VBProject.Protection = 1
    excel = FakeExcel(workbook)
    com_runtime = FakeComRuntime()
    monkeypatch.setattr(
        exporter, "_load_com_runtime", lambda: (com_runtime, lambda _: excel)
    )

    with pytest.raises(ExportError, match="パスワードで保護"):
        export_workbook(source, tmp_path)

    assert workbook.closed
    assert excel.quit
    assert com_runtime.uninitialized


def test_export_plan_failure_is_not_reported_as_trust_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """出力計画のエラーをVBAアクセス拒否と誤報しない。"""
    source = tmp_path / "source.xlsm"
    source.touch()
    workbook = FakeWorkbook([FakeComponent("Module1", 1)])
    excel = FakeExcel(workbook)
    com_runtime = FakeComRuntime()
    monkeypatch.setattr(
        exporter, "_load_com_runtime", lambda: (com_runtime, lambda _: excel)
    )
    monkeypatch.setattr(
        exporter, "build_export_plan",
        lambda *_: (_ for _ in ()).throw(OSError("output unavailable")),
    )

    with pytest.raises(OSError, match="output unavailable"):
        export_workbook(source, tmp_path)

    assert workbook.closed
    assert excel.quit
    assert com_runtime.uninitialized
