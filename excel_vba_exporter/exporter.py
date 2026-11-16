"""ExcelブックからVBAコンポーネントをエクスポートする。"""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

SUPPORTED_EXTENSIONS = {".xls", ".xlsm", ".xlsb", ".xlam"}
COMPONENT_EXTENSIONS = {1: ".bas", 2: ".cls", 3: ".frm"}
AUTOMATION_SECURITY_FORCE_DISABLE = 3
INVALID_FILENAME_CHARACTERS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
RESERVED_FILENAME_STEMS = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{number}" for number in range(1, 10)),
    *(f"LPT{number}" for number in range(1, 10)),
}


class ExportError(Exception):
    """利用者に表示するエクスポートエラー。"""


@dataclass(frozen=True)
class ExportPlanItem:
    """エクスポート対象のコンポーネントと出力先。"""

    component: Any
    destination: Path


@dataclass(frozen=True)
class ExportResult:
    """エクスポートされたファイルの一覧。"""

    files: tuple[Path, ...]


def _safe_filename(name: str) -> str:
    """コンポーネント名をWindowsで使用できるファイル名にする。

    Args:
        name: VBAコンポーネント名。

    Returns:
        ファイル名として安全なコンポーネント名。

    Raises:
        ExportError: ファイル名に変換できない場合。

    """
    safe_name = INVALID_FILENAME_CHARACTERS.sub("_", name).rstrip(" .")
    if not safe_name or safe_name in {".", ".."}:
        raise ExportError(f"コンポーネント名をファイル名に変換できません: {name!r}")

    if safe_name.split(".", maxsplit=1)[0].upper() in RESERVED_FILENAME_STEMS:
        safe_name = f"_{safe_name}"
    return safe_name


def build_export_plan(
    components: Iterable[Any], output_directory: Path
) -> list[ExportPlanItem]:
    """対応するVBAコンポーネントの出力先を検証して作成する。

    標準モジュール、クラスモジュール、フォームだけを対象にし、既存ファイルや
    出力名の衝突がある場合はエクスポート開始前にエラーにする。

    Args:
        components: Excel VBAプロジェクトのコンポーネント一覧。
        output_directory: エクスポート先フォルダー。

    Returns:
        検証済みのエクスポート対象一覧。

    Raises:
        ExportError: 出力ファイルが既に存在する場合や名前が衝突する場合。

    """
    plan: list[ExportPlanItem] = []
    planned_names: set[str] = set()

    for component in components:
        extension = COMPONENT_EXTENSIONS.get(component.Type)
        if extension is None:
            continue

        filename = f"{_safe_filename(str(component.Name))}{extension}"
        destination = output_directory / filename
        comparison_name = filename.casefold()
        names_to_check = {comparison_name}
        if extension == ".frm":
            names_to_check.add(f"{Path(filename).stem}.frx".casefold())

        if planned_names.intersection(names_to_check):
            raise ExportError(f"出力ファイル名が重複しています: {filename}")
        for output_name in names_to_check:
            existing_path = output_directory / output_name
            if existing_path.exists():
                raise ExportError(
                    "同名のファイルが既に存在します。上書きしないため処理を"
                    f"中止しました: "
                    f"{existing_path}"
                )

        planned_names.update(names_to_check)
        plan.append(ExportPlanItem(component, destination))

    return plan


def _load_com_runtime() -> tuple[Any, Any]:
    """Excel COM操作に必要なWindowsモジュールを読み込む。

    Returns:
        pythoncomモジュールとExcel起動関数。

    Raises:
        ExportError: Windowsまたはpywin32を利用できない場合。

    """
    try:
        import pythoncom
        from win32com.client import DispatchEx
    except ImportError as exc:
        raise ExportError(
            "Excelとの連携に必要なpywin32を読み込めません。"
            "Windows上で依存関係をインストールしてください。"
        ) from exc

    return pythoncom, DispatchEx


def export_workbook(
    workbook_path: str | Path, output_directory: str | Path
) -> ExportResult:
    """Excelブックを読み取り専用で開き、対象VBAコンポーネントを出力する。

    Args:
        workbook_path: エクスポート元のExcelファイル。
        output_directory: 出力先として選択された既存フォルダー。

    Returns:
        作成されたファイルのパス一覧。

    Raises:
        ExportError: 入力、出力先、Excel、またはVBAプロジェクトに問題がある場合。

    """
    source = Path(workbook_path).expanduser()
    destination = Path(output_directory).expanduser()
    if source.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ExportError(
            "対応しているファイル形式は .xls、.xlsm、.xlsb、.xlam です。"
        )
    if not source.is_file():
        raise ExportError(f"対象のExcelファイルが見つかりません: {source}")
    if not destination.is_dir():
        raise ExportError(f"出力先フォルダーが見つかりません: {destination}")

    pythoncom, dispatch_excel = _load_com_runtime()
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = dispatch_excel("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = AUTOMATION_SECURITY_FORCE_DISABLE
        try:
            workbook = excel.Workbooks.Open(
                str(source.resolve()),
                UpdateLinks=0,
                ReadOnly=True,
                AddToMru=False,
            )
        except Exception as exc:
            raise ExportError(
                f"Excelで対象ファイルを開けませんでした: {source}"
            ) from exc

        try:
            project = workbook.VBProject
        except Exception as exc:
            raise ExportError(
                "ExcelがVBAプロジェクトへのアクセスを拒否しました。\n\n"
                "Excelの［ファイル］→［オプション］→［トラスト センター］→"
                "［トラスト センターの設定］→［マクロの設定］で、"
                "［VBA プロジェクト オブジェクト モデルへのアクセスを信頼する］"
                "を有効にしてください。\n"
                "作業中のブックを保存してExcelをすべて終了し、"
                "再度エクスポートしてください。\n"
                "設定が禁止されている場合は管理者に確認してください。"
                "ツールはセキュリティ設定を自動変更しません。\n\n"
                f"Excelのエラー詳細:\n{exc}"
            ) from exc

        if project.Protection == 1:
            raise ExportError(
                "VBAプロジェクトがパスワードで保護されています。"
                "ExcelのVBAエディターで保護を解除して保存し、"
                "対象ブックを閉じてから再実行してください。"
            )
        plan = build_export_plan(project.VBComponents, destination)

        exported_files: list[Path] = []
        for item in plan:
            try:
                item.component.Export(str(item.destination))
            except Exception as exc:
                raise ExportError(
                    f"コンポーネントをエクスポートできませんでした: "
                    f"{item.destination.name}"
                ) from exc
            if not item.destination.is_file():
                raise ExportError(
                    f"エクスポート後のファイルが見つかりません: "
                    f"{item.destination}"
                )
            exported_files.append(item.destination)

        return ExportResult(tuple(exported_files))
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        finally:
            workbook = None
            try:
                if excel is not None:
                    excel.Quit()
            finally:
                excel = None
                pythoncom.CoUninitialize()
