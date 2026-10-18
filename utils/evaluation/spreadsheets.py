"""Read computed spreadsheet values using a real, isolated calculation engine."""
from pathlib import Path
import shutil
import re
from zipfile import ZipFile
import subprocess
import tempfile

import openpyxl


class UnsupportedExternalWorkbook(ValueError):
    """Cannot safely recalculate this artifact in the isolated grader."""


def _require_internal_calculation(path, workbook):
    # OOXML external-link metadata can expose grader-only files when Calc
    # refreshes a workbook. These task outputs require self-contained tables.
    with ZipFile(path) as archive:
        parts = [name.lower() for name in archive.namelist()]
    if any(name.startswith(("xl/externallinks/", "xl/querytables/"))
           or name == "xl/connections.xml" or name.endswith("vbaproject.bin")
           for name in parts):
        raise UnsupportedExternalWorkbook("External workbook links, data connections and VBA are unsupported")

    formulas = [cell.value for sheet in workbook for row in sheet for cell in row
                if cell.data_type == "f"]
    formulas.extend(name.attr_text for name in workbook.defined_names.values() if name.attr_text)
    for formula in formulas:
        from openpyxl.formula import Tokenizer
        tokens = Tokenizer(formula if formula.startswith("=") else "=" + formula).items
        for token in tokens:
            if token.subtype == "RANGE" and re.search(r"\[[^]]+\].*!", token.value):
                raise UnsupportedExternalWorkbook("External workbook reference is unsupported")
            if token.type == "FUNC" and token.subtype == "OPEN":
                function = token.value.rstrip("(").upper().split(".")[-1]
                if function in {"WEBSERVICE", "DDE", "RTD"}:
                    raise UnsupportedExternalWorkbook("External data functions are unsupported")


def _write_calculation_profile(profile):
    # These are documented schema properties in LibreOffice's main.xcd:
    # org.openoffice.Office.Common/Security/Scripting. Disable all script
    # runtimes and active OLE/DDE content, independently of default settings.
    profile.mkdir()
    (profile / "registrymodifications.xcu").write_text(
        '<oor:items xmlns:oor="http://openoffice.org/2001/registry">'
        '<item oor:path="/org.openoffice.Office.Common/Security/Scripting">'
        '<prop oor:name="DisableMacrosExecution" oor:op="fuse"><value>true</value></prop>'
        '<prop oor:name="DisableActiveContent" oor:op="fuse"><value>true</value></prop>'
        '<prop oor:name="BlockUntrustedRefererLinks" oor:op="fuse"><value>true</value></prop>'
        '</item></oor:items>', encoding="utf-8")


def load_calculated_workbook(path):
    """Recalculate formula workbooks without mutating the agent's artifact.

    A fresh LibreOffice profile prevents cross-task state and lock contention.
    The evaluation image supplies LibreOffice Calc; the original workbook is
    never overwritten and conversion has a bounded runtime.
    """
    path = Path(path)
    workbook = openpyxl.load_workbook(path, data_only=False)
    _require_internal_calculation(path, workbook)
    has_formulas = any(cell.data_type == "f" for sheet in workbook
                       for row in sheet for cell in row)
    workbook.close()
    if not has_formulas:
        return openpyxl.load_workbook(path, data_only=True)

    engine = shutil.which("soffice") or shutil.which("libreoffice")
    if engine is None:
        mac_engine = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        engine = str(mac_engine) if mac_engine.is_file() else None
    if engine is None:
        raise RuntimeError("Formula workbook requires LibreOffice Calc (soffice)")

    with tempfile.TemporaryDirectory(prefix="cowork-calc-") as temporary:
        temporary = Path(temporary)
        source = temporary / "input"
        output = temporary / "output"
        source.mkdir()
        output.mkdir()
        shutil.copyfile(path, source / "workbook.xlsx")
        profile = temporary / "profile"
        _write_calculation_profile(profile)
        result = subprocess.run(
            [engine, "-env:UserInstallation=" + profile.as_uri(), "--headless",
             "--convert-to", "xlsx", "--outdir", str(output), str(source / "workbook.xlsx")],
            capture_output=True, text=True, timeout=60,
        )
        calculated = output / "workbook.xlsx"
        if result.returncode or not calculated.is_file():
            raise RuntimeError("LibreOffice could not recalculate workbook: " +
                               (result.stderr or result.stdout).strip()[:500])
        # The non-read-only loader materializes all values before temp cleanup.
        return openpyxl.load_workbook(calculated, data_only=True)
