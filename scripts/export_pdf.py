"""
가통문 xlsx → 학생별 PDF 일괄 변환 (Windows + Excel 필요)

사용법:
  python scripts/export_pdf.py <연도> <월>
  python scripts/export_pdf.py 2026 4

출력:
  output/pdf/{학년_반}_{이름}.pdf (시트당 1개)
  output/pdf/가통문_{연도}_{월}.zip (전체 묶음)
"""

from __future__ import annotations

import os
import site
import sys
import zipfile
from pathlib import Path

# pywin32 DLL 경로를 PATH에 등록 (Streamlit 등 서브프로세스에서 못 찾는 문제 방지)
for _sp in site.getsitepackages():
    _pywin32_dir = Path(_sp) / "pywin32_system32"
    if _pywin32_dir.is_dir():
        os.add_dll_directory(str(_pywin32_dir))
        os.environ["PATH"] = str(_pywin32_dir) + os.pathsep + os.environ.get("PATH", "")
        break

import pythoncom
import win32com.client

XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0


def export_all_sheets_to_pdf(xlsx_path: Path, out_dir: Path,
                              progress_cb=None) -> list[Path]:
    """xlsx의 각 시트를 개별 PDF로 내보냄."""
    out_dir.mkdir(parents=True, exist_ok=True)
    pythoncom.CoInitialize()

    # 새 Excel 인스턴스 (사용자가 열어둔 Excel에 영향 없도록)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    pdf_paths: list[Path] = []
    try:
        wb = excel.Workbooks.Open(str(xlsx_path.resolve()), ReadOnly=True)
        n = wb.Sheets.Count
        for i in range(1, n + 1):
            ws = wb.Sheets(i)
            name = str(ws.Name).strip()
            # 파일명에 쓸 수 없는 문자 치환
            safe = "".join(c if c.isalnum() or c in "_-가-힣" else "_" for c in name)
            pdf_path = out_dir / f"{safe}.pdf"
            ws.ExportAsFixedFormat(
                Type=XL_TYPE_PDF,
                Filename=str(pdf_path.resolve()),
                Quality=XL_QUALITY_STANDARD,
                IncludeDocProperties=False,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
            pdf_paths.append(pdf_path)
            if progress_cb:
                progress_cb(i, n, name)
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

    return pdf_paths


def zip_pdfs(pdf_paths: list[Path], zip_path: Path) -> Path:
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in pdf_paths:
            zf.write(p, arcname=p.name)
    return zip_path


def main():
    if len(sys.argv) < 3:
        print("사용법: python export_pdf.py <연도> <월>")
        sys.exit(1)
    year = int(sys.argv[1])
    month = int(sys.argv[2])

    base = Path(__file__).parent.parent
    xlsx = base / "output" / f"가통문_{year}_{month:02d}.xlsx"
    if not xlsx.exists():
        print(f"X {xlsx.name} 없음")
        sys.exit(1)

    pdf_dir = base / "output" / "pdf"
    # 기존 PDF 정리
    for p in pdf_dir.glob("*.pdf"):
        p.unlink()

    def cb(i, n, name):
        print(f"  [{i}/{n}] {name}")

    print(f"PDF 변환 시작 ({xlsx.name})")
    pdfs = export_all_sheets_to_pdf(xlsx, pdf_dir, progress_cb=cb)
    zip_path = base / "output" / f"가통문_{year}_{month:02d}.zip"
    zip_pdfs(pdfs, zip_path)
    print(f"\n완료: {len(pdfs)}개 PDF → {pdf_dir}")
    print(f"      ZIP: {zip_path}")


if __name__ == "__main__":
    main()
