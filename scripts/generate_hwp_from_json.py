from __future__ import annotations

import argparse
import sys
from pathlib import Path

from generate_docx_from_json import build_docx


class HwpConversionError(RuntimeError):
    pass


def _import_win32():
    try:
        import win32com.client  # type: ignore
    except Exception as exc:  # pragma: no cover - environment dependent
        raise HwpConversionError(
            "pywin32(win32com.client) import failed. "
            "Install dependency with: pip install pywin32"
        ) from exc
    return win32com.client


def is_hwp_com_available() -> bool:
    try:
        win32 = _import_win32()
        return bool(win32.gencache.GetClassForProgID("HWPFrame.HwpObject"))
    except Exception:
        return False


def convert_docx_to_hwp(
    input_docx_path: Path,
    output_hwp_path: Path,
    *,
    visible: bool = False,
) -> None:
    if not input_docx_path.exists():
        raise HwpConversionError(f"Input DOCX not found: {input_docx_path}")

    win32 = _import_win32()

    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    except Exception as exc:  # pragma: no cover - environment dependent
        raise HwpConversionError(
            "Cannot create HWP COM object (HWPFrame.HwpObject). "
            "Check that Hancom Office (Hangul) is installed."
        ) from exc

    try:
        hwp.XHwpWindows.Item(0).Visible = bool(visible)
    except Exception:
        pass

    # FilePath checker registration is optional and environment-dependent.
    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    except Exception:
        pass

    try:
        open_ok = hwp.Open(str(input_docx_path))
        if open_ok is False:
            raise HwpConversionError(f"HWP failed to open DOCX: {input_docx_path}")

        output_hwp_path.parent.mkdir(parents=True, exist_ok=True)
        save_ok = hwp.SaveAs(str(output_hwp_path), "HWP", "")
        if save_ok is False:
            raise HwpConversionError(f"HWP SaveAs failed: {output_hwp_path}")
    except HwpConversionError:
        raise
    except Exception as exc:  # pragma: no cover - environment dependent
        raise HwpConversionError(
            "DOCX->HWP conversion failed during HWP automation. "
            "Try opening DOCX in Hangul manually and verify compatibility."
        ) from exc
    finally:
        try:
            hwp.Quit()
        except Exception:
            pass


def build_hwp_from_json(
    input_json_path: Path,
    output_hwp_path: Path,
    *,
    intermediate_docx_path: Path | None = None,
    keep_intermediate_docx: bool = True,
    visible: bool = False,
) -> Path:
    input_json_path = input_json_path.resolve()
    output_hwp_path = output_hwp_path.resolve()

    if intermediate_docx_path is None:
        intermediate_docx_path = output_hwp_path.with_suffix(".docx")
    intermediate_docx_path = intermediate_docx_path.resolve()

    build_docx(input_json_path, intermediate_docx_path)
    convert_docx_to_hwp(intermediate_docx_path, output_hwp_path, visible=visible)

    if not keep_intermediate_docx:
        try:
            intermediate_docx_path.unlink(missing_ok=True)
        except Exception:
            pass

    return output_hwp_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate HWP from structured JSON (via intermediate DOCX + Hancom COM automation)."
    )
    parser.add_argument(
        "--input",
        default="output/structured_questions.json",
        help="Path to input JSON file.",
    )
    parser.add_argument(
        "--output",
        default="output/math_problem_book.hwp",
        help="Path to output HWP file.",
    )
    parser.add_argument(
        "--docx-intermediate",
        default=None,
        help="Intermediate DOCX path. Default is output path with .docx suffix.",
    )
    parser.add_argument(
        "--no-keep-docx",
        action="store_true",
        help="Delete intermediate DOCX after successful HWP conversion.",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Show Hangul window during automation.",
    )
    parser.add_argument(
        "--check-only",
        action="store_true",
        help="Only check whether HWP COM automation is available.",
    )
    args = parser.parse_args()

    if args.check_only:
        ok = is_hwp_com_available()
        print("HWP COM availability:", "AVAILABLE" if ok else "NOT_AVAILABLE")
        return

    input_path = Path(args.input)
    output_path = Path(args.output)
    intermediate_path = Path(args.docx_intermediate) if args.docx_intermediate else None

    try:
        result = build_hwp_from_json(
            input_json_path=input_path,
            output_hwp_path=output_path,
            intermediate_docx_path=intermediate_path,
            keep_intermediate_docx=not args.no_keep_docx,
            visible=args.visible,
        )
        print(f"Generated HWP: {result}")
    except HwpConversionError as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        print(
            "Prerequisites: Hancom Office (Hangul) installed and COM ProgID "
            "'HWPFrame.HwpObject' available.",
            file=sys.stderr,
        )
        raise SystemExit(1)


if __name__ == "__main__":
    main()
