from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path


def _find_java_exec(name: str) -> str:
    path = shutil.which(name)
    if path:
        return path

    java_home = os.environ.get("JAVA_HOME")
    if java_home:
        candidate = Path(java_home) / "bin" / f"{name}.exe"
        if candidate.exists():
            return str(candidate)

    raise RuntimeError(f"Cannot find '{name}'. Ensure JDK is installed and JAVA_HOME/PATH is configured.")


def _classpath(lib_dir: Path) -> str:
    jars = sorted(lib_dir.glob("*.jar"))
    if not jars:
        raise RuntimeError(f"No JAR files found in {lib_dir}")
    return os.pathsep.join(str(j.resolve()) for j in jars)


def compile_java_sources(src_root: Path, build_dir: Path, lib_dir: Path) -> None:
    javac = _find_java_exec("javac")
    classpath = _classpath(lib_dir)

    java_files = sorted(src_root.rglob("*.java"))
    if not java_files:
        raise RuntimeError(f"No Java source files found under {src_root}")

    build_dir.mkdir(parents=True, exist_ok=True)

    cmd = [
        javac,
        "-encoding",
        "UTF-8",
        "-cp",
        classpath,
        "-d",
        str(build_dir),
        *[str(p) for p in java_files],
    ]
    subprocess.run(cmd, check=True)


def run_generator(
    build_dir: Path,
    lib_dir: Path,
    input_json: Path,
    output_hwp: Path,
    visual_dir: Path,
) -> None:
    java = _find_java_exec("java")
    classpath = os.pathsep.join([str(build_dir.resolve()), _classpath(lib_dir)])

    cmd = [
        java,
        "-cp",
        classpath,
        "com.mpbook.hwp.HwpBookGenerator",
        "--input",
        str(input_json.resolve()),
        "--output",
        str(output_hwp.resolve()),
        "--visual-dir",
        str(visual_dir.resolve()),
    ]
    subprocess.run(cmd, check=True)


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Generate HWP in compatibility-first mode (equation controls and visual image controls enabled)."
        )
    )
    parser.add_argument("--input", default="output/structured_questions.json", help="Input JSON path.")
    parser.add_argument("--output", default="output/math_problem_book_oss.hwp", help="Output HWP path.")
    parser.add_argument(
        "--visual-dir",
        default="output/derived_visuals",
        help="Directory containing cropped visual images to place in HWP.",
    )
    parser.add_argument(
        "--skip-compile",
        action="store_true",
        help="Skip javac compile step and run existing classes.",
    )
    args = parser.parse_args()

    root = Path(__file__).resolve().parent
    java_module_root = root / "hwp-java"
    src_root = java_module_root / "src" / "main" / "java"
    build_dir = java_module_root / "build" / "classes"
    lib_dir = java_module_root / "lib"

    input_json = Path(args.input)
    output_hwp = Path(args.output)
    visual_dir = Path(args.visual_dir)

    if not input_json.exists():
        raise FileNotFoundError(f"Input JSON not found: {input_json}")

    if not args.skip_compile:
        compile_java_sources(src_root=src_root, build_dir=build_dir, lib_dir=lib_dir)

    output_hwp.parent.mkdir(parents=True, exist_ok=True)
    run_generator(
        build_dir=build_dir,
        lib_dir=lib_dir,
        input_json=input_json,
        output_hwp=output_hwp,
        visual_dir=visual_dir,
    )

    print(f"Generated HWP (safe+equation mode): {output_hwp.resolve()}")


if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as exc:
        print(f"[ERROR] command failed: {' '.join(exc.cmd)}", file=sys.stderr)
        raise
