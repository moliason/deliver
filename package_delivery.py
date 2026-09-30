"""归档交付目录，保留清单、缺失项和 HTML 报告。"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from delivery_report_packager import generate_html_report, html_report_parse_log


def main():
    parser = argparse.ArgumentParser(description="归档 deliver.py 生成的交付目录")
    parser.add_argument("--source", type=Path, default=Path("build/workspace"))
    parser.add_argument("--output", type=Path, default=Path("build/package"))
    parser.add_argument("--log", type=Path, default=Path("build/deliver_output.txt"))
    parser.add_argument("--report", type=Path, default=Path("build/html_report/index.html"))
    args = parser.parse_args()
    source = args.source.resolve()
    output = args.output.resolve()
    error_message = ""
    try:
        if not source.is_dir():
            raise ValueError(f"交付目录不存在: {source}")
        if output == source or source in output.parents:
            raise ValueError("归档输出不能位于交付源目录内")
        if output.exists() and any(output.iterdir()):
            raise ValueError("归档输出目录非空，请指定新的 --output")
        password = os.getenv("ZIP_PASSWORD", "")
        seven_zip = shutil.which("7z") or shutil.which("7zz")
        if password and not seven_zip:
            raise ValueError("ZIP_PASSWORD 已设置，加密打包需要安装 7-Zip 并加入 PATH")
        output.mkdir(parents=True, exist_ok=True)
        package_count = 0
        markers = ("Deliver List To ", "Change To ", "Example ", "Firmware ", "XML Files ", "Common Files ")
        for item in sorted(source.iterdir()):
            if not any(marker in item.name for marker in markers):
                continue
            if item.is_file() and item.suffix.casefold() == ".xlsx":
                shutil.copy2(item, output / item.name)
                package_count += 1
                continue
            if not item.is_dir():
                continue
            with tempfile.TemporaryDirectory(prefix="delivery-package-") as temporary:
                staged = Path(temporary) / item.name
                shutil.copytree(item, staged, ignore=shutil.ignore_patterns(".git", ".gitlab"))
                for path in staged.rglob("*"):
                    if not path.is_file():
                        continue
                    relative = path.relative_to(staged)
                    remove_markdown = path.suffix.casefold() == ".md" and not {"FreeRTOS", "LWIP"}.intersection(relative.parts)
                    remove_hidden = path.name.startswith(".") and path.name not in {".project", ".cproject"}
                    if remove_markdown or remove_hidden:
                        path.unlink()
                    elif path.suffix.casefold() in {".xml", ".c", ".h"}:
                        content = path.read_bytes().replace(b"\r\n", b"\n").replace(b"\r", b"\n")
                        path.write_bytes(content.replace(b"\n", b"\r\n"))
                if "Common Files " in item.name:
                    shutil.copytree(staged, output / item.name)
                    continue
                zip_path = output / f"{item.name}.zip"
                if password:
                    result = subprocess.run(
                        [seven_zip, "a", "-tzip", f"-p{password}", "-mem=AES256", "-mcu=on", str(zip_path), item.name],
                        cwd=temporary, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    )
                    if result.returncode:
                        raise ValueError(f"7-Zip 打包失败，退出码 {result.returncode}: {item.name}")
                else:
                    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                        for path in sorted(staged.rglob("*")):
                            if path.is_file():
                                archive.write(path, path.relative_to(temporary).as_posix())
                package_count += 1
                print(f"ZIP_PACKAGE={zip_path}")
        log_text = args.log.read_text(encoding="utf-8-sig", errors="replace") if args.log.exists() else ""
        data = html_report_parse_log(log_text)
        missing_lines = [line.split("=", 1)[1] for line in log_text.splitlines() if line.startswith("MISSING_DELIVER_FILE=")]
        (output / "missing_file.txt").write_text("\n".join(missing_lines) or "NO_MISSING_DELIVER_FILE", encoding="utf-8")
        if not package_count:
            raise ValueError("未生成 ZIP 或 Excel：源目录中没有匹配的交付产物")
        return 1 if data["missing_items"] or data["error_lines"] else 0
    except (OSError, ValueError) as error:
        error_message = str(error)
        print(f"归档失败: {error}", file=sys.stderr)
        return 1
    finally:
        generate_html_report(args.log, output, args.report, error=error_message)


if __name__ == "__main__":
    raise SystemExit(main())
