"""Package a directory for delivery using only the Python standard library."""

import argparse
import fnmatch
import sys
import zipfile
from pathlib import Path


DEFAULT_IGNORED_DIRS = {".git", ".svn", "__pycache__"}
DEFAULT_IGNORED_FILES = {".DS_Store", "Thumbs.db"}


def package(source: Path, output: Path, required: list[str], excludes: list[str]) -> int:
    source = source.resolve()
    output = output.resolve()
    if not source.is_dir():
        raise ValueError(f"源目录不存在: {source}")
    if output == source or source in output.parents:
        raise ValueError("输出文件不能放在源目录内")

    files = []
    for path in source.rglob("*"):
        relative = path.relative_to(source)
        if any(part in DEFAULT_IGNORED_DIRS for part in relative.parts):
            continue
        if path.name in DEFAULT_IGNORED_FILES or path.suffix == ".pyc":
            continue
        if path.is_symlink():
            raise ValueError(f"源目录包含符号链接: {relative}")
        if path.is_file() and not any(
            fnmatch.fnmatchcase(relative.as_posix(), pattern)
            or fnmatch.fnmatchcase(path.name, pattern)
            for pattern in excludes
        ):
            files.append(relative)

    included = {path.as_posix() for path in files}
    for item in required:
        required_path = Path(item)
        if required_path.is_absolute() or ".." in required_path.parts:
            raise ValueError(f"必需路径必须位于源目录内: {item}")
        required_name = required_path.as_posix().rstrip("/")
        if required_name not in included and not any(
            name.startswith(required_name + "/") for name in included
        ):
            raise ValueError(f"缺少必需文件或目录: {item}")

    if not files:
        raise ValueError("没有可打包的文件")
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for relative in sorted(files, key=lambda item: item.as_posix()):
            archive.write(source / relative, relative.as_posix())
    return len(files)


def main() -> int:
    parser = argparse.ArgumentParser(description="将目录打包为交付 ZIP")
    parser.add_argument("source", type=Path, help="待交付目录")
    parser.add_argument("output", type=Path, help="输出 ZIP 路径")
    parser.add_argument("--required", action="append", default=[], metavar="PATH", help="必需的相对文件或目录，可重复")
    parser.add_argument("--exclude", action="append", default=[], metavar="GLOB", help="排除文件的 glob 模式，可重复")
    args = parser.parse_args()
    try:
        count = package(args.source, args.output, args.required, args.exclude)
    except (OSError, ValueError, zipfile.BadZipFile) as error:
        print(f"打包失败: {error}", file=sys.stderr)
        return 1
    print(f"已打包 {count} 个文件: {args.output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
