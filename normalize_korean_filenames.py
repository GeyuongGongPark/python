#!/usr/bin/env python3
"""도깨비 한글 자모 문제 해결용 파일명 정규화 스크립트."""

from __future__ import annotations

import argparse
import os
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Tuple


@dataclass
class RenameResult:
    """작업 결과를 집계하기 위한 자료구조."""

    renamed: int = 0
    skipped: int = 0
    conflicts: int = 0

    def log_progress(self, renamed: bool, conflict: bool) -> None:
        if conflict:
            self.conflicts += 1
        elif renamed:
            self.renamed += 1
        else:
            self.skipped += 1


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Mac에서 생성된 분리 자모 파일명을 Windows에서도 정상적으로 보이도록 정규화합니다.",
    )
    parser.add_argument(
        "target",
        nargs="?",
        default=".",
        help="정규화할 루트 폴더 (기본값: 현재 디렉터리)",
    )
    parser.add_argument(
        "--form",
        default="NFC",
        choices=("NFC", "NFD", "NFKC", "NFKD"),
        help="적용할 유니코드 정규화 방식 (기본값: NFC)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="실제 이름을 바꾸지 않고 변경 예정 목록만 출력합니다.",
    )
    parser.add_argument(
        "--include-dirs",
        action="store_true",
        help="폴더 이름도 함께 정규화합니다.",
    )
    return parser.parse_args()


def collect_paths(root: Path, include_dirs: bool) -> Iterable[Path]:
    """파일 및 (옵션) 디렉터리 경로를 후위 순회하면서 만들어낸다."""

    for current, dirs, files in os.walk(root, topdown=False):
        current_path = Path(current)
        for filename in files:
            yield current_path / filename
        if include_dirs:
            for dirname in dirs:
                yield current_path / dirname


def build_unique_name(path: Path, normalized: str) -> Tuple[str, bool]:
    """이름 충돌 시 뒤에 숫자를 붙여 고유한 이름을 만들어 준다."""

    if normalized == path.name:
        return normalized, False

    parent = path.parent
    candidate = normalized
    counter = 1

    while True:
        target = parent / candidate
        try:
            same_file = target.exists() and target.samefile(path)
        except FileNotFoundError:
            same_file = False

        if not target.exists() or same_file:
            return candidate, False

        stem, suffix = os.path.splitext(normalized)
        candidate = f"{stem}_{counter}{suffix}"
        counter += 1

        # 과도하게 많은 충돌 방지를 위한 안전장치
        if counter > 9999:
            return normalized, True


def normalize_path(path: Path, form: str, dry_run: bool) -> Tuple[bool, bool]:
    """단일 경로에 대해 이름을 정규화한다."""

    normalized = unicodedata.normalize(form, path.name)
    if normalized == path.name:
        return False, False

    unique_name, conflict = build_unique_name(path, normalized)
    if conflict:
        return False, True

    target = path.with_name(unique_name)
    if dry_run:
        print(f"[DRY-RUN] {path} -> {target}")
        return True, False

    path.rename(target)
    print(f"RENAMED: {path} -> {target}")
    return True, False


def main() -> int:
    args = parse_args()
    root = Path(args.target).expanduser().resolve()

    if not root.exists():
        print(f"지정한 경로가 없습니다: {root}", file=sys.stderr)
        return 1

    result = RenameResult()

    for path in collect_paths(root, args.include_dirs):
        renamed, conflict = normalize_path(path, args.form, args.dry_run)
        result.log_progress(renamed, conflict)

    print(
        f"완료 - 변경: {result.renamed}, 그대로: {result.skipped}, 충돌: {result.conflicts}"
    )

    if result.conflicts:
        print(
            "충돌이 발생한 파일은 직접 확인 후 다시 실행하세요.",
            file=sys.stderr,
        )
        return 2

    return 0


if __name__ == "__main__":
    sys.exit(main())

