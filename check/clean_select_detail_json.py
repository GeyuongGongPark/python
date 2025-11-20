#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
selectDetail.json 파일의 null, true, false 등 보기 어려운 값들을 정리하는 스크립트
"""

import json
import os
from pathlib import Path
from typing import Any, Dict, List, Optional
import re


def clean_value(value: Any, remove_nulls: bool = True, convert_booleans: bool = True) -> Any:
    """
    값들을 정리하는 함수
    
    Args:
        value: 정리할 값
        remove_nulls: null 값 제거 여부
        convert_booleans: boolean 값을 한글로 변환할지 여부
    
    Returns:
        정리된 값
    """
    if value is None:
        if remove_nulls:
            return None  # 제거 대상
        else:
            return "N/A"
    
    if isinstance(value, bool):
        if convert_booleans:
            return "예" if value else "아니오"
        return value
    
    if isinstance(value, str):
        # 빈 문자열 처리
        if value == "":
            return None if remove_nulls else ""
        return value
    
    if isinstance(value, (int, float)):
        return value
    
    if isinstance(value, list):
        cleaned_list = []
        for item in value:
            cleaned_item = clean_value(item, remove_nulls, convert_booleans)
            if cleaned_item is not None or not remove_nulls:
                cleaned_list.append(cleaned_item)
        return cleaned_list if cleaned_list or not remove_nulls else None
    
    if isinstance(value, dict):
        cleaned_dict = {}
        for k, v in value.items():
            cleaned_v = clean_value(v, remove_nulls, convert_booleans)
            if cleaned_v is not None or not remove_nulls:
                cleaned_dict[k] = cleaned_v
        return cleaned_dict if cleaned_dict or not remove_nulls else None
    
    return value


def clean_json_data(data: Dict[str, Any], remove_nulls: bool = True, convert_booleans: bool = True) -> Dict[str, Any]:
    """
    JSON 데이터를 정리하는 함수
    
    Args:
        data: 원본 JSON 데이터
        remove_nulls: null 값 제거 여부
        convert_booleans: boolean 값을 한글로 변환할지 여부
    
    Returns:
        정리된 JSON 데이터
    """
    cleaned_data = {}
    
    for key, value in data.items():
        cleaned_value = clean_value(value, remove_nulls, convert_booleans)
        if cleaned_value is not None or not remove_nulls:
            cleaned_data[key] = cleaned_value
    
    return cleaned_data


def process_json_file(input_path: str, output_path: str = None, 
                      remove_nulls: bool = True, convert_booleans: bool = True,
                      indent: int = 2, ensure_ascii: bool = False):
    """
    JSON 파일을 처리하는 함수
    
    Args:
        input_path: 입력 JSON 파일 경로
        output_path: 출력 JSON 파일 경로 (None이면 자동 생성)
        remove_nulls: null 값 제거 여부
        convert_booleans: boolean 값을 한글로 변환할지 여부
        indent: JSON 들여쓰기
        ensure_ascii: ASCII만 사용할지 여부
    """
    input_path = Path(input_path)
    
    if not input_path.exists():
        print(f"❌ 파일을 찾을 수 없습니다: {input_path}")
        return
    
    # 출력 경로 설정
    if output_path is None:
        # 원본 파일명에 _cleaned 추가
        output_path = input_path.parent / f"{input_path.stem}_cleaned{input_path.suffix}"
    else:
        output_path = Path(output_path)
    
    print(f"📖 파일 읽는 중: {input_path}")
    
    try:
        with open(input_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        print(f"❌ JSON 파싱 오류: {e}")
        return
    except Exception as e:
        print(f"❌ 파일 읽기 오류: {e}")
        return
    
    print(f"🧹 데이터 정리 중...")
    cleaned_data = clean_json_data(data, remove_nulls, convert_booleans)
    
    print(f"💾 파일 저장 중: {output_path}")
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(cleaned_data, f, indent=indent, ensure_ascii=ensure_ascii, sort_keys=False)
        print(f"✅ 완료: {output_path}")
    except Exception as e:
        print(f"❌ 파일 저장 오류: {e}")


def extract_manage_no_from_path(path: Path) -> Optional[str]:
    """
    경로에서 관리번호(CYYYYMMDD-####)를 추출.
    """
    pattern = re.compile(r"^C\d{8}-\d{4}$")
    for part in path.parts[::-1]:
        if pattern.match(part):
            return part
    return None


def extract_manage_no_from_json(data: Dict[str, Any]) -> Optional[str]:
    """
    JSON 내용에서 관리번호(ManageNo)를 추출.
    """
    manage_no = data.get("ManageNo")
    if isinstance(manage_no, str) and manage_no:
        return manage_no
    return None


def extract_company_from_path(path: Path) -> Optional[str]:
    """
    경로에서 기업명(raw_data/<기업명>/...)을 추출.
    """
    parts = list(path.parts)
    for i, part in enumerate(parts):
        if part == "raw_data" and i + 1 < len(parts):
            return parts[i + 1]
    return None


def process_directory(directory_path: str, pattern: str = "selectDetail.json",
                     remove_nulls: bool = True, convert_booleans: bool = True,
                     indent: int = 2, ensure_ascii: bool = False,
                     collect_dir: Optional[str] = None):
    """
    디렉토리 내의 모든 selectDetail.json 파일을 처리하는 함수
    
    Args:
        directory_path: 디렉토리 경로
        pattern: 검색할 파일 패턴
        remove_nulls: null 값 제거 여부
        convert_booleans: boolean 값을 한글로 변환할지 여부
        indent: JSON 들여쓰기
        ensure_ascii: ASCII만 사용할지 여부
    """
    directory_path = Path(directory_path)
    
    if not directory_path.exists():
        print(f"❌ 디렉토리를 찾을 수 없습니다: {directory_path}")
        return
    
    # 파일 찾기
    json_files = list(directory_path.rglob(pattern))
    
    if not json_files:
        print(f"⚠️  '{pattern}' 파일을 찾을 수 없습니다.")
        return
    
    print(f"📁 {len(json_files)}개의 파일을 찾았습니다.")

    # 수집 루트 폴더 준비
    if collect_dir is None:
        # 기본 수집 경로: /Users/ggpark/Desktop/python/check/clean
        base_root = Path(__file__).resolve().parent
        collect_root = (base_root / "clean").resolve()
    else:
        collect_root = Path(collect_dir).resolve()
    collect_root.mkdir(parents=True, exist_ok=True)

    for json_file in json_files:
        print(f"\n{'='*60}")
        # 파일 읽어 관리번호 확인
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                original_data = json.load(f)
        except Exception as e:
            print(f"❌ 원본 읽기 실패, 건너뜀: {json_file} ({e})")
            continue

        manage_no = extract_manage_no_from_json(original_data)
        if not manage_no:
            manage_no = extract_manage_no_from_path(Path(json_file)) or "UNKNOWN"

        # 파일 이름: %관리번호%_selectetail_clean.json (요청 명칭 유지)
        out_name = f"{manage_no}_selectetail_clean.json"
        # 기업명 하위 폴더 결정
        company = extract_company_from_path(Path(json_file)) or "UNKNOWN"
        collect_path = (collect_root / company)
        collect_path.mkdir(parents=True, exist_ok=True)
        out_path = collect_path / out_name

        # 중복 방지: 이미 존재하면 접미사 부여
        if out_path.exists():
            counter = 2
            while True:
                candidate = collect_path / f"{manage_no}_selectetail_clean_{counter}.json"
                if not candidate.exists():
                    out_path = candidate
                    break
                counter += 1

        # 정리 및 저장
        process_json_file(str(json_file), str(out_path), remove_nulls, convert_booleans, indent, ensure_ascii)
    
    print(f"\n{'='*60}")
    print(f"✅ 모든 파일 처리 완료!")


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(
        description="selectDetail.json 파일의 null, true, false 등 보기 어려운 값들을 정리합니다."
    )
    parser.add_argument(
        "input",
        nargs='?',
        default=None,
        help="입력 JSON 파일 경로 또는 디렉토리 경로 (미지정 시 check/raw_data 전체 처리)"
    )
    parser.add_argument(
        "-o", "--output",
        help="출력 JSON 파일 경로 (파일 처리 시에만 사용)"
    )
    parser.add_argument(
        "--keep-nulls",
        action="store_true",
        help="null 값을 유지하고 'N/A'로 변환"
    )
    parser.add_argument(
        "--keep-booleans",
        action="store_true",
        help="boolean 값을 한글로 변환하지 않고 유지"
    )
    parser.add_argument(
        "--indent",
        type=int,
        default=2,
        help="JSON 들여쓰기 (기본값: 2)"
    )
    parser.add_argument(
        "--pattern",
        default="selectDetail.json",
        help="디렉토리 처리 시 검색할 파일 패턴 (기본값: selectDetail.json)"
    )
    
    args = parser.parse_args()
    
    # 입력 경로 결정: 인자가 없으면 스크립트 기준 'check/raw_data'를 기본 사용
    if args.input is None:
        default_dir = (Path(__file__).resolve().parent / "raw_data").resolve()
        if default_dir.exists():
            input_path = default_dir
            print(f"🔎 입력 미지정: 기본 경로 사용 -> {input_path}")
        else:
            print("❌ 기본 경로 'check/raw_data'를 찾지 못했습니다. 입력 경로를 지정해 주세요.")
            raise SystemExit(1)
    else:
        input_path = Path(args.input)
    
    if input_path.is_file():
        # 단일 파일 처리
        process_json_file(
            str(input_path),
            args.output,
            remove_nulls=not args.keep_nulls,
            convert_booleans=not args.keep_booleans,
            indent=args.indent
        )
    elif input_path.is_dir():
        # 디렉토리 처리
        process_directory(
            str(input_path),
            args.pattern,
            remove_nulls=not args.keep_nulls,
            convert_booleans=not args.keep_booleans,
            indent=args.indent
        )
    else:
        print(f"❌ 파일 또는 디렉토리를 찾을 수 없습니다: {input_path}")

