import pandas as pd
import xml.etree.ElementTree as ET
from pathlib import Path
from tempfile import TemporaryDirectory
import zipfile

COLUMNS_TO_COMPARE = ["계약명", "상대 계약자", "요청자", "검토담당자", "계약 시작일", "계약 종료일"]
LIST_FILE = "계약서리스트_양식_대주산업_모두싸인.xlsx"
SIGNED_FILE = "체결계약서조회_2025-11-18.xlsx"
OUTPUT_FILE = "비교결과.xlsx"


def sanitize_styles(src_path: Path, tmp_dir: Path) -> Path:
    """openpyxl이 읽을 수 있도록 styles.xml의 비어있는 fill을 보정한다."""
    sanitized_path = tmp_dir / src_path.name
    with zipfile.ZipFile(src_path, "r") as src, zipfile.ZipFile(
        sanitized_path, "w"
    ) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "xl/styles.xml":
                root = ET.fromstring(data)
                ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
                fills = root.find("m:fills", ns)
                if fills is not None:
                    for fill in fills.findall("m:fill", ns):
                        pattern = fill.find("m:patternFill", ns)
                        gradient = fill.find("m:gradientFill", ns)
                        if pattern is None and gradient is None:
                            pattern = ET.SubElement(
                                fill,
                                "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}patternFill",
                            )
                            pattern.set("patternType", "none")
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            dst.writestr(info, data)
    return sanitized_path


def normalize_text(value: str) -> str:
    if pd.isna(value):
        return ""
    return str(value).replace("\n", " ").strip()

def normalize_contract_name(value: str) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    return (
        text.replace("(모두사인)", "")
        .replace("(모두싸인)", "")
        .strip()
    )


def normalize_date(value) -> str:
    if pd.isna(value):
        return ""
    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        return normalize_text(value)
    return dt.strftime("%Y-%m-%d")


def normalize_people(value: str) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    parts = [part.strip() for part in text.replace(";", ",").split(",")]
    parts = [part for part in parts if part]
    return ", ".join(sorted(parts))


def load_signed_contracts(base_path: Path) -> pd.DataFrame:
    file_path = base_path / SIGNED_FILE
    with TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        sanitized = sanitize_styles(file_path, tmp_dir)
        df = pd.read_excel(sanitized, header=0)
    df = df.rename(
        columns={
            "관리번호": "관리번호",
            "상대 계약자": "상대 계약자",
            "요청자": "요청자",
            "검토담당자": "검토담당자",
            "계약 종료": "계약 종료일",
        }
    )
    keep_cols = ["관리번호"] + COLUMNS_TO_COMPARE
    df = df[keep_cols]
    df["계약명"] = df["계약명"].apply(normalize_contract_name)
    df = df.dropna(subset=["관리번호"]).drop_duplicates(subset=["관리번호"])
    return df


def load_contract_list(base_path: Path) -> pd.DataFrame:
    file_path = base_path / LIST_FILE
    df = pd.read_excel(file_path, header=1)
    df = df.rename(
        columns={
            "관리 번호": "관리번호",
            "담당자(요청자)": "요청자",
            "법무 검토 담당자(,로 구분)": "검토담당자",
            "계약 시작 일자": "계약 시작일",
            "계약 종료 일자": "계약 종료일",
        }
    )
    keep_cols = ["관리번호"] + COLUMNS_TO_COMPARE
    df = df[keep_cols]
    df = df.dropna(subset=["관리번호"]).drop_duplicates(subset=["관리번호"])
    return df


def compare_datasets(signed_df: pd.DataFrame, list_df: pd.DataFrame) -> pd.DataFrame:
    merged = signed_df.merge(
        list_df, on="관리번호", how="outer", suffixes=("_체결", "_리스트")
    )

    def row_status(row):
        if pd.isna(row["관리번호"]):
            return "UNKNOWN"
        in_signed = not row.isna().get("계약명_체결", False) or isinstance(
            row["계약명_체결"], str
        )
        in_list = not row.isna().get("계약명_리스트", False) or isinstance(
            row["계약명_리스트"], str
        )
        if pd.isna(row["계약명_체결"]):
            return "LIST_ONLY"
        if pd.isna(row["계약명_리스트"]):
            return "SIGNED_ONLY"
        return "BOTH"

    def compare_cols(row):
        diffs = []
        for col in COLUMNS_TO_COMPARE:
            left = row[f"{col}_체결"]
            right = row[f"{col}_리스트"]
            if col in ("요청자", "검토담당자"):
                equal = normalize_people(left) == normalize_people(right)
            elif col == "계약 시작일":
                equal = normalize_date(left) == normalize_date(right)
            elif col == "계약 종료일":
                left_val = normalize_date(left)
                right_val = normalize_date(right)
                if not left_val and not right_val:
                    continue
                if not left_val or not right_val:
                    continue
                equal = left_val == right_val
            else:
                equal = normalize_text(left) == normalize_text(right)
            if not equal:
                diffs.append(col)
        return diffs

    merged["비교결과"] = merged.apply(row_status, axis=1)
    merged["불일치항목"] = merged.apply(compare_cols, axis=1)
    merged["모든항목일치"] = merged["불일치항목"].apply(lambda x: len(x) == 0)
    return merged


def save_report(merged: pd.DataFrame, output_path: Path):
    summary = [
        ("총 계약 수 (체결)", merged["관리번호"].where(~merged["계약명_체결"].isna()).nunique()),
        ("총 계약 수 (계약서리스트)", merged["관리번호"].where(~merged["계약명_리스트"].isna()).nunique()),
        ("양쪽 모두 존재", (merged["비교결과"] == "BOTH").sum()),
        ("양쪽 모두 일치", merged["모든항목일치"].sum()),
        (
            "양쪽 모두 있으나 불일치",
            ((merged["비교결과"] == "BOTH") & (merged["모든항목일치"] == False)).sum(),
        ),
        ("체결파일에만 있음", (merged["비교결과"] == "SIGNED_ONLY").sum()),
        ("계약서리스트에만 있음", (merged["비교결과"] == "LIST_ONLY").sum()),
    ]
    summary_df = pd.DataFrame(summary, columns=["항목", "값"])

    matched = merged[(merged["비교결과"] == "BOTH") & merged["모든항목일치"]]
    mismatched = merged[(merged["비교결과"] == "BOTH") & (~merged["모든항목일치"])]
    signed_only = merged[merged["비교결과"] == "SIGNED_ONLY"]
    list_only = merged[merged["비교결과"] == "LIST_ONLY"]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        matched.to_excel(writer, sheet_name="PerfectMatch", index=False)
        mismatched.to_excel(writer, sheet_name="Mismatch", index=False)
        signed_only.to_excel(writer, sheet_name="OnlySigned", index=False)
        list_only.to_excel(writer, sheet_name="OnlyList", index=False)


def main():
    base_path = Path(__file__).parent
    signed_df = load_signed_contracts(base_path)
    list_df = load_contract_list(base_path)

    merged = compare_datasets(signed_df, list_df)
    output_path = base_path / OUTPUT_FILE
    save_report(merged, output_path)

    total = len(merged)
    matched = merged["모든항목일치"].sum()
    mismatched = ((merged["비교결과"] == "BOTH") & (~merged["모든항목일치"])).sum()
    signed_only = (merged["비교결과"] == "SIGNED_ONLY").sum()
    list_only = (merged["비교결과"] == "LIST_ONLY").sum()

    print("비교 완료")
    print(f"총 비교 건수: {total}")
    print(f"완전 일치: {matched}")
    print(f"불일치: {mismatched}")
    print(f"체결 파일에만 존재: {signed_only}")
    print(f"계약서 리스트에만 존재: {list_only}")
    print(f"결과 파일: {output_path}")


if __name__ == "__main__":
    main()

