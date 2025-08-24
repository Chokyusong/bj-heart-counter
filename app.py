# app.py
import io, re, csv
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="BJ별 하트 정리", layout="centered")
st.title("BJ별 하트 정리 자동화")
st.caption("CSV/XLSX 업로드 → 참여BJ별 시트 생성 → ID/닉네임 분리 → 일반/제휴(한 줄 표기) + 합계(E열)")

uploaded = st.file_uploader("CSV 또는 엑셀(.xlsx) 업로드", type=["csv", "xlsx"])
sheet_name = st.text_input("시트 이름 (엑셀일 때만, 비우면 첫 시트)", value="")

# ---------- utils ----------
def autosize_columns(wb):
    for ws in wb.worksheets:
        for col in ws.columns:
            w = 0
            letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    w = max(w, len(str(cell.value)))
            ws.column_dimensions[letter].width = min(w + 2, 50)

def sanitize_sheet(name: str) -> str:
    return re.sub(r'[\\/*?:\[\]]', "_", str(name))[:31] or "BJ"

def read_any_table(uploaded_file, sheet: str | int | None):
    """xlsx/csv 모두 DataFrame으로 읽기 (CSV 인코딩/구분자 자동 감지)."""
    name = (uploaded_file.name or "").lower()
    if name.endswith(".xlsx"):
        return pd.read_excel(uploaded_file, sheet_name=(sheet if str(sheet).strip() else 0))

    raw = uploaded_file.read()
    uploaded_file.seek(0)
    for enc in ["utf-8", "utf-8-sig", "cp949", "euc-kr"]:
        try:
            text = raw.decode(enc)
            try:
                dialect = csv.Sniffer().sniff(text[:4000], delimiters=[",", "\t", ";", "|"])
                sep = dialect.delimiter
            except Exception:
                sep = ","
            return pd.read_csv(io.StringIO(text), sep=sep)
        except Exception:
            continue
    raise ValueError("CSV 인코딩/구분자 해석 실패")

# ---------- core ----------
def build_output_excel(df: pd.DataFrame) -> bytes:
    # 1) 컬럼명 고정 (실제 파일의 열 이름 기준)
    df.columns = [str(c).strip() for c in df.columns]
    col_bj    = next((c for c in df.columns if c == "참여BJ"), None)
    col_heart = next((c for c in df.columns if c == "후원하트"), None)
    col_mix   = next((c for c in df.columns if c == "후원 아이디(닉네임)"), None)
    if not (col_bj and col_heart and col_mix):
        raise ValueError(f"필수 컬럼 누락: 참여BJ={col_bj}, 후원하트={col_heart}, 후원 아이디(닉네임)={col_mix}")

    # 2) 정규화
    df[col_bj] = df[col_bj].astype(str).str.strip()
    df[col_heart] = df[col_heart].astype(str).str.replace(",", "", regex=False)
    df[col_heart] = pd.to_numeric(df[col_heart], errors="coerce").fillna(0).astype(int)
    df[col_mix] = df[col_mix].astype(str).str.strip()

    # 3) "후원 아이디(닉네임)" → ID/닉네임 분리 + 잡스러운 문자 정리
    split = df[col_mix].str.extract(r'^\s*(?P<ID>[^()]+?)(?:\((?P<NICK>.*)\))?\s*$')
    df["ID"] = (split["ID"].fillna("")
                .str.replace("\u200b","",regex=False)   # zero-width space
                .str.replace("\ufeff","",regex=False)    # BOM
                .str.replace("＠","@",regex=False)        # 전각@ → @
                .str.strip())
    df["닉네임"] = split["NICK"].fillna("").str.strip()

    # 4) BJ×ID×닉네임 합산(같은 ID는 합치기). 원본 행 그대로 쓰려면 아래 groupby 제거.
    base = (
        df.groupby([col_bj, "ID", "닉네임"], as_index=False)[col_heart]
          .sum()
          .rename(columns={col_bj:"참여BJ", col_heart:"후원하트"})
    )

    # 5) 엑셀 작성
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        # 요약 시트
        summary = base.groupby("참여BJ", as_index=False)["후원하트"].sum().sort_values("후원하트", ascending=False)
        summary.to_excel(writer, sheet_name="요약", index=False)

        for bj in summary["참여BJ"]:
            sub = base[base["참여BJ"] == bj].copy()
            sub["is_aff"] = sub["ID"].str.contains("@")

            gen = sub[~sub["is_aff"]].sort_values("후원하트", ascending=False)[["ID","닉네임","후원하트"]].copy()
            aff = sub[ sub["is_aff"]].sort_values("후원하트", ascending=False)[["ID","닉네임","후원하트"]].copy()

            gsum = int(gen["후원하트"].sum()) if not gen.empty else 0
            asum = int(aff["후원하트"].sum()) if not aff.empty else 0
            total = gsum + asum

            sheet = sanitize_sheet(bj)

            # --- 1행: B=참여BJ, C=총합 (E가 아님 주의) ---
            row1 = pd.DataFrame([[ "", bj, total, "", "" ]],
                                columns=["ID","닉네임","후원하트","구분","합계"])
            row1.to_excel(writer, sheet_name=sheet, index=False, header=False, startrow=0)

            # --- 2행: 컬럼 헤더만 출력 ---
            header_only = pd.DataFrame(columns=["ID","닉네임","후원하트","구분","합계"])
            header_only.to_excel(writer, sheet_name=sheet, index=False, startrow=1)  # header=True 기본값

            row = 2  # 데이터는 3행부터

            # --- 일반 블록: 첫 행 D/E만 채움, 나머지 공란 ---
            if not gen.empty:
                gen_block = gen.copy()
                gen_block["구분"] = ""
                gen_block["합계"] = ""
                gen_block.iloc[0, gen_block.columns.get_loc("구분")] = "일반하트"
                gen_block.iloc[0, gen_block.columns.get_loc("합계")] = gsum
                gen_block.to_excel(writer, sheet_name=sheet, index=False, header=False, startrow=row)
                row += len(gen_block)

            # --- 제휴 블록: 첫 행 D/E만 채움, 중간 헤더/구분줄 없음 ---
            if not aff.empty:
                aff_block = aff.copy()
                aff_block["구분"] = ""
                aff_block["합계"] = ""
                aff_block.iloc[0, aff_block.columns.get_loc("구분")] = "제휴하트"
                aff_block.iloc[0, aff_block.columns.get_loc("합계")] = asum
                aff_block.to_excel(writer, sheet_name=sheet, index=False, header=False, startrow=row)

    # 자동 열 너비
    bio.seek(0)
    wb = load_workbook(bio)
    autosize_columns(wb)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()

# ---------- run ----------
if uploaded:
    try:
        df_in = read_any_table(uploaded, sheet_name if uploaded.name.lower().endswith(".xlsx") else None)
        result_bytes = build_output_excel(df_in)
        st.success("완료! 아래 버튼으로 결과 엑셀을 다운로드하세요.")
        st.download_button(
            label="결과 엑셀 다운로드",
            data=result_bytes,
            file_name="BJ별_정리_하트내역.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"오류: {e}")
else:
    st.info("CSV/XLSX 파일을 업로드하면 자동으로 처리합니다.")
