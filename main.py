import os
import re
import base64
from io import BytesIO
from copy import copy
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from urllib.parse import urlparse
from collections import defaultdict

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font


# =========================================================
# 업로드 결과 양식 (업로드샘플(수정용).xlsx 내장)
# =========================================================
# 외부 템플릿 파일이 없어도 EXE 하나만으로 동일한 양식을 생성하도록
# 사용자가 제공한 업로드 샘플 워크북 자체를 프로그램 안에 내장합니다.
UPLOAD_TEMPLATE_B64 = """
UEsDBBQABgAIAAAAIQBBN4LPbgEAAAQFAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACsVMluwjAQvVfqP0S+Vomhh6qqCBy6HFsk6AeYeJJYJLblGSj8fSdmUVWxCMElUWzPWybzPBit2iZZQkDj
bC76WU8kYAunja1y8T39SJ9FgqSsVo2zkIs1oBgN7+8G07UHTLjaYi5qIv8iJRY1tAoz58HyTulCq4g/QyW9KuaqAvnY6z3JwlkCSyl1GGI4eINSLRpK3le8
vFEyM1Ykr5tzHVUulPeNKRSxULm0+h9J6srSFKBdsWgZOkMfQGmsAahtMh8MM4YJELExFPIgZ4AGLyPdusq4MgrD2nh8YOtHGLqd4662dV/8O4LRkIxVoE/V
sne5auSPC/OZc/PsNMilrYktylpl7E73Cf54GGV89W8spPMXgc/oIJ4xkPF5vYQIc4YQad0A3rrtEfQcc60C6Anx9FY3F/AX+5QOjtQ4OI+c2gCXd2EXka46
9QwEgQzsQ3Jo2PaMHPmr2w7dnaJBH+CW8Q4b/gIAAP//AwBQSwMEFAAGAAgAAAAhALVVMCP0AAAATAIAAAsACAJfcmVscy8ucmVscyCiBAIooAACAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACskk1PwzAMhu9I/IfI99XdkBBCS3dBSLshVH6ASdwPtY2j
JBvdvyccEFQagwNHf71+/Mrb3TyN6sgh9uI0rIsSFDsjtnethpf6cXUHKiZylkZxrOHEEXbV9dX2mUdKeSh2vY8qq7iooUvJ3yNG0/FEsRDPLlcaCROlHIYW
PZmBWsZNWd5i+K4B1UJT7a2GsLc3oOqTz5t/15am6Q0/iDlM7NKZFchzYmfZrnzIbCH1+RpVU2g5abBinnI6InlfZGzA80SbvxP9fC1OnMhSIjQS+DLPR8cl
oPV/WrQ08cudecQ3CcOryPDJgosfqN4BAAD//wMAUEsDBBQABgAIAAAAIQCBPpSX8wAAALoCAAAaAAgBeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHMgogQB
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
"""


# =========================================================
# 기본 유틸
# =========================================================

def clean_id(value):
    """엑셀 숫자형 ID의 .0 제거"""
    if pd.isna(value):
        return ""

    text = str(value).strip()

    if text.endswith(".0"):
        text = text[:-2]

    return text


def parse_date(value):
    """
    '2026-08-26 (전자심의)' 같은 값에서도 날짜만 추출.
    """
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value

    text = str(value).strip()

    # YYYY-MM-DD / YYYY.MM.DD / YYYY/MM/DD
    match = re.search(
        r"(\d{4}[-./]\d{1,2}[-./]\d{1,2})",
        text
    )

    if match:
        date_text = match.group(1).replace(".", "-").replace("/", "-")
        return pd.to_datetime(date_text, errors="coerce")

    return pd.to_datetime(value, errors="coerce")


def format_date(value):
    dt = parse_date(value)

    if pd.isna(dt):
        return "날짜없음"

    return dt.strftime("%Y-%m-%d")


def extract_domain(url):
    """
    기존 중복검사와 동일한 방식:
    URL 전체가 아니라 호스트(도메인)가 같으면 중복으로 판정.

    www.example.com -> example.com
    example.com/a   -> example.com
    """
    if pd.isna(url):
        return ""

    text = str(url).strip()

    if not text:
        return ""

    if not text.lower().startswith(("http://", "https://")):
        text = "http://" + text

    try:
        parsed = urlparse(text)
        netloc = parsed.netloc.lower().strip()

        if netloc.startswith("www."):
            netloc = netloc[4:]

        # 포트 제거
        return netloc.split(":")[0]

    except Exception:
        return ""


def normalize_alias(text):
    """사이트명 비교용 정규화"""
    if text is None or pd.isna(text):
        return ""

    text = str(text).strip().lower()

    # 공백 제거
    text = re.sub(r"\s+", "", text)

    # 앞뒤 구분기호 제거
    text = re.sub(r"^[\-_,:;|/]+|[\-_,:;|/]+$", "", text)

    return text


# =========================================================
# 사이트명 / 별칭 처리
# =========================================================

def extract_current_site_expression(info_name):
    """
    처리대기 파일의 정보명에서 실제 사이트명 부분만 추출.

    예:
    대체-티비위키             -> 티비위키
    링크정보2-늑대닷컴        -> 늑대닷컴
    링크정보-2, 짭플릭스      -> 짭플릭스
    대체-좋은티비(밴드티비, 티비조타)
                             -> 좋은티비(밴드티비, 티비조타)
    """
    if pd.isna(info_name):
        return ""

    text = str(info_name).strip()

    # 대체-, 대체02-, 대체-02-, 신규- 등
    # 앞쪽 업무구분 숫자는 제거하되 사이트명 뒤의 숫자는 그대로 유지합니다.
    # 예: 대체02-늑대닷컴2 -> 늑대닷컴2
    match = re.match(
        r"^\s*(?:대체|신규)\s*(?:[-,:]?\s*\d+)?\s*[-,:]?\s*(.+)$",
        text,
        flags=re.IGNORECASE
    )

    if match:
        return match.group(1).strip()

    # 링크정보2-, 링크정보-2, 등
    match = re.match(
        r"^\s*링크정보\s*-?\s*\d*\s*[-,:]?\s*(.+)$",
        text,
        flags=re.IGNORECASE
    )

    if match:
        value = match.group(1).strip()
        value = re.sub(r"^[,\-:]\s*", "", value)
        return value

    # 예외적인 링크정보 형식
    if "," in text:
        left, right = text.split(",", 1)

        if "링크정보" in left:
            return right.strip()

    # 일반적인 '-' 형식의 마지막 부분
    if "-" in text:
        return text.rsplit("-", 1)[-1].strip()

    return text


def extract_keyword_site_expression(keyword):
    """
    기준파일 키워드에서 사이트명 부분만 추출.

    예:
    02대체, 티비위키
        -> 티비위키

    02대체, 웹툰365(weptoon), 한국만화출판협회
        -> 웹툰365(weptoon)

    02대체, 좋은티비(밴드티비, 티비조타)
        -> 좋은티비(밴드티비, 티비조타)
    """
    if pd.isna(keyword):
        return ""

    text = str(keyword).strip()

    index = text.find("02대체")

    if index == -1:
        return ""

    text = text[index + len("02대체"):]
    text = text.lstrip(" ,)]}>-").strip()

    # 괄호 밖에서 처음 나오는 쉼표까지만 사이트 표현으로 봄
    depth = 0
    result = []

    for char in text:

        if char == "(":
            depth += 1

        elif char == ")" and depth > 0:
            depth -= 1

        if char == "," and depth == 0:
            break

        result.append(char)

    return "".join(result).strip()


def extract_primary_alias(expression):
    """
    대표 사이트명 추출.

    (코믹툰툰)뚜툰 -> 뚜툰
    뚜툰(코믹툰툰) -> 뚜툰
    """
    if not expression:
        return ""

    text = str(expression).strip()

    outside = re.sub(r"\([^()]*\)", "", text).strip()

    if outside:
        # 혹시 밖에 복수 표기가 있으면 첫 항목 사용
        first = re.split(r"[,/|;]+", outside)[0]
        return normalize_alias(first)

    # 괄호 안밖 구조가 특이한 경우 별칭 중 하나라도 반환
    aliases = extract_aliases(text)

    return sorted(aliases)[0] if aliases else ""


def extract_aliases(expression):
    """
    하나의 사이트 표현을 동일사이트 별칭 집합으로 변환.

    예:
    '(코믹툰툰)뚜툰'
        -> {'코믹툰툰', '뚜툰'}

    '좋은티비(밴드티비, 티비조타)'
        -> {'좋은티비', '밴드티비', '티비조타'}
    """
    if expression is None or pd.isna(expression):
        return set()

    text = str(expression).strip()

    if not text:
        return set()

    aliases = set()

    # 괄호 내부
    inside_list = re.findall(r"\(([^()]*)\)", text)

    # 괄호를 제거한 바깥쪽 이름
    outside = re.sub(r"\([^()]*\)", "", text).strip()

    parts = [outside] + inside_list

    for part in parts:

        # 괄호 안 "A, B" 형태도 각각 별칭 처리
        tokens = re.split(r"[,/|;]+", part)

        for token in tokens:
            alias = normalize_alias(token)

            if alias:
                aliases.add(alias)

    return aliases


# =========================================================
# 동일사이트 그룹 연결용 Union-Find
# =========================================================

class UnionFind:

    def __init__(self):
        self.parent = {}
        self.rank = {}

    def add(self, item):

        if item not in self.parent:
            self.parent[item] = item
            self.rank[item] = 0

    def find(self, item):

        self.add(item)

        if self.parent[item] != item:
            self.parent[item] = self.find(self.parent[item])

        return self.parent[item]

    def union(self, a, b):

        root_a = self.find(a)
        root_b = self.find(b)

        if root_a == root_b:
            return

        if self.rank[root_a] < self.rank[root_b]:
            root_a, root_b = root_b, root_a

        self.parent[root_b] = root_a

        if self.rank[root_a] == self.rank[root_b]:
            self.rank[root_a] += 1


# =========================================================
# 기준파일 전처리
# =========================================================

def prepare_reference_data(df_ref):

    required_columns = [
        "접수번호",
        "URL",
        "결정일",
        "키워드"
    ]

    missing = [
        col for col in required_columns
        if col not in df_ref.columns
    ]

    if missing:
        raise ValueError(
            "기준파일에 필요한 열이 없습니다.\n"
            f"누락된 열: {', '.join(missing)}"
        )

    df_ref = df_ref.copy()

    # 중복검사용 도메인
    df_ref["_domain"] = (
        df_ref["URL"]
        .apply(extract_domain)
    )

    # 날짜/접수번호 정렬용
    df_ref["_date"] = (
        df_ref["결정일"]
        .apply(parse_date)
    )

    df_ref["_id_num"] = (
        pd.to_numeric(
            df_ref["접수번호"],
            errors="coerce"
        )
        .fillna(-1)
    )

    # 사이트명/별칭 추출
    df_ref["_site_expr"] = (
        df_ref["키워드"]
        .apply(extract_keyword_site_expression)
    )

    df_ref["_site_primary"] = (
        df_ref["_site_expr"]
        .apply(extract_primary_alias)
    )

    df_ref["_aliases"] = (
        df_ref["_site_expr"]
        .apply(extract_aliases)
    )

    # -----------------------------------------------------
    # 1) 도메인 -> 기준파일의 가장 최근 이력
    # -----------------------------------------------------

    valid_domain = df_ref[
        df_ref["_domain"] != ""
    ].copy()

    valid_domain = valid_domain.sort_values(
        by=["_date", "_id_num"],
        ascending=[False, False],
        na_position="last"
    )

    latest_domain_record = {}

    for idx, row in valid_domain.iterrows():

        domain = row["_domain"]

        if domain not in latest_domain_record:
            latest_domain_record[domain] = idx

    # -----------------------------------------------------
    # 2) 사이트 별칭 그룹 구성
    # -----------------------------------------------------

    uf = UnionFind()

    alias_to_rows = defaultdict(set)

    for idx, row in df_ref.iterrows():

        aliases = row["_aliases"]

        if not aliases:
            continue

        alias_list = list(aliases)

        for alias in alias_list:
            uf.add(alias)
            alias_to_rows[alias].add(idx)

        # 같은 키워드 안의 괄호 별칭들을 동일사이트로 연결
        first = alias_list[0]

        for alias in alias_list[1:]:
            uf.union(first, alias)

    # root -> 관련 기준파일 행
    component_rows = defaultdict(set)

    for alias, rows in alias_to_rows.items():

        root = uf.find(alias)
        component_rows[root].update(rows)

    return (
        df_ref,
        latest_domain_record,
        uf,
        alias_to_rows,
        component_rows
    )


def select_latest_history(
    df_ref,
    current_expression,
    uf,
    alias_to_rows,
    component_rows
):
    """
    현재 사이트와 동일사이트로 연결된 과거 이력 중 가장 최근 건 선택.

    우선순위:
    1. 결정일 최신
    2. 같은 결정일이면 현재 대표 사이트명과 정확히 같은 키워드 우선
    3. 그 다음 현재 별칭이 직접 포함된 키워드
    4. 접수번호가 큰 건 우선
    """

    current_aliases = extract_aliases(
        current_expression
    )

    current_primary = extract_primary_alias(
        current_expression
    )

    if not current_aliases:
        return None

    # -------------------------------------------------
    # 1순위: 현재 사이트명이 기준파일 키워드에 직접 존재하는 이력
    # -------------------------------------------------
    # 예: 현재가 '늑대닷컴2'이면 '늑대닷컴'과 과거 별칭 관계가
    # 있더라도 먼저 '늑대닷컴2'가 직접 적힌 이력만 검색합니다.
    direct_candidate_indices = set()

    for alias in current_aliases:
        direct_candidate_indices.update(
            alias_to_rows.get(alias, set())
        )

    if direct_candidate_indices:
        candidate_indices = direct_candidate_indices

    else:
        # -------------------------------------------------
        # 2순위: 직접 이름 이력이 전혀 없을 때만 동일사이트 별칭 그룹 확장
        # -------------------------------------------------
        # 예: '(코믹툰툰)뚜툰'처럼 과거/현재 명칭이 같은 사이트로
        # 연결되어 있을 경우 해당 그룹의 이력을 후보로 사용합니다.
        candidate_indices = set()

        for alias in current_aliases:
            if alias in uf.parent:
                root = uf.find(alias)
                candidate_indices.update(
                    component_rows.get(root, set())
                )

    if not candidate_indices:
        return None

    candidates = df_ref.loc[
        list(candidate_indices)
    ].copy()

    # 대표 사이트명 정확일치
    candidates["_match_exact"] = (
        candidates["_site_primary"]
        == current_primary
    ).astype(int)

    # 현재 별칭 중 하나가 해당 행에 직접 존재
    candidates["_match_direct"] = (
        candidates["_aliases"]
        .apply(
            lambda row_aliases:
                bool(current_aliases & row_aliases)
        )
    ).astype(int)

    # 최신 날짜를 가장 먼저 본다.
    # 같은 날짜라면 정확 매칭 > 직접 별칭 > 큰 접수번호
    candidates = candidates.sort_values(
        by=[
            "_date",
            "_match_exact",
            "_match_direct",
            "_id_num"
        ],
        ascending=[
            False,
            False,
            False,
            False
        ],
        na_position="last"
    )

    if candidates.empty:
        return None

    return candidates.iloc[0]


# =========================================================
# 메인 처리
# =========================================================

def run_process(reference_path, daily_path):

    try:
        df_ref = pd.read_excel(
            reference_path
        )

        df_daily = pd.read_excel(
            daily_path
        )

    except Exception as e:
        raise ValueError(
            "엑셀 파일을 읽는 중 오류가 발생했습니다.\n"
            f"{e}"
        )

    # -----------------------------------------------------
    # 일일파일 열 확인
    # -----------------------------------------------------

    if df_daily.shape[1] == 0:
        raise ValueError(
            "일일파일에 열이 없습니다."
        )

    # 사용자의 기준:
    # 모니터면 A열=등록번호
    # 일반건이면 A열=접수번호
    # 따라서 무조건 첫 번째 열(A열)을 ID로 사용
    daily_id_col = df_daily.columns[0]

    # 작업파일에서 반드시 필요한 열은 URL과 정보명뿐입니다.
    # '정보내용'은 링크정보 건의 처리비고 생성에만 사용하므로 선택사항입니다.
    required_daily_columns = [
        "URL",
        "정보명"
    ]

    missing_daily = [
        col for col in required_daily_columns
        if col not in df_daily.columns
    ]

    if missing_daily:
        raise ValueError(
            "일일파일에 필요한 열이 없습니다.\n"
            f"누락된 열: {', '.join(missing_daily)}"
        )

    (
        df_ref,
        latest_domain_record,
        uf,
        alias_to_rows,
        component_rows
    ) = prepare_reference_data(df_ref)

    # -----------------------------------------------------
    # 일일파일 내부 최초 도메인 저장
    # -----------------------------------------------------

    seen_daily = {}

    # 화면 복사용 중복 ID
    duplicate_ids = []

    results = []

    normal_count = 0
    duplicate_count = 0
    review_count = 0

    prefix_text = (
        "해당 정보는 차단된 사이트와 동일한 불법 저작물을 제공하면서, "
        "차단을 회피하기 위해 접속 URL만 변경하고 있는 대체 사이트 내용임.\n\n"
    )

    # -----------------------------------------------------
    # 한 행씩 처리
    # -----------------------------------------------------

    for _, row in df_daily.iterrows():

        current_id = clean_id(
            row[daily_id_col]
        )

        current_url = (
            str(row["URL"]).strip()
            if pd.notna(row["URL"])
            else ""
        )

        current_info = (
            str(row["정보명"]).strip()
            if pd.notna(row["정보명"])
            else ""
        )

        # 링크정보 건은 대체사이트가 아니라 링크 제공 사이트로 별도 처리
        is_link_info = bool(
            re.match(
                r"^\s*링크정보",
                current_info,
                flags=re.IGNORECASE
            )
        )

        # 링크정보 건의 처리비고에는 작업파일의 '정보내용'을 원문 그대로 사용합니다.
        # 다만 처리대기 내보내기 형식에 따라 '정보내용' 열이 없을 수도 있으므로
        # 열 자체가 없을 때 프로그램 전체가 중단되지 않도록 선택적으로 읽습니다.
        current_content = ""

        if "정보내용" in df_daily.columns:
            content_value = row["정보내용"]

            if pd.notna(content_value):
                current_content = str(content_value)

        current_domain = extract_domain(
            current_url
        )

        current_site_expr = (
            extract_current_site_expression(
                current_info
            )
        )

        if is_link_info:
            new_keyword = (
                f"05-2링크, {current_site_expr}"
                if current_site_expr
                else "05-2링크"
            )
        else:
            new_keyword = (
                f"02대체, {current_site_expr}"
                if current_site_expr
                else ""
            )

        duplicate_details = []

        # =================================================
        # ① 전체 접속차단 기준파일과 도메인 중복
        # =================================================

        if (
            current_domain
            and current_domain
            in latest_domain_record
        ):

            ref_idx = latest_domain_record[
                current_domain
            ]

            ref_row = df_ref.loc[ref_idx]

            ref_id = clean_id(
                ref_row["접수번호"]
            )

            ref_url = (
                str(ref_row["URL"]).strip()
                if pd.notna(ref_row["URL"])
                else ""
            )

            ref_date = format_date(
                ref_row["결정일"]
            )

            duplicate_details.append(
                "기준파일 중복: "
                f"[URL {ref_url} / "
                f"접수번호 {ref_id} / "
                f"결정일 {ref_date}]"
            )

        # =================================================
        # ② 오늘 처리대기 파일 내부 도메인 중복
        # =================================================

        if current_domain:

            if current_domain in seen_daily:

                first_data = seen_daily[
                    current_domain
                ]

                duplicate_details.append(
                    "오늘파일 내부 중복: "
                    f"[최초번호 {first_data['id']} / "
                    f"URL {first_data['url']}]"
                )

            else:

                seen_daily[current_domain] = {
                    "id": current_id,
                    "url": current_url
                }

        # =================================================
        # ③ 중복이면 처리비고/키워드 생성하지 않음
        # =================================================

        if duplicate_details:

            duplicate_count += 1

            if current_id:
                duplicate_ids.append(
                    current_id
                )

            results.append({
                daily_id_col:
                    current_id,

                "처리비고":
                    "",

                "키워드":
                    "",

                "판정":
                    "중복",

                "중복상세":
                    "\n".join(
                        duplicate_details
                    ),

                "URL":
                    current_url,

                "정보명":
                    current_info,

                "참조접수번호":
                    "",

                "참조URL":
                    "",

                "참조결정일":
                    ""
            })

            continue

        # =================================================
        # ④ 링크정보 건: 정보내용 원문 + 전용 키워드
        # =================================================

        if is_link_info:

            # 링크정보 건은 작업파일의 '정보내용'을 처리비고로 그대로 사용합니다.
            # 정보내용 열이 없거나 값이 비어 있으면 프로그램을 중단하지 않고
            # 해당 건만 '확인필요'로 표시합니다.
            if current_content.strip():

                normal_count += 1
                link_judgment = "정상"
                link_detail = ""

            else:

                review_count += 1
                link_judgment = "확인필요"
                link_detail = (
                    "링크정보 건이지만 작업파일에 '정보내용'이 없습니다. "
                    "정보내용이 포함된 내보내기 파일인지 확인하세요."
                )

            results.append({
                daily_id_col:
                    current_id,

                "처리비고":
                    current_content,

                "키워드":
                    new_keyword,

                "판정":
                    link_judgment,

                "중복상세":
                    link_detail,

                "URL":
                    current_url,

                "정보명":
                    current_info,

                "참조접수번호":
                    "",

                "참조URL":
                    "",

                "참조결정일":
                    ""
            })

            continue

        # =================================================
        # ⑤ 대체사이트 건: 과거 이력 검색
        # =================================================

        best_match = select_latest_history(
            df_ref=df_ref,
            current_expression=current_site_expr,
            uf=uf,
            alias_to_rows=alias_to_rows,
            component_rows=component_rows
        )

        if best_match is not None:

            orig_url = (
                str(best_match["URL"]).strip()
                if pd.notna(best_match["URL"])
                else ""
            )

            orig_id = clean_id(
                best_match["접수번호"]
            )

            date_str = format_date(
                best_match["결정일"]
            )

            memo_text = (
                f"{prefix_text}"
                f"[원사이트 {orig_url} / "
                f"접수번호 {orig_id} / "
                f"결정일자 {date_str}]"
            )

            judgment = "정상"
            normal_count += 1

            ref_id = orig_id
            ref_url = orig_url
            ref_date = date_str

        else:

            # 과거 사이트 이력이 없으면 잘못된 URL을 임의로 넣지 않고
            # 기존 프로그램처럼 확인용 문구만 생성
            memo_text = (
                f"[과거 정보열람 이력 없음: "
                f"{current_site_expr}]"
            )

            judgment = "확인필요"
            review_count += 1

            ref_id = ""
            ref_url = ""
            ref_date = ""

        results.append({
            daily_id_col:
                current_id,

            "처리비고":
                memo_text,

            "키워드":
                new_keyword,

            "판정":
                judgment,

            "중복상세":
                "",

            "URL":
                current_url,

            "정보명":
                current_info,

            "참조접수번호":
                ref_id,

            "참조URL":
                ref_url,

            "참조결정일":
                ref_date
        })

    # 순서 유지 중복 제거
    duplicate_ids = list(
        dict.fromkeys(
            duplicate_ids
        )
    )

    # -----------------------------------------------------
    # 결과 엑셀 생성 - 사용자가 제공한 업로드 샘플 원본을 템플릿으로 사용
    # -----------------------------------------------------

    directory = os.path.dirname(
        daily_path
    )

    base_name = os.path.splitext(
        os.path.basename(
            daily_path
        )
    )[0]

    out_path = os.path.join(
        directory,
        f"{base_name}_통합처리완료.xlsx"
    )

    # 프로그램에 내장된 '업로드샘플(수정용).xlsx'를 그대로 불러옵니다.
    # 따라서 제목, 열 순서, 글꼴, 열 너비, 배경색, 테두리,
    # 시트명, 인쇄설정 등이 원본 샘플과 동일하게 유지됩니다.
    template_bytes = base64.b64decode(
        UPLOAD_TEMPLATE_B64
    )

    workbook = load_workbook(
        BytesIO(template_bytes)
    )

    if "2023 하반기" in workbook.sheetnames:
        worksheet = workbook["2023 하반기"]
    else:
        worksheet = workbook[workbook.sheetnames[0]]

    upload_headers = [
        "고유번호",
        "정보명",
        "URL",
        "인지방법",
        "단체명",
        "신청자명",
        "위반내용",
        "주제",
        "정보유형",
        "유통형태",
        "유통망",
        "발생장소",
        "업체명",
        "부서",
        "분과",
        "서버위치",
        "서버IP",
        "키워드",
        "정보내용",
        "처리비고",
        "증거자료1",
        "증거자료2",
        "증거자료3",
        "증거자료4",
        "증거자료5",
        "대표이미지(썸네일)"
    ]

    # 혹시 내장 템플릿이 훼손됐는지 헤더를 확인하고,
    # 다를 경우에도 정확한 업로드 제목으로 복구합니다.
    for col_idx, header in enumerate(
        upload_headers,
        start=1
    ):
        worksheet.cell(
            row=1,
            column=col_idx
        ).value = header

    # -----------------------------------------------------
    # 템플릿 데이터행 스타일 확보
    # -----------------------------------------------------

    # 샘플의 후반부 행은 업로드용 기본 데이터 형식(A/R/T 중심)이므로
    # 추가 행을 만들 때 이 행의 스타일을 복사합니다.
    source_style_row = min(
        max(worksheet.max_row - 1, 2),
        23
    )

    source_styles = []

    for col_idx in range(1, 27):
        source_cell = worksheet.cell(
            row=source_style_row,
            column=col_idx
        )
        source_styles.append(
            copy(source_cell._style)
        )

    source_row_height = (
        worksheet.row_dimensions[
            source_style_row
        ].height
        or 214.5
    )

    # -----------------------------------------------------
    # 샘플에 들어 있던 예시 데이터만 삭제 (서식은 유지)
    # -----------------------------------------------------

    original_max_row = worksheet.max_row

    for row_idx in range(
        2,
        original_max_row + 1
    ):
        for col_idx in range(1, 27):
            worksheet.cell(
                row=row_idx,
                column=col_idx
            ).value = None

    required_last_row = len(results) + 1

    # 결과가 샘플 행수보다 많으면 동일 스타일로 행 추가
    if required_last_row > original_max_row:

        for row_idx in range(
            original_max_row + 1,
            required_last_row + 1
        ):
            for col_idx in range(1, 27):
                cell = worksheet.cell(
                    row=row_idx,
                    column=col_idx
                )
                cell._style = copy(
                    source_styles[col_idx - 1]
                )

            worksheet.row_dimensions[
                row_idx
            ].height = source_row_height

    # 결과가 샘플보다 적으면 남는 예시 행 제거
    elif required_last_row < original_max_row:
        worksheet.delete_rows(
            required_last_row + 1,
            original_max_row - required_last_row
        )

    # -----------------------------------------------------
    # 결과 입력
    # 샘플과 동일하게 A=고유번호 / R=키워드 / T=처리비고만 채움
    # -----------------------------------------------------

    duplicate_fill = PatternFill(
        fill_type="solid",
        fgColor="FFC7CE"
    )

    duplicate_font = Font(
        name="맑은 고딕",
        size=11,
        color="9C0006",
        bold=True
    )

    review_fill = PatternFill(
        fill_type="solid",
        fgColor="FFF2CC"
    )

    for result_index, result in enumerate(
        results,
        start=2
    ):

        judgment = result.get(
            "판정",
            ""
        )

        current_id = str(
            result.get(
                daily_id_col,
                ""
            )
            or ""
        )

        keyword = str(
            result.get(
                "키워드",
                ""
            )
            or ""
        )

        memo = str(
            result.get(
                "처리비고",
                ""
            )
            or ""
        )

        detail = str(
            result.get(
                "중복상세",
                ""
            )
            or ""
        )

        # 중복건은 바로 각하할 수 있도록
        # 키워드는 만들지 않고 T열에 중복 상대만 표시합니다.
        if judgment == "중복":
            keyword = ""
            memo = (
                "[중복]\n" + detail
                if detail
                else "[중복]"
            )

        # 확인필요인데 처리비고가 비어 있는 경우 사유를 T열에 표시
        elif judgment == "확인필요" and not memo.strip():
            memo = detail

        # A열 = 고유번호
        worksheet.cell(
            row=result_index,
            column=1
        ).value = current_id
        worksheet.cell(
            row=result_index,
            column=1
        ).number_format = "@"

        # R열 = 키워드
        worksheet.cell(
            row=result_index,
            column=18
        ).value = keyword

        # T열 = 처리비고 또는 중복상세
        worksheet.cell(
            row=result_index,
            column=20
        ).value = memo

        # 샘플의 행 높이 범위(198 / 214.5)를 유지하면서
        # 긴 처리비고는 조금 더 높은 행을 사용합니다.
        memo_length = len(memo)
        memo_lines = memo.count("\n") + 1

        if memo_length >= 175 or memo_lines >= 4:
            worksheet.row_dimensions[
                result_index
            ].height = 214.5
        else:
            worksheet.row_dimensions[
                result_index
            ].height = 198

        # 기존 요청: 중복건 전체 빨간색
        if judgment == "중복":
            for col_idx in range(1, 27):
                cell = worksheet.cell(
                    row=result_index,
                    column=col_idx
                )
                cell.fill = duplicate_fill
                cell.font = duplicate_font

        # 확인필요 전체 노란색
        elif judgment == "확인필요":
            for col_idx in range(1, 27):
                worksheet.cell(
                    row=result_index,
                    column=col_idx
                ).fill = review_fill

    # 행 추가/삭제 후 사용 영역에 맞춰 자동필터 정의가 남아있다면 갱신
    if worksheet.auto_filter.ref:
        worksheet.auto_filter.ref = (
            f"A1:Z{required_last_row}"
        )

    workbook.save(
        out_path
    )

    return {
        "out_path":
            out_path,

        "duplicate_ids":
            duplicate_ids,

        "total_count":
            len(results),

        "duplicate_count":
            duplicate_count,

        "normal_count":
            normal_count,

        "review_count":
            review_count,

        "id_column":
            daily_id_col
    }


# =========================================================
# GUI
# =========================================================

class App:

    def __init__(self, root):

        self.root = root

        self.root.title(
            "02대체 통합 자동처리기 v1.5"
        )

        self.root.geometry(
            "700x650"
        )

        self.root.resizable(
            False,
            False
        )

        self.reference_path = ""
        self.daily_path = ""

        # -------------------------------------------------
        # 기준파일
        # -------------------------------------------------

        tk.Label(
            root,
            text=(
                "[1] 기준 파일: "
                "02대체(접속차단)전체.xlsx"
            ),
            font=(
                "맑은 고딕",
                9,
                "bold"
            )
        ).pack(
            anchor="w",
            padx=20,
            pady=(15, 2)
        )

        self.lbl_reference = tk.Label(
            root,
            text="파일을 선택하세요...",
            fg="gray",
            anchor="w"
        )

        self.lbl_reference.pack(
            fill="x",
            padx=20
        )

        tk.Button(
            root,
            text="기준파일 선택",
            command=self.select_reference
        ).pack(
            anchor="w",
            padx=20,
            pady=(2, 12)
        )

        # -------------------------------------------------
        # 일일파일
        # -------------------------------------------------

        tk.Label(
            root,
            text=(
                "[2] 작업할 파일: "
                "모니터/일반 처리대기 내보내기"
            ),
            font=(
                "맑은 고딕",
                9,
                "bold"
            )
        ).pack(
            anchor="w",
            padx=20,
            pady=(5, 2)
        )

        self.lbl_daily = tk.Label(
            root,
            text="파일을 선택하세요...",
            fg="gray",
            anchor="w"
        )

        self.lbl_daily.pack(
            fill="x",
            padx=20
        )

        tk.Button(
            root,
            text="처리대기 파일 선택",
            command=self.select_daily
        ).pack(
            anchor="w",
            padx=20,
            pady=(2, 15)
        )

        # -------------------------------------------------
        # 실행버튼
        # -------------------------------------------------

        tk.Button(
            root,
            text=(
                "중복검사 + 처리비고 + "
                "키워드 한 번에 생성"
            ),
            bg="#0078D7",
            fg="white",
            font=(
                "맑은 고딕",
                11,
                "bold"
            ),
            height=2,
            command=self.process
        ).pack(
            fill="x",
            padx=20
        )

        # -------------------------------------------------
        # 설명
        # -------------------------------------------------

        tk.Label(
            root,
            text=(
                "※ 중복건: 빨간색 / "
                "처리비고·키워드 미생성\n"
                "※ 정상건: 처리비고·키워드 자동 생성 / "
                "과거이력 없음: 노란색"
            ),
            justify="left",
            fg="#555555",
            font=(
                "맑은 고딕",
                9
            )
        ).pack(
            anchor="w",
            padx=20,
            pady=(12, 8)
        )

        # -------------------------------------------------
        # 결과창
        # -------------------------------------------------

        tk.Label(
            root,
            text=(
                "[검사 결과 / 중복번호 복사용]"
            ),
            font=(
                "맑은 고딕",
                9,
                "bold"
            )
        ).pack(
            anchor="w",
            padx=20,
            pady=(5, 2)
        )

        self.result_text = (
            scrolledtext.ScrolledText(
                root,
                height=16,
                width=90,
                font=(
                    "맑은 고딕",
                    10
                )
            )
        )

        self.result_text.pack(
            padx=20,
            pady=(0, 10)
        )

        self.result_text.insert(
            tk.END,
            "두 파일을 선택한 뒤 실행버튼을 누르세요.\n"
        )

    def select_reference(self):

        file = filedialog.askopenfilename(
            filetypes=[
                (
                    "Excel Files",
                    "*.xlsx *.xlsm"
                )
            ]
        )

        if file:

            self.reference_path = file

            self.lbl_reference.config(
                text=os.path.basename(file),
                fg="black"
            )

    def select_daily(self):

        file = filedialog.askopenfilename(
            filetypes=[
                (
                    "Excel Files",
                    "*.xlsx *.xlsm"
                )
            ]
        )

        if file:

            self.daily_path = file

            self.lbl_daily.config(
                text=os.path.basename(file),
                fg="black"
            )

    def process(self):

        if (
            not self.reference_path
            or not self.daily_path
        ):

            messagebox.showwarning(
                "경고",
                "두 파일을 모두 선택해주세요."
            )

            return

        try:

            self.result_text.delete(
                1.0,
                tk.END
            )

            self.result_text.insert(
                tk.END,
                "처리 중입니다...\n"
            )

            self.root.update()

            result = run_process(
                self.reference_path,
                self.daily_path
            )

            self.result_text.delete(
                1.0,
                tk.END
            )

            summary = (
                f"전체: {result['total_count']}건\n"
                f"정상: {result['normal_count']}건\n"
                f"중복: {result['duplicate_count']}건\n"
                f"확인필요: {result['review_count']}건\n\n"
            )

            self.result_text.insert(
                tk.END,
                summary
            )

            if result["duplicate_ids"]:

                self.result_text.insert(
                    tk.END,
                    "[중복 번호 - 그대로 복사 가능]\n"
                )

                self.result_text.insert(
                    tk.END,
                    ",".join(
                        result["duplicate_ids"]
                    )
                )

                self.result_text.insert(
                    tk.END,
                    "\n\n"
                )

            else:

                self.result_text.insert(
                    tk.END,
                    "중복 번호 없음\n\n"
                )

            self.result_text.insert(
                tk.END,
                "[결과파일]\n"
                f"{result['out_path']}"
            )

            messagebox.showinfo(
                "완료",
                "통합 처리가 완료되었습니다.\n\n"
                f"전체 {result['total_count']}건\n"
                f"중복 {result['duplicate_count']}건\n"
                f"정상 {result['normal_count']}건\n"
                f"확인필요 {result['review_count']}건\n\n"
                "중복행은 빨간색으로 표시되며 "
                "처리비고와 키워드는 비워집니다."
            )

        except Exception as e:

            self.result_text.delete(
                1.0,
                tk.END
            )

            self.result_text.insert(
                tk.END,
                f"오류 발생:\n{str(e)}"
            )

            messagebox.showerror(
                "오류",
                f"작업 중 오류가 발생했습니다.\n\n{str(e)}"
            )


if __name__ == "__main__":

    root = tk.Tk()

    app = App(root)

    root.mainloop()
