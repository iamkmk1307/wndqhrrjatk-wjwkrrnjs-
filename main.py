import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from urllib.parse import urlsplit
import pandas as pd


def extract_domain(url):
    """
    기준 차단목록과 일일파일을 비교할 때 사용하는 도메인 추출 함수.

    예:
      https://www.t.me/s/funbe_next -> t.me
      t.me/s/toonkor_com            -> t.me

    기준 차단목록 비교에서만 이 값을 사용한다.
    """
    if pd.isna(url):
        return ""

    url = str(url).strip()
    if not url:
        return ""

    if not url.lower().startswith(("http://", "https://")):
        url = "http://" + url

    try:
        parsed = urlsplit(url)
        netloc = parsed.netloc.lower().strip()

        # 사용자정보가 있는 특수 URL 방어
        if "@" in netloc:
            netloc = netloc.rsplit("@", 1)[-1]

        # www. 제거
        if netloc.startswith("www."):
            netloc = netloc[4:]

        # 포트번호 제거
        # 일반적인 host:port 형식 기준
        if ":" in netloc and not netloc.startswith("["):
            netloc = netloc.split(":", 1)[0]

        return netloc

    except Exception:
        return ""


def normalize_full_url(url):
    """
    일일파일 내부 중복 비교 전용.

    도메인만 비교하지 않고 하위 경로/쿼리까지 포함한 전체 URL을 비교한다.

    아래 차이만 무시:
      - http:// 와 https://
      - 맨 앞 www.
      - URL 마지막의 /

    예:
      https://www.t.me/s/funbe_next/
        -> t.me/s/funbe_next

      http://t.me/s/funbe_next
        -> t.me/s/funbe_next

      t.me/s/toonkor_com
        -> t.me/s/toonkor_com

    따라서
      t.me/s/funbe_next != t.me/s/toonkor_com
    이므로 서로 중복이 아니다.
    """
    if pd.isna(url):
        return ""

    value = str(url).strip()
    if not value:
        return ""

    # scheme 제거
    value = re.sub(r"(?i)^https?://", "", value)

    # www. 제거
    value = re.sub(r"(?i)^www\.", "", value)

    # 마지막 /만 제거
    value = re.sub(r"/+$", "", value).strip()

    if not value:
        return ""

    # host 부분만 소문자로 통일하고 path/query/fragment는 그대로 유지
    m = re.match(r"^([^/?#]+)(.*)$", value)
    if not m:
        return value

    host = m.group(1).lower()
    rest = m.group(2)

    return host + rest


def clean_id(value):
    """등록번호/접수번호 값 정리"""
    if pd.isna(value):
        return ""

    value = str(value).strip()

    if value.endswith(".0"):
        value = value[:-2]

    return value


def run_process(daily_path, block_path):

    # -----------------------------
    # 1. 엑셀 읽기
    # -----------------------------
    df_daily = pd.read_excel(daily_path)

    df_block = pd.read_excel(
        block_path,
        sheet_name="차단목록"
    )

    # -----------------------------
    # 2. URL 열 검색
    # -----------------------------
    def find_col(df, keywords):
        for col in df.columns:
            cleaned = str(col).replace(" ", "").lower()

            for keyword in keywords:
                if keyword.lower() in cleaned:
                    return col

        return None

    daily_url_col = find_col(
        df_daily,
        ["URL", "주소", "링크"]
    )

    block_url_col = find_col(
        df_block,
        ["URL", "주소", "링크"]
    )

    block_id_col = find_col(
        df_block,
        ["접수번호", "기존접수번호", "관리번호"]
    )

    # -----------------------------
    # 3. 일일파일 번호는 무조건 A열
    # -----------------------------
    daily_id_col = df_daily.columns[0]

    if daily_url_col is None:
        raise ValueError(
            "일일파일에서 URL 열을 찾을 수 없습니다."
        )

    if block_url_col is None:
        raise ValueError(
            "기준파일에서 URL 열을 찾을 수 없습니다."
        )

    # -----------------------------
    # 4. 비교키 생성
    # -----------------------------

    # 기준파일 ↔ 일일파일 비교용:
    # 기존대로 메인 도메인 기준
    df_daily["검사용_도메인"] = (
        df_daily[daily_url_col]
        .apply(extract_domain)
    )

    df_block["검사용_도메인"] = (
        df_block[block_url_col]
        .apply(extract_domain)
    )

    # 일일파일 내부 비교용:
    # 도메인이 아니라 하위 경로를 포함한 전체 URL 기준
    df_daily["검사용_전체URL"] = (
        df_daily[daily_url_col]
        .apply(normalize_full_url)
    )

    # -----------------------------
    # 5. 기존 차단목록 딕셔너리 생성
    # -----------------------------
    mapping_dict = {}

    for _, row in df_block.iterrows():

        dom = row["검사용_도메인"]

        if not dom:
            continue

        if (
            block_id_col is not None
            and pd.notna(row[block_id_col])
        ):
            block_id = clean_id(
                row[block_id_col]
            )
        else:
            block_id = "접수번호없음"

        if pd.notna(row[block_url_col]):
            block_url = str(
                row[block_url_col]
            ).strip()
        else:
            block_url = "URL없음"

        # 동일 도메인이 기준파일에 여러 개 있어도
        # 기존 프로그램과 동일하게 최초 데이터 하나만 저장
        if dom not in mapping_dict:
            mapping_dict[dom] = {
                "url": block_url,
                "id": block_id
            }

    # -----------------------------
    # 6. 중복 검사
    # -----------------------------

    dup_ids = []
    excel_data = []

    # ★ 중요:
    # 예전 EXE에서는 seen_daily에 '도메인'을 저장해서
    # t.me/AAA 와 t.me/BBB도 중복으로 처리했음.
    #
    # 수정본에서는 '정규화된 전체 URL'을 저장한다.
    seen_daily_urls = {}

    for _, row in df_daily.iterrows():

        dom = row["검사용_도메인"]
        full_url_key = row["검사용_전체URL"]

        if not dom:
            continue

        current_id = clean_id(
            row[daily_id_col]
        )

        current_url = (
            str(row[daily_url_col]).strip()
            if pd.notna(row[daily_url_col])
            else ""
        )

        duplicate_types = []
        duplicate_infos = []

        # =====================================
        # ① 기존 차단목록과 중복 검사
        #    → 기존대로 도메인 기준
        # =====================================

        if dom in mapping_dict:

            block_info = mapping_dict[dom]

            duplicate_types.append(
                "기존 차단목록 중복"
            )

            duplicate_infos.append(
                f"[{block_info['url']} / "
                f"접수번호 {block_info['id']}]"
            )

        # =====================================
        # ② 일일파일 내부 중복 검사
        #    → 전체 URL 기준
        # =====================================

        if full_url_key:

            if full_url_key in seen_daily_urls:

                first_info = seen_daily_urls[full_url_key]

                duplicate_types.append(
                    "일일파일 내부 동일 URL 중복"
                )

                duplicate_infos.append(
                    f"[최초번호 {first_info['id']} / "
                    f"URL {first_info['url']}]"
                )

            else:

                # 최초 등장 전체 URL 기록
                seen_daily_urls[full_url_key] = {
                    "id": current_id,
                    "url": current_url
                }

        # =====================================
        # 중복이면 결과 저장
        # =====================================

        if duplicate_types:

            if current_id:
                dup_ids.append(current_id)

            excel_data.append({

                "중복된 등록/접수번호":
                    current_id,

                "일일파일 URL":
                    current_url,

                "도메인":
                    dom,

                "중복유형":
                    " + ".join(
                        duplicate_types
                    ),

                "중복 상세정보":
                    " / ".join(
                        duplicate_infos
                    )
            })

    # -----------------------------
    # 7. 중복번호 중복제거
    # -----------------------------

    # 순서는 유지하면서 번호 중복 제거
    dup_ids = list(
        dict.fromkeys(dup_ids)
    )

    # -----------------------------
    # 8. 결과 엑셀 생성
    # -----------------------------

    df_out = pd.DataFrame(
        excel_data,
        columns=[
            "중복된 등록/접수번호",
            "일일파일 URL",
            "도메인",
            "중복유형",
            "중복 상세정보"
        ]
    )

    dir_name = os.path.dirname(
        daily_path
    )

    base_name = os.path.splitext(
        os.path.basename(daily_path)
    )[0]

    out_path = os.path.join(
        dir_name,
        f"{base_name}_중복검사완료.xlsx"
    )

    with pd.ExcelWriter(
        out_path,
        engine="openpyxl"
    ) as writer:

        df_out.to_excel(
            writer,
            index=False,
            sheet_name="중복결과"
        )

        worksheet = writer.sheets[
            "중복결과"
        ]

        worksheet.column_dimensions[
            "A"
        ].width = 22

        worksheet.column_dimensions[
            "B"
        ].width = 40

        worksheet.column_dimensions[
            "C"
        ].width = 30

        worksheet.column_dimensions[
            "D"
        ].width = 28

        worksheet.column_dimensions[
            "E"
        ].width = 70

    return (
        out_path,
        dup_ids,
        len(excel_data)
    )


# ==========================================
# 화면 UI
# ==========================================

class App:

    def __init__(self, root):

        self.root = root

        self.root.title(
            "저작권 URL 중복검사기 v1.4"
        )

        self.root.geometry(
            "520x500"
        )

        self.root.resizable(
            False,
            False
        )

        self.block_path = ""
        self.daily_path = ""

        # -----------------------------------
        # 기준파일
        # -----------------------------------

        tk.Label(
            root,
            text="[1] 기준 파일: 중복검사(저작권).xlsm",
            font=("맑은 고딕", 9, "bold")
        ).pack(
            anchor="w",
            padx=20,
            pady=(15, 2)
        )

        self.lbl_block = tk.Label(
            root,
            text="파일을 선택하세요...",
            fg="gray",
            anchor="w"
        )

        self.lbl_block.pack(
            fill="x",
            padx=20
        )

        tk.Button(
            root,
            text="차단목록 파일 선택",
            command=self.select_block
        ).pack(
            anchor="w",
            padx=20,
            pady=(2, 10)
        )

        # -----------------------------------
        # 일일파일
        # -----------------------------------

        tk.Label(
            root,
            text="[2] 일일 파일: 모니터/일반 내보내기 파일",
            font=("맑은 고딕", 9, "bold")
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
            text="일일 파일 선택",
            command=self.select_daily
        ).pack(
            anchor="w",
            padx=20,
            pady=(2, 15)
        )

        # -----------------------------------
        # 검사버튼
        # -----------------------------------

        tk.Button(
            root,
            text="중복 검사 시작하기",
            bg="#0078D7",
            fg="white",
            font=("맑은 고딕", 11, "bold"),
            height=2,
            command=self.process
        ).pack(
            fill="x",
            padx=20
        )

        # -----------------------------------
        # 결과창
        # -----------------------------------

        tk.Label(
            root,
            text="[중복된 등록번호 / 접수번호]",
            font=("맑은 고딕", 9, "bold")
        ).pack(
            anchor="w",
            padx=20,
            pady=(15, 2)
        )

        self.result_text = (
            scrolledtext.ScrolledText(
                root,
                height=6,
                width=60,
                font=("맑은 고딕", 10)
            )
        )

        self.result_text.pack(
            padx=20,
            pady=(0, 10)
        )

        self.result_text.insert(
            tk.END,
            "검사를 완료하면 중복된 등록번호/"
            "접수번호가 표시됩니다.\n"
            "(전체 선택 후 복사해서 시스템에서 "
            "검색하세요)"
        )

    # ---------------------------------------
    # 기준파일 선택
    # ---------------------------------------

    def select_block(self):

        file = filedialog.askopenfilename(
            filetypes=[
                (
                    "Excel Files",
                    "*.xlsx *.xlsm"
                )
            ]
        )

        if file:

            self.block_path = file

            self.lbl_block.config(
                text=os.path.basename(file),
                fg="black"
            )

    # ---------------------------------------
    # 일일파일 선택
    # ---------------------------------------

    def select_daily(self):

        file = filedialog.askopenfilename(
            filetypes=[
                (
                    "Excel Files",
                    "*.xlsx *.xls"
                )
            ]
        )

        if file:

            self.daily_path = file

            self.lbl_daily.config(
                text=os.path.basename(file),
                fg="black"
            )

    # ---------------------------------------
    # 중복 검사 실행
    # ---------------------------------------

    def process(self):

        if (
            not self.block_path
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
                "검사 중입니다..."
            )

            self.root.update()

            (
                out_file,
                dup_ids,
                dup_count
            ) = run_process(
                self.daily_path,
                self.block_path
            )

            self.result_text.delete(
                1.0,
                tk.END
            )

            if dup_ids:

                result_str = ",".join(
                    dup_ids
                )

                self.result_text.insert(
                    tk.END,
                    result_str
                )

                messagebox.showinfo(
                    "완료",
                    f"총 {dup_count}건의 "
                    f"중복이 발견되었습니다!\n\n"
                    f"등록/접수번호 {len(dup_ids)}개가 "
                    f"표시되었습니다.\n\n"
                    f"생성된 엑셀 파일에서 "
                    f"상세내역을 확인하세요."
                )

            else:

                self.result_text.insert(
                    tk.END,
                    "중복된 등록번호/"
                    "접수번호가 없습니다."
                )

                messagebox.showinfo(
                    "완료",
                    "검사 결과 중복건이 없습니다!"
                )

        except Exception as e:

            self.result_text.delete(
                1.0,
                tk.END
            )

            self.result_text.insert(
                tk.END,
                f"오류 발생: {str(e)}"
            )

            messagebox.showerror(
                "오류 발생",
                f"작업 중 오류가 발생했습니다:\n"
                f"{str(e)}"
            )


if __name__ == "__main__":

    root = tk.Tk()

    app = App(root)

    root.mainloop()
