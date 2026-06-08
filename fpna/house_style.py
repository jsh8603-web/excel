"""
fpna.house_style — 맥킨지/회계법인 룩 단일 SSOT.

모든 템플릿 빌더·렌더러는 색·폰트·숫자서식·보더를 *직접* 지정하지 않고
이 모듈의 상수와 헬퍼만 사용한다. 룩을 바꾸려면 여기 한 곳만 고친다.

관찰한 공통 규약(CFI / Macabacus / Vertex42)을 토큰화한 것이며,
외부 템플릿 파일은 커밋하지 않는다(개념만 재현).

핵심 규약
---------
- 숫자: 음수 괄호 `#,##0;(#,##0)`, % `0.0%`. 단위는 헤더에 명시(₩mn 등).
- 무채색 본문 + 단일 액센트 1색.
- 입력셀=파랑 글씨, 계산셀=검정 글씨(IB 관례), 링크셀=초록.
- gridlines off, 본문 9~10pt sans(맑은 고딕/Calibri).
- 항목 좌측·숫자 우측·들여쓰기 계층. 합계는 상단 단일 보더.
- variance bridge = 누적 막대 + base 시리즈 투명(fill=none) 트릭.
"""
from __future__ import annotations

import fpna._bootstrap  # noqa: F401  (vendor/ 주입)

from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# --------------------------------------------------------------------------
# 1. 색 팔레트 (무채색 + 단일 액센트)
# --------------------------------------------------------------------------
# ARGB hex (앞 2자리 alpha 생략 시 openpyxl 이 FF 보정)
INK = "1A1A1A"          # 본문 검정(약간 누그러뜨린)
INK_SOFT = "595959"     # 보조 텍스트 회색
RULE = "BFBFBF"         # 가는 구분선 회색
RULE_STRONG = "808080"  # 합계 보더 진회색
BAND = "F2F2F2"         # 줄무늬/헤더 배경 옅은 회색
HEADER_BG = "404040"    # 헤더 진회색 배경
HEADER_FG = "FFFFFF"    # 헤더 흰 글씨
WHITE = "FFFFFF"

ACCENT = "2E5A87"       # 단일 액센트(차분한 네이비블루)
ACCENT_SOFT = "9DB7CE"  # 액센트 옅은 톤

INPUT_FG = "0070C0"     # 입력셀 파랑(IB 관례)
CALC_FG = INK           # 계산셀 검정
LINK_FG = "008000"      # 시트간 링크 초록

POS_FG = "1F7A1F"       # 양(개선) 초록
NEG_FG = "C00000"       # 음(악화) 빨강

# --------------------------------------------------------------------------
# 2. 폰트 (회사 PC 확정 보유 폰트만)
# --------------------------------------------------------------------------
# 맑은 고딕 = 한국어 Windows 기본 탑재. Calibri = Office 기본. 둘 다 안전.
FONT_NAME = "맑은 고딕"
FONT_NAME_LATIN = "Calibri"
SIZE_BODY = 10
SIZE_SMALL = 9
SIZE_TITLE = 16
SIZE_SUBTITLE = 11
SIZE_SECTION = 11

# --------------------------------------------------------------------------
# 3. 숫자 서식 코드 (Excel 범용)
# --------------------------------------------------------------------------
FMT_INT = "#,##0;(#,##0)"
FMT_INT_DASH = "#,##0;(#,##0);\"-\""       # 0을 대시로
FMT_NUM1 = "#,##0.0;(#,##0.0)"
FMT_NUM2 = "#,##0.00;(#,##0.00)"
FMT_PCT1 = "0.0%;(0.0%)"
FMT_PCT2 = "0.00%;(0.00%)"
FMT_PCT0 = "0%;(0%)"
FMT_MULT = '0.0"x"'                          # 배수(LTV/CAC 등)
FMT_DATE = "yyyy-mm-dd"
FMT_MONTH = "yyyy-mm"
FMT_MONEY_MN = '#,##0;(#,##0)'               # 단위는 헤더 표기(₩mn)

# --------------------------------------------------------------------------
# 4. 정렬
# --------------------------------------------------------------------------
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
CENTER = Alignment(horizontal="center", vertical="center")
LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)


def indent(level: int) -> Alignment:
    """들여쓰기 계층용 좌측 정렬(level 만큼 indent)."""
    return Alignment(horizontal="left", vertical="center", indent=max(0, level))


# --------------------------------------------------------------------------
# 5. 보더 (합계 상단 단일 보더가 핵심 IB 관례)
# --------------------------------------------------------------------------
_thin = Side(style="thin", color=RULE)
_med = Side(style="medium", color=RULE_STRONG)
_dbl = Side(style="double", color=RULE_STRONG)

BORDER_NONE = Border()
BORDER_TOP = Border(top=_thin)                       # 소계 상단
BORDER_TOP_STRONG = Border(top=_med)                 # 합계 상단
BORDER_TOP_BOTTOM = Border(top=_thin, bottom=_thin)
BORDER_TOTAL = Border(top=_med, bottom=_dbl)         # 최종 합계(위 single·아래 double)
BORDER_BOTTOM = Border(bottom=_thin)


# --------------------------------------------------------------------------
# 6. Font 팩토리
# --------------------------------------------------------------------------
def font(color: str = INK, *, bold: bool = False, size: int = SIZE_BODY,
         italic: bool = False) -> Font:
    return Font(name=FONT_NAME, size=size, bold=bold, italic=italic, color=color)


F_BODY = font()
F_BODY_BOLD = font(bold=True)
F_INPUT = font(INPUT_FG)
F_CALC = font(CALC_FG)
F_LINK = font(LINK_FG)
F_SOFT = font(INK_SOFT, size=SIZE_SMALL)
F_TITLE = font(INK, bold=True, size=SIZE_TITLE)
F_SUBTITLE = font(INK_SOFT, size=SIZE_SUBTITLE)
F_SECTION = font(ACCENT, bold=True, size=SIZE_SECTION)
F_HEADER = font(HEADER_FG, bold=True, size=SIZE_SMALL)
F_TOTAL = font(INK, bold=True)

FILL_HEADER = PatternFill("solid", fgColor=HEADER_BG)
FILL_BAND = PatternFill("solid", fgColor=BAND)
FILL_ACCENT = PatternFill("solid", fgColor=ACCENT)
FILL_NONE = PatternFill(fill_type=None)


# --------------------------------------------------------------------------
# 7. 셀 역할 헬퍼 — 빌더는 이걸로 셀 의미(입력/계산/링크)를 표현
# --------------------------------------------------------------------------
def set_cell(ws: Worksheet, row: int, col: int, value=None, *,
             role: str = "calc", number_format: str | None = None,
             align: Alignment | None = None, bold: bool = False,
             border: Border | None = None, fill: PatternFill | None = None,
             size: int = SIZE_BODY):
    """단일 셀에 값+서식 일괄 적용. role ∈ {calc,input,link,label,header,total}.

    재무 관례상 셀의 '역할'을 색으로 구분(입력 파랑·계산 검정·링크 초록).
    """
    cell = ws.cell(row=row, column=col, value=value)
    role_font = {
        "calc": font(CALC_FG, bold=bold, size=size),
        "input": font(INPUT_FG, bold=bold, size=size),
        "link": font(LINK_FG, bold=bold, size=size),
        "label": font(INK, bold=bold, size=size),
        "soft": font(INK_SOFT, bold=bold, size=size),
        "header": F_HEADER,
        "total": font(INK, bold=True, size=size),
    }.get(role, font(CALC_FG, bold=bold, size=size))
    cell.font = role_font
    if role == "header":
        cell.fill = FILL_HEADER
        cell.alignment = align or CENTER
    else:
        cell.alignment = align or (LEFT if role in ("label", "soft") else RIGHT)
    if fill is not None:
        cell.fill = fill
    if number_format:
        cell.number_format = number_format
    if border is not None:
        cell.border = border
    return cell


# --------------------------------------------------------------------------
# 8. 시트 외관 (gridlines off, 여백, freeze, 열폭)
# --------------------------------------------------------------------------
def style_sheet(ws: Worksheet, *, freeze: str | None = "A2",
                zoom: int = 100) -> None:
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = zoom
    if freeze:
        ws.freeze_panes = freeze
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True if ws.sheet_properties.pageSetUpPr else None


def set_widths(ws: Worksheet, widths: dict[int, float]) -> None:
    """{열번호: 폭} 일괄 적용."""
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def title_block(ws: Worksheet, title: str, subtitle: str = "",
                *, row: int = 1, last_col: int = 6) -> int:
    """제목/부제 블록. 다음에 쓸 시작 행을 반환."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    c = ws.cell(row=row, column=1, value=title)
    c.font = F_TITLE
    c.alignment = LEFT
    r = row + 1
    if subtitle:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        c2 = ws.cell(row=r, column=1, value=subtitle)
        c2.font = F_SUBTITLE
        c2.alignment = LEFT
        r += 1
    return r + 1  # 한 줄 띄움


def section_header(ws: Worksheet, row: int, text: str, last_col: int = 6) -> int:
    c = ws.cell(row=row, column=1, value=text)
    c.font = F_SECTION
    c.alignment = LEFT
    for col in range(1, last_col + 1):
        ws.cell(row=row, column=col).border = BORDER_BOTTOM
    return row + 1


def safe_sheet_title(name: str) -> str:
    """Excel 시트명 제약: ≤31자, 금지문자 []:*?/\\ 제거."""
    for ch in "[]:*?/\\":
        name = name.replace(ch, " ")
    name = name.strip() or "Sheet"
    return name[:31]


# --------------------------------------------------------------------------
# 9. 워터폴(variance bridge) — 누적막대 + base 투명 트릭
# --------------------------------------------------------------------------
def add_waterfall(ws: Worksheet, *, anchor: str, data_min_row: int,
                  data_max_row: int, base_col: int, value_col: int,
                  cat_col: int, title: str = "Bridge",
                  height: float = 8.0, width: float = 18.0) -> BarChart:
    """워터폴 차트를 stacked bar 로 그린다.

    레이아웃 전제(빌더가 보조 컬럼을 미리 채워둠):
      cat_col   : 카테고리 라벨
      base_col  : 투명 받침(floating bar 의 바닥 높이)
      value_col : 실제 표시 막대 높이(증감분/기둥)

    base 시리즈는 fill=none + line=none 으로 투명 처리해 막대가 떠 있게 만든다.
    """
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = title
    chart.height = height
    chart.width = width
    chart.gapWidth = 40

    cats = Reference(ws, min_col=cat_col, min_row=data_min_row, max_row=data_max_row)

    base_ref = Reference(ws, min_col=base_col, min_row=data_min_row - 1,
                         max_row=data_max_row)
    val_ref = Reference(ws, min_col=value_col, min_row=data_min_row - 1,
                        max_row=data_max_row)
    chart.add_data(base_ref, titles_from_data=True)
    chart.add_data(val_ref, titles_from_data=True)
    chart.set_categories(cats)

    # series[0] = base → 투명
    base_series: Series = chart.series[0]
    base_series.graphicalProperties.noFill = True
    base_series.graphicalProperties.line.noFill = True

    # series[1] = value → 액센트 채움
    val_series: Series = chart.series[1]
    val_series.graphicalProperties.solidFill = ACCENT

    chart.legend = None
    ws.add_chart(chart, anchor)
    return chart


def add_line_chart(ws: Worksheet, *, anchor: str, data_min_col: int,
                   data_max_col: int, data_min_row: int, data_max_row: int,
                   cat_col: int, title: str = "", height: float = 7.5,
                   width: float = 16.0) -> LineChart:
    """라인 차트(추이/포캐스트). data 영역 1행 위에 시리즈명 가정."""
    ch = LineChart()
    ch.title = title or None
    ch.height = height
    ch.width = width
    data = Reference(ws, min_col=data_min_col, max_col=data_max_col,
                     min_row=data_min_row - 1, max_row=data_max_row)
    cats = Reference(ws, min_col=cat_col, min_row=data_min_row, max_row=data_max_row)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    for s in ch.series:
        s.smooth = False
    ws.add_chart(ch, anchor)
    return ch


def add_bar_chart(ws: Worksheet, *, anchor: str, data_min_col: int,
                  data_max_col: int, data_min_row: int, data_max_row: int,
                  cat_col: int, title: str = "", height: float = 7.5,
                  width: float = 16.0, stacked: bool = False) -> BarChart:
    ch = BarChart()
    ch.type = "col"
    if stacked:
        ch.grouping = "stacked"
        ch.overlap = 100
    ch.title = title or None
    ch.height = height
    ch.width = width
    data = Reference(ws, min_col=data_min_col, max_col=data_max_col,
                     min_row=data_min_row - 1, max_row=data_max_row)
    cats = Reference(ws, min_col=cat_col, min_row=data_min_row, max_row=data_max_row)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ws.add_chart(ch, anchor)
    return ch


def write_matrix(ws: Worksheet, top: int, left: int, headers: list[str],
                 rows: list[tuple], *, value_fmt: str = FMT_INT,
                 label_width: bool = True) -> int:
    """간단 표 작성 헬퍼. headers[0]=라벨열. rows=(label, v1, v2, ...).

    숫자 셀은 value_fmt, 라벨은 좌측. 마지막 행이 합계면 호출측이 보더 처리.
    반환: 표 다음 행번호.
    """
    for j, h in enumerate(headers):
        set_cell(ws, top, left + j, h, role="header",
                 align=LEFT if j == 0 else CENTER)
    r = top + 1
    for row in rows:
        set_cell(ws, r, left, row[0], role="label", align=LEFT)
        for j, v in enumerate(row[1:], start=1):
            is_pct = isinstance(v, float) and -2 < v < 2 and value_fmt == FMT_PCT1
            set_cell(ws, r, left + j, v, role="calc",
                     number_format=value_fmt)
        r += 1
    return r


__all__ = [name for name in dir() if not name.startswith("_")]
