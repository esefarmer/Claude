"""単位認定会議資料 生成スクリプト

使い方:
    python build_report.py 入力フォルダ 出力.xlsx [--club クラブ一覧.xlsx] [--master 科目マスタ.xlsx]

入力フォルダ内の「学期末評定」帳票(.xlsx、1クラス1ファイル)をすべて読み込み、
学年全体の会議資料を1冊のExcelにまとめる。

判定ルール(教務部の指定による):
  - 履修の痕跡(欠時数の入力)があるのに「評定」「観点別」が空欄 → 履修不認定
  - 全科目で評定が空欄の生徒 → 履修不認定または籍のない生徒(要確認)。順位・平均から除外
  - 欠時超過: 欠時数 ≥ 授業時数 × 1/3（1/3以上）
  - 皆勤: 欠席・遅刻・早退がすべて0
  - 登校不調: 欠席10以上 または 遅刻10以上
閾値は出力ブックの「設定」シートで変更でき、数式で全シートに反映される。
"""

import argparse
import csv
import datetime
import io
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

FONT = "游ゴシック"
ATTEND_LABELS = ["授業日数", "出停・忌引日数", "出席すべき日数", "欠席日数", "出席日数", "遅刻", "早退", "学期末評定平均"]
TOP_N = 30
TOP_ROWS = 45          # 同順位で30名を超える場合に備えた表示行数
EXCESS_ROWS = 300      # 欠時超過一覧の最大表示行数
MAX_CLUBS = 60

# ---------------------------------------------------------------- 読み込み


def _num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return int(v) if float(v).is_integer() else float(v)
    s = str(v).strip()
    try:
        f = float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return None


def parse_class_file(path):
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    grade = cls = None
    for r in rows[:15]:
        for v in r:
            m = re.fullmatch(r"\s*(\d+)年(\S+?)組\s*", unicodedata.normalize("NFKC", str(v or "")))
            if m:
                grade, cls = int(m.group(1)), normalize_cls(m.group(2))
                break
        if grade:
            break
    if grade is None:
        raise ValueError(f"{path.name}: 「○年○組」の表記が見つかりません")

    students = {}
    order = []
    header_rows = [i for i, r in enumerate(rows) if len(r) > 1 and r[1] == "科目名"]
    for h in header_rows:
        hdr = rows[h]
        subj_cols = [(c, str(hdr[c]).strip()) for c in range(3, min(61, len(hdr)), 3)
                     if hdr[c] not in (None, "")]
        att_cols = {str(v).strip(): c for c, v in enumerate(hdr) if str(v or "").strip() in ATTEND_LABELS}
        i = h + 1
        while i < len(rows) and rows[i][1] != "No.":
            i += 1
        i += 1
        while i < len(rows):
            r = rows[i]
            if str(r[2] or "").startswith("学期末評定") or r[1] == "科目名":
                break
            no, name = _num(r[1]), r[2]
            if no is not None and name not in (None, ""):
                st = students.get(no)
                if st is None:
                    st = students[no] = {"学年": grade, "組": cls, "番号": no,
                                         "氏名": str(name).strip(), "科目": [], "出欠": {}}
                    order.append(no)
                for c, subj in subj_cols:
                    ev, kan, ab = r[c], r[c + 1], r[c + 2]
                    ev_n, ab_n = _num(ev), _num(ab)
                    kan = str(kan).strip() if kan not in (None, "") else ""
                    if ev_n is None and kan == "" and ab_n is None:
                        continue  # 履修していない科目
                    st["科目"].append({"科目": subj, "評定": ev_n, "観点別": kan, "欠時数": ab_n,
                                       "状態": "認定" if ev_n is not None else "不認定"})
                for label, c in att_cols.items():
                    if label not in st["出欠"] or st["出欠"][label] is None:
                        st["出欠"][label] = _num(r[c])
            i += 1
    result = [students[n] for n in order]
    for st in result:
        st["評定なし"] = not any(s["評定"] is not None for s in st["科目"])
    return result


def normalize_cls(v):
    """組の表記をそろえる: 「1」「1組」「１」→ 1、「FA１」→ "FA1"（数字以外を含む組は文字列のまま）。"""
    c = unicodedata.normalize("NFKC", str(v if v is not None else "")).strip()
    c = re.sub(r"組$", "", c)
    n = _num(c)
    return n if n is not None else c


def cls_order(c):
    return (0, c, "") if isinstance(c, (int, float)) else (1, 0, str(c))


def _read_rows(path):
    path = Path(path)
    if path.suffix.lower() == ".csv":
        text = path.read_bytes()
        for enc in ("utf-8-sig", "cp932"):
            try:
                return list(csv.reader(io.StringIO(text.decode(enc))))
            except UnicodeDecodeError:
                continue
        raise ValueError(f"{path.name}: 文字コードを判別できません")
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def read_club_list(path):
    """クラブ一覧(学年・組・番号・クラブの列を持つ表。.xlsx / .csv)を読む。列名で自動判別する。

    「◎」などの印は取り除く。兼部は「、」「,」区切りで複数行に展開する。
    """
    rows = _read_rows(path)
    # 見出しの候補（先に書いたものを優先）
    alias = {"学年": ["学年", "年"], "組": ["組", "クラス"], "番号": ["番号", "番", "出席番号"],
             "氏名": ["氏名", "名前", "生徒名"],
             "クラブ": ["クラブ", "所属クラブ", "クラブ名", "部活動", "部活", "所属", "部"]}
    for hi, r in enumerate(rows[:20]):
        cells = [str(v or "").strip() for v in r]
        found = {}
        for k, names in alias.items():
            for name in names:
                if name in cells:
                    found[k] = cells.index(name)
                    break
        if {"クラブ", "学年", "組", "番号"} <= found.keys():
            break
    else:
        raise ValueError("クラブ一覧に「学年(年)」「組」「番号(番)」「クラブ(所属)」の見出し行が見つかりません")
    out = []
    for r in rows[hi + 1:]:
        get = lambda k: r[found[k]] if k in found and found[k] < len(r) else None
        club = str(get("クラブ") or "").strip()
        if not club:
            continue
        g = _num(unicodedata.normalize("NFKC", str(get("学年") or "")).replace("年", ""))
        c = normalize_cls(get("組"))
        n = _num(unicodedata.normalize("NFKC", str(get("番号") or "")))
        for cl in re.split(r"[、,，]", club):
            cl = cl.strip().lstrip("◎○●☆★◇◆").strip()
            if cl:
                out.append({"学年": g, "組": c, "番号": n, "氏名": str(get("氏名") or "").strip(), "クラブ": cl})
    return out


def read_master(path):
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    master = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r and r[0] not in (None, "") and len(r) > 1 and _num(r[1]) is not None:
            master[str(r[0]).strip()] = _num(r[1])
    return master

# ---------------------------------------------------------------- 書式


THIN = Side(style="thin", color="808080")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
NOTE_FONT = Font(name=FONT, size=9, color="595959")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def f(size=10, bold=False, color="000000"):
    return Font(name=FONT, size=size, bold=bold, color=color)


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = f(14, True)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = NOTE_FONT


def header(ws, row, labels, widths=None):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=lab)
        c.font = f(10, True)
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = CENTER
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


def body(ws, r1, r2, ncol, fmt=None):
    for r in range(r1, r2 + 1):
        for c in range(1, ncol + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.font = f()
            if fmt and c in fmt:
                cell.number_format = fmt[c]
            if cell.alignment.horizontal is None:
                cell.alignment = Alignment(vertical="center")


def live_rows(ws, r1, r2, ncol, key_col, count_ref, fmt=None):
    """数式で行が埋まる一覧: 罫線は値のある行だけ（条件付き書式）、印刷範囲は該当件数に合わせて伸縮。"""
    body(ws, r1, r2, ncol, fmt)
    for r in range(r1, r2 + 1):
        for c in range(1, ncol + 1):
            ws.cell(row=r, column=c).border = Border()
    rng = f"A{r1}:{get_column_letter(ncol)}{r2}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'${key_col}{r1}<>""'], border=BORDER))
    q = f"'{ws.title}'!"
    ws.defined_names["_xlnm.Print_Area"] = DefinedName(
        "_xlnm.Print_Area", localSheetId=ws.parent.index(ws),
        attr_text=f"OFFSET({q}$A$1,0,0,MAX({r1 - 1},{r1 - 1}+{q}{count_ref}),{ncol})")


def page_setup(ws, header_row, landscape=True):
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{header_row}:{header_row}"
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    ws.oddFooter.center.text = "&P / &N"
    ws.oddHeader.right.text = "取扱注意（個人情報）"

# ---------------------------------------------------------------- ブック作成

# 生徒一覧の列
S = dict(学年="A", 組="B", 番号="C", 氏名="D", キー="E", 授業日数="F", 出停忌引="G", 出席すべき日数="H",
         欠席="I", 出席="J", 遅刻="K", 早退="L", 評定合計="M", 評定科目数="N", 評定平均="O",
         不認定="P", 皆勤="Q", 登校不調="R", 状態="S", 元平均="T", 順位="U", 順位キー="V",
         皆勤連番="W", 不調連番="X")
# 履修明細の列
R = dict(学年="A", 組="B", 番号="C", 氏名="D", キー="E", 科目="F", 評定="G", 観点別="H", 欠時数="I",
         状態="J", 授業時数="K", 上限="L", 判定="M", 超過連番="N")


def build(students, out_path, clubs=None, master=None, sources=()):
    master = master or {}
    wb = openpyxl.Workbook()
    ws_sum = wb.active
    ws_sum.title = "概要"
    names = ["履修不認定一覧", "欠時超過一覧", "評定平均上位30名", "皆勤者一覧", "登校不調者一覧",
             "クラブ別評定平均", "科目別評定分布", "クラス別比較", "クラブ所属", "生徒一覧", "履修明細",
             "科目マスタ", "設定"]
    sh = {n: wb.create_sheet(n) for n in names}
    grades = sorted({s["学年"] for s in students})
    classes = sorted({(s["学年"], s["組"]) for s in students}, key=lambda k: (k[0], cls_order(k[1])))
    subjects = []
    for s in students:
        for x in s["科目"]:
            if x["科目"] not in subjects:
                subjects.append(x["科目"])
    # 在籍生徒（評定のある生徒）が3名以上履修し、誰にも評定がない科目は「評定を付けない科目」とみなす
    enrolled, graded = {}, {}
    for s in students:
        if s["評定なし"]:
            continue
        for x in s["科目"]:
            enrolled[x["科目"]] = enrolled.get(x["科目"], 0) + 1
            graded[x["科目"]] = graded.get(x["科目"], 0) + (x["評定"] is not None)
    no_grade_subj = {k for k, n in enrolled.items() if n >= 3 and graded[k] == 0}

    # ---- 設定
    ws = sh["設定"]
    title(ws, "設定（判定基準）", "黄色のセルを変更すると、すべてのシートの判定・一覧に反映されます。")
    settings = [("欠時超過とする割合（授業時数に対して）", 1 / 3, "欠時数がこの割合以上なら欠時超過。教務部指定: 1/3以上"),
                ("登校不調：欠席日数の基準（以上）", 10, "教務部指定"),
                ("登校不調：遅刻回数の基準（以上）", 10, "教務部指定"),
                ("評定平均上位の人数", TOP_N, "同順位は全員掲載"),
                ("皆勤の定義", "欠席・遅刻・早退がすべて0", "変更する場合は生徒一覧Q列の数式を修正")]
    header(ws, 4, ["項目", "値", "備考"], [40, 16, 50])
    for i, (k, v, note) in enumerate(settings, 5):
        ws.cell(row=i, column=1, value=k)
        c = ws.cell(row=i, column=2, value=v)
        ws.cell(row=i, column=3, value=note)
        if i < 9:
            c.fill = INPUT_FILL
            c.font = f(color="0000FF")
    body(ws, 5, 9, 3)
    for i in range(5, 8):
        ws.cell(row=i, column=2).fill = INPUT_FILL
        ws.cell(row=i, column=2).font = f(color="0000FF")
    ws.cell(row=8, column=2).fill = INPUT_FILL
    ws["B5"].number_format = "0.000"
    RATIO, ABS_T, LATE_T, TOPN = "設定!$B$5", "設定!$B$6", "設定!$B$7", "設定!$B$8"

    # ---- 科目マスタ
    ws = sh["科目マスタ"]
    title(ws, "科目マスタ（授業時数）",
          "黄色の「授業時数」を入力すると、欠時超過の判定が自動で行われます（上限＝授業時数×設定!B5）。")
    header(ws, 3, ["科目名", "授業時数", "欠時上限（これ以上で欠時超過）", "履修者数", "欠時数の最大", "評定の扱い"],
           [26, 12, 22, 12, 14, 14])
    dv = DataValidation(type="list", formula1='"評定あり,評定なし"', allow_blank=False)
    ws.add_data_validation(dv)
    for i, subj in enumerate(subjects, 4):
        ws.cell(row=i, column=1, value=subj)
        c = ws.cell(row=i, column=2, value=master.get(subj))
        c.fill = INPUT_FILL
        c.font = f(color="0000FF")
        ws.cell(row=i, column=3, value=f'=IF(B{i}="","",B{i}*{RATIO})')
        ws.cell(row=i, column=4, value=f"=COUNTIF(履修明細!$F:$F,A{i})")
        ws.cell(row=i, column=5, value=f"=IFERROR(_xlfn.MAXIFS(履修明細!$I:$I,履修明細!$F:$F,A{i}),0)")
        ws.cell(row=i, column=6, value="評定なし" if subj in no_grade_subj else "評定あり")
        dv.add(f"F{i}")
    last_m = 3 + len(subjects)
    body(ws, 4, last_m, 6, {3: "0.0"})
    for i in range(4, last_m + 1):
        for col in (2, 6):
            ws.cell(row=i, column=col).fill = INPUT_FILL
            ws.cell(row=i, column=col).font = f(color="0000FF")
    ws.cell(row=last_m + 2, column=1, value=(
        "※「評定の扱い」＝評定なし の科目（総合的な探究の時間・ホームルーム等）は、評定が空欄でも履修不認定としません"
        "（欠時超過の判定は行います）。初期値は、在籍生徒3名以上が履修し誰にも評定がない科目を「評定なし」としています。")).font = NOTE_FONT
    page_setup(ws, 3, landscape=False)

    # ---- 履修明細
    ws = sh["履修明細"]
    title(ws, "履修明細（生徒×科目）", "評定・観点別が空欄で欠時数が入っている履修＝履修不認定。授業時数は科目マスタから参照。")
    header(ws, 3, ["学年", "組", "番号", "氏名", "生徒キー", "科目", "評定", "観点別", "欠時数", "状態",
                   "授業時数", "欠時上限", "欠時判定", "超過連番", "生徒区分", "不認定連番"],
           [6, 5, 6, 20, 10, 22, 6, 8, 8, 12, 9, 9, 12, 9, 12, 9])
    r = 4
    for s in students:
        key = f'{s["学年"]}-{s["組"]}-{s["番号"]}'
        for x in s["科目"]:
            vals = [s["学年"], s["組"], s["番号"], s["氏名"], key, x["科目"], x["評定"], x["観点別"],
                    x["欠時数"], None]
            for c, v in enumerate(vals, 1):
                ws.cell(row=r, column=c, value=v)
            ws.cell(row=r, column=11, value=f'=IFERROR(IF(INDEX(科目マスタ!$B:$B,MATCH(F{r},科目マスタ!$A:$A,0))="","",'
                                            f'INDEX(科目マスタ!$B:$B,MATCH(F{r},科目マスタ!$A:$A,0))),"")')
            ws.cell(row=r, column=12, value=f'=IF(K{r}="","",K{r}*{RATIO})')
            ws.cell(row=r, column=13, value=f'=IF(K{r}="","時数未入力",IF(N(I{r})>=L{r},"欠時超過",""))')
            ws.cell(row=r, column=14, value=f'=IF(M{r}="欠時超過",COUNTIF(M$4:M{r},"欠時超過"),"")')
            ws.cell(row=r, column=10, value=f'=IF(G{r}<>"","認定",IF(IFERROR(INDEX(科目マスタ!$F:$F,'
                                            f'MATCH(F{r},科目マスタ!$A:$A,0)),"")="評定なし","評定なし科目","不認定"))')
            ws.cell(row=r, column=15, value="評定なし生徒" if s["評定なし"] else "")
            ws.cell(row=r, column=16, value=f'=IF(AND(J{r}="不認定",O{r}=""),'
                                            f'COUNTIFS(J$4:J{r},"不認定",O$4:O{r},""),"")')
            r += 1
    last_r = r - 1
    body(ws, 4, last_r, 16, {12: "0.0"})
    ws.auto_filter.ref = f"A3:P{last_r}"
    page_setup(ws, 3)

    # ---- 生徒一覧
    ws = sh["生徒一覧"]
    title(ws, "生徒一覧（出欠・評定平均値）", "評定平均値＝評定のついた科目の評定の平均（小数第4位で丸め）。")
    header(ws, 3, ["学年", "組", "番号", "氏名", "生徒キー", "授業日数", "出停・忌引", "出席すべき日数", "欠席",
                   "出席", "遅刻", "早退", "評定合計", "評定科目数", "評定平均値", "不認定科目数", "皆勤",
                   "登校不調", "状態", "帳票の評定平均", "学年内順位", "順位キー", "皆勤連番", "不調連番"],
           [6, 5, 6, 20, 10, 8, 8, 9, 6, 6, 6, 6, 8, 8, 9, 9, 6, 8, 26, 10, 8, 10, 8, 8])
    first_s = 4
    for i, s in enumerate(students, first_s):
        a = s["出欠"]
        key = f'{s["学年"]}-{s["組"]}-{s["番号"]}'
        vals = [s["学年"], s["組"], s["番号"], s["氏名"], key, a.get("授業日数"), a.get("出停・忌引日数"),
                a.get("出席すべき日数"), a.get("欠席日数"), a.get("出席日数"), a.get("遅刻"), a.get("早退")]
        for c, v in enumerate(vals, 1):
            ws.cell(row=i, column=c, value=v)
        ws[f"M{i}"] = f"=SUMIFS(履修明細!$G:$G,履修明細!$E:$E,E{i})"
        ws[f"N{i}"] = f'=COUNTIFS(履修明細!$E:$E,E{i},履修明細!$J:$J,"認定")'
        ws[f"O{i}"] = f'=IF(N{i}=0,"",ROUND(M{i}/N{i},4))'
        ws[f"P{i}"] = f'=COUNTIFS(履修明細!$E:$E,E{i},履修明細!$J:$J,"不認定")'
        ws[f"Q{i}"] = f'=IF(AND(N{i}>0,I{i}=0,K{i}=0,L{i}=0),"○","")'
        ws[f"R{i}"] = f'=IF(OR(I{i}>={ABS_T},K{i}>={LATE_T}),"○","")'
        ws[f"S{i}"] = "評定なし（履修不認定または籍なし・要確認）" if s["評定なし"] else ""
        ws[f"T{i}"] = s["出欠"].get("学期末評定平均")
        ws[f"U{i}"] = f'=IF(O{i}="","",COUNTIFS($A:$A,A{i},$O:$O,">"&O{i})+1)'
        ws[f"V{i}"] = (f'=IF(O{i}="","",A{i}&"-"&(U{i}+COUNTIFS($A${first_s}:A{i},A{i},'
                       f'$O${first_s}:O{i},O{i})-1))')
        ws[f"W{i}"] = f'=IF(Q{i}="○",COUNTIF($Q${first_s}:Q{i},"○"),"")'
        ws[f"X{i}"] = f'=IF(R{i}="○",COUNTIF($R${first_s}:R{i},"○"),"")'
    last_s = first_s + len(students) - 1
    body(ws, first_s, last_s, 24, {15: "0.00", 20: "0.00"})
    for c in "QR":
        for i in range(first_s, last_s + 1):
            ws[f"{c}{i}"].alignment = CENTER
    ws.auto_filter.ref = f"A3:X{last_s}"
    page_setup(ws, 3)

    def pick(col_letter, match_expr, fallback='""'):
        return f'IFERROR(INDEX(生徒一覧!${col_letter}:${col_letter},{match_expr}),{fallback})'

    # ---- 評定平均上位30名（学年ごと）
    ws = sh["評定平均上位30名"]
    title(ws, "評定平均値 上位30名（学年別）",
          "評定のない生徒（履修不認定・籍なし）は除外。同順位は全員掲載。「不認定科目」は不認定の科目数。")
    widths = [6, 6, 5, 6, 20, 10, 8, 8, 10]
    labels = ["順位", "学年", "組", "番号", "氏名", "評定平均値", "評定合計", "科目数", "不認定科目"]
    row = 4
    for g in grades:
        ws.cell(row=row, column=1, value=f"{g}年").font = f(12, True)
        row += 1
        header(ws, row, labels, widths)
        hdr = row
        for k in range(1, TOP_ROWS + 1):
            r = hdr + k
            m = f'MATCH("{g}-{k}",生徒一覧!$V:$V,0)'
            ws[f"A{r}"] = f'=IFERROR(IF({pick("U", m)}<={TOPN},{pick("U", m)},""),"")'
            ws[f"B{r}"] = f'=IF($A{r}="","",{pick("A", m)})'
            ws[f"C{r}"] = f'=IF($A{r}="","",{pick("B", m)})'
            ws[f"D{r}"] = f'=IF($A{r}="","",{pick("C", m)})'
            ws[f"E{r}"] = f'=IF($A{r}="","",{pick("D", m)})'
            ws[f"F{r}"] = f'=IF($A{r}="","",{pick("O", m)})'
            ws[f"G{r}"] = f'=IF($A{r}="","",{pick("M", m)})'
            ws[f"H{r}"] = f'=IF($A{r}="","",{pick("N", m)})'
            ws[f"I{r}"] = f'=IF($A{r}="","",IF({pick("P", m)}=0,"",{pick("P", m)}))'
        body(ws, hdr + 1, hdr + TOP_ROWS, 9, {6: "0.00"})
        row = hdr + TOP_ROWS + 2
    page_setup(ws, 5, landscape=False)
    ws.print_title_rows = None

    # ---- 皆勤者一覧 / 登校不調者一覧
    def live_list(ws, ttl, sub, seq_col, cols, count_formula):
        title(ws, ttl, sub)
        ws["A3"] = "該当者数"
        ws["B3"] = count_formula
        ws["A3"].font = f(10, True)
        ws["B3"].font = f(10, True)
        header(ws, 4, ["No."] + [c[0] for c in cols], [5] + [c[2] for c in cols])
        n = len(students)
        for k in range(1, n + 1):
            r = 4 + k
            m = f"MATCH({k},生徒一覧!${seq_col}:${seq_col},0)"
            ws[f"A{r}"] = f'=IF({k}<=$B$3,{k},"")'
            for j, (_, src, _w) in enumerate(cols, 2):
                v = pick(src, m) if src not in ("D", "S") else f'IFERROR(INDEX(生徒一覧!${src}:${src},{m})&"","")'
                ws.cell(row=r, column=j, value=f'=IF($A{r}="","",{v})')
        fmt = {j: "0.00" for j, c in enumerate(cols, 2) if c[1] == "O"}
        live_rows(ws, 5, 4 + n, len(cols) + 1, "A", "$B$3", fmt)
        page_setup(ws, 4, landscape=False)

    live_list(sh["皆勤者一覧"], "皆勤者一覧（評定平均値付き）",
              "皆勤＝欠席・遅刻・早退がすべて0（評定のない生徒を除く）。学年・組・番号順。", "W",
              [("学年", "A", 6), ("組", "B", 5), ("番号", "C", 6), ("氏名", "D", 20), ("欠席", "I", 6),
               ("遅刻", "K", 6), ("早退", "L", 6), ("評定平均値", "O", 10)],
              '=COUNTIF(生徒一覧!$Q:$Q,"○")')
    ws = sh["登校不調者一覧"]
    live_list(ws, "登校不調者一覧",
              "欠席日数が設定値以上、または遅刻回数が設定値以上の生徒（設定シート参照）。学年・組・番号順。", "X",
              [("学年", "A", 6), ("組", "B", 5), ("番号", "C", 6), ("氏名", "D", 20),
               ("出席すべき日数", "H", 9), ("欠席", "I", 6), ("遅刻", "K", 6), ("早退", "L", 6),
               ("評定平均値", "O", 10), ("不認定科目数", "P", 9), ("状態", "S", 30)],
              '=COUNTIF(生徒一覧!$R:$R,"○")')
    # 該当理由の列を追加
    ws["M4"] = "該当理由"
    ws["M4"].font = f(10, True)
    ws["M4"].fill = HEAD_FILL
    ws["M4"].border = BORDER
    ws["M4"].alignment = CENTER
    ws.column_dimensions["M"].width = 16
    for k in range(1, len(students) + 1):
        r = 4 + k
        ws[f"M{r}"] = (f'=IF(A{r}="","",IF(AND(G{r}>={ABS_T},H{r}>={LATE_T}),"欠席・遅刻",'
                       f'IF(G{r}>={ABS_T},"欠席","遅刻")))')
        ws[f"M{r}"].font = f()
    live_rows(ws, 5, 4 + len(students), 13, "A", "$B$3", {10: "0.00"})

    # ---- 欠時超過一覧
    ws = sh["欠時超過一覧"]
    title(ws, "欠時超過一覧（欠時数が授業時数の1/3以上の履修）",
          "授業時数は「科目マスタ」で入力。学年・組・番号順。")
    ws["A3"] = "該当件数"
    ws["B3"] = '=COUNTIF(履修明細!$M:$M,"欠時超過")'
    ws["D3"] = '=IF(COUNTBLANK(科目マスタ!$B$4:$B$' + str(last_m) + ')>0,"※科目マスタで授業時数が未入力の科目が"&' \
               'COUNTBLANK(科目マスタ!$B$4:$B$' + str(last_m) + ')&"件あります（未入力の科目は判定されません）","")'
    ws["D3"].font = f(10, True, "C00000")
    for c in ("A3", "B3"):
        ws[c].font = f(10, True)
    cols = [("学年", "A", 6), ("組", "B", 5), ("番号", "C", 6), ("氏名", "D", 20), ("科目", "F", 22),
            ("評定", "G", 6), ("欠時数", "I", 8), ("授業時数", "K", 9), ("欠時上限", "L", 9)]
    header(ws, 4, ["No."] + [c[0] for c in cols], [5] + [c[2] for c in cols])
    for k in range(1, EXCESS_ROWS + 1):
        r = 4 + k
        m = f"MATCH({k},履修明細!$N:$N,0)"
        ws[f"A{r}"] = f'=IF({k}<=$B$3,{k},"")'
        for j, (_, src, _w) in enumerate(cols, 2):
            ws.cell(row=r, column=j,
                    value=f'=IF($A{r}="","",IFERROR(INDEX(履修明細!${src}:${src},{m})&"",""))'
                    if src in ("D", "F", "G") else
                    f'=IF($A{r}="","",IFERROR(INDEX(履修明細!${src}:${src},{m}),""))')
    live_rows(ws, 5, 4 + EXCESS_ROWS, 10, "A", "$B$3", {10: "0.0"})
    page_setup(ws, 4, landscape=False)

    # ---- 履修不認定一覧（帳票の内容から確定するため値で出力）
    ws = sh["履修不認定一覧"]
    title(ws, "履修不認定一覧", "評定・観点別が空欄の履修。A: 全科目で評定がない生徒（籍の有無を確認）／B: 科目ごとの不認定。")
    ws["A4"] = "A. 全科目で評定のない生徒（履修不認定または籍なし・要確認）"
    ws["A4"].font = f(11, True)
    header(ws, 5, ["学年", "組", "番号", "氏名", "履修科目数", "欠席日数", "出席すべき日数", "確認欄"],
           [6, 5, 6, 20, 22, 9, 9, 18])
    r = 6
    none_st = [s for s in students if s["評定なし"]]
    for s in none_st:
        for c, v in enumerate([s["学年"], s["組"], s["番号"], s["氏名"], len(s["科目"]),
                               s["出欠"].get("欠席日数"), s["出欠"].get("出席すべき日数"), ""], 1):
            ws.cell(row=r, column=c, value=v)
        r += 1
    if not none_st:
        ws.cell(row=r, column=1, value="該当なし")
        r += 1
    body(ws, 6, r - 1, 8)
    r += 1
    ws.cell(row=r, column=1, value="B. 科目ごとの履修不認定（Aの生徒を除く）").font = f(11, True)
    r += 1
    ws.cell(row=r - 1, column=5, value="該当件数").font = f(10, True)
    ws.cell(row=r - 1, column=6, value='=COUNT(履修明細!$P:$P)').font = f(10, True)
    header(ws, r, ["学年", "組", "番号", "氏名", "科目", "欠時数", "欠席日数", "確認欄"])
    hb = r
    est = sum(1 for st in students if not st["評定なし"] for x in st["科目"]
              if x["評定"] is None and x["科目"] not in no_grade_subj)
    nrow = est + 30
    cnt = f"$F${hb - 1}"
    for k in range(1, nrow + 1):
        rr = hb + k
        m = f"MATCH({k},履修明細!$P:$P,0)"
        for j, src in enumerate(["A", "B", "C", "D", "F", "I"], 1):
            v = f'INDEX(履修明細!${src}:${src},{m})' + ('&""' if src in ("D", "F") else "")
            ws.cell(row=rr, column=j, value=f'=IF({k}>{cnt},"",IFERROR({v},""))')
        ws.cell(row=rr, column=7, value=f'=IF(C{rr}="","",IFERROR(INDEX(生徒一覧!$I:$I,'
                                        f'MATCH(A{rr}&"-"&B{rr}&"-"&C{rr},生徒一覧!$E:$E,0)),""))')
    live_rows(ws, hb + 1, hb + nrow, 8, "A", cnt)
    page_setup(ws, 5, landscape=False)
    ws.print_title_rows = None
    ws.freeze_panes = None

    # ---- クラブ所属（入力・照合）
    ws = sh["クラブ所属"]
    title(ws, "クラブ所属（別資料から転記）",
          "黄色の列（学年・組・番号・クラブ）を入力。組・番号は数字のみ。兼部は1行ずつ。氏名照合が「該当なし」の行は要確認。")
    header(ws, 3, ["学年", "組", "番号", "氏名（名簿）", "クラブ", "生徒キー", "氏名照合", "評定平均値", "皆勤",
                   "登校不調", "不認定科目数"], [6, 5, 6, 18, 18, 10, 18, 10, 6, 8, 10])
    clubs = clubs or []
    n_rows = max(len(clubs), 1) + 200
    for i in range(4, 4 + n_rows):
        c = clubs[i - 4] if i - 4 < len(clubs) else None
        if c:
            for col, k in enumerate(["学年", "組", "番号", "氏名", "クラブ"], 1):
                ws.cell(row=i, column=col, value=c[k])
        elif i == 4:
            pass
        ws[f"F{i}"] = f'=IF(OR(A{i}="",B{i}="",C{i}=""),"",A{i}&"-"&B{i}&"-"&C{i})'
        ws[f"G{i}"] = f'=IF(F{i}="","",IFERROR(INDEX(生徒一覧!$D:$D,MATCH(F{i},生徒一覧!$E:$E,0)),"該当なし"))'
        ws[f"H{i}"] = f'=IF(F{i}="","",IFERROR(INDEX(生徒一覧!$O:$O,MATCH(F{i},生徒一覧!$E:$E,0)),""))'
        ws[f"I{i}"] = f'=IF(F{i}="","",IFERROR(INDEX(生徒一覧!$Q:$Q,MATCH(F{i},生徒一覧!$E:$E,0)),""))'
        ws[f"J{i}"] = f'=IF(F{i}="","",IFERROR(INDEX(生徒一覧!$R:$R,MATCH(F{i},生徒一覧!$E:$E,0)),""))'
        ws[f"K{i}"] = f'=IF(F{i}="","",IFERROR(INDEX(生徒一覧!$P:$P,MATCH(F{i},生徒一覧!$E:$E,0)),""))'
    body(ws, 4, 3 + n_rows, 11, {8: "0.00"})
    for i in range(4, 4 + n_rows):
        for col in (1, 2, 3, 5):
            ws.cell(row=i, column=col).fill = INPUT_FILL
            ws.cell(row=i, column=col).font = f(color="0000FF")
    page_setup(ws, 3, landscape=False)

    # ---- クラブ別評定平均
    ws = sh["クラブ別評定平均"]
    title(ws, "クラブ別 評定平均値一覧",
          "クラブ所属シートから集計。平均は評定のある所属生徒のみ。クラブ名（黄色）は追加・修正できます。")
    header(ws, 3, ["クラブ", "所属人数", "評定平均値（平均）", "最高", "最低", "学年平均との差", "皆勤者数",
                   "登校不調者数", "不認定科目のある生徒"], [20, 9, 12, 8, 8, 12, 9, 10, 12])
    club_names = []
    for c in clubs:
        if c["クラブ"] not in club_names:
            club_names.append(c["クラブ"])
    CR = "クラブ所属!"
    for i in range(4, 4 + MAX_CLUBS):
        nm = club_names[i - 4] if i - 4 < len(club_names) else None
        ws[f"A{i}"] = nm
        ws[f"A{i}"].fill = INPUT_FILL
        ws[f"B{i}"] = f'=IF(A{i}="","",COUNTIF({CR}$E:$E,A{i}))'
        ws[f"C{i}"] = f'=IF(A{i}="","",IFERROR(AVERAGEIFS({CR}$H:$H,{CR}$E:$E,A{i}),""))'
        ws[f"D{i}"] = f'=IF(C{i}="","",_xlfn.MAXIFS({CR}$H:$H,{CR}$E:$E,A{i}))'
        ws[f"E{i}"] = f'=IF(C{i}="","",_xlfn.MINIFS({CR}$H:$H,{CR}$E:$E,A{i}))'
        ws[f"F{i}"] = f'=IF(C{i}="","",C{i}-$C${4 + MAX_CLUBS + 2})'
        ws[f"G{i}"] = f'=IF(A{i}="","",COUNTIFS({CR}$E:$E,A{i},{CR}$I:$I,"○"))'
        ws[f"H{i}"] = f'=IF(A{i}="","",COUNTIFS({CR}$E:$E,A{i},{CR}$J:$J,"○"))'
        ws[f"I{i}"] = f'=IF(A{i}="","",COUNTIFS({CR}$E:$E,A{i},{CR}$K:$K,">0"))'
    tot = 4 + MAX_CLUBS + 1
    ws[f"A{tot}"] = "クラブ所属者全体"
    ws[f"B{tot}"] = f'=COUNTIF({CR}$F:$F,"?*")'
    ws[f"C{tot}"] = f'=IFERROR(AVERAGE({CR}$H:$H),"")'
    ws[f"A{tot + 1}"] = "（参考）学年全体"
    ws[f"B{tot + 1}"] = f'=COUNT(生徒一覧!$O:$O)'
    ws[f"C{tot + 1}"] = "=IFERROR(AVERAGE(生徒一覧!$O:$O),\"\")"
    body(ws, 4, tot + 1, 9, {3: "0.00", 4: "0.00", 5: "0.00", 6: "+0.00;-0.00;0.00"})
    for i in range(4, 4 + MAX_CLUBS):
        ws[f"A{i}"].fill = INPUT_FILL
    for c in ("A", "B", "C"):
        for rr in (tot, tot + 1):
            ws[f"{c}{rr}"].font = f(10, True)
    ws[f"A{tot + 3}"] = ("※所属人数は兼部を含む延べ人数。「学年全体」は評定のある全生徒の平均。"
                         + ("" if clubs else "クラブ一覧が未取り込みのため空欄です。"))
    ws[f"A{tot + 3}"].font = NOTE_FONT
    page_setup(ws, 3, landscape=False)

    # ---- 科目別評定分布
    ws = sh["科目別評定分布"]
    title(ws, "科目別 評定分布", "学年全体。割合は評定のついた人数に対する割合。")
    header(ws, 3, ["科目", "履修者数", "認定", "不認定", "評定5", "評定4", "評定3", "評定2", "評定1", "評定平均",
                   "5の割合", "1・2の割合", "授業時数", "欠時超過"],
           [22, 8, 7, 7, 7, 7, 7, 7, 7, 9, 9, 9, 9, 9])
    RM = "履修明細!"
    for i, subj in enumerate(subjects, 4):
        ws[f"A{i}"] = subj
        ws[f"B{i}"] = f"=COUNTIF({RM}$F:$F,A{i})"
        ws[f"C{i}"] = f'=COUNTIFS({RM}$F:$F,A{i},{RM}$J:$J,"認定")'
        ws[f"D{i}"] = f'=COUNTIFS({RM}$F:$F,A{i},{RM}$J:$J,"不認定")'
        for j, g in enumerate([5, 4, 3, 2, 1]):
            col = get_column_letter(5 + j)
            ws[f"{col}{i}"] = f"=COUNTIFS({RM}$F:$F,$A{i},{RM}$G:$G,{g})"
        ws[f"J{i}"] = f'=IFERROR(AVERAGEIFS({RM}$G:$G,{RM}$F:$F,A{i}),"")'
        ws[f"K{i}"] = f'=IF(C{i}=0,"",E{i}/C{i})'
        ws[f"L{i}"] = f'=IF(C{i}=0,"",(H{i}+I{i})/C{i})'
        ws[f"M{i}"] = f'=IFERROR(INDEX(科目マスタ!$B:$B,MATCH(A{i},科目マスタ!$A:$A,0))&"","")'
        ws[f"N{i}"] = f'=COUNTIFS({RM}$F:$F,A{i},{RM}$M:$M,"欠時超過")'
    last_sub = 3 + len(subjects)
    body(ws, 4, last_sub, 14, {10: "0.00", 11: "0%", 12: "0%"})
    page_setup(ws, 3)

    # ---- クラス別比較
    ws = sh["クラス別比較"]
    title(ws, "クラス別比較")
    header(ws, 3, ["学年", "組", "在籍（名簿）", "評定のある生徒", "評定平均値（平均）", "不認定件数",
                   "評定なし生徒", "欠時超過件数", "皆勤者", "登校不調者", "欠席10日以上"],
           [6, 5, 9, 9, 12, 9, 9, 9, 8, 9, 9])
    SL = "生徒一覧!"
    for i, (g, c) in enumerate(classes, 4):
        ws[f"A{i}"], ws[f"B{i}"] = g, c
        cond = f"{SL}$A:$A,A{i},{SL}$B:$B,B{i}"
        ws[f"C{i}"] = f"=COUNTIFS({cond})"
        ws[f"D{i}"] = f'=COUNTIFS({cond},{SL}$N:$N,">0")'
        ws[f"E{i}"] = f'=IFERROR(AVERAGEIFS({SL}$O:$O,{cond}),"")'
        ws[f"F{i}"] = f'=COUNTIFS({RM}$A:$A,A{i},{RM}$B:$B,B{i},{RM}$J:$J,"不認定")'
        ws[f"G{i}"] = f'=COUNTIFS({cond},{SL}$S:$S,"評定なし*")'
        ws[f"H{i}"] = f'=COUNTIFS({RM}$A:$A,A{i},{RM}$B:$B,B{i},{RM}$M:$M,"欠時超過")'
        ws[f"I{i}"] = f'=COUNTIFS({cond},{SL}$Q:$Q,"○")'
        ws[f"J{i}"] = f'=COUNTIFS({cond},{SL}$R:$R,"○")'
        ws[f"K{i}"] = f'=COUNTIFS({cond},{SL}$I:$I,">="&10)'
    lc = 3 + len(classes)
    t = lc + 1
    ws[f"A{t}"] = "合計"
    for col in "CDFGHIJK":
        ws[f"{col}{t}"] = f"=SUM({col}4:{col}{lc})"
    ws[f"E{t}"] = f'=IFERROR(AVERAGE({SL}$O:$O),"")'
    body(ws, 4, t, 11, {5: "0.00"})
    for col in "ABCDEFGHIJK":
        ws[f"{col}{t}"].font = f(10, True)
    page_setup(ws, 3)

    # ---- 概要
    ws = ws_sum
    today = datetime.date.today()
    title(ws, "単位認定会議資料", f"作成日：{today.year}年{today.month}月{today.day}日　／　取扱注意（個人情報を含む）")
    ws["A4"] = "対象"
    ws["B4"] = "、".join(f"{g}年{c}組" for g, c in classes)
    items = [
        ("在籍者数（名簿上）", f"=COUNTA({SL}$E$4:$E${last_s})", "人", "生徒一覧"),
        ("評定のない生徒（履修不認定または籍なし・要確認）", f'=COUNTIF({SL}$S:$S,"評定なし*")', "人", "履修不認定一覧 A"),
        ("履修不認定（科目数・延べ）", f'=COUNTIF({RM}$J:$J,"不認定")', "件", "履修不認定一覧"),
        ("欠時超過（授業時数の1/3以上・延べ）", '=COUNTIF(履修明細!$M:$M,"欠時超過")', "件", "欠時超過一覧"),
        ("授業時数が未入力の科目", f"=COUNTBLANK(科目マスタ!$B$4:$B${last_m})", "科目", "科目マスタ"),
        ("評定平均値（学年全体の平均）", f'=IFERROR(AVERAGE({SL}$O:$O),"")', "", "生徒一覧"),
        ("皆勤者", f'=COUNTIF({SL}$Q:$Q,"○")', "人", "皆勤者一覧"),
        ("登校不調者（欠席10以上または遅刻10以上）", f'=COUNTIF({SL}$R:$R,"○")', "人", "登校不調者一覧"),
        ("クラブ所属（延べ）", f'=COUNTIF(クラブ所属!$F:$F,"?*")', "人", "クラブ別評定平均"),
    ]
    header(ws, 6, ["項目", "値", "単位", "詳細シート"], [46, 12, 6, 20])
    for i, (k, v, u, s) in enumerate(items, 7):
        ws.cell(row=i, column=1, value=k)
        ws.cell(row=i, column=2, value=v)
        ws.cell(row=i, column=3, value=u)
        ws.cell(row=i, column=4, value=s)
    body(ws, 7, 6 + len(items), 4)
    ws["B12"].number_format = "0.00"
    notes = [
        "判定基準",
        "・履修不認定：履修している（欠時数の入力がある）が、評定・観点別が空欄のもの"
        "（評定を付けない科目＝科目マスタで「評定なし」の科目を除く）。",
        "・欠時超過：欠時数が授業時数の1/3以上のもの（授業時数は科目マスタに入力。割合は設定シートで変更可）。",
        "・評定平均値：評定のついた科目の評定の単純平均。評定のない生徒は順位・平均から除外。",
        "・皆勤：欠席・遅刻・早退がすべて0。　・登校不調：欠席10日以上 または 遅刻10回以上。",
        "資料の構成：概要 → 履修不認定一覧 → 欠時超過一覧 → 評定平均上位30名 → 皆勤者一覧 → 登校不調者一覧 → "
        "クラブ別評定平均 → 科目別評定分布 → クラス別比較（以降は基礎データ）",
        "元データ：" + "、".join(sources),
    ]
    for i, n in enumerate(notes, 8 + len(items)):
        ws.cell(row=i, column=1, value=n).font = f(10, i == 8 + len(items))
    page_setup(ws, 6, landscape=False)
    ws.print_title_rows = None
    ws.freeze_panes = None

    # 列全体参照（$A:$A）をデータ範囲に限定する（再計算の高速化・見出し行の誤集計防止）
    extents = {"履修明細": last_r, "生徒一覧": last_s, "クラブ所属": 3 + n_rows, "科目マスタ": last_m}
    for w in wb.worksheets:
        for row_cells in w.iter_rows():
            for cell in row_cells:
                v = cell.value
                if not (isinstance(v, str) and v.startswith("=") and ":$" in v):
                    continue
                for sheet_name, last in extents.items():
                    v = re.sub(rf"{sheet_name}!\$([A-Z]+):\$\1\b",
                               lambda m, sn=sheet_name, l=last: f"{sn}!${m.group(1)}$4:${m.group(1)}${max(l, 4)}", v)
                if w.title in extents:
                    last = max(extents[w.title], 4)
                    v = re.sub(r"(?<![!\w$])\$([A-Z]+):\$\1\b",
                               lambda m: f"${m.group(1)}$4:${m.group(1)}${last}", v)
                cell.value = v
        w.sheet_view.zoomScale = 90
    wb.save(out_path)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input_dir")
    ap.add_argument("output")
    ap.add_argument("--club", help="クラブ一覧(.xlsx)")
    ap.add_argument("--master", help="科目マスタ(.xlsx: A列 科目名, B列 授業時数)")
    a = ap.parse_args()
    files = sorted(p for p in Path(a.input_dir).glob("*.xlsx") if not p.name.startswith("~$"))
    students, sources = [], []
    for p in files:
        try:
            st = parse_class_file(p)
        except ValueError as e:
            print(f"スキップ: {e}", file=sys.stderr)
            continue
        students += st
        sources.append(p.name)
        print(f"{p.name}: {st[0]['学年']}年{st[0]['組']}組 {len(st)}名", file=sys.stderr)
    students.sort(key=lambda s: (s["学年"], cls_order(s["組"]), s["番号"]))
    clubs = read_club_list(a.club) if a.club else None
    master = read_master(a.master) if a.master else None
    build(students, a.output, clubs, master, sources)
    print(f"出力: {a.output}", file=sys.stderr)


if __name__ == "__main__":
    main()
