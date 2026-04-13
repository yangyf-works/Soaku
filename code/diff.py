from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from difflib import SequenceMatcher
from patiencediff import PatienceSequenceMatcher
import json

def row_similarity(row_a, row_b):
    """行同士の類似度"""
    return SequenceMatcher(None, row_a, row_b).ratio()


def match_rows(sub_a, sub_b, threshold=0.5):
    """
    行同士を類似度でマッチングする
    戻り値:
        matched: [(i, j or None)]
        unmatched_b: [j, ...]
    """
    matched = []
    used_b = set()

    for i, row_a in enumerate(sub_a):
        best_j = None
        best_score = 0

        for j, row_b in enumerate(sub_b):
            if j in used_b:
                continue

            score = row_similarity(row_a, row_b)
            if score > best_score:
                best_score = score
                best_j = j

        if best_score >= threshold:
            matched.append((i, best_j))
            used_b.add(best_j)
        else:
            matched.append((i, None))

    unmatched_b = [j for j in range(len(sub_b)) if j not in used_b]

    return matched, unmatched_b

def compare_sheets(wb1, wb2):
    sheets1 = set(wb1.sheetnames)
    sheets2 = set(wb2.sheetnames)

    added = sheets2 - sheets1      # file2 にのみ存在
    deleted = sheets1 - sheets2    # file1 にのみ存在
    common = sheets1 & sheets2     # 両方に存在

    return {
        "added": added,
        "deleted": deleted,
        "common": common
    }

def get_header_map(ws, headrow=1):
    """
    ヘッダ行を取得して
    {ヘッダ名: 列番号} の辞書を返す
    """
    header_map = {}

    for col_idx, cell in enumerate(ws[headrow], start=1):
        if cell.value is not None:
            header_map[str(cell.value)] = col_idx

    return header_map

def compare_columns(wb1, wb2, sheet_name, headrow=1):
    ws1 = wb1[sheet_name]
    ws2 = wb2[sheet_name]

    header1 = get_header_map(ws1, headrow)  # {name: col_index}
    header2 = get_header_map(ws2, headrow)

    set1 = set(header1.keys())
    set2 = set(header2.keys())

    # ★ ここを変更
    added = [header2[col] for col in (set2 - set1)]
    deleted = [header1[col] for col in (set1 - set2)]
    common = [
        (header1[col], header2[col])
        for col in (set1 & set2)
    ]

    return {
        "added": added,       # [col_index_in_file2]
        "deleted": deleted,   # [col_index_in_file1]
        "common": common,     # [(idx1, idx2)]
    }

def normalize_row(row, common_cols):
    return tuple(
        "" if row[col-1].value is None else str(row[col-1].value)
        for col in common_cols
    )

def trim_tail_empty(rows):
    while rows and all(cell == "" or cell is None for cell in rows[-1]):
        rows.pop()
    return rows

def to_excel_row(idx, base):
    return idx + base

def to_excel_cell(col_idx, row_idx):
    return f"{get_column_letter(col_idx)}{row_idx}"

def exceldiff(path1 , 
         path2 ,
         dataonly=False,
         headrow=1,
         threshold=0.5):
    
    base = headrow + 1
    wb1 = load_workbook(path1, data_only=dataonly)
    wb2 = load_workbook(path2, data_only=dataonly)

    result = compare_sheets(wb1, wb2)

    diff_result = {}
    if result["added"]:
        diff_result["sheet_added"] = sorted(result["added"])
    if result["deleted"]:
        diff_result["sheet_deleted"] = sorted(result["deleted"])

    sheet_modified= {}
    for sheet_name in result["common"]:
        colresult = compare_columns(wb1, wb2, sheet_name, headrow)

        if colresult["added"]:
            sheet_modified.setdefault(sheet_name, {}).setdefault("columns", {})["added"] = [
                    get_column_letter(col)
                    for col in sorted(colresult["added"])
                ]
            
        if colresult["deleted"]:
            sheet_modified.setdefault(sheet_name, {}).setdefault("columns", {})["deleted"] = [
                    get_column_letter(col)
                    for col in sorted(colresult["deleted"])
                ]

        common_pairs = sorted(colresult["common"], key=lambda x: x[1])
        common_cols_a = [v[0] for v in common_pairs]
        common_cols_b = [v[1] for v in common_pairs]
        rows_a = [
            normalize_row(r, common_cols_a)
            for r in wb1[sheet_name].iter_rows(min_row=headrow+1)
        ]
        rows_b = [
            normalize_row(r, common_cols_b)
            for r in wb2[sheet_name].iter_rows(min_row=headrow+1)
        ]
        rows_a = trim_tail_empty(rows_a)
        rows_b = trim_tail_empty(rows_b)

        sm = PatienceSequenceMatcher(None, rows_a, rows_b)
        added_rows = []
        removed_rows = []
        cellsdiff = []
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "replace":
                sub_a = rows_a[i1:i2]
                sub_b = rows_b[j1:j2]

                """
                replaceブロックを解析して
                行変更・行追加・行削除・セル変更を再構築する
                """
                matched, unmatched_b = match_rows(sub_a, sub_b, threshold)
                cell_changes = []
                # 行削除・行変更
                for i, j in matched:
                    abs_i = i1 + i
                    excel_i = to_excel_row(abs_i, base)
                    if j is None:
                        removed_rows.append(excel_i)
                    else:
                        abs_j = j1 + j
                        excel_j = to_excel_row(abs_j, base)

                        row_a = sub_a[i]
                        row_b = sub_b[j]

                        if row_a == row_b:
                            continue

                        for col_idx, (a, b) in enumerate(zip(row_a, row_b)):
                            if a != b:
                                cell_changes.append({
                                    "from": to_excel_cell(common_pairs[col_idx][0], excel_i),
                                    "to": to_excel_cell(common_pairs[col_idx][1], excel_j)
                                })

                # 行追加
                for j in unmatched_b:
                    abs_j = j1 + j
                    excel_j = to_excel_row(abs_j, base)
                    added_rows.append(excel_j)

                cellsdiff.extend(cell_changes)

            elif tag == "delete":
                removed_rows.extend(to_excel_row(i, base) for i in range(i1, i2))

            elif tag == "insert":
                added_rows.extend(to_excel_row(j, base) for j in range(j1, j2))

        if added_rows:
            sheet_modified.setdefault(sheet_name, {}).setdefault("rows", {})["added"] = sorted(added_rows)

        if removed_rows:
            sheet_modified.setdefault(sheet_name, {}).setdefault("rows", {})["deleted"] = sorted(removed_rows)
            
        if cellsdiff:
            sheet_modified.setdefault(sheet_name, {}).setdefault("rows", {}).update({
                "modified": cellsdiff
            })
    if sheet_modified:
        diff_result.setdefault("sheet_modified", {}).update(sheet_modified)
    return diff_result
