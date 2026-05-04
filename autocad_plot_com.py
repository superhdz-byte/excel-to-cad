"""
節点・要素番号プロット → AutoCAD 2024 COM自動描画

使い方：
  python autocad_plot_com.py

依存ライブラリ：
  pip install pywin32 pandas openpyxl
"""

import math
import os
import time
import pandas as pd
import win32com.client
import pythoncom


# ============================================================
# ★ 設定
# ============================================================
EXCEL_PATH = r"F:\1.SF-Works\CAD\常時全土圧.xlsx"
DWG_PATH   = r"F:\1.SF-Works\CAD\output.dwg"

SIZE = 300.0

COLOR_ELEMENT_LINE = 3
COLOR_NODE_TEXT    = 4
COLOR_ELEM_TEXT    = 1

AC_ALIGN_LEFT = 0
AC_ALIGN_MIDDLE_CENTER = 4
# ============================================================


def make_point(x, y, z=0.0):
    return win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        [float(x), float(y), float(z)]
    )


def ensure_layer(doc, name, color=7):
    try:
        layer = doc.Layers.Item(name)
    except Exception:
        layer = doc.Layers.Add(name)
        layer.Color = color
    return layer


def connect_autocad():
    pythoncom.CoInitialize()

    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        print("✅ 起動中の AutoCAD に接続しました")
    except Exception:
        print("🚀 AutoCAD を起動中...")
        acad = win32com.client.Dispatch("AutoCAD.Application")
        time.sleep(5)
        print("✅ AutoCAD 起動完了")

    acad.Visible = True
    return acad


def load_data_format_excel(excel_path):
    import os
    import re
    import pandas as pd

    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excelファイルが見つかりません: {excel_path}")

    df = pd.read_excel(excel_path, header=None, dtype=str)

    grid_start = None
    member_start = None

    # A列だけでなく、行全体から検索する
    for i in range(len(df)):
        row_text = " ".join(
            str(v) for v in df.iloc[i].tolist()
            if pd.notna(v)
        )

        if "格点データ" in row_text:
            grid_start = i

        if "部材データ" in row_text:
            member_start = i

    if grid_start is None:
        raise Exception("❌ ■格点データ が見つかりません")

    if member_start is None:
        raise Exception("❌ ■部材データ が見つかりません")

    print(f"✅ ■格点データ 行: {grid_start + 1}")
    print(f"✅ ■部材データ 行: {member_start + 1}")

    # ============================================================
    # 1. 格点データ解析
    #    形式：
    #    1  4.3840  8.5500   16  11.4740  1.3550
    # ============================================================
    nodes = {}

    for i in range(grid_start + 1, member_start):
        row_text = " ".join(
            str(v) for v in df.iloc[i].tolist()
            if pd.notna(v)
        )

        nums = re.findall(r"[-+]?\d+(?:\.\d+)?(?:[Ee][-+]?\d+)?", row_text)

        # 左側：格点番号 X Y
        if len(nums) >= 3:
            try:
                node_no = int(float(nums[0]))
                x = float(nums[1])
                y = float(nums[2])
                nodes[node_no] = (x, y)
            except Exception:
                pass

        # 右側：格点番号 X Y
        if len(nums) >= 6:
            try:
                node_no = int(float(nums[3]))
                x = float(nums[4])
                y = float(nums[5])
                nodes[node_no] = (x, y)
            except Exception:
                pass

    # ============================================================
    # 2. 部材データ解析
    #    形式：
    #    部材番号 i端 j端 部材長 ...
    # ============================================================
    elements = []

    for i in range(member_start + 1, len(df)):
        row_text = " ".join(
            str(v) for v in df.iloc[i].tolist()
            if pd.notna(v)
        )

        # 次のデータブロックに入ったら終了
        if i > member_start + 3 and "■" in row_text:
            break

        nums = re.findall(r"[-+]?\d+(?:\.\d+)?(?:[Ee][-+]?\d+)?", row_text)

        if len(nums) >= 3:
            try:
                eno = int(float(nums[0]))
                i_node = int(float(nums[1]))
                j_node = int(float(nums[2]))
                elements.append((eno, i_node, j_node))
            except Exception:
                pass

    print(f"✅ 格点読込: {len(nodes)} 個")
    print(f"✅ 部材読込: {len(elements)} 本")

    if not nodes:
        raise Exception("❌ 格点データを読み取れませんでした")

    if not elements:
        raise Exception("❌ 部材データを読み取れませんでした")

    return nodes, elements


def draw_all(doc, nodes, elements, size=300.0):
    ms = doc.ModelSpace
    errors = []

    line_layer = ensure_layer(doc, "要素線", COLOR_ELEMENT_LINE)
    node_layer = ensure_layer(doc, "節点番号", COLOR_NODE_TEXT)
    elem_text_layer = ensure_layer(doc, "要素番号", COLOR_ELEM_TEXT)

    # ============================================================
    # 1. 要素線を描画
    # ============================================================
    print("📐 要素線を描画中...")
    doc.ActiveLayer = line_layer

    for idx, (eno, i_node, j_node) in enumerate(elements):
        if i_node not in nodes:
            errors.append(f"要素{eno}: i端節点 {i_node} が座標シートにありません")
            continue
        if j_node not in nodes:
            errors.append(f"要素{eno}: j端節点 {j_node} が座標シートにありません")
            continue

        xi, yi = nodes[i_node]
        xj, yj = nodes[j_node]

        try:
            line = ms.AddLine(make_point(xi, yi), make_point(xj, yj))
            line.Layer = "要素線"
        except Exception as e:
            errors.append(f"要素{eno} 線分描画エラー: {e}")

    # ============================================================
    # 2. 節点円・節点番号を描画
    # ============================================================
    print("⭕ 節点番号を描画中...")
    doc.ActiveLayer = node_layer

    for node_no, (x, y) in nodes.items():
        try:
            circle = ms.AddCircle(make_point(x, y), size)
            circle.Layer = "節点番号"

            txt = ms.AddText(str(node_no), make_point(x + size * 0.3, y + size * 0.3), size)
            txt.Layer = "節点番号"
            txt.Alignment = AC_ALIGN_LEFT
        except Exception as e:
            errors.append(f"節点{node_no} 描画エラー: {e}")

    # ============================================================
    # 3. 要素番号を描画
    # ============================================================
    print("🔢 要素番号を描画中...")
    doc.ActiveLayer = elem_text_layer

    total = len(elements)

    for idx, (eno, i_node, j_node) in enumerate(elements):
        if i_node not in nodes or j_node not in nodes:
            continue

        xi, yi = nodes[i_node]
        xj, yj = nodes[j_node]

        xmid = (xi + xj) / 2.0
        ymid = (yi + yj) / 2.0

        try:
            txt = ms.AddText(str(eno), make_point(xmid, ymid), size)
            txt.Layer = "要素番号"
            txt.Alignment = AC_ALIGN_MIDDLE_CENTER
            txt.TextAlignmentPoint = make_point(xmid, ymid)
        except Exception as e:
            errors.append(f"要素{eno} 番号描画エラー: {e}")

        if (idx + 1) % 10 == 0 or (idx + 1) == total:
            pct = (idx + 1) / total * 100
            print(f"  進捗: {idx + 1}/{total} ({pct:.0f}%)")

    return errors


def save_dwg(doc, dwg_path):
    if not dwg_path:
        print("💡 DWG_PATH が未設定のため保存しません")
        return

    os.makedirs(os.path.dirname(dwg_path), exist_ok=True)

    if os.path.exists(dwg_path):
        try:
            os.remove(dwg_path)
        except Exception:
            print(f"⚠️ 既存DWGを削除できません。AutoCADで開いている可能性があります: {dwg_path}")
            print("   別名で保存します。")
            base, ext = os.path.splitext(dwg_path)
            dwg_path = base + "_new" + ext

    doc.SaveAs(dwg_path)
    print(f"💾 保存完了: {dwg_path}")


def main():
    print("=" * 60)
    print("  節点・要素番号プロット → AutoCAD 2024 COM自動描画")
    print("=" * 60)

    nodes, elements = load_data_format_excel(EXCEL_PATH)

    acad = connect_autocad()

    print("📄 新規図面を作成中...")
    doc = acad.Documents.Add()

    t0 = time.time()

    try:
        # 環境によって効かない場合があるため try
        try:
            acad.ScreenUpdating = False
        except Exception:
            pass

        errors = draw_all(doc, nodes, elements, SIZE)

        try:
            acad.ScreenUpdating = True
        except Exception:
            pass

        acad.ZoomExtents()

        save_dwg(doc, DWG_PATH)

    finally:
        try:
            acad.ScreenUpdating = True
        except Exception:
            pass

    elapsed = time.time() - t0
    print(f"✅ 描画処理完了: {elapsed:.1f} 秒")

    if errors:
        print()
        print(f"⚠️ エラー・スキップ: {len(errors)}件")
        for e in errors[:30]:
            print(" -", e)

        if len(errors) > 30:
            print(f" ... 他 {len(errors) - 30} 件")

    print()
    print("完了。AutoCAD画面を確認してください。")


if __name__ == "__main__":
    main()
