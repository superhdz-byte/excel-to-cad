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
EXCEL_PATH = r"F:\1.SF-Works\CAD\節点・要素番号プロット(IJCAD)1rev2.xlsx"
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


def load_data(excel_path):
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excelファイルが見つかりません: {excel_path}")

    nodes_df = pd.read_excel(excel_path, sheet_name="座標", header=None)
    elems_df = pd.read_excel(excel_path, sheet_name="要素", header=None)

    nodes = {}
    for _, r in nodes_df.iterrows():
        if pd.isna(r[0]) or pd.isna(r[1]) or pd.isna(r[2]):
            continue
        node_no = int(r[0])
        x = float(r[1])
        y = float(r[2])
        nodes[node_no] = (x, y)

    elements = []
    for _, r in elems_df.iterrows():
        if pd.isna(r[0]) or pd.isna(r[1]) or pd.isna(r[2]):
            continue
        eno = int(r[0])
        i_node = int(r[1])
        j_node = int(r[2])
        elements.append((eno, i_node, j_node))

    print(f"✅ データ読み込み完了: 節点 {len(nodes)}個 / 要素 {len(elements)}個")
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

    nodes, elements = load_data(EXCEL_PATH)

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
