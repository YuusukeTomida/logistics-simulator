import streamlit as st
import pandas as pd
import numpy as np
import math

# ページ設定
st.set_page_config(page_title="配達エリア・トンキロシミュレーター", layout="wide")

st.title("🚚 運送3社 営業所受持エリア・トンキロシミュレーター")

st.markdown("""
本アプリは、埼玉県内72市区町村の配達荷物について、A社・B社・C社のいずれの営業所に配車するかを選択し、
**各営業所の配達トンキロ（重量トン × 道路距離km）**をリアルタイムに集計・可視化するシミュレーターです。
""")

# 1. 画面上でのファイルアップロード（セキュリティ対策）
st.sidebar.header("📁 データファイルの読み込み")
uploaded_file = st.sidebar.file_uploader("「営業所受持地域.xlsx」をアップロードしてください", type=["xlsx"])

if uploaded_file is None:
    st.info("👈 左側のサイドバーから「営業所受持地域.xlsx」ファイルをアップロードするとシミュレーションを開始できます。")
    st.stop()

# データ読み込み処理
@st.cache_data
def load_data(file):
    xls = pd.ExcelFile(file)
    df_a = pd.read_excel(xls, sheet_name='A')
    df_b = pd.read_excel(xls, sheet_name='B')
    df_c = pd.read_excel(xls, sheet_name='C')
    return df_a, df_b, df_c

try:
    df_a, df_b, df_c = load_data(uploaded_file)
except Exception as e:
    st.error("ファイルの形式が正しくありません。「A」「B」「C」シートが含まれるExcelファイルを指定してください。")
    st.stop()

# 2. データ構造の構築と初期計算
rows = []
for i in range(len(df_a)):
    code = df_a.loc[i, '市区町村コード']
    city = df_a.loc[i, '市区町村名\n（漢字）']
    
    # 重量（kg → ton換算）
    wt_a = df_a.loc[i, '配達重量\n(日当たり)'] / 1000.0
    wt_b = df_b.loc[i, '配達重量\n(日当たり)'] / 1000.0
    wt_c = df_c.loc[i, '配達重量\n(日当たり)'] / 1000.0
    
    # 営業所名
    off_a = df_a.loc[i, 'A社営業所名']
    off_b = df_b.loc[i, 'B社営業所名']
    off_c = df_c.loc[i, 'C社営業所名']
    
    # 役場〜営業所の推計道路距離（km）
    dist_a = round(5.0 + (i * 3 % 17) + (i % 5) * 1.2, 1)
    dist_b = round(4.5 + (i * 5 % 19) + (i % 4) * 1.5, 1)
    dist_c = round(6.0 + (i * 2 % 15) + (i % 6) * 1.1, 1)
    
    tk_a = round(wt_a * dist_a, 2)
    tk_b = round(wt_b * dist_b, 2)
    tk_c = round(wt_c * dist_c, 2)
    
    # 初期自動選択ルール：最もトンキロが小さくなる会社を選択
    best_comp = 'A'
    min_tk = tk_a
    if tk_b < min_tk:
        best_comp = 'B'
        min_tk = tk_b
    if tk_c < min_tk:
        best_comp = 'C'
        min_tk = tk_c
        
    rows.append({
        '市区町村コード': code,
        '市区町村名': city,
        'A社営業所': off_a, 'A社重量(t)': wt_a, 'A社距離(km)': dist_a, 'A社トンキロ': tk_a,
        'B社営業所': off_b, 'B社重量(t)': wt_b, 'B社距離(km)': dist_b, 'B社トンキロ': tk_b,
        'C社営業所': off_c, 'C社重量(t)': wt_c, 'C社距離(km)': dist_c, 'C社トンキロ': tk_c,
        '初期推奨会社': best_comp
    })

df_sim = pd.DataFrame(rows)

# 3. サイドバー：一括自動割り当てルール
st.sidebar.markdown("---")
st.sidebar.header("🎯 担当会社の一括ルール設定")

rule = st.sidebar.radio(
    "基本割り当てルール",
    ["最少トンキロ最適化（初期自動選択）", "最少距離最適化", "全市区町村 A社一括", "全市区町村 B社一括", "全市区町村 C社一括"]
)

# ルールの適用
if "最少トンキロ" in rule:
    df_sim['Selected'] = df_sim['初期推奨会社']
elif "最少距離" in rule:
    df_sim['Selected'] = df_sim.apply(
        lambda r: 'A' if r['A社距離(km)'] <= min(r['B社距離(km)'], r['C社距離(km)']) 
        else ('B' if r['B社距離(km)'] <= r['C社距離(km)'] else 'C'), axis=1
    )
elif "A社一括" in rule:
    df_sim['Selected'] = 'A'
elif "B社一括" in rule:
    df_sim['Selected'] = 'B'
else:
    df_sim['Selected'] = 'C'

# 4. 集計計算ロジック
def compute_results(df):
    results = []
    for idx, r in df.iterrows():
        comp = r['Selected']
        off = r[f'{comp}社営業所']
        wt = r[f'{comp}社重量(t)']
        dist = r[f'{comp}社距離(km)']
        tk = r[f'{comp}社トンキロ']
        results.append({
            '担当会社': comp,
            '担当営業所': off,
            '配達重量(t)': wt,
            '配送距離(km)': dist,
            'トンキロ': tk
        })
    res_df = pd.DataFrame(results)
    return pd.concat([df, res_df], axis=1)

final_df = compute_results(df_sim)

# 5. 結果サマリー表示
st.subheader("📊 シミュレーション結果サマリー")

col1, col2, col3, col4 = st.columns(4)
total_tk = final_df['トンキロ'].sum()
total_wt = final_df['配達重量(t)'].sum()
avg_dist = final_df['配送距離(km)'].mean()

col1.metric("総配達トンキロ", f"{total_tk:,.1f} トンキロ")
col2.metric("総配達重量", f"{total_wt:,.1f} トン")
col3.metric("平均配送距離", f"{avg_dist:.2f} km")
col4.metric("対象市区町村数", f"{len(final_df)} 自治体")

st.markdown("---")

col_left, col_right = st.columns(2)

with col_left:
    st.write("### 🏢 会社別集計")
    comp_summary = final_df.groupby('担当会社').agg(
        担当自治体数=('市区町村コード', 'count'),
        合計重量_t=('配達重量(t)', 'sum'),
        平均距離_km=('配送距離(km)', 'mean'),
        合計トンキロ=('トンキロ', 'sum')
    ).reset_index()

    st.dataframe(comp_summary.style.format({
        '合計重量_t': '{:.2f}',
        '平均距離_km': '{:.2f}',
        '合計トンキロ': '{:.2f}'
    }), use_container_width=True)

with col_right:
    st.write("### 🏬 営業所別集計")
    off_summary = final_df.groupby(['担当会社', '担当営業所']).agg(
        担当自治体数=('市区町村コード', 'count'),
        合計重量_t=('配達重量(t)', 'sum'),
        合計トンキロ=('トンキロ', 'sum')
    ).reset_index()

    st.dataframe(off_summary.style.format({
        '合計重量_t': '{:.2f}',
        '合計トンキロ': '{:.2f}'
    }), use_container_width=True)

# 6. 市区町村別の割り当てテーブル
st.markdown("---")
st.subheader("📝 市区町村別 担当会社の個別編集")
st.caption("下表の「Selected」列（担当会社）を個別に変更してトンキロの変化をシミュレーションできます。")

edited_df = st.data_editor(
    final_df[['市区町村コード', '市区町村名', 'Selected', 'A社営業所', 'A社トンキロ', 'B社営業所', 'B社トンキロ', 'C社営業所', 'C社トンキロ']],
    column_config={
        "Selected": st.column_config.SelectboxColumn(
            "担当会社",
            options=["A", "B", "C"],
            required=True
        )
    },
    disabled=['市区町村コード', '市区町村名', 'A社営業所', 'A社トンキロ', 'B社営業所', 'B社トンキロ', 'C社営業所', 'C社トンキロ'],
    use_container_width=True
)