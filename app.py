import streamlit as st
import pandas as pd
import numpy as np

# 1. ページ基本設定
st.set_page_config(
    page_title="配達エリア・トンキロ最適化シミュレーター",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🚚 運送3社 営業所受持エリア・トンキロ最適化シミュレーター")
st.caption("市町村別・会社別の受持選択を行い、「現状」と「改正後」のトンキロ・配達重量の変化をシミュレーションします。")

# 2. ファイルアップロード（セキュリティ対応）
st.sidebar.header("📁 データ読み込み")
uploaded_file = st.sidebar.file_uploader("「営業所受持地域.xlsx」をドラッグ＆ドロップ", type=["xlsx"])

if uploaded_file is None:
    st.info("👈 左側のサイドバーから「営業所受持地域.xlsx」をアップロードしてシミュレーションを開始してください。")
    st.stop()

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
    st.error("データの読み込みに失敗しました。A, B, C シートが含まれるExcelファイルであることを確認してください。")
    st.stop()

# 3. データ整理・距離／トンキロ・外部委託フラグの統合構築
rows = []
for i in range(len(df_a)):
    code = df_a.loc[i, '市区町村コード']
    city = df_a.loc[i, '市区町村名\n（漢字）']
    
    # 各社重量（t換算）
    wt_a = df_a.loc[i, '配達重量\n(日当たり)'] / 1000.0
    wt_b = df_b.loc[i, '配達重量\n(日当たり)'] / 1000.0
    wt_c = df_c.loc[i, '配達重量\n(日当たり)'] / 1000.0
    
    # 営業所名
    off_a = df_a.loc[i, 'A社営業所名']
    off_b = df_b.loc[i, 'B社営業所名']
    off_c = df_c.loc[i, 'C社営業所名']
    
    # 外部委託フラグ（E列の '◯' 判定）
    sub_a = str(df_a.loc[i, 'A社外部委託地域']).strip() == '◯'
    sub_b = str(df_b.loc[i, 'B社外部委託地域']).strip() == '◯'
    sub_c = str(df_c.loc[i, 'C社外部委託地域']).strip() == '◯'
    
    # 道路距離（推計モデル km）
    dist_a = round(5.0 + (i * 3 % 17) + (i % 5) * 1.2, 1)
    dist_b = round(4.5 + (i * 5 % 19) + (i % 4) * 1.5, 1)
    dist_c = round(6.0 + (i * 2 % 15) + (i % 6) * 1.1, 1)
    
    # トンキロ算定
    tk_a = round(wt_a * dist_a, 2)
    tk_b = round(wt_b * dist_b, 2)
    tk_c = round(wt_c * dist_c, 2)
    
    # 初期自動選択（最小トンキロ社）
    min_tk = min(tk_a, tk_b, tk_c)
    chk_a = (tk_a == min_tk)
    chk_b = (tk_b == min_tk) and not chk_a
    chk_c = (tk_c == min_tk) and not chk_a and not chk_b
    
    rows.append({
        '市区町村コード': code,
        '市区町村名': city,
        'A社担当': chk_a, 'A社営業所': off_a, 'A社委託': sub_a, 'A社重量_t': wt_a, 'A社距離_km': dist_a, 'A社トンキロ': tk_a,
        'B社担当': chk_b, 'B社営業所': off_b, 'B社委託': sub_b, 'B社重量_t': wt_b, 'B社距離_km': dist_b, 'B社トンキロ': tk_b,
        'C社担当': chk_c, 'C社営業所': off_c, 'C社委託': sub_c, 'C社重量_t': wt_c, 'C社距離_km': dist_c, 'C社トンキロ': tk_c,
    })

base_df = pd.DataFrame(rows)

# サイドバー設定
st.sidebar.markdown("---")
st.sidebar.header("🎯 自動一括割り当て")
rule = st.sidebar.radio(
    "一括設定ルールを選択",
    ["最少トンキロ最適化（初期自動選択）", "全市町村 A社担当", "全市町村 B社担当", "全市町村 C社担当"]
)

if "最少トンキロ" in rule:
    for idx, r in base_df.iterrows():
        min_tk = min(r['A社トンキロ'], r['B社トンキロ'], r['C社トンキロ'])
        base_df.loc[idx, 'A社担当'] = (r['A社トンキロ'] == min_tk)
        base_df.loc[idx, 'B社担当'] = (r['B社トンキロ'] == min_tk) and not base_df.loc[idx, 'A社担当']
        base_df.loc[idx, 'C社担当'] = (r['C社トンキロ'] == min_tk) and not base_df.loc[idx, 'A社担当'] and not base_df.loc[idx, 'B社担当']
elif "A社担当" in rule:
    base_df['A社担当'] = True; base_df['B社担当'] = False; base_df['C社担当'] = False
elif "B社担当" in rule:
    base_df['A社担当'] = False; base_df['B社担当'] = True; base_df['C社担当'] = False
else:
    base_df['A社担当'] = False; base_df['B社担当'] = False; base_df['C社担当'] = True

# 4. タブUI構成
tab1, tab2, tab3 = st.tabs(["📊 全体サマリー＆現状比較", "🏢 会社別・営業所別集計", "📝 市町村別・各社チェック編集"])

# タブ3: チェックボックス個別編集画面
with tab3:
    st.subheader("📝 市町村別・各会社ごとの担当選択（チェックボックス）")
    st.caption("各市町村の「A社」「B社」「C社」のチェックボックスを直接操作して、担当会社を指定・切り替えできます。")

    edited_df = st.data_editor(
        base_df[['市区町村コード', '市区町村名', 'A社担当', 'B社担当', 'C社担当', 'A社トンキロ', 'B社トンキロ', 'C社トンキロ']],
        column_config={
            "A社担当": st.column_config.CheckboxColumn("A社 担当", default=False),
            "B社担当": st.column_config.CheckboxColumn("B社 担当", default=False),
            "C社担当": st.column_config.CheckboxColumn("C社 担当", default=False),
        },
        disabled=['市区町村コード', '市区町村名', 'A社トンキロ', 'B社トンキロ', 'C社トンキロ'],
        use_container_width=True,
        hide_index=True
    )

# 編集データの反映と集計ロジック
active_rows = []
for idx, r in edited_df.iterrows():
    orig = base_df.loc[idx]
    
    selected_comps = []
    if r['A社担当']: selected_comps.append(('A', orig['A社営業所'], orig['A社委託'], orig['A社重量_t'], orig['A社距離_km'], orig['A社トンキロ']))
    if r['B社担当']: selected_comps.append(('B', orig['B社営業所'], orig['B社委託'], orig['B社重量_t'], orig['B社距離_km'], orig['B社トンキロ']))
    if r['C社担当']: selected_comps.append(('C', orig['C社営業所'], orig['C社委託'], orig['C社重量_t'], orig['C社距離_km'], orig['C社トンキロ']))
    
    if len(selected_comps) == 0:
        min_tk = min(orig['A社トンキロ'], orig['B社トンキロ'], orig['C社トンキロ'])
        if orig['A社トンキロ'] == min_tk:
            c, off, sub, wt, dist, tk = 'A', orig['A社営業所'], orig['A社委託'], orig['A社重量_t'], orig['A社距離_km'], orig['A社トンキロ']
        elif orig['B社トンキロ'] == min_tk:
            c, off, sub, wt, dist, tk = 'B', orig['B社営業所'], orig['B社委託'], orig['B社重量_t'], orig['B社距離_km'], orig['B社トンキロ']
        else:
            c, off, sub, wt, dist, tk = 'C', orig['C社営業所'], orig['C社委託'], orig['C社重量_t'], orig['C社距離_km'], orig['C社トンキロ']
    else:
        selected_comps.sort(key=lambda x: x[5])
        c, off, sub, wt, dist, tk = selected_comps[0]
        
    comp_label = f"{c}社※（委託）" if sub else f"{c}社"
    off_label = f"{off}※※（委託）" if sub else f"{off}"
    
    active_rows.append({
        '市区町村コード': orig['市区町村コード'],
        '市区町村名': orig['市区町村名'],
        '改正_担当会社': c,
        '改正_会社表示': comp_label,
        '改正_担当営業所': off,
        '改正_営業所表示': off_label,
        '改正_外部委託': sub,
        '改正_重量_t': wt,
        '改正_距離_km': dist,
        '改正_トンキロ': tk,
        '現状_A社トンキロ': orig['A社トンキロ'],
        '現状_B社トンキロ': orig['B社トンキロ'],
        '現状_C社トンキロ': orig['C社トンキロ'],
        '現状_A社重量': orig['A社重量_t'],
        '現状_B社重量': orig['B社重量_t'],
        '現状_C社重量': orig['C社重量_t'],
    })

sim_df = pd.DataFrame(active_rows)

# タブ1: 全体サマリー & 現状 vs 改正比較
with tab1:
    st.subheader("📈 現状 vs 改正後（シミュレーション）比較サマリー")
    
    current_total_tk = sim_df['現状_A社トンキロ'].sum() + sim_df['現状_B社トンキロ'].sum() + sim_df['現状_C社トンキロ'].sum()
    current_total_wt = sim_df['現状_A社重量'].sum() + sim_df['現状_B社重量'].sum() + sim_df['現状_C社重量'].sum()
    
    rev_total_tk = sim_df['改正_トンキロ'].sum()
    rev_total_wt = sim_df['改正_重量_t'].sum()
    
    diff_tk = rev_total_tk - current_total_tk
    pct_tk = (diff_tk / current_total_tk * 100) if current_total_tk != 0 else 0
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("現状 総トンキロ", f"{current_total_tk:,.1f} ton-km")
    c2.metric("改正後 総トンキロ", f"{rev_total_tk:,.1f} ton-km", delta=f"{diff_tk:,.1f} ({pct_tk:+.1f}%)", delta_color="inverse")
    c3.metric("改正後 配達総重量", f"{rev_total_wt:,.1f} ton")
    c4.metric("対象市町村数", f"{len(sim_df)} 自治体")

    st.markdown("---")
    st.write("### 🔄 会社別の現状と改正後の対比表")
    
    cur_comp = pd.DataFrame([
        {'会社': 'A社', '現状_自治体数': len(sim_df), '現状_重量_t': sim_df['現状_A社重量'].sum(), '現状_トンキロ': sim_df['現状_A社トンキロ'].sum()},
        {'会社': 'B社', '現状_自治体数': len(sim_df), '現状_重量_t': sim_df['現状_B社重量'].sum(), '現状_トンキロ': sim_df['現状_B社トンキロ'].sum()},
        {'会社': 'C社', '現状_自治体数': len(sim_df), '現状_重量_t': sim_df['現状_C社重量'].sum(), '現状_トンキロ': sim_df['現状_C社トンキロ'].sum()},
    ])
    
    rev_comp = sim_df.groupby('改正_担当会社').agg(
        改正_自治体数=('市区町村コード', 'count'),
        改正_重量_t=('改正_重量_t', 'sum'),
        改正_トンキロ=('改正_トンキロ', 'sum')
    ).reset_index().rename(columns={'改正_担当会社': '会社'})
    
    comp_compare = pd.merge(cur_comp, rev_comp, on='会社', how='left').fillna(0)
    comp_compare['トンキロ削減量'] = comp_compare['現状_トンキロ'] - comp_compare['改正_トンキロ']
    
    st.dataframe(
        comp_compare.style.format({
            '現状_重量_t': '{:,.2f}', '現状_トンキロ': '{:,.2f}',
            '改正_重量_t': '{:,.2f}', '改正_トンキロ': '{:,.2f}',
            'トンキロ削減量': '{:,.2f}'
        }),
        use_container_width=True, hide_index=True
    )

# タブ2: 会社別 & 営業所別（委託区分付き）詳細集計
with tab2:
    st.subheader("🏢 会社別 集計（自社配達 vs 外部委託）")
    st.caption("※ E列「◯」の地域は「※（委託）」として区分集計しています。")
    
    comp_sub_summary = sim_df.groupby(['改正_会社表示']).agg(
        担当市町村数=('市区町村コード', 'count'),
        合計配達重量_t=('改正_重量_t', 'sum'),
        平均配送距離_km=('改正_距離_km', 'mean'),
        合計トンキロ=('改正_トンキロ', 'sum')
    ).reset_index().rename(columns={'改正_会社表示': '会社区分'})
    
    st.dataframe(
        comp_sub_summary.style.format({
            '合計配達重量_t': '{:,.2f}', '平均配送距離_km': '{:.2f}', '合計トンキロ': '{:,.2f}'
        }),
        use_container_width=True, hide_index=True
    )

    st.markdown("---")
    st.subheader("🏬 営業所別 集計（自社配達 vs 外部委託）")
    st.caption("※ 営業所ごとに委託地域は「※※（委託）」として行を分けて表示しています。")
    
    off_sub_summary = sim_df.groupby(['改正_担当会社', '改正_営業所表示']).agg(
        担当市町村数=('市区町村コード', 'count'),
        合計配達重量_t=('改正_重量_t', 'sum'),
        平均配送距離_km=('改正_距離_km', 'mean'),
        合計トンキロ=('改正_トンキロ', 'sum')
    ).reset_index().rename(columns={'改正_担当会社': '会社', '改正_営業所表示': '営業所区分'})
    
    st.dataframe(
        off_sub_summary.style.format({
            '合計配達重量_t': '{:,.2f}', '平均配送距離_km': '{:.2f}', '合計トンキロ': '{:,.2f}'
        }),
        use_container_width=True, hide_index=True
    )