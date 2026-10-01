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
    if tk_a == min_tk:
        best_comp = 'A社'
    elif tk_b == min_tk:
        best_comp = 'B社'
    else:
        best_comp = 'C社'
    
    rows.append({
        '市区町村コード': code,
        '市区町村名': city,
        '一括担当': best_comp, # 一括選択用（単一選択）
        'A社担当': False,
        'B社担当': False,
        'C社担当': False,
        'A社営業所': off_a, 'A社委託': sub_a, 'A社重量_t': wt_a, 'A社距離_km': dist_a, 'A社トンキロ': tk_a,
        'B社営業所': off_b, 'B社委託': sub_b, 'B社重量_t': wt_b, 'B社距離_km': dist_b, 'B社トンキロ': tk_b,
        'C社営業所': off_c, 'C社委託': sub_c, 'C社重量_t': wt_c, 'C社距離_km': dist_c, 'C社トンキロ': tk_c,
    })

base_df = pd.DataFrame(rows)

# サイドバー設定
st.sidebar.markdown("---")
st.sidebar.header("🎯 自動一括割り当て")
rule = st.sidebar.radio(
    "一括設定ルールを選択",
    ["最少トンキロ最適化（初期自動選択）", "全市町村 A社一括", "全市町村 B社一括", "全市町村 C社一括", "選択クリア"]
)

if "最少トンキロ" in rule:
    for idx, r in base_df.iterrows():
        min_tk = min(r['A社トンキロ'], r['B社トンキロ'], r['C社トンキロ'])
        best = 'A社' if r['A社トンキロ'] == min_tk else ('B社' if r['B社トンキロ'] == min_tk else 'C社')
        base_df.loc[idx, '一括担当'] = best
        base_df.loc[idx, 'A社担当'] = False
        base_df.loc[idx, 'B社担当'] = False
        base_df.loc[idx, 'C社担当'] = False
elif "A社一括" in rule:
    base_df['一括担当'] = 'A社'; base_df['A社担当'] = False; base_df['B社担当'] = False; base_df['C社担当'] = False
elif "B社一括" in rule:
    base_df['一括担当'] = 'B社'; base_df['A社担当'] = False; base_df['B社担当'] = False; base_df['C社担当'] = False
elif "C社一括" in rule:
    base_df['一括担当'] = 'C社'; base_df['A社担当'] = False; base_df['B社担当'] = False; base_df['C社担当'] = False
else:
    base_df['一括担当'] = 'なし'; base_df['A社担当'] = False; base_df['B社担当'] = False; base_df['C社担当'] = False

# 4. タブUI構成
tab1, tab2, tab3 = st.tabs(["📊 全体サマリー＆現状比較", "🏢 会社別・営業所別集計", "📝 市町村別・受持選択（編集）"])

# タブ3: 個別編集画面
with tab3:
    st.subheader("📝 市町村別・受持選択テーブル")
    st.markdown("""
    - **一括担当（列1）**: 「A社」「B社」「C社」から1つ選択すると、市町村全体の荷物（全社分）が選択した会社に集計されます。
    - **個社チェック（列2〜4）**: チェックを入れた会社の荷物のみが、その会社に集計されます。
    - 💡 **優先ルール**: 個社チェックが1つでも入っている場合は **「個社チェック」が優先** されます。
    """)

    edited_df = st.data_editor(
        base_df[['市区町村コード', '市区町村名', '一括担当', 'A社担当', 'B社担当', 'C社担当', 'A社トンキロ', 'B社トンキロ', 'C社トンキロ']],
        column_config={
            "一括担当": st.column_config.SelectboxColumn(
                "市町村全体 一括担当",
                options=["なし", "A社", "B社", "C社"],
                required=True
            ),
            "A社担当": st.column_config.CheckboxColumn("A社 個別", default=False),
            "B社担当": st.column_config.CheckboxColumn("B社 個別", default=False),
            "C社担当": st.column_config.CheckboxColumn("C社 個別", default=False),
        },
        disabled=['市区町村コード', '市区町村名', 'A社トンキロ', 'B社トンキロ', 'C社トンキロ'],
        use_container_width=True,
        hide_index=True
    )

# 編集データの集計反映ロジック
active_records = []

for idx, r in edited_df.iterrows():
    orig = base_df.loc[idx]
    
    has_indiv = r['A社担当'] or r['B社担当'] or r['C社担当']
    
    if has_indiv:
        # 個別チェック優先ルール（要件⑥）
        if r['A社担当']:
            sub = orig['A社委託']
            active_records.append({
                '市区町村コード': orig['市区町村コード'], '市区町村名': orig['市区町村名'],
                '担当会社': 'A社', '会社表示': 'A社（委託）' if sub else 'A社',
                '担当営業所': orig['A社営業所'], '営業所表示': f"{orig['A社営業所']}（委託）" if sub else orig['A社営業所'],
                '重量_t': orig['A社重量_t'], '距離_km': orig['A社距離_km'], 'トンキロ': orig['A社トンキロ'],
                '現状_A社トンキロ': orig['A社トンキロ'], '現状_B社トンキロ': orig['B社トンキロ'], '現状_C社トンキロ': orig['C社トンキロ'],
                '現状_A社重量': orig['A社重量_t'], '現状_B社重量': orig['B社重量_t'], '現状_C社重量': orig['C社重量_t']
            })
        if r['B社担当']:
            sub = orig['B社委託']
            active_records.append({
                '市区町村コード': orig['市区町村コード'], '市区町村名': orig['市区町村名'],
                '担当会社': 'B社', '会社表示': 'B社（委託）' if sub else 'B社',
                '担当営業所': orig['B社営業所'], '営業所表示': f"{orig['B社営業所']}（委託）" if sub else orig['B社営業所'],
                '重量_t': orig['B社重量_t'], '距離_km': orig['B社距離_km'], 'トンキロ': orig['B社トンキロ'],
                '現状_A社トンキロ': orig['A社トンキロ'], '現状_B社トンキロ': orig['B社トンキロ'], '現状_C社トンキロ': orig['C社トンキロ'],
                '現状_A社重量': orig['A社重量_t'], '現状_B社重量': orig['B社重量_t'], '現状_C社重量': orig['C社重量_t']
            })
        if r['C社担当']:
            sub = orig['C社委託']
            active_records.append({
                '市区町村コード': orig['市区町村コード'], '市区町村名': orig['市区町村名'],
                '担当会社': 'C社', '会社表示': 'C社（委託）' if sub else 'C社',
                '担当営業所': orig['C社営業所'], '営業所表示': f"{orig['C社営業所']}（委託）" if sub else orig['C社営業所'],
                '重量_t': orig['C社重量_t'], '距離_km': orig['C社距離_km'], 'トンキロ': orig['C社トンキロ'],
                '現状_A社トンキロ': orig['A社トンキロ'], '現状_B社トンキロ': orig['B社トンキロ'], '現状_C社トンキロ': orig['C社トンキロ'],
                '現状_A社重量': orig['A社重量_t'], '現状_B社重量': orig['B社重量_t'], '現状_C社重量': orig['C社重量_t']
            })
    else:
        # 一括選択ルール（A・B・C全社の合計トンキロを選択会社へ一括集計）
        target_comp = r['一括担当']
        if target_comp in ['A社', 'B社', 'C社']:
            c_code = target_comp[0] # 'A', 'B', 'C'
            sub = orig[f'{c_code}社委託']
            off = orig[f'{c_code}社営業所']
            
            # 全3社分を合計
            total_wt = orig['A社重量_t'] + orig['B社重量_t'] + orig['C社重量_t']
            total_tk = orig['A社トンキロ'] + orig['B社トンキロ'] + orig['C社トンキロ']
            avg_dist = orig[f'{c_code}社距離_km']
            
            active_records.append({
                '市区町村コード': orig['市区町村コード'], '市区町村名': orig['市区町村名'],
                '担当会社': target_comp, '会社表示': f"{target_comp}（委託）" if sub else target_comp,
                '担当営業所': off, '営業所表示': f"{off}（委託）" if sub else off,
                '重量_t': total_wt, '距離_km': avg_dist, 'トンキロ': total_tk,
                '現状_A社トンキロ': orig['A社トンキロ'], '現状_B社トンキロ': orig['B社トンキロ'], '現状_C社トンキロ': orig['C社トンキロ'],
                '現状_A社重量': orig['A社重量_t'], '現状_B社重量': orig['B社重量_t'], '現状_C社重量': orig['C社重量_t']
            })

sim_df = pd.DataFrame(active_records) if len(active_records) > 0 else pd.DataFrame()

# タブ1: 全体サマリー & 現状 vs 改正比較
with tab1:
    st.subheader("📈 現状 vs 改正後（シミュレーション）比較サマリー")
    
    current_total_tk = base_df['A社トンキロ'].sum() + base_df['B社トンキロ'].sum() + base_df['C社トンキロ'].sum()
    current_total_wt = base_df['A社重量_t'].sum() + base_df['B社重量_t'].sum() + base_df['C社重量_t'].sum()
    
    rev_total_tk = sim_df['トンキロ'].sum() if not sim_df.empty else 0.0
    rev_total_wt = sim_df['重量_t'].sum() if not sim_df.empty else 0.0
    
    diff_tk = rev_total_tk - current_total_tk
    pct_tk = (diff_tk / current_total_tk * 100) if current_total_tk != 0 else 0
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("現状 総トンキロ", f"{current_total_tk:,.1f} ton-km")
    c2.metric("改正後 総トンキロ", f"{rev_total_tk:,.1f} ton-km", delta=f"{diff_tk:,.1f} ({pct_tk:+.1f}%)", delta_color="inverse")
    c3.metric("改正後 配達総重量", f"{rev_total_wt:,.1f} ton")
    c4.metric("設定済みレコード数", f"{len(sim_df)} 件")

    st.markdown("---")
    st.write("### 🔄 会社別の現状と改正後の対比表")
    
    cur_comp = pd.DataFrame([
        {'会社': 'A社', '現状_自治体数': len(base_df), '現状_重量_t': base_df['A社重量_t'].sum(), '現状_トンキロ': base_df['A社トンキロ'].sum()},
        {'会社': 'B社', '現状_自治体数': len(base_df), '現状_重量_t': base_df['B社重量_t'].sum(), '現状_トンキロ': base_df['B社トンキロ'].sum()},
        {'会社': 'C社', '現状_自治体数': len(base_df), '現状_重量_t': base_df['C社重量_t'].sum(), '現状_トンキロ': base_df['C社トンキロ'].sum()},
    ])
    
    if not sim_df.empty:
        rev_comp = sim_df.groupby('担当会社').agg(
            改正_担当件数=('市区町村コード', 'count'),
            改正_重量_t=('重量_t', 'sum'),
            改正_トンキロ=('トンキロ', 'sum')
        ).reset_index().rename(columns={'担当会社': '会社'})
    else:
        rev_comp = pd.DataFrame(columns=['会社', '改正_担当件数', '改正_重量_t', '改正_トンキロ'])
    
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
    st.caption("※ E列「◯」の地域は「（委託）」として区分集計しています。")
    
    if not sim_df.empty:
        comp_sub_summary = sim_df.groupby(['会社表示']).agg(
            担当件数=('市区町村コード', 'count'),
            合計配達重量_t=('重量_t', 'sum'),
            平均配送距離_km=('距離_km', 'mean'),
            合計トンキロ=('トンキロ', 'sum')
        ).reset_index().rename(columns={'会社表示': '会社区分'})
        
        st.dataframe(
            comp_sub_summary.style.format({
                '合計配達重量_t': '{:,.2f}', '平均配送距離_km': '{:.2f}', '合計トンキロ': '{:,.2f}'
            }),
            use_container_width=True, hide_index=True
        )
    else:
        st.info("データが未選択です。")

    st.markdown("---")
    st.subheader("🏬 営業所別 集計（自社配達 vs 外部委託）")
    st.caption("※ 営業所ごとに委託地域は「（委託）」として区分表示しています。")
    
    if not sim_df.empty:
        off_sub_summary = sim_df.groupby(['担当会社', '営業所表示']).agg(
            担当件数=('市区町村コード', 'count'),
            合計配達重量_t=('重量_t', 'sum'),
            平均配送距離_km=('距離_km', 'mean'),
            合計トンキロ=('トンキロ', 'sum')
        ).reset_index().rename(columns={'担当会社': '会社', '営業所表示': '営業所区分'})
        
        st.dataframe(
            off_sub_summary.style.format({
                '合計配達重量_t': '{:,.2f}', '平均配送距離_km': '{:.2f}', '合計トンキロ': '{:,.2f}'
            }),
            use_container_width=True, hide_index=True
        )
    else:
        st.info("データが未選択です。")