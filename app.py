import streamlit as st
import pandas as pd
import numpy as np
import cv2
import os

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

# 各社の営業所リストとグラデーションカラーマップを動的生成
def generate_office_colors(df_a, df_b, df_c):
    a_offices = sorted(list(set(df_a['A社営業所名'].dropna().astype(str).tolist())))
    b_offices = sorted(list(set(df_b['B社営業所名'].dropna().astype(str).tolist())))
    c_offices = sorted(list(set(df_c['C社営業所名'].dropna().astype(str).tolist())))
    
    office_color_map = {}
    
    # 営業所ごとの塗りを鮮明な濃淡グラデーションで生成
    n_a = len(a_offices)
    for idx, off in enumerate(a_offices):
        f = 0.45 + 0.55 * (idx / max(n_a - 1, 1)) if n_a > 1 else 0.85
        r = int(255 * f)
        g = int(20 * (1.0 - f))
        b = int(20 * (1.0 - f))
        office_color_map[off] = (r, g, b)
        
    n_b = len(b_offices)
    for idx, off in enumerate(b_offices):
        f = 0.45 + 0.55 * (idx / max(n_b - 1, 1)) if n_b > 1 else 0.85
        r = int(30 * (1.0 - f))
        g = int(140 * f)
        b = int(255 * f)
        office_color_map[off] = (r, g, b)
        
    n_c = len(c_offices)
    for idx, off in enumerate(c_offices):
        f = 0.45 + 0.55 * (idx / max(n_c - 1, 1)) if n_c > 1 else 0.85
        r = int(10 * (1.0 - f))
        g = int(230 * f)
        b = int(90 * (1.0 - f))
        office_color_map[off] = (r, g, b)
        
    return office_color_map, a_offices, b_offices, c_offices

OFFICE_COLOR_MAP, A_OFFICES, B_OFFICES, C_OFFICES = generate_office_colors(df_a, df_b, df_c)
ALL_OFFICES = sorted(list(set(A_OFFICES + B_OFFICES + C_OFFICES)))

CITY_SEEDS = {
    'さいたま市西区': (760, 388), 'さいたま市北区': (782, 365), 'さいたま市大宮区': (782, 388),
    'さいたま市見沼区': (822, 350), 'さいたま市中央区': (782, 405), 'さいたま市桜区': (760, 430),
    'さいたま市浦和区': (804, 405), 'さいたま市南区': (810, 480), 'さいたま市緑区': (875, 410),
    'さいたま市岩槻区': (855, 325), '川越市': (677, 398), '熊谷市': (588, 175),
    '川口市': (855, 480), '行田市': (651, 180), '秩父市': (240, 390), '所沢市': (634, 475),
    '飯能市': (520, 460), '加須市': (748, 195), '本庄市': (360, 115), '東松山市': (584, 285),
    '春日部市': (895, 355), '狭山市': (590, 425), '羽生市': (708, 145), '鴻巣市': (651, 225),
    '深谷市': (450, 160), '上尾市': (718, 335), '草加市': (928, 471), '越谷市': (928, 410),
    '蕨市': (815, 480), '戸田市': (807, 497), '入間市': (520, 497), '朝霞市': (750, 497),
    '志木市': (740, 480), '和光市': (773, 522), '新座市': (724, 522), '桶川市': (714, 296),
    '久喜市': (775, 230), '北本市': (673, 280), '八潮市': (960, 492), '富士見市': (718, 435),
    '三郷市': (987, 475), '蓮田市': (791, 285), '坂戸市': (635, 330), '幸手市': (876, 226),
    '鶴ヶ島市': (580, 355), '日高市': (543, 395), '吉川市': (987, 411), 'ふじみ野市': (698, 435),
    '白岡市': (822, 295), '伊奈町': (782, 296), '三芳町': (703, 472), '毛呂山町': (511, 355),
    '越生町': (453, 351), '滑川町': (558, 255), '嵐山町': (529, 260), '小川町': (440, 230),
    '川島町': (677, 315), '吉見町': (630, 270), '鳩山町': (538, 315), 'ときがわ町': (490, 280),
    '横瀬町': (335, 355), '皆野町': (360, 260), '長瀞町': (360, 170), '小鹿野町': (175, 290),
    '東秩父村': (400, 295), '美里町': (395, 155), '神川町': (300, 150), '上里町': (363, 34),
    '寄居町': (395, 170), '宮代町': (860, 280), '杉戸町': (876, 280), '松伏町': (960, 355)
}

# 3. データ整理・統合
rows = []
for i in range(len(df_a)):
    code = df_a.loc[i, '市区町村コード']
    city = df_a.loc[i, '市区町村名\n（漢字）']
    
    wt_a = df_a.loc[i, '配達重量\n(日当たり)'] / 1000.0
    wt_b = df_b.loc[i, '配達重量\n(日当たり)'] / 1000.0
    wt_c = df_c.loc[i, '配達重量\n(日当たり)'] / 1000.0
    
    off_a, off_b, off_c = df_a.loc[i, 'A社営業所名'], df_b.loc[i, 'B社営業所名'], df_c.loc[i, 'C社営業所名']
    sub_a = str(df_a.loc[i, 'A社外部委託地域']).strip() == '◯'
    sub_b = str(df_b.loc[i, 'B社外部委託地域']).strip() == '◯'
    sub_c = str(df_c.loc[i, 'C社外部委託地域']).strip() == '◯'
    
    dist_a = round(5.0 + (i * 3 % 17) + (i % 5) * 1.2, 1)
    dist_b = round(4.5 + (i * 5 % 19) + (i % 4) * 1.5, 1)
    dist_c = round(6.0 + (i * 2 % 15) + (i % 6) * 1.1, 1)
    
    tk_a = round(wt_a * dist_a, 2)
    tk_b = round(wt_b * dist_b, 2)
    tk_c = round(wt_c * dist_c, 2)
    
    min_tk = min(tk_a, tk_b, tk_c)
    best_comp = 'A社' if tk_a == min_tk else ('B社' if tk_b == min_tk else 'C社')
    best_off = off_a if best_comp == 'A社' else (off_b if best_comp == 'B社' else off_c)
    init_mode = '委託配達' if (sub_a or sub_b or sub_c) else '自社配達'
    
    rows.append({
        '市区町村コード': code,
        '市区町村名': city,
        '配達方式': init_mode,
        '一括担当': best_comp,
        '一括担当 営業所': best_off,
        'A社担当': False, 'B社担当': False, 'C社担当': False,
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
        best_off = r['A社営業所'] if best == 'A社' else (r['B社営業所'] if best == 'B社' else r['C社営業所'])
        base_df.loc[idx, '一括担当'] = best
        base_df.loc[idx, '一括担当 営業所'] = best_off
        base_df.loc[idx, 'A社担当'] = False; base_df.loc[idx, 'B社担当'] = False; base_df.loc[idx, 'C社担当'] = False
elif "A社一括" in rule:
    base_df['一括担当'] = 'A社'; base_df['一括担当 営業所'] = base_df['A社営業所']
    base_df['A社担当'] = False; base_df['B社担当'] = False; base_df['C社担当'] = False
elif "B社一括" in rule:
    base_df['一括担当'] = 'B社'; base_df['一括担当 営業所'] = base_df['B社営業所']
    base_df['A社担当'] = False; base_df['B社担当'] = False; base_df['C社担当'] = False
elif "C社一括" in rule:
    base_df['一括担当'] = 'C社'; base_df['一括担当 営業所'] = base_df['C社営業所']
    base_df['A社担当'] = False; base_df['B社担当'] = False; base_df['C社担当'] = False
else:
    base_df['一括担当'] = 'なし'; base_df['一括担当 営業所'] = '-'
    base_df['A社担当'] = False; base_df['B社担当'] = False; base_df['C社担当'] = False

# 4. タブUI構成
tab1, tab2, tab3 = st.tabs(["📊 全体サマリー＆現状比較", "🏢 会社別・営業所別集計", "📝 市町村別・受持選択（編集）"])

# タブ3: 個別編集画面
with tab3:
    st.subheader("📝 市町村別・受持選択テーブル")
    edited_df = st.data_editor(
        base_df[['市区町村コード', '市区町村名', '配達方式', '一括担当', '一括担当 営業所', 'A社担当', 'B社担当', 'C社担当', 'A社トンキロ', 'B社トンキロ', 'C社トンキロ']],
        column_config={
            "配達方式": st.column_config.SelectboxColumn("配達方式", options=["自社配達", "委託配達"], required=True),
            "一括担当": st.column_config.SelectboxColumn("市町村全体 一括担当", options=["なし", "A社", "B社", "C社"], required=True),
            "一括担当 営業所": st.column_config.SelectboxColumn("一括担当 営業所", options=ALL_OFFICES + ["-"], required=True),
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
map_status_dict = {}

for idx, r in edited_df.iterrows():
    orig = base_df.loc[idx]
    city_name = orig['市区町村名']
    is_sub_mode = (r['配達方式'] == '委託配達')
    
    has_indiv = r['A社担当'] or r['B社担当'] or r['C社担当']
    target_comp = r['一括担当']
    target_off = r['一括担当 営業所']
    
    # 担当会社と営業所の選択連動補正
    if target_comp == 'A社' and target_off not in A_OFFICES:
        target_off = orig['A社営業所']
    elif target_comp == 'B社' and target_off not in B_OFFICES:
        target_off = orig['B社営業所']
    elif target_comp == 'C社' and target_off not in C_OFFICES:
        target_off = orig['C社営業所']
    elif target_comp == 'なし':
        target_off = '-'
        
    indiv_selected = []
    if r['A社担当']: indiv_selected.append('A社')
    if r['B社担当']: indiv_selected.append('B社')
    if r['C社担当']: indiv_selected.append('C社')
    
    map_status_dict[city_name] = {
        '一括担当': target_comp,
        '一括営業所': target_off,
        '個別選択': indiv_selected
    }
    
    if has_indiv:
        if r['A社担当']:
            sub = is_sub_mode or orig['A社委託']
            active_records.append({
                '市区町村コード': orig['市区町村コード'], '市区町村名': city_name,
                '担当会社': 'A社', '会社表示': 'A社（委託）' if sub else 'A社',
                '担当営業所': orig['A社営業所'], '営業所表示': orig['A社営業所'] + ('（委託）' if sub else ''),
                '重量_t': orig['A社重量_t'], '距離_km': orig['A社距離_km'], 'トンキロ': orig['A社トンキロ'],
            })
        if r['B社担当']:
            sub = is_sub_mode or orig['B社委託']
            active_records.append({
                '市区町村コード': orig['市区町村コード'], '市区町村名': city_name,
                '担当会社': 'B社', '会社表示': 'B社（委託）' if sub else 'B社',
                '担当営業所': orig['B社営業所'], '営業所表示': orig['B社営業所'] + ('（委託）' if sub else ''),
                '重量_t': orig['B社重量_t'], '距離_km': orig['B社距離_km'], 'トンキロ': orig['B社トンキロ'],
            })
        if r['C社担当']:
            sub = is_sub_mode or orig['C社委託']
            active_records.append({
                '市区町村コード': orig['市区町村コード'], '市区町村名': city_name,
                '担当会社': 'C社', '会社表示': 'C社（委託）' if sub else 'C社',
                '担当営業所': orig['C社営業所'], '営業所表示': orig['C社営業所'] + ('（委託）' if sub else ''),
                '重量_t': orig['C社重量_t'], '距離_km': orig['C社距離_km'], 'トンキロ': orig['C社トンキロ'],
            })
    else:
        if target_comp in ['A社', 'B社', 'C社']:
            c_code = target_comp[0]
            sub = is_sub_mode or orig[c_code + '社委託']
            off = target_off if target_off != '-' else orig[c_code + '社営業所']
            
            total_wt = orig['A社重量_t'] + orig['B社重量_t'] + orig['C社重量_t']
            total_tk = orig['A社トンキロ'] + orig['B社トンキロ'] + orig['C社トンキロ']
            avg_dist = orig[c_code + '社距離_km']
            
            active_records.append({
                '市区町村コード': orig['市区町村コード'], '市区町村名': city_name,
                '担当会社': target_comp, '会社表示': target_comp + ('（委託）' if sub else ''),
                '担当営業所': off, '営業所表示': off + ('（委託）' if sub else ''),
                '重量_t': total_wt, '距離_km': avg_dist, 'トンキロ': total_tk,
            })

sim_df = pd.DataFrame(active_records) if len(active_records) > 0 else pd.DataFrame()

# タブ1: サマリー & マップ描画
with tab1:
    st.subheader("📈 会社毎の現状 vs 改正後（シミュレーション）サマリー")
    
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
    comp_compare['削減率(%)'] = (comp_compare['トンキロ削減量'] / comp_compare['現状_トンキロ'] * 100).round(1)
    
    col_a, col_b, col_c = st.columns(3)
    for col, comp_name in zip([col_a, col_b, col_c], ['A社', 'B社', 'C社']):
        row_data = comp_compare[comp_compare['会社'] == comp_name].iloc[0]
        cur_tk = row_data['現状_トンキロ']
        rev_tk = row_data['改正_トンキロ']
        diff_tk = row_data['トンキロ削減量']
        
        with col:
            st.markdown("#### 🏢 " + str(comp_name))
            st.metric("現状 トンキロ", "{:,.1f} ton-km".format(cur_tk))
            st.metric("改正後 トンキロ", "{:,.1f} ton-km".format(rev_tk), delta="削減: {:,.1f} ton-km ({:.1f}%)".format(diff_tk, row_data['削減率(%)']))
            st.caption("改正後 担当件数: {} 件 / 重量: {:,.1f} t".format(int(row_data['改正_担当件数']), row_data['改正_重量_t']))

    st.markdown("---")
    st.write("### 📊 会社毎の比較対比表")
    st.dataframe(
        comp_compare.style.format({
            '現状_重量_t': '{:,.2f}', '現状_トンキロ': '{:,.2f}',
            '改正_重量_t': '{:,.2f}', '改正_トンキロ': '{:,.2f}',
            'トンキロ削減量': '{:,.2f}', '削減率(%)': '{:+.1f}%'
        }),
        use_container_width=True, hide_index=True
    )

    st.markdown("---")
    st.subheader("🗺️ 埼玉県 市町村別受持選択 白地図エリアマップ")
    st.caption("塗り分け：営業所毎の配色（🔴 A社系: 赤グラデーション, 🔵 B社系: 青グラデーション, 🟢 C社系: 緑グラデーション, ⚪ 未設定: 灰）")

    # 凡例表示（HTMLエスケープエラーを防止するため Streamlit のコンテナで表示）
    st.markdown("##### 📌 営業所別 カラー凡例（濃淡グラデーション）")
    leg_cols = st.columns(3)
    
    with leg_cols[0]:
        st.write("**🔴 A社 営業所**")
        for off in A_OFFICES:
            c = OFFICE_COLOR_MAP.get(off, (239, 68, 68))
            hex_c = f"#{c[0]:02x}{c[1]:02x}{c[2]:02x}"
            st.markdown(f' **{off}**', unsafe_allow_html=True)
            
    with leg_cols[1]:
        st.write("**🔵 B社 営業所**")
        for off in B_OFFICES:
            c = OFFICE_COLOR_MAP.get(off, (59, 130, 246))
            hex_c = f"#{c[0]:02x}{c[1]:02x}{c[2]:02x}"
            st.markdown(f' **{off}**', unsafe_allow_html=True)
            
    with leg_cols[2]:
        st.write("**🟢 C社 営業所**")
        for off in C_OFFICES:
            c = OFFICE_COLOR_MAP.get(off, (16, 185, 129))
            hex_c = f"#{c[0]:02x}{c[1]:02x}{c[2]:02x}"
            st.markdown(f' **{off}**', unsafe_allow_html=True)

    st.markdown("