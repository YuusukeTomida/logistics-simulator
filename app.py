import streamlit as st
import pandas as pd
import numpy as np
import folium
from streamlit_folium import st_folium

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

# 埼玉県72市区町村の推計境界ボックス（ポリゴン塗りつぶし用）座標データ
CITY_BOUNDS = {
    'さいたま市西区': [35.90, 139.56, 35.93, 139.60], 'さいたま市北区': [35.92, 139.60, 35.95, 139.64],
    'さいたま市大宮区': [35.89, 139.61, 35.92, 139.64], 'さいたま市見沼区': [35.91, 139.64, 35.95, 139.68],
    'さいたま市中央区': [35.87, 139.60, 35.90, 139.63], 'さいたま市桜区': [35.84, 139.59, 35.88, 139.63],
    'さいたま市浦和区': [35.85, 139.63, 35.89, 139.67], 'さいたま市南区': [35.83, 139.64, 35.86, 139.68],
    'さいたま市緑区': [35.85, 139.68, 35.89, 139.72], 'さいたま市岩槻区': [35.93, 139.67, 35.97, 139.72],
    '川越市': [35.89, 139.44, 35.95, 139.52], '熊谷市': [36.10, 139.34, 36.19, 139.43],
    '川口市': [35.78, 139.69, 35.84, 139.76], '行田市': [36.10, 139.42, 36.17, 139.49],
    '秩父市': [35.90, 138.98, 36.08, 139.18], '所沢市': [35.77, 139.42, 35.83, 139.51],
    '飯能市': [35.81, 139.18, 35.90, 139.36], '加須市': [36.08, 139.55, 36.17, 139.64],
    '本庄市': [36.19, 139.14, 36.28, 139.23], '東松山市': [36.00, 139.35, 36.08, 139.44],
    '春日部市': [35.94, 139.71, 36.01, 139.80], '狭山市': [35.82, 139.37, 35.88, 139.45],
    '羽生市': [36.13, 139.50, 36.21, 139.58], '鴻巣市': [36.02, 139.47, 36.10, 139.55],
    '深谷市': [36.14, 139.23, 36.24, 139.33], '上尾市': [35.94, 139.55, 36.01, 139.63],
    '草加市': [35.80, 139.77, 35.85, 139.83], '越谷市': [35.76, 139.76, 35.83, 139.83],
    '蕨市': [35.81, 139.66, 35.84, 139.70], '戸田市': [35.79, 139.65, 35.83, 139.70],
    '入間市': [35.80, 139.34, 35.87, 139.42], '朝霞市': [35.79, 139.57, 35.84, 139.62],
    '志木市': [35.81, 139.55, 35.84, 139.60], '和光市': [35.76, 139.58, 35.80, 139.63],
    '新座市': [35.77, 139.52, 35.82, 139.58], '桶川市': [35.98, 139.52, 36.03, 139.59],
    '久喜市': [36.02, 139.62, 36.10, 139.71], '北本市': [36.01, 139.50, 36.05, 139.56],
    '八潮市': [35.80, 139.81, 35.84, 139.86], '富士見市': [35.83, 139.52, 35.88, 139.57],
    '三郷市': [35.80, 139.84, 35.86, 139.90], '蓮田市': [35.95, 139.62, 36.01, 139.68],
    '坂戸市': [35.93, 139.35, 35.98, 139.43], '幸手市': [36.04, 139.69, 36.11, 139.76],
    '鶴ヶ島市': [35.91, 139.36, 35.95, 139.42], '日高市': [35.86, 139.29, 35.92, 139.37],
    '吉川市': [35.86, 139.81, 35.92, 139.87], 'ふじみ野市': [35.85, 139.49, 35.90, 139.55],
    '白岡市': [35.99, 139.63, 36.04, 139.69], '伊奈町': [35.97, 139.60, 36.02, 139.64],
    '三芳町': [35.81, 139.50, 35.85, 139.55], '毛呂山町': [35.91, 139.26, 35.97, 139.34],
    '越生町': [35.93, 139.24, 35.99, 139.32], '滑川町': [36.02, 139.35, 36.07, 139.41],
    '嵐山町': [36.02, 139.30, 36.07, 139.36], '小川町': [36.03, 139.22, 36.09, 139.30],
    '川島町': [35.96, 139.44, 36.01, 139.52], '吉見町': [36.01, 139.41, 36.07, 139.49],
    '鳩山町': [35.95, 139.29, 36.01, 139.36], 'ときがわ町': [35.97, 139.23, 36.03, 139.31],
    '横瀬町': [35.95, 139.07, 36.01, 139.14], '皆野町': [36.04, 139.06, 36.10, 139.13],
    '長瀞町': [36.09, 139.08, 36.14, 139.14], '小鹿野町': [35.98, 138.92, 36.05, 139.05],
    '東秩父村': [36.03, 139.14, 36.09, 139.22], '美里町': [36.15, 139.15, 36.21, 139.22],
    '神川町': [36.13, 139.04, 36.22, 139.14], '上里町': [36.22, 139.10, 36.28, 139.17],
    '寄居町': [36.08, 139.15, 36.15, 139.23], '宮代町': [36.00, 139.69, 36.05, 139.75],
    '杉戸町': [36.00, 139.75, 36.06, 139.83], '松伏町': [35.90, 139.79, 35.96, 139.85]
}

# 3. データ整理・距離／トンキロ・外部委託フラグの統合構築
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
    
    # 初期配列
    init_mode = '委託配達' if (sub_a or sub_b or sub_c) else '自社配達'
    
    rows.append({
        '市区町村コード': code,
        '市区町村名': city,
        '配達方式': init_mode, # 新規追加列（2列目と3列目の間）
        '一括担当': best_comp,
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
    - **配達方式（列2）**: 「自社配達」または「委託配達」を選択できます。
    - **一括担当（列3）**: 「A社」「B社」「C社」から1つ選択すると、市町村全体の荷物が選択した会社に一括集計されます。
    - **個社チェック（列4〜6）**: チェックを入れた会社の荷物のみが集計されます（個別選択が最優先されます）。
    """)

    edited_df = st.data_editor(
        base_df[['市区町村コード', '市区町村名', '配達方式', '一括担当', 'A社担当', 'B社担当', 'C社担当', 'A社トンキロ', 'B社トンキロ', 'C社トンキロ']],
        column_config={
            "配達方式": st.column_config.SelectboxColumn(
                "配達方式",
                options=["自社配達", "委託配達"],
                required=True
            ),
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
map_status_list = []

for idx, r in edited_df.iterrows():
    orig = base_df.loc[idx]
    city_name = orig['市区町村名']
    deliv_mode = r['配達方式'] # 追加された配達方式
    is_sub_mode = (deliv_mode == '委託配達')
    
    has_indiv = r['A社担当'] or r['B社担当'] or r['C社担当']
    target_comp = r['一括担当']
    
    indiv_selected = []
    if r['A社担当']: indiv_selected.append('A社')
    if r['B社担当']: indiv_selected.append('B社')
    if r['C社担当']: indiv_selected.append('C社')
    
    map_status_list.append({
        '市区町村名': city_name,
        '一括担当': target_comp,
        '個別選択': indiv_selected
    })
    
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
            off = orig[c_code + '社営業所']
            
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

# タブ1: 全体サマリー & 現状 vs 改正比較 & 白地図マップ
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
    st.caption("エリア塗りつぶし：一括担当（赤: A社, 青: B社, 緑: C社, 灰: なし） / ドット：個別選択された会社の色")

    # APIキー不要の白地図（CartoDB light_nolabels）を使用
    m = folium.Map(
        location=[35.98, 139.40],
        zoom_start=10,
        min_zoom=9,
        max_zoom=12,
        max_bounds=True,
        min_lat=35.6,
        max_lat=36.4,
        min_lon=138.7,
        max_lon=140.0,
        tiles="https://{s}.basemaps.cartocdn.com/light_nolabels/{z}/{x}/{y}{r}.png",
        attr="© [OpenStreetMap](https://www.openstreetmap.org/copyright) © [CARTO](https://carto.com/attributions)"
    )
    
    COLOR_MAP = {
        'A社': '#EF4444',
        'B社': '#3B82F6',
        'C社': '#10B981',
        'なし': '#CBD5E1'
    }

    for item in map_status_list:
        c_name = item['市区町村名']
        bulk = item['一括担当']
        indivs = item['個別選択']
        
        bounds = CITY_BOUNDS.get(c_name, [35.9, 139.4, 35.95, 139.45])
        
        fill_col = COLOR_MAP.get(bulk, '#CBD5E1')
        if bulk == 'なし' and len(indivs) > 0:
            fill_col = '#CBD5E1'
            
        indiv_str = ', '.join(indivs) if len(indivs) > 0 else 'なし'
        
        # 市町村全体のエリア矩形ポリゴン（塗りつぶし白地図表現）
        folium.Rectangle(
            bounds=[[bounds[0], bounds[1]], [bounds[2], bounds[3]]],
            color="#64748B",
            weight=1,
            fill=True,
            fill_color=fill_col,
            fill_opacity=0.65,
            popup=str(c_name) + " | 一括: " + str(bulk) + " | 個別: " + str(indiv_str),
            tooltip=str(c_name)
        ).add_to(m)
        
        # ドット表示（個別選択ありの場合）
        if len(indivs) > 0:
            center_lat = (bounds[0] + bounds[2]) / 2.0
            center_lon = (bounds[1] + bounds[3]) / 2.0
            
            for d_idx, comp_indiv in enumerate(indivs):
                dot_color = COLOR_MAP.get(comp_indiv, '#000000')
                offset_lat = center_lat + (d_idx - (len(indivs)-1)/2.0) * 0.012
                folium.CircleMarker(
                    location=[offset_lat, center_lon],
                    radius=5,
                    color='#FFFFFF',
                    weight=1,
                    fill=True,
                    fill_color=dot_color,
                    fill_opacity=1.0,
                    popup=str(c_name) + " - 個別: " + str(comp_indiv),
                    tooltip=str(c_name) + " (" + str(comp_indiv) + ")"
                ).add_to(m)

    st_folium(m, width="100%", height=520)

# タブ2: 会社別 & 営業所別 詳細集計
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