import streamlit as st
import pandas as pd
import numpy as np
import cv2
import os

# 1. ページ基本設定
st.set_page_config(
    page_title="配達エリア・トンキロ最適化シミュレーター（大阪府版）",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🚚 運送2社 営業所受持エリア・トンキロ最適化シミュレーター（大阪府版）")
st.caption("大阪府の市区町村別・会社別の受持選択を行い、「現状」と「改正後」のトンキロ・配達重量の変化をシミュレーションします。")

# 2. ファイルアップロード
st.sidebar.header("📁 データ読み込み")
uploaded_file = st.sidebar.file_uploader("「営業所受持地域.xlsx」をドラッグ＆ドロップ", type=["xlsx"])

if uploaded_file is None:
    st.info("👈 左側のサイドバーから「営業所受持地域.xlsx」をアップロードしてシミュレーションを開始してください。")
    st.stop()

@st.cache_data
def load_data(file):
    xls = pd.ExcelFile(file)
    df_j = pd.read_excel(xls, sheet_name='J')
    df_t = pd.read_excel(xls, sheet_name='T')
    return df_j, df_t

try:
    df_j, df_t = load_data(uploaded_file)
except Exception as e:
    st.error("データの読み込みに失敗しました。J, T シートが含まれるExcelファイルであることを確認してください。")
    st.stop()

# 各社の営業所リストとグラデーションカラーマップを社別に分離生成（同名営業所の色の混同防止）
def generate_office_colors(df_j, df_t):
    j_offices = sorted(list(set(df_j['J社営業所名'].dropna().astype(str).tolist())))
    t_offices = sorted(list(set(df_t['T社営業所名'].dropna().astype(str).tolist())))
    
    j_office_color_map = {}
    t_office_color_map = {}
    
    n_j = len(j_offices)
    for idx, off in enumerate(j_offices):
        f = 0.45 + 0.55 * (idx / max(n_j - 1, 1)) if n_j > 1 else 0.85
        r = int(255 * f)
        g = int(20 * (1.0 - f))
        b = int(20 * (1.0 - f))
        j_office_color_map[off] = (r, g, b)
        
    n_t = len(t_offices)
    for idx, off in enumerate(t_offices):
        f = 0.45 + 0.55 * (idx / max(n_t - 1, 1)) if n_t > 1 else 0.85
        r = int(30 * (1.0 - f))
        g = int(140 * f)
        b = int(255 * f)
        t_office_color_map[off] = (r, g, b)
        
    return j_office_color_map, t_office_color_map, j_offices, t_offices

J_OFFICE_COLOR_MAP, T_OFFICE_COLOR_MAP, J_OFFICES, T_OFFICES = generate_office_colors(df_j, df_t)
ALL_OFFICES = sorted(list(set(J_OFFICES + T_OFFICES)))

# 『座標_5.xlsx』E列（改正③）準拠の基本座標 (1024x1400解像度基準)
CITY_SEEDS = {
    '能勢町': (542, 143),
    '豊能町': (635, 232),
    '島本町': (897, 293),
    '高槻市': (816, 331),
    '枚方市': (923, 331),
    '池田市': (553, 351),
    '箕面市': (618, 359),
    '茨木市': (702, 292),
    '交野市': (932, 517),
    '豊中市': (592, 473),
    '吹田市': (702, 497),
    '摂津市': (750, 517),
    '寝屋川市': (832, 504),
    '守口市': (750, 583),
    '門真市': (818, 584),
    '四條畷市': (923, 584),
    '大東市': (818, 600),
    '東大阪市': (867, 672),
    '大阪市西淀川区': (566, 619),
    '大阪市東淀川区': (670, 550),
    '大阪市淀川区': (603, 593),
    '大阪市北区': (637, 625),
    '大阪市都島区': (680, 594),
    '大阪市旭区': (715, 588),
    '大阪市城東区': (714, 638),
    '大阪市鶴見区': (761, 619),
    '大阪市福島区': (594, 647),
    '大阪市此花区': (545, 673),
    '大阪市西区': (617, 673),
    '大阪市中央区': (653, 671),
    '大阪市東成区': (719, 676),
    '大阪市生野区': (708, 712),
    '大阪市港区': (555, 713),
    '大阪市浪速区': (628, 702),
    '大阪市天王寺区': (665, 707),
    '大阪市大正区': (582, 740),
    '大阪市西成区': (620, 736),
    '大阪市阿倍野区': (660, 741),
    '大阪市平野区': (724, 795),
    '大阪市住之江区': (533, 748),
    '大阪市住吉区': (660, 799),
    '大阪市東住吉区': (680, 799),
    '八尾市': (818, 769),
    '柏原市': (881, 837),
    '松原市': (723, 842),
    '藤井寺市': (817, 853),
    '羽曳野市': (811, 911),
    '太子町': (890, 962),
    '富田林市': (811, 1011),
    '大阪狭山市': (724, 1011),
    '河南町': (866, 1011),
    '千早赤阪村': (881, 1104),
    '河内長野市': (743, 1147),
    '堺市堺区': (607, 848),
    '堺市北区': (665, 850),
    '堺市西区': (565, 930),
    '堺市中区': (640, 930),
    '堺市東区': (682, 911),
    '堺市美原区': (754, 932),
    '堺市南区': (641, 1011),
    '高石市': (535, 930),
    '泉大津市': (504, 1000),
    '忠岡町': (504, 1011),
    '和泉市': (565, 1096),
    '岸和田市': (504, 1179),
    '貝塚市': (494, 1220),
    '熊取町': (426, 1179),
    '泉佐野市': (370, 1179),
    '田尻町': (308, 1187),
    '泉南市': (318, 1251),
    '阪南市': (239, 1280),
    '岬町': (118, 1336)
}

# 連動して着色する離れ島（夢洲・舞洲／関西国際空港島など）の追加座標
EXTRA_SEEDS = {
    '大阪市此花区': [(477, 710), (503, 713)], # 夢洲・舞洲エリア
    '泉佐野市': [(236, 1101)]                 # 関西国際空港島エリア
}

# 3. データ整理・統合
rows = []
for i in range(len(df_j)):
    code = df_j.loc[i, '市区町村コード']
    city = df_j.loc[i, '市区町村名\n（漢字）']
    
    wt_j = df_j.loc[i, '配達重量\n(日当たり)'] / 1000.0
    wt_t = df_t.loc[i, '配達重量\n(日当たり)'] / 1000.0
    
    off_j, off_t = df_j.loc[i, 'J社営業所名'], df_t.loc[i, 'T社営業所名']
    sub_j = str(df_j.loc[i, 'J社外部委託地域']).strip() == '◯'
    sub_t = str(df_t.loc[i, 'T社外部委託地域']).strip() == '◯'
    
    dist_j = round(5.0 + (i * 3 % 17) + (i % 5) * 1.2, 1)
    dist_t = round(4.5 + (i * 5 % 19) + (i % 4) * 1.5, 1)
    
    tk_j = round(wt_j * dist_j, 2)
    tk_t = round(wt_t * dist_t, 2)
    
    best_comp = 'J社' if tk_j <= tk_t else 'T社'
    best_off = off_j if best_comp == 'J社' else off_t
    init_mode = '委託配達' if (sub_j or sub_t) else '自社配達'
    
    rows.append({
        '市区町村コード': code,
        '市区町村名': city,
        '配達方式': init_mode,
        '一括担当': best_comp,
        '一括担当 営業所': best_off,
        'J社担当': False, 'T社担当': False,
        'J社営業所': off_j, 'J社委託': sub_j, 'J社重量_t': wt_j, 'J社距離_km': dist_j, 'J社トンキロ': tk_j,
        'T社営業所': off_t, 'T社委託': sub_t, 'T社重量_t': wt_t, 'T社距離_km': dist_t, 'T社トンキロ': tk_t,
    })

base_df = pd.DataFrame(rows)

# サイドバー設定
st.sidebar.markdown("---")
st.sidebar.header("🎯 自動一括割り当て")
rule = st.sidebar.radio(
    "一括設定ルールを選択",
    ["最少トンキロ最適化（初期自動選択）", "全市町村 J社一括", "全市町村 T社一括", "選択クリア"]
)

if "最少トンキロ" in rule:
    for idx, r in base_df.iterrows():
        best = 'J社' if r['J社トンキロ'] <= r['T社トンキロ'] else 'T社'
        best_off = r['J社営業所'] if best == 'J社' else r['T社営業所']
        base_df.loc[idx, '一括担当'] = best
        base_df.loc[idx, '一括担当 営業所'] = best_off
        base_df.loc[idx, 'J社担当'] = False; base_df.loc[idx, 'T社担当'] = False
elif "J社一括" in rule:
    base_df['一括担当'] = 'J社'; base_df['一括担当 営業所'] = base_df['J社営業所']
    base_df['J社担当'] = False; base_df['T社担当'] = False
elif "T社一括" in rule:
    base_df['一括担当'] = 'T社'; base_df['一括担当 営業所'] = base_df['T社営業所']
    base_df['J社担当'] = False; base_df['T社担当'] = False
else:
    base_df['一括担当'] = 'なし'; base_df['一括担当 営業所'] = '-'
    base_df['J社担当'] = False; base_df['T社担当'] = False

# 4. タブUI構成
tab1, tab2, tab3 = st.tabs(["📊 全体サマリー＆現状比較", "🏢 会社別・営業所別集計", "📝 市町村別・受持選択（編集）"])

active_records = []
map_status_dict = {}

# タブ3: 個別編集画面 ＆ マップ（下部配置）
with tab3:
    st.subheader("📝 市町村別・受持選択テーブル")
    edited_df = st.data_editor(
        base_df[['市区町村コード', '市区町村名', '配達方式', '一括担当', '一括担当 営業所', 'J社担当', 'T社担当', 'J社トンキロ', 'T社トンキロ']],
        column_config={
            "配達方式": st.column_config.SelectboxColumn("配達方式", options=["自社配達", "委託配達"], required=True),
            "一括担当": st.column_config.SelectboxColumn("市町村全体 一括担当", options=["なし", "J社", "T社"], required=True),
            "一括担当 営業所": st.column_config.SelectboxColumn("一括担当 営業所", options=ALL_OFFICES + ["-"], required=True),
            "J社担当": st.column_config.CheckboxColumn("J社 個別", default=False),
            "T社担当": st.column_config.CheckboxColumn("T社 個別", default=False),
        },
        disabled=['市区町村コード', '市区町村名', 'J社トンキロ', 'T社トンキロ'],
        use_container_width=True,
        hide_index=True
    )

    # 編集データの処理
    for idx, r in edited_df.iterrows():
        orig = base_df.loc[idx]
        city_name = orig['市区町村名']
        is_sub_mode = (r['配達方式'] == '委託配達')
        
        has_indiv = r['J社担当'] or r['T社担当']
        target_comp = r['一括担当']
        target_off = r['一括担当 営業所']
        
        if target_comp == 'J社' and target_off not in J_OFFICES:
            target_off = orig['J社営業所']
        elif target_comp == 'T社' and target_off not in T_OFFICES:
            target_off = orig['T社営業所']
        elif target_comp == 'なし':
            target_off = '-'
            
        indiv_selected = []
        if r['J社担当']: indiv_selected.append('J社')
        if r['T社担当']: indiv_selected.append('T社')
        
        map_status_dict[city_name] = {
            '一括担当': target_comp,
            '一括営業所': target_off,
            '個別選択': indiv_selected
        }
        
        if has_indiv:
            if r['J社担当']:
                sub = is_sub_mode or orig['J社委託']
                active_records.append({
                    '市区町村コード': orig['市区町村コード'], '市区町村名': city_name,
                    '担当会社': 'J社', '会社表示': 'J社（委託）' if sub else 'J社',
                    '担当営業所': orig['J社営業所'], '営業所表示': orig['J社営業所'] + ('（委託）' if sub else ''),
                    '重量_t': orig['J社重量_t'], '距離_km': orig['J社距離_km'], 'トンキロ': orig['J社トンキロ'],
                })
            if r['T社担当']:
                sub = is_sub_mode or orig['T社委託']
                active_records.append({
                    '市区町村コード': orig['市区町村コード'], '市区町村名': city_name,
                    '担当会社': 'T社', '会社表示': 'T社（委託）' if sub else 'T社',
                    '担当営業所': orig['T社営業所'], '営業所表示': orig['T社営業所'] + ('（委託）' if sub else ''),
                    '重量_t': orig['T社重量_t'], '距離_km': orig['T社距離_km'], 'トンキロ': orig['T社トンキロ'],
                })
        else:
            if target_comp in ['J社', 'T社']:
                c_code = target_comp[0]
                sub = is_sub_mode or orig[c_code + '社委託']
                off = target_off if target_off != '-' else orig[c_code + '社営業所']
                
                total_wt = orig['J社重量_t'] + orig['T社重量_t']
                total_tk = orig['J社トンキロ'] + orig['T社トンキロ']
                avg_dist = orig[c_code + '社距離_km']
                
                active_records.append({
                    '市区町村コード': orig['市区町村コード'], '市区町村名': city_name,
                    '担当会社': target_comp, '会社表示': target_comp + ('（委託）' if sub else ''),
                    '担当営業所': off, '営業所表示': off + ('（委託）' if sub else ''),
                    '重量_t': total_wt, '距離_km': avg_dist, 'トンキロ': total_tk,
                })

    st.markdown("---")
    st.subheader("🗺️ 大阪府 市区町村別受持選択 白地図エリアマップ")
    st.caption("塗り分け：営業所毎の配色（🔴 J社系: 赤グラデーション, 🔵 T社系: 青グラデーション, ⚪ 未設定: 灰）")

    # 凡例表示
    leg_cols = st.columns(2)
    with leg_cols[0]:
        st.write("**🔴 J社 営業所**")
        for off in J_OFFICES:
            c = J_OFFICE_COLOR_MAP.get(off, (239, 68, 68))
            hex_c = f"#{c[0]:02x}{c[1]:02x}{c[2]:02x}"
            st.write("■ " + str(off) + " (" + str(hex_c) + ")")
            
    with leg_cols[1]:
        st.write("**🔵 T社 営業所**")
        for off in T_OFFICES:
            c = T_OFFICE_COLOR_MAP.get(off, (59, 130, 246))
            hex_c = f"#{c[0]:02x}{c[1]:02x}{c[2]:02x}"
            st.write("■ " + str(off) + " (" + str(hex_c) + ")")

    st.write("")

    # 画像読み込み
    map_img_path = None
    for target_path in ["画像4_2.png", "画像4_2.jpg", "画像4.png", "画像3.jpg"]:
        if os.path.exists(target_path):
            map_img_path = target_path
            break

    if map_img_path and os.path.exists(map_img_path):
        src_img = cv2.imread(map_img_path)
        src_img = cv2.resize(src_img, (1024, 1400), interpolation=cv2.INTER_AREA)
        h, w, _ = src_img.shape
        
        gray = cv2.cvtColor(src_img, cv2.COLOR_BGR2GRAY)
        line_bin = (gray > 40).astype(np.uint8) * 255
            
        kernel = np.ones((3, 3), np.uint8)
        line_bin_dilated = cv2.dilate(line_bin, kernel, iterations=1)
        
        canvas_rgb = np.full((h, w, 3), 255, dtype=np.uint8)
        canvas_rgb[line_bin_dilated == 255] = (0, 0, 0)
        
        bg_protection_mask = np.zeros((h + 2, w + 2), np.uint8)
        cv2.floodFill(canvas_rgb.copy(), bg_protection_mask, (0, 0), (255, 255, 255), (15, 15, 15), (15, 15, 15), cv2.FLOODFILL_FIXED_RANGE)

        for c_name, (sx, sy) in CITY_SEEDS.items():
            info = map_status_dict.get(c_name, {'一括担当': 'なし', '一括営業所': '-', '個別選択': []})
            bulk = info['一括担当']
            bulk_off = info['一括営業所']
            indivs = info['個別選択']
            
            # 社別にカラーマップを参照（同名営業所の色混同を完全防止）
            if bulk == 'J社' and bulk_off in J_OFFICE_COLOR_MAP:
                fill_rgb = J_OFFICE_COLOR_MAP[bulk_off]
            elif bulk == 'T社' and bulk_off in T_OFFICE_COLOR_MAP:
                fill_rgb = T_OFFICE_COLOR_MAP[bulk_off]
            else:
                DEFAULT_RGB = {'J社': (239, 68, 68), 'T社': (59, 130, 246), 'なし': (220, 225, 230)}
                fill_rgb = DEFAULT_RGB.get(bulk, (220, 225, 230))
                
            if bulk == 'なし' and len(indivs) > 0:
                fill_rgb = (220, 225, 230)

            # 塗りつぶし対象の全座標（本土＋離島追加分）
            target_coords = [(sx, sy)]
            if c_name in EXTRA_SEEDS:
                target_coords.extend(EXTRA_SEEDS[c_name])
                
            for cs_x, cs_y in target_coords:
                if 0 <= cs_x < w and 0 <= cs_y < h:
                    # 黒枠線の上の場合は周囲の空白ピクセルを自動探索
                    if line_bin_dilated[cs_y, cs_x] == 255:
                        found = False
                        for r in range(1, 15):
                            for dy in range(-r, r+1):
                                for dx in range(-r, r+1):
                                    nx, ny = cs_x + dx, cs_y + dy
                                    if 0 <= nx < w and 0 <= ny < h and line_bin_dilated[ny, nx] == 0:
                                        cs_x, cs_y = nx, ny
                                        found = True
                                        break
                                if found: break
                            if found: break

                    if line_bin_dilated[cs_y, cs_x] == 0:
                        m_curr = bg_protection_mask.copy()
                        cv2.floodFill(canvas_rgb, m_curr, (cs_x, cs_y), fill_rgb, (15, 15, 15), (15, 15, 15), cv2.FLOODFILL_FIXED_RANGE)
                    
            if 0 <= sx < w and 0 <= sy < h and len(indivs) > 0:
                DEFAULT_RGB = {'J社': (239, 68, 68), 'T社': (59, 130, 246)}
                for d_idx, comp_indiv in enumerate(indivs):
                    dot_rgb = DEFAULT_RGB.get(comp_indiv, (0, 0, 0))
                    dot_x = sx + (d_idx - (len(indivs)-1)/2.0) * 16
                    cv2.circle(canvas_rgb, (int(dot_x), sy), 7, (255, 255, 255), -1)
                    cv2.circle(canvas_rgb, (int(dot_x), sy), 6, dot_rgb, -1)

        canvas_rgb[line_bin_dilated == 255] = (0, 0, 0)

        gray_check = cv2.cvtColor(canvas_rgb, cv2.COLOR_RGB2GRAY)
        non_bg = np.where(gray_check < 250)
        
        if len(non_bg[0]) > 0:
            min_y, max_y = np.min(non_bg[0]), np.max(non_bg[0])
            min_x, max_x = np.min(non_bg[1]), np.max(non_bg[1])
            pad = 12
            crop_min_y = max(0, min_y - pad)
            crop_max_y = min(h, max_y + pad)
            crop_min_x = max(0, min_x - pad)
            crop_max_x = min(w, max_x + pad)
            cropped_img = canvas_rgb[crop_min_y:crop_max_y, crop_min_x:crop_max_x]
        else:
            cropped_img = canvas_rgb

        # 地図表示
        map_col1, map_col2, map_col3 = st.columns([1, 4, 1])
        with map_col2:
            st.image(cropped_img, width=600)
    else:
        st.warning("マップ画像が見つかりません。`画像4_2.png` を配置してください。")

sim_df = pd.DataFrame(active_records) if len(active_records) > 0 else pd.DataFrame()

# タブ1: サマリー
with tab1:
    st.subheader("📈 会社毎の現状 vs 改正後（シミュレーション）サマリー")
    
    cur_comp = pd.DataFrame([
        {'会社': 'J社', '現状_自治体数': len(base_df), '現状_重量_t': base_df['J社重量_t'].sum(), '現状_トンキロ': base_df['J社トンキロ'].sum()},
        {'会社': 'T社', '現状_自治体数': len(base_df), '現状_重量_t': base_df['T社重量_t'].sum(), '現状_トンキロ': base_df['T社トンキロ'].sum()},
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
    
    col_j, col_t = st.columns(2)
    for col, comp_name in zip([col_j, col_t], ['J社', 'T社']):
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

# タブ2: 会社別 & 営業所別 詳細集計
with tab2:
    st.subheader("🏢 会社別 集計（自社配達 vs 外部委託）")
    st.caption("※ 「（委託）」指定の地域は区分して集計しています。")
    
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
    st.caption("※ 営業所ごとに「（委託）」区分で行を分けて表示しています。")
    
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