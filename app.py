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

# 「座標3.xlsx」のB列（改正列）のデータに基づく最新の領域シード座標 (X, Y)
CITY_SEEDS = {
    'さいたま市西区': (745, 388),
    'さいたま市北区': (782, 365),
    'さいたま市大宮区': (760, 388),
    'さいたま市見沼区': (822, 350),
    'さいたま市中央区': (782, 405),
    'さいたま市桜区': (760, 430),
    'さいたま市浦和区': (804, 405),
    'さいたま市南区': (810, 480),
    'さいたま市緑区': (875, 410),
    'さいたま市岩槻区': (855, 325),
    '川越市': (677, 398),
    '熊谷市': (588, 175),
    '川口市': (855, 480),
    '行田市': (651, 180),
    '秩父市': (240, 390),
    '所沢市': (634, 475),
    '飯能市': (520, 460),
    '加須市': (748, 195),
    '本庄市': (360, 115),
    '東松山市': (584, 285),
    '春日部市': (895, 355),
    '狭山市': (590, 425),
    '羽生市': (708, 145),
    '鴻巣市': (651, 225),
    '深谷市': (450, 160),
    '上尾市': (718, 335),
    '草加市': (928, 471),
    '越谷市': (928, 410),
    '蕨市': (804, 497),
    '戸田市': (807, 497),
    '入間市': (520, 497),
    '朝霞市': (750, 497),
    '志木市': (740, 480),
    '和光市': (773, 522),
    '新座市': (724, 522),
    '桶川市': (714, 296),
    '久喜市': (775, 230),
    '北本市': (673, 280),
    '八潮市': (960, 492),
    '富士見市': (718, 435),
    '三郷市': (987, 475),
    '蓮田市': (791, 285),
    '坂戸市': (635, 330),
    '幸手市': (876, 226),
    '鶴ヶ島市': (580, 355),
    '日高市': (543, 395),
    '吉川市': (987, 411),
    'ふじみ野市': (698, 435),
    '白岡市': (822, 295),
    '伊奈町': (782, 296),
    '三芳町': (703, 472),
    '毛呂山町': (511, 355),
    '越生町': (460, 330),
    '滑川町': (558, 255),
    '嵐山町': (529, 260),
    '小川町': (440, 230),
    '川島町': (677, 315),
    '吉見町': (630, 270),
    '鳩山町': (538, 315),
    'ときがわ町': (490, 280),
    '横瀬町': (335, 355),
    '皆野町': (360, 260),
    '長瀞町': (360, 170),
    '小鹿野町': (175, 290),
    '東秩父村': (400, 295),
    '美里町': (395, 180),
    '神川町': (300, 150),
    '上里町': (363, 34),
    '寄居町': (395, 170),
    '宮代町': (860, 280),
    '杉戸町': (876, 280),
    '松伏町': (960, 355)
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
    init_mode = '委託配達' if (sub_a or sub_b or sub_c) else '自社配達'
    
    rows.append({
        '市区町村コード': code,
        '市区町村名': city,
        '配達方式': init_mode,
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
    edited_df = st.data_editor(
        base_df[['市区町村コード', '市区町村名', '配達方式', '一括担当', 'A社担当', 'B社担当', 'C社担当', 'A社トンキロ', 'B社トンキロ', 'C社トンキロ']],
        column_config={
            "配達方式": st.column_config.SelectboxColumn("配達方式", options=["自社配達", "委託配達"], required=True),
            "一括担当": st.column_config.SelectboxColumn("市町村全体 一括担当", options=["なし", "A社", "B社", "C社"], required=True),
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
    
    indiv_selected = []
    if r['A社担当']: indiv_selected.append('A社')
    if r['B社担当']: indiv_selected.append('B社')
    if r['C社担当']: indiv_selected.append('C社')
    
    map_status_dict[city_name] = {
        '一括担当': target_comp,
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
    st.caption("塗り分け：一括担当（🔴 A社: 赤, 🔵 B社: 青, 🟢 C社: 緑, ⚪ なし: 灰） / ドット：個別選択された会社の色")

    map_img_path = None
    for target_path in ["20261001_bc6e30f7720a548fb561a31_2.jpg", "20261001_bc6e30f7720a548fb561a31_2.png", "20261001_bc6e30f7720a548fb561a31.jpg", "20261001_bc6e30f7720a548fb561a31.png"]:
        if os.path.exists(target_path):
            map_img_path = target_path
            break

    if map_img_path and os.path.exists(map_img_path):
        src_img = cv2.imread(map_img_path)
        h, w, _ = src_img.shape
        
        # 1. 境界線の二値化と線強調
        gray = cv2.cvtColor(src_img, cv2.COLOR_BGR2GRAY)
        
        if np.mean(gray) < 100:
            line_bin = (gray > 25).astype(np.uint8) * 255
        else:
            line_bin = (gray < 200).astype(np.uint8) * 255
            
        # 純粋な RGB キャンバス作成
        kernel = np.ones((3, 3), np.uint8)
        line_bin_dilated = cv2.dilate(line_bin, kernel, iterations=1)
        
        canvas_rgb = np.full((h, w, 3), 255, dtype=np.uint8)
        canvas_rgb[line_bin_dilated == 255] = (0, 0, 0) # 黒色境界線
        
        # RGB カラーマップ定義
        RGB_MAP = {
            'A社': (239, 68, 68),
            'B社': (59, 130, 246),
            'C社': (16, 185, 129),
            'なし': (220, 225, 230)
        }
        
        # 県外（マップ外側の背景領域）ペイント保護マスク
        bg_protection_mask = np.zeros((h + 2, w + 2), np.uint8)
        cv2.floodFill(canvas_rgb.copy(), bg_protection_mask, (0, 0), (255, 255, 255), (15, 15, 15), (15, 15, 15), cv2.FLOODFILL_FIXED_RANGE)

        # 1. 各72市区町村のメイン領域をFloodFillペイント
        for c_name, (sx, sy) in CITY_SEEDS.items():
            info = map_status_dict.get(c_name, {'一括担当': 'なし', '個別選択': []})
            bulk = info['一括担当']
            indivs = info['個別選択']
            
            fill_rgb = RGB_MAP.get(bulk, (220, 225, 230))
            if bulk == 'なし' and len(indivs) > 0:
                fill_rgb = (220, 225, 230)
                
            if 0 <= sx < w and 0 <= sy < h:
                # 境界線上（黒色）にシードが当たっている場合は近傍の広い白地（>200）へ全自動退避
                if np.mean(canvas_rgb[sy, sx]) < 150:
                    found = False
                    for r in range(1, 20):
                        for dy in range(-r, r+1, 2):
                            for dx in range(-r, r+1, 2):
                                nx, ny = sx + dx, sy + dy
                                if 0 <= nx < w and 0 <= ny < h and np.mean(canvas_rgb[ny, nx]) > 200:
                                    sx, sy = nx, ny
                                    found = True
                                    break
                            if found: break
                        if found: break
                
                # 保護マスクをコピー（.copy()）してペイント実行
                m_curr = bg_protection_mask.copy()
                cv2.floodFill(canvas_rgb, m_curr, (sx, sy), fill_rgb, (15, 15, 15), (15, 15, 15), cv2.FLOODFILL_FIXED_RANGE)
                
                # ドット描画（個別選択時）
                if len(indivs) > 0:
                    for d_idx, comp_indiv in enumerate(indivs):
                        dot_rgb = RGB_MAP.get(comp_indiv, (0, 0, 0))
                        dot_x = sx + (d_idx - (len(indivs)-1)/2.0) * 16
                        cv2.circle(canvas_rgb, (int(dot_x), sy), 7, (255, 255, 255), -1)
                        cv2.circle(canvas_rgb, (int(dot_x), sy), 6, dot_rgb, -1)

        # 2. 埼玉県内の「小さな未塗り白地スペース」を精密に自動補填（area < 1200）
        gray_temp = cv2.cvtColor(canvas_rgb, cv2.COLOR_RGB2GRAY)
        white_holes = (gray_temp > 250).astype(np.uint8)
        num_holes, labels_holes, stats_holes, _ = cv2.connectedComponentsWithStats(white_holes)
        bg_hole_label = labels_holes[0, 0]

        for i in range(1, num_holes):
            area = stats_holes[i, cv2.CC_STAT_AREA]
            cx = int(stats_holes[i, cv2.CC_STAT_LEFT] + stats_holes[i, cv2.CC_STAT_WIDTH]/2)
            cy = int(stats_holes[i, cv2.CC_STAT_TOP] + stats_holes[i, cv2.CC_STAT_HEIGHT]/2)
            
            if i != bg_hole_label and area < 1200:
                found_color = None
                for r in range(1, 25):
                    for dy in range(-r, r+1, 3):
                        for dx in range(-r, r+1, 3):
                            nx, ny = cx + dx, cy + dy
                            if 0 <= nx < w and 0 <= ny < h:
                                p_col = canvas_rgb[ny, nx]
                                if np.mean(p_col) < 230 and np.mean(p_col) > 10:
                                    found_color = tuple(int(c) for c in p_col)
                                    break
                        if found_color is not None: break
                    if found_color is not None: break
                
                if found_color is not None:
                    m_hole = bg_protection_mask.copy()
                    cv2.floodFill(canvas_rgb, m_hole, (cx, cy), found_color, (15, 15, 15), (15, 15, 15), cv2.FLOODFILL_FIXED_RANGE)

        # 3. 不要な背景余白を自動クロップしてフィット表示
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

        st.image(cropped_img, use_container_width=True)
    else:
        st.warning("マップ画像が見つかりません。リポジトリに画像を配置してください。")

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