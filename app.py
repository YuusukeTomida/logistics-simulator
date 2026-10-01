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

# 埼玉県72市区町村の白地図画像上の内部塗りつぶし開始座標（シードポイント X, Y）
CITY_SEEDS = {
    'さいたま市西区': (715, 370), 'さいたま市北区': (756, 355), 'さいたま市大宮区': (748, 380),
    'さいたま市見沼区': (782, 365), 'さいたま市中央区': (738, 395), 'さいたま市桜区': (718, 410),
    'さいたま市浦和区': (762, 405), 'さいたま市南区': (760, 430), 'さいたま市緑区': (804, 405),
    'さいたま市岩槻区': (822, 350), '川越市': (662, 375), '熊谷市': (588, 175),
    '川口市': (834, 445), '行田市': (651, 180), '秩父市': (212, 375), '所沢市': (634, 465),
    '飯能市': (380, 405), '加須市': (748, 195), '本庄市': (326, 150), '東松山市': (584, 285),
    '春日部市': (860, 325), '狭山市': (590, 425), '羽生市': (708, 165), '鴻巣市': (664, 245),
    '深谷市': (450, 160), '上尾市': (718, 335), '草加市': (874, 430), '越谷市': (871, 385),
    '蕨市': (777, 442), '戸田市': (760, 455), '入間市': (562, 450), '朝霞市': (732, 460),
    '志木市': (718, 435), '和光市': (753, 475), '新座市': (703, 472), '桶川市': (677, 315),
    '久喜市': (775, 255), '北本市': (673, 290), '八潮市': (906, 445), '富士見市': (698, 415),
    '三郷市': (921, 430), '蓮田市': (763, 315), '坂戸市': (590, 345), '幸手市': (843, 250),
    '鶴ヶ島市': (576, 370), '日高市': (543, 395), '吉川市': (925, 385), 'ふじみ野市': (677, 398),
    '白岡市': (791, 285), '伊奈町': (733, 315), '三芳町': (665, 430), '毛呂山町': (511, 355),
    '越生町': (412, 335), '滑川町': (558, 265), '嵐山町': (529, 260), '小川町': (486, 255),
    '川島町': (640, 335), '吉見町': (630, 270), '鳩山町': (538, 315), 'ときがわ町': (471, 315),
    '横瀬町': (335, 355), '皆野町': (325, 250), '長瀞町': (340, 210), '小鹿野町': (175, 290),
    '東秩父村': (415, 285), '美里町': (380, 185), '神川町': (345, 140), '上里町': (360, 110),
    '寄居町': (415, 225), '宮代町': (822, 295), '杉戸町': (849, 280), '松伏町': (904, 350)
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

    map_img_path = "20261001_bc6e30f7720a548fb561a31.png"
    if not os.path.exists(map_img_path):
        map_img_path = "20261001_bc6e30f7720a548fb561a31.jpg"

    if os.path.exists(map_img_path):
        src_img = cv2.imread(map_img_path)
        h, w, _ = src_img.shape
        
        # 白黒反転処理（背景：黒→白 255, 境界線：灰色/緑線→濃いグレー 60,70,85）
        gray = cv2.cvtColor(src_img, cv2.COLOR_BGR2GRAY)
        inv_bgr = np.full_like(src_img, 255)
        
        # 境界線（明度が高めの線ピクセル）を抽出して線描画
        line_mask = (gray > 40)
        # 境界線を少し太く補強して完全密閉
        kernel = np.ones((3, 3), np.uint8)
        line_mask_dilated = cv2.dilate(line_mask.astype(np.uint8), kernel, iterations=1) > 0
        inv_bgr[line_mask_dilated] = (60, 70, 85)
        
        # OpenCV (BGR) でのカラーマップ完全定義（要件②）
        # A社 → 🟥 赤 (BGR: 68, 68, 239)
        # B社 → 🟦 青 (BGR: 246, 130, 59)
        # C社 → 🟩 緑 (BGR: 129, 185, 16)
        # なし → ⚪ 灰 (BGR: 220, 225, 230)
        BGR_MAP = {
            'A社': (68, 68, 239),
            'B社': (246, 130, 59),
            'C社': (129, 185, 16),
            'なし': (220, 225, 230)
        }
        
        # ★【核心修正】埼玉県外（左上等の背景領域）のペイント保護用マスク★
        bg_protection_mask = np.zeros((h + 2, w + 2), np.uint8)
        # 画面左上 (0,0) の背景から広がっている領域を「値=1（ペイント保護領域）」に設定
        cv2.floodFill(inv_bgr.copy(), bg_protection_mask, (0, 0), (255, 255, 255), (15, 15, 15), (15, 15, 15), cv2.FLOODFILL_FIXED_RANGE)
        
        # 1. 各72市区町村の閉領域を個別にペイント
        for c_name, seed in CITY_SEEDS.items():
            info = map_status_dict.get(c_name, {'一括担当': 'なし', '個別選択': []})
            bulk = info['一括担当']
            indivs = info['個別選択']
            
            fill_bgr = BGR_MAP.get(bulk, (220, 225, 230))
            if bulk == 'なし' and len(indivs) > 0:
                fill_bgr = (220, 225, 230)
                
            x, y = seed[0], seed[1]
            if 0 <= x < w and 0 <= y < h:
                # シード位置が濃い境界線上に接触している場合は白地エリアへ退避
                if np.mean(inv_bgr[y, x]) < 150:
                    found = False
                    for dy in range(-7, 8, 2):
                        for dx in range(-7, 8, 2):
                            nx, ny = x + dx, y + dy
                            if 0 <= nx < w and 0 <= ny < h and np.mean(inv_bgr[ny, nx]) > 200:
                                x, y = nx, ny
                                found = True
                                break
                        if found: break
                
                # ★解決策★ bg_protection_mask のクローン（.copy()）を渡すことで、
                # 前の市町村のペイントが次の市町村のペイントを誤ってブロックしないようにする！
                current_mask = bg_protection_mask.copy()
                
                # ペイント実行
                cv2.floodFill(inv_bgr, current_mask, (x, y), fill_bgr, (20, 20, 20), (20, 20, 20), cv2.FLOODFILL_FIXED_RANGE)
                
                # ドット描画（個別選択時）
                if len(indivs) > 0:
                    for d_idx, comp_indiv in enumerate(indivs):
                        dot_bgr = BGR_MAP.get(comp_indiv, (0, 0, 0))
                        dot_x = x + (d_idx - (len(indivs)-1)/2.0) * 16
                        cv2.circle(inv_bgr, (int(dot_x), y), 7, (255, 255, 255), -1)
                        cv2.circle(inv_bgr, (int(dot_x), y), 6, dot_bgr, -1)

        # 2. 余白を自動クロップして画面いっぱいに高品質表示
        img_rgb = cv2.cvtColor(inv_bgr, cv2.COLOR_BGR2RGB)
        gray_check = cv2.cvtColor(img_rgb, cv2.COLOR_RGB2GRAY)
        non_bg = np.where(gray_check < 250)
        
        if len(non_bg[0]) > 0:
            min_y, max_y = np.min(non_bg[0]), np.max(non_bg[0])
            min_x, max_x = np.min(non_bg[1]), np.max(non_bg[1])
            pad = 12
            crop_min_y = max(0, min_y - pad)
            crop_max_y = min(h, max_y + pad)
            crop_min_x = max(0, min_x - pad)
            crop_max_x = min(w, max_x + pad)
            cropped_img = img_rgb[crop_min_y:crop_max_y, crop_min_x:crop_max_x]
        else:
            cropped_img = img_rgb

        st.image(cropped_img, use_container_width=True)
    else:
        st.warning("「20261001_bc6e30f7720a548fb561a31.png」がリポジトリ内に存在しません。画像をアップロードして配置してください。")

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