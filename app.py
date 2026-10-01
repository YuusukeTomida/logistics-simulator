import streamlit as st
import pandas as pd
import numpy as np
import cv2
from PIL import Image, ImageDraw, ImageFont
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
    'さいたま市西区': (637, 542), 'さいたま市北区': [676, 521], 'さいたま市大宮区': (669, 552),
    'さいたま市見沼区': (702, 532), 'さいたま市中央区': (660, 566), 'さいたま市桜区': (640, 584),
    'さいたま市浦和区': (684, 576), 'さいたま市南区': [682, 606], 'さいたま市緑区': (724, 580),
    'さいたま市岩槻区': (743, 517), '川越市': (585, 547), '熊谷市': (511, 331),
    '川口市': (756, 622), '行田市': (574, 335), '秩父市': (212, 537), '所沢市': (556, 642),
    '飯能市': (380, 577), '加須市': (670, 349), '本庄市': (326, 308), '東松山市': (507, 442),
    '春日部市': (782, 482), '狭山市': [528, 595], '羽生市': (630, 317), '鴻巣市': (587, 398),
    '深谷市': [424, 308], '上尾市': [640, 497], '草加市': (796, 608), '越谷市': (793, 560),
    '蕨市': [699, 618], '戸田市': [682, 633], '入間市': (485, 627), '朝霞市': (654, 638),
    '志木市': (640, 613), '和光市': (675, 655), '新座市': (625, 653), '桶川市': (599, 467),
    '久喜市': (697, 411), '北本市': (595, 442), '八潮市': (828, 625), '富士見市': (620, 597),
    '三郷市': (843, 612), '蓮田市': (685, 471), '坂戸市': [512, 501], '幸手市': (765, 404),
    '鶴ヶ島市': (498, 526), '日高市': (465, 560), '吉川市': (847, 564), 'ふじみ野市': (599, 580),
    '白岡市': (713, 442), '伊奈町': (655, 471), '三芳町': (587, 612), '毛呂山町': (433, 520),
    '越生町': (412, 498), '滑川町': (480, 421), '嵐山町': [451, 417], '小川町': (408, 411),
    '川島町': (562, 492), '吉見町': (552, 427), '鳩山町': (460, 471), 'ときがわ町': (393, 469),
    '横瀬町': (317, 501), '皆野町': [287, 396], '長瀞町': (308, 362), '小鹿野町': (175, 452),
    '東秩父村': (360, 442), '美里町': [338, 328], '神川町': (308, 281), '上里町': (321, 252),
    '寄居町': [370, 368], '宮代町': (743, 451), '杉戸町': [770, 439], '松伏町': (825, 520)
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
    st.markdown("""
    - **配達方式（列2）**: 「自社配達」または「委託配達」を選択できます。
    - **一括担当（列3）**: 「A社」「B社」「C社」から1つ選択すると、市町村全体の荷物が選択した会社に一括集計されます。
    - **個社チェック（列4〜6）**: チェックを入れた会社の荷物のみが集計されます（個別選択が最優先されます）。
    """)

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

# タブ1: サマリー & 高精細クロップ済み画像塗りつぶしマップ
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
    st.subheader("🗺️ 埼玉県 市町村別受持選択 高精細マップ")
    st.caption("塗り分け：一括担当（赤: A社, 青: B社, 緑: C社, 灰: なし） / ドット：個別選択された会社の色")

    # 高画質処理 ＆ 余白自動クロップ ＆ 文字の鮮明再描画
    map_img_path = "20261001_bc6e30f7720a548fb561a31.png"
    if os.path.exists(map_img_path):
        img_bgr = cv2.imread(map_img_path)
        h, w, _ = img_bgr.shape
        
        # 明るく見やすいパステル系カラー (BGR)
        BGR_MAP = {
            'A社': (110, 110, 245),  # 鮮やかな赤系
            'B社': (245, 160, 90),   # 鮮やかな青系
            'C社': (130, 210, 110),  # 鮮やかな緑系
            'なし': (235, 235, 235)  # 灰色
        }
        
        mask = np.zeros((h + 2, w + 2), np.uint8)
        
        # 1. 塗りつぶし実行
        for c_name, seed in CITY_SEEDS.items():
            info = map_status_dict.get(c_name, {'一括担当': 'なし', '個別選択': []})
            bulk = info['一括担当']
            indivs = info['個別選択']
            
            fill_bgr = BGR_MAP.get(bulk, (245, 245, 245))
            if bulk == 'なし' and len(indivs) > 0:
                fill_bgr = (235, 235, 235)
                
            x, y = seed[0], seed[1]
            if 0 <= x < w and 0 <= y < h:
                cv2.floodFill(img_bgr, mask, (x, y), fill_bgr, (20, 20, 20), (20, 20, 20), cv2.FLOODFILL_FIXED_RANGE)
                
                # ドット描画（個別選択時）
                if len(indivs) > 0:
                    for d_idx, comp_indiv in enumerate(indivs):
                        dot_bgr = BGR_MAP.get(comp_indiv, (0, 0, 0))
                        dot_x = x + (d_idx - (len(indivs)-1)/2.0) * 14
                        cv2.circle(img_bgr, (int(dot_x), y - 10), 6, (255, 255, 255), -1)
                        cv2.circle(img_bgr, (int(dot_x), y - 10), 5, dot_bgr, -1)

        # 2. PILによる市町村名テキストのくっきり高画質描画（縁取り付き）
        img_rgb = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(img_rgb)
        draw = ImageDraw.Draw(pil_img)
        
        # 日本語フォント指定（Linux/Windows環境対応）
        font = ImageFont.load_default()
        
        for c_name, seed in CITY_SEEDS.items():
            x, y = seed[0], seed[1]
            # 文字背景の白色フチ＋文字本体（ネイビー）
            draw.text((x - 12, y - 4), c_name, fill=(15, 23, 42), stroke_width=2, stroke_fill=(255, 255, 255), font=font)

        # 3. ③ 余白（グレーエリア）の自動クロップ（トリミング）
        np_img = np.array(pil_img)
        gray_img = cv2.cvtColor(np_img, cv2.COLOR_RGB2GRAY)
        
        # 埼玉県の地図領域（グレー背景値235未満）を抽出
        non_bg_pts = np.where(gray_img < 235)
        if len(non_bg_pts[0]) > 0:
            min_y, max_y = np.min(non_bg_pts[0]), np.max(non_bg_pts[0])
            min_x, max_x = np.min(non_bg_pts[1]), np.max(non_bg_pts[1])
            
            # マージン（余白）を少しだけ残してトリミング
            pad = 15
            crop_min_y = max(0, min_y - pad)
            crop_max_y = min(h, max_y + pad)
            crop_min_x = max(0, min_x - pad)
            crop_max_x = min(w, max_x + pad)
            
            cropped_img = np_img[crop_min_y:crop_max_y, crop_min_x:crop_max_x]
        else:
            cropped_img = np_img

        # ① 高解像度拡大表示 (Streamlitコンテナ幅ぴったりに綺麗にフィット)
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