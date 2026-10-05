import streamlit as st
import pandas as pd
import numpy as np
import cv2
import os
import math

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

# 3. 役所・役場座標 ＆ 営業所初期座標定義
MUNICIPAL_HALL_COORDS = {
    '大阪市都島区': (34.7011, 135.5303), '大阪市福島区': (34.6922, 135.4727), '大阪市此花区': (34.6853, 135.4619),
    '大阪市西区': (34.6756, 135.4872), '大阪市港区': (34.6644, 135.4608), '大阪市大正区': (34.6528, 135.4711),
    '大阪市天王寺区': (34.6561, 135.5208), '大阪市浪速区': (34.6558, 135.4981), '大阪市西淀川区': (34.7119, 135.4578),
    '大阪市東淀川区': (34.7394, 135.5347), '大阪市東成区': (34.6694, 135.5472), '大阪市生野区': (34.6550, 135.5439),
    '大阪市旭区': (34.7214, 135.5428), '大阪市城東区': (34.6989, 135.5458), '大阪市阿倍野区': (34.6386, 135.5175),
    '大阪市住吉区': (34.6086, 135.4958), '大阪市東住吉区': (34.6228, 135.5275), '大阪市西成区': (34.6358, 135.4947),
    '大阪市淀川区': (34.7203, 135.4831), '大阪市鶴見区': (34.7036, 135.5714), '大阪市住之江区': (34.6097, 135.4722),
    '大阪市平野区': (34.6214, 135.5489), '大阪市北区': (34.7056, 135.5103), '大阪市中央区': (34.6822, 135.5108),
    '堺市堺区': (34.5731, 135.4831), '堺市中区': (34.5369, 135.4967), '堺市東区': (34.5328, 135.5222),
    '堺市西区': (34.5483, 135.4608), '堺市南区': (34.4925, 135.5019), '堺市北区': (34.5661, 135.5133),
    '堺市美原区': (34.5269, 135.5583), '岸和田市': (34.4597, 135.3719), '豊中市': (34.7814, 135.4703),
    '池田市': (34.8214, 135.4278), '吹田市': (34.7583, 135.5167), '泉大津市': (34.5028, 135.4056),
    '高槻市': (34.8483, 135.6181), '貝塚市': (34.4369, 135.3583), '守口市': (34.7358, 135.5653),
    '枚方市': (34.8161, 135.6508), '茨木市': (34.8161, 135.5683), '八尾市': (34.6269, 135.6008),
    '泉佐野市': (34.4086, 135.3283), '富田林市': (34.5008, 135.6008), '寝屋川市': (34.7656, 135.6269),
    '河内長野市': (34.4558, 135.5639), '松原市': (34.5778, 135.5539), '大東市': (34.7119, 135.6239),
    '和泉市': (34.4883, 135.4267), '箕面市': (34.8269, 135.4703), '柏原市': (34.5808, 135.6289),
    '羽曳野市': (34.5583, 135.6067), '門真市': (34.7336, 135.5886), '摂津市': (34.7792, 135.5622),
    '高石市': (34.5208, 135.4383), '藤井寺市': (34.5739, 135.5967), '東大阪市': (34.6794, 135.6008),
    '泉南市': (34.3639, 135.2858), '四條畷市': (34.7369, 135.6419), '交野市': (34.7878, 135.6886),
    '大阪狭山市': (34.5036, 135.5539), '阪南市': (34.3583, 135.2389), '島本町': (34.8869, 135.6639),
    '豊能町': (34.9383, 135.4739), '能勢町': (34.9669, 135.3989), '忠岡町': (34.4869, 135.3919),
    '熊取町': (34.4036, 135.3539), '田尻町': (34.3983, 135.2919), '岬町': (34.3214, 135.1539),
    '太子町': (34.5222, 135.6489), '河南町': (34.4883, 135.6358), '千早赤阪村': (34.4569, 135.6189)
}

EXISTING_OFFICES_INFO = {
    ('J社', '大阪'): {'address': '大阪府茨木市宿久庄2-10-2', 'coords': (34.8583, 135.5383)},
    ('J社', '尼崎'): {'address': '兵庫県尼崎市西高洲町16-19', 'coords': (34.7083, 135.4083)},
    ('J社', '大阪南港'): {'address': '大阪府大阪市住之江区南港南1-1-125', 'coords': (34.6183, 135.4083)},
    ('J社', '東大阪'): {'address': '大阪府八尾市西高安町4-57-1', 'coords': (34.6308, 135.6383)},
    ('J社', '松原'): {'address': '大阪府松原市三宅西6-891-1', 'coords': (34.5883, 135.5383)},
    ('J社', '貝塚'): {'address': '大阪府貝塚市港17-8', 'coords': (34.4483, 135.3483)},
    ('J社', '京都'): {'address': '京都府久世郡久御山町下津屋上ノ浜28-2', 'coords': (34.8883, 135.7383)},
    ('T社', '大阪中央'): {'address': '大阪府大阪市鶴見区焼野3-2-11', 'coords': (34.7139, 135.5839)},
    ('T社', '尼崎'): {'address': '兵庫県尼崎市東海岸町21-10', 'coords': (34.6983, 135.4183)},
    ('T社', '南大阪'): {'address': '大阪府堺市堺区築港八幡町1-1', 'coords': (34.6083, 135.4383)},
    ('T社', '北大阪'): {'address': '大阪府茨木市宮島2-5-1', 'coords': (34.8039, 135.5839)},
    ('T社', '東大阪'): {'address': '大阪府東大阪市本庄中1-4-90', 'coords': (34.6883, 135.6083)},
    ('T社', '泉佐野'): {'address': '大阪府泉佐野市下瓦屋町2-2-50', 'coords': (34.4239, 135.3339)},
    ('T社', '京都'): {'address': '京都府京都市伏見区横大路一本木22', 'coords': (34.9283, 135.7483)},
}

def calc_haversine_distance(lat1, lon1, lat2, lon2):
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return round(R * c * 1.2, 1)

# セッション状態での新規営業所管理
if "custom_offices" not in st.session_state:
    st.session_state.custom_offices = []

# カラーマップ生成
def generate_office_colors(df_j, df_t, custom_offices):
    j_offices = sorted(list(set(df_j['J社営業所名'].dropna().astype(str).tolist() + [o['name'] for o in custom_offices if o['company'] == 'J社'])))
    t_offices = sorted(list(set(df_t['T社営業所名'].dropna().astype(str).tolist() + [o['name'] for o in custom_offices if o['company'] == 'T社'])))
    
    j_color_map = {}
    t_color_map = {}
    
    n_j = len(j_offices)
    for idx, off in enumerate(j_offices):
        f = 0.45 + 0.55 * (idx / max(n_j - 1, 1)) if n_j > 1 else 0.85
        j_color_map[off] = (int(255 * f), int(20 * (1.0 - f)), int(20 * (1.0 - f)))
        
    n_t = len(t_offices)
    for idx, off in enumerate(t_offices):
        f = 0.45 + 0.55 * (idx / max(n_t - 1, 1)) if n_t > 1 else 0.85
        t_color_map[off] = (int(30 * (1.0 - f)), int(140 * f), int(255 * f))
        
    return j_color_map, t_color_map, j_offices, t_offices

J_OFFICE_COLOR_MAP, T_OFFICE_COLOR_MAP, J_OFFICES, T_OFFICES = generate_office_colors(df_j, df_t, st.session_state.custom_offices)
ALL_OFFICES = sorted(list(set(J_OFFICES + T_OFFICES)))

# 白地図座標
CITY_SEEDS = {
    '能勢町': (542, 143), '豊能町': (635, 232), '島本町': (897, 293), '高槻市': (816, 331), '枚方市': (923, 331),
    '池田市': (553, 351), '箕面市': (618, 359), '茨木市': (702, 292), '交野市': (932, 517), '豊中市': (592, 473),
    '吹田市': (702, 497), '摂津市': (750, 517), '寝屋川市': (832, 504), '守口市': (750, 583), '門真市': (818, 584),
    '四條畷市': (923, 584), '大東市': (818, 600), '東大阪市': (867, 672), '大阪市西淀川区': (566, 619),
    '大阪市東淀川区': (670, 550), '大阪市淀川区': (603, 593), '大阪市北区': (637, 625), '大阪市都島区': (680, 594),
    '大阪市旭区': (715, 588), '大阪市城東区': (714, 638), '大阪市鶴見区': (761, 619), '大阪市福島区': (594, 647),
    '大阪市此花区': (545, 673), '大阪市西区': (617, 673), '大阪市中央区': (653, 671), '大阪市東成区': (719, 676),
    '大阪市生野区': (708, 712), '大阪市港区': (555, 713), '大阪市浪速区': (628, 702), '大阪市天王寺区': (665, 707),
    '大阪市大正区': (582, 740), '大阪市西成区': (620, 736), '大阪市阿倍野区': (660, 741), '大阪市平野区': (724, 795),
    '大阪市住之江区': (533, 748), '大阪市住吉区': (660, 799), '大阪市東住吉区': (680, 799), '八尾市': (818, 769),
    '柏原市': (881, 837), '松原市': (723, 842), '藤井寺市': (817, 853), '羽曳野市': (811, 911), '太子町': (890, 962),
    '富田林市': (811, 1011), '大阪狭山市': (724, 1011), '河南町': (866, 1011), '千早赤阪村': (881, 1104),
    '河内長野市': (743, 1147), '堺市堺区': (607, 848), '堺市北区': (665, 850), '堺市西区': (565, 930),
    '堺市中区': (640, 930), '堺市東区': (682, 911), '堺市美原区': (754, 932), '堺市南区': (641, 1011),
    '高石市': (535, 930), '泉大津市': (504, 1000), '忠岡町': (504, 1011), '和泉市': (565, 1096),
    '岸和田市': (504, 1179), '貝塚市': (494, 1220), '熊取町': (426, 1179), '泉佐野市': (370, 1179),
    '田尻町': (308, 1187), '泉南市': (318, 1251), '阪南市': (239, 1280), '岬町': (118, 1336)
}

EXTRA_SEEDS = {
    '大阪市此花区': [(477, 710), (503, 713)],
    '泉佐野市': [(236, 1101)]
}

# 4. ベースデータ作成（距離・重量・トンキロ算出手順を含む）
all_office_info = dict(EXISTING_OFFICES_INFO)
for cust in st.session_state.custom_offices:
    all_office_info[(cust['company'], cust['name'])] = {'address': cust['address'], 'coords': cust['coords']}

rows = []
for i in range(len(df_j)):
    code = df_j.loc[i, '市区町村コード']
    city = df_j.loc[i, '市区町村名\n（漢字）']
    
    wt_j_kg = df_j.loc[i, '配達重量\n(日当たり)']
    wt_t_kg = df_t.loc[i, '配達重量\n(日当たり)']
    
    off_j_def, off_t_def = df_j.loc[i, 'J社営業所名'], df_t.loc[i, 'T社営業所名']
    sub_j = str(df_j.loc[i, 'J社外部委託地域']).strip() == '◯'
    sub_t = str(df_t.loc[i, 'T社外部委託地域']).strip() == '◯'
    
    c_lat, c_lon = MUNICIPAL_HALL_COORDS.get(city, (34.6853, 135.5208))
    j_coords = all_office_info.get(('J社', off_j_def), {}).get('coords', (34.6853, 135.5208))
    t_coords = all_office_info.get(('T社', off_t_def), {}).get('coords', (34.6853, 135.5208))
    
    dist_j = calc_haversine_distance(j_coords[0], j_coords[1], c_lat, c_lon)
    dist_t = calc_haversine_distance(t_coords[0], t_coords[1], c_lat, c_lon)
    
    tk_j = round(dist_j * (wt_j_kg / 1000.0), 2)
    tk_t = round(dist_t * (wt_t_kg / 1000.0), 2)
    
    best_comp = 'J社' if tk_j <= tk_t else 'T社'
    best_off = off_j_def if best_comp == 'J社' else off_t_def
    init_mode = '委託配達' if (sub_j or sub_t) else '自社配達'
    
    rows.append({
        '市町村コード': code,
        '市町村名': city,
        '配達方式': init_mode,
        '市町村一括担当': best_comp,
        '市町村一括担当営業所': best_off,
        'J社個別': False, 'T社個別': False,
        'J社 距離(km)': dist_j, 'J社 配達重量(kg)': wt_j_kg, 'J社 トンキロ(t・km)': tk_j,
        'T社 距離(km)': dist_t, 'T社 配達重量(kg)': wt_t_kg, 'T社 トンキロ(t・km)': tk_t,
        '_j_def_off': off_j_def, '_t_def_off': off_t_def,
        '_sub_j': sub_j, '_sub_t': sub_t
    })

base_df = pd.DataFrame(rows)

# サイドバールール
st.sidebar.markdown("---")
st.sidebar.header("🎯 自動一括割り当て")
rule = st.sidebar.radio(
    "一括設定ルールを選択",
    ["最少トンキロ最適化（初期自動選択）", "全市町村 J社一括", "全市町村 T社一括", "選択クリア"]
)

if "最少トンキロ" in rule:
    for idx, r in base_df.iterrows():
        best = 'J社' if r['J社 トンキロ(t・km)'] <= r['T社 トンキロ(t・km)'] else 'T社'
        best_off = r['_j_def_off'] if best == 'J社' else r['_t_def_off']
        base_df.loc[idx, '市町村一括担当'] = best
        base_df.loc[idx, '市町村一括担当営業所'] = best_off
        base_df.loc[idx, 'J社個別'] = False; base_df.loc[idx, 'T社個別'] = False
elif "J社一括" in rule:
    base_df['市町村一括担当'] = 'J社'; base_df['市町村一括担当営業所'] = base_df['_j_def_off']
    base_df['J社個別'] = False; base_df['T社個別'] = False
elif "T社一括" in rule:
    base_df['市町村一括担当'] = 'T社'; base_df['市町村一括担当営業所'] = base_df['_t_def_off']
    base_df['J社個別'] = False; base_df['T社個別'] = False
else:
    base_df['市町村一括担当'] = 'なし'; base_df['市町村一括担当営業所'] = '-'
    base_df['J社個別'] = False; base_df['T社個別'] = False

# 5. タブUI構成
tab1, tab2, tab3 = st.tabs(["📊 全体サマリー＆現状比較", "🏢 会社別・営業所別集計", "📝 市町村別・受持選択（編集）"])

active_records = []
map_status_dict = {}

# タブ3: 個別編集画面
with tab3:
    st.subheader("➕ ③ 新規営業所の登録（J社 / T社）")
    st.caption("新規営業所名と住所を入力して追加すると、下のテーブルおよびシミュレーション選択肢に即座に反映されます。")
    
    col_add_j, col_add_t = st.columns(2)
    with col_add_j:
        st.markdown("##### 🔴 J社 新規営業所")
        j_new_name = st.text_input("J社 営業所名", key="j_off_name_input", placeholder="例: 茨木西")
        j_new_addr = st.text_input("J社 所在地住所", key="j_off_addr_input", placeholder="例: 大阪府茨木市駅前1-1-1")
        if st.button("J社 営業所を追加", key="btn_add_j"):
            if j_new_name and j_new_addr:
                # 簡易位置設定（市役所中心位置を基準）
                st.session_state.custom_offices.append({
                    'company': 'J社',
                    'name': j_new_name,
                    'address': j_new_addr,
                    'coords': (34.8161, 135.5683)
                })
                st.success("J社新規営業所「" + j_new_name + "」を追加しました！")
                st.rerun()

    with col_add_t:
        st.markdown("##### 🔵 T社 新規営業所")
        t_new_name = st.text_input("T社 営業所名", key="t_off_name_input", placeholder="例: 堺中央")
        t_new_addr = st.text_input("T社 所在地住所", key="t_off_addr_input", placeholder="例: 大阪府堺市堺区南瓦町3-1")
        if st.button("T社 営業所を追加", key="btn_add_t"):
            if t_new_name and t_new_addr:
                st.session_state.custom_offices.append({
                    'company': 'T社',
                    'name': t_new_name,
                    'address': t_new_addr,
                    'coords': (34.5731, 135.4831)
                })
                st.success("T社新規営業所「" + t_new_name + "」を追加しました！")
                st.rerun()

    st.markdown("---")
    st.subheader("📝 市町村別・受持選択テーブル")
    st.caption("※ 添付出力イメージ構成：市町村別の営業所距離（km）、配達重量（kg）、トンキロ（t・km）を動的計算します。")
    
    display_cols = [
        '市町村コード', '市町村名', '配達方式', '市町村一括担当', '市町村一括担当営業所',
        'J社個別', 'T社個別',
        'J社 距離(km)', 'J社 配達重量(kg)', 'J社 トンキロ(t・km)',
        'T社 距離(km)', 'T社 配達重量(kg)', 'T社 トンキロ(t・km)'
    ]
    
    edited_df = st.data_editor(
        base_df[display_cols],
        column_config={
            "配達方式": st.column_config.SelectboxColumn("配達方式", options=["自社配達", "委託配達"], required=True),
            "市町村一括担当": st.column_config.SelectboxColumn("市町村一括担当", options=["なし", "J社", "T社"], required=True),
            "市町村一括担当営業所": st.column_config.SelectboxColumn("市町村一括担当営業所", options=ALL_OFFICES + ["-"], required=True),
            "J社個別": st.column_config.CheckboxColumn("J社個別", default=False),
            "T社個別": st.column_config.CheckboxColumn("T社個別", default=False),
            "J社 距離(km)": st.column_config.NumberColumn("J社 距離 (km)", format="%.1f"),
            "J社 配達重量(kg)": st.column_config.NumberColumn("J社 配達重量 (kg)", format="%d"),
            "J社 トンキロ(t・km)": st.column_config.NumberColumn("J社 トンキロ (t・km)", format="%.2f"),
            "T社 距離(km)": st.column_config.NumberColumn("T社 距離 (km)", format="%.1f"),
            "T社 配達重量(kg)": st.column_config.NumberColumn("T社 配達重量 (kg)", format="%d"),
            "T社 トンキロ(t・km)": st.column_config.NumberColumn("T社 トンキロ (t・km)", format="%.2f"),
        },
        disabled=['市町村コード', '市町村名', 'J社 距離(km)', 'J社 配達重量(kg)', 'J社 トンキロ(t・km)', 'T社 距離(km)', 'T社 配達重量(kg)', 'T社 トンキロ(t・km)'],
        use_container_width=True,
        hide_index=True
    )

    # リアルタイム再計算処理
    for idx, r in edited_df.iterrows():
        orig = base_df.loc[idx]
        city_name = orig['市町村名']
        city_code = orig['市町村コード']
        is_sub_mode = (r['配達方式'] == '委託配達')
        
        bulk_comp = r['市町村一括担当']
        bulk_off = r['市町村一括担当営業所']
        j_indiv = r['J社個別']
        t_indiv = r['T社個別']
        
        j_off_name = orig['_j_def_off']
        t_off_name = orig['_t_def_off']
        
        if bulk_comp == 'J社' and bulk_off != '-':
            j_off_name = bulk_off
        elif bulk_comp == 'T社' and bulk_off != '-':
            t_off_name = bulk_off
            
        c_lat, c_lon = MUNICIPAL_HALL_COORDS.get(city_name, (34.6853, 135.5208))
        j_coords = all_office_info.get(('J社', j_off_name), {}).get('coords', (34.6853, 135.5208))
        t_coords = all_office_info.get(('T社', t_off_name), {}).get('coords', (34.6853, 135.5208))
        
        dist_j = calc_haversine_distance(j_coords[0], j_coords[1], c_lat, c_lon)
        dist_t = calc_haversine_distance(t_coords[0], t_coords[1], c_lat, c_lon)
        
        wt_j_kg = orig['J社 配達重量(kg)']
        wt_t_kg = orig['T社 配達重量(kg)']
        wt_j_t = wt_j_kg / 1000.0
        wt_t_t = wt_t_kg / 1000.0
        
        tk_j = round(dist_j * wt_j_t, 2)
        tk_t = round(dist_t * wt_t_t, 2)
        
        indiv_selected = []
        if j_indiv: indiv_selected.append('J社')
        if t_indiv: indiv_selected.append('T社')
        
        map_status_dict[city_name] = {
            '一括担当': bulk_comp,
            '一括営業所': bulk_off if bulk_off != '-' else (j_off_name if bulk_comp == 'J社' else t_off_name),
            '個別選択': indiv_selected
        }
        
        has_indiv = j_indiv or t_indiv
        if has_indiv:
            if j_indiv:
                sub = is_sub_mode or orig['_sub_j']
                active_records.append({
                    '市区町村コード': city_code, '市区町村名': city_name,
                    '担当会社': 'J社', '会社表示': 'J社（委託）' if sub else 'J社',
                    '担当営業所': j_off_name, '営業所表示': j_off_name + ('（委託）' if sub else ''),
                    '重量_t': wt_j_t, '距離_km': dist_j, 'トンキロ': tk_j,
                })
            if t_indiv:
                sub = is_sub_mode or orig['_sub_t']
                active_records.append({
                    '市区町村コード': city_code, '市区町村名': city_name,
                    '担当会社': 'T社', '会社表示': 'T社（委託）' if sub else 'T社',
                    '担当営業所': t_off_name, '営業所表示': t_off_name + ('（委託）' if sub else ''),
                    '重量_t': wt_t_t, '距離_km': dist_t, 'トンキロ': tk_t,
                })
        else:
            if bulk_comp in ['J社', 'T社']:
                assigned_off = bulk_off if bulk_off != '-' else (j_off_name if bulk_comp == 'J社' else t_off_name)
                c_code = bulk_comp[0]
                sub = is_sub_mode or (orig['_sub_j'] if c_code == 'J' else orig['_sub_t'])
                
                total_wt = wt_j_t + wt_t_t
                avg_dist = dist_j if c_code == 'J' else dist_t
                total_tk = round(avg_dist * total_wt, 2)
                
                active_records.append({
                    '市区町村コード': city_code, '市区町村名': city_name,
                    '担当会社': bulk_comp, '会社表示': bulk_comp + ('（委託）' if sub else ''),
                    '担当営業所': assigned_off, '営業所表示': assigned_off + ('（委託）' if sub else ''),
                    '重量_t': total_wt, '距離_km': avg_dist, 'トンキロ': total_tk,
                })

    st.markdown("---")
    st.subheader("🗺️ 大阪府 市区町村別受持選択 白地図エリアマップ")
    st.caption("塗り分け：営業所毎の配色（🔴 J社系: 赤グラデーション, 🔵 T社系: 青グラデーション, ⚪ 未設定: 灰）")

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

    # 画像描画
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
            
            if bulk == 'J社' and bulk_off in J_OFFICE_COLOR_MAP:
                fill_rgb = J_OFFICE_COLOR_MAP[bulk_off]
            elif bulk == 'T社' and bulk_off in T_OFFICE_COLOR_MAP:
                fill_rgb = T_OFFICE_COLOR_MAP[bulk_off]
            else:
                DEFAULT_RGB = {'J社': (239, 68, 68), 'T社': (59, 130, 246), 'なし': (220, 225, 230)}
                fill_rgb = DEFAULT_RGB.get(bulk, (220, 225, 230))
                
            if bulk == 'なし' and len(indivs) > 0:
                fill_rgb = (220, 225, 230)

            target_coords = [(sx, sy)]
            if c_name in EXTRA_SEEDS:
                target_coords.extend(EXTRA_SEEDS[c_name])
                
            for cs_x, cs_y in target_coords:
                if 0 <= cs_x < w and 0 <= cs_y < h:
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
        {'会社': 'J社', '現状_自治体数': len(base_df), '現状_重量_t': base_df['J社 配達重量(kg)'].sum() / 1000.0, '現状_トンキロ': base_df['J社 トンキロ(t・km)'].sum()},
        {'会社': 'T社', '現状_自治体数': len(base_df), '現状_重量_t': base_df['T社 配達重量(kg)'].sum() / 1000.0, '現状_トンキロ': base_df['T社 トンキロ(t・km)'].sum()},
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