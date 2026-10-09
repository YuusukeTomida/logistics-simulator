import streamlit as st
import pandas as pd
import numpy as np
import cv2
import os
import math
import io
import openpyxl
from openpyxl.drawing.image import Image as OpenPyxlImage

# 1. ページ基本設定
st.set_page_config(
    page_title="配達エリア・トンキロ最適化シミュレーター（大阪府版）",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🚚 運送2社 営業所受持エリア・トンキロ最適化シミュレーター（大阪府版）")
st.caption("大阪府の市区町村別・会社別の受持選択を行い、「現状」と「4つの改正案」のトンキロ・配達重量の変化を一括比較・シミュレーションします。")

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

FULL_CITY_COORDS_LOOKUP = {
    '東大阪市': (34.6794, 135.6008), '東大阪': (34.6794, 135.6008),
    '大阪市都島区': (34.7011, 135.5303), '都島区': (34.7011, 135.5303), '都島': (34.7011, 135.5303),
    '大阪市福島区': (34.6922, 135.4727), '福島区': (34.6922, 135.4727),
    '大阪市此花区': (34.6853, 135.4619), '此花区': (34.6853, 135.4619),
    '大阪市西区': (34.6756, 135.4872), '大阪市港区': (34.6644, 135.4608),
    '大阪市大正区': (34.6528, 135.4711), '大正区': (34.6528, 135.4711),
    '大阪市天王寺区': (34.6561, 135.5208), '天王寺区': (34.6561, 135.5208), '天王寺': (34.6561, 135.5208),
    '大阪市浪速区': (34.6558, 135.4981), '浪速区': (34.6558, 135.4981),
    '大阪市西淀川区': (34.7119, 135.4578), '西淀川区': (34.7119, 135.4578),
    '大阪市東淀川区': (34.7394, 135.5347), '東淀川区': (34.7394, 135.5347),
    '大阪市東成区': (34.6694, 135.5472), '東成区': (34.6694, 135.5472),
    '大阪市生野区': (34.6550, 135.5439), '生野区': (34.6550, 135.5439),
    '大阪市旭区': (34.7214, 135.5428), '大阪市城東区': (34.6989, 135.5458),
    '城東区': (34.6989, 135.5458), '大阪市阿倍野区': (34.6386, 135.5175),
    '阿倍野区': (34.6386, 135.5175), '阿倍野': (34.6386, 135.5175),
    '大阪市住吉区': (34.6086, 135.4958), '大阪市東住吉区': (34.6228, 135.5275),
    '東住吉区': (34.6228, 135.5275), '大阪市西成区': (34.6358, 135.4947),
    '西成区': (34.6358, 135.4947), '大阪市淀川区': (34.7203, 135.4831),
    '淀川区': (34.7203, 135.4831), '大阪市鶴見区': (34.7036, 135.5714),
    '鶴見区': (34.7036, 135.5714), '大阪市住之江区': (34.6097, 135.4722),
    '住之江区': (34.6097, 135.4722), '住之江': (34.6097, 135.4722),
    '大阪市平野区': (34.6214, 135.5489), '平野区': (34.6214, 135.5489),
    '大阪市北区': (34.7056, 135.5103), '大阪市中央区': (34.6822, 135.5108),
    '堺市堺区': (34.5731, 135.4831), '堺区': (34.5731, 135.4831),
    '堺市中区': (34.5369, 135.4967), '堺市東区': (34.5328, 135.5222),
    '堺市西区': (34.5483, 135.4608), '堺市南区': (34.4925, 135.5019),
    '堺市北区': (34.5661, 135.5133), '堺市美原区': (34.5269, 135.5583),
    '美原区': (34.5269, 135.5583), '美原': (34.5269, 135.5583),
    '岸和田市': (34.4597, 135.3719), '岸和田': (34.4597, 135.3719),
    '豊中市': (34.7814, 135.4703), '豊中': (34.7814, 135.4703),
    '池田市': (34.8214, 135.4278), '池田': (34.8214, 135.4278),
    '吹田市': (34.7583, 135.5167), '吹田': (34.7583, 135.5167),
    '泉大津市': (34.5028, 135.4056), '泉大津': (34.5028, 135.4056),
    '高槻市': (34.8483, 135.6181), '高槻': (34.8483, 135.6181),
    '貝塚市': (34.4369, 135.3583), '貝塚': (34.4369, 135.3583),
    '守口市': (34.7358, 135.5653), '守口': (34.7358, 135.5653),
    '枚方市': (34.8161, 135.6508), '枚方': (34.8161, 135.6508),
    '茨木市': (34.8161, 135.5683), '茨木': (34.8161, 135.5683),
    '八尾市': (34.6269, 135.6008), '八尾': (34.6269, 135.6008),
    '泉佐野市': (34.4086, 135.3283), '泉佐野': (34.4086, 135.3283),
    '富田林市': (34.5008, 135.6008), '富田林': (34.5008, 135.6008),
    '寝屋川市': (34.7656, 135.6269), '寝屋川': (34.7656, 135.6269),
    '河内長野市': (34.4558, 135.5639), '河内長野': (34.4558, 135.5639),
    '松原市': (34.5778, 135.5539), '松原': (34.5778, 135.5539),
    '大東市': (34.7119, 135.6239), '大東': (34.7119, 135.6239),
    '和泉市': (34.4883, 135.4267), '和泉': (34.4883, 135.4267),
    '箕面市': (34.8269, 135.4703), '箕面': (34.8269, 135.4703),
    '柏原市': (34.5808, 135.6289), '柏原': (34.5808, 135.6289),
    '羽曳野市': (34.5583, 135.6067), '羽曳野': (34.5583, 135.6067),
    '門真市': (34.7336, 135.5886), '門真': (34.7336, 135.5886),
    '摂津市': (34.7792, 135.5622), '摂津': (34.7792, 135.5622),
    '高石市': (34.5208, 135.4383), '高石': (34.5208, 135.4383),
    '藤井寺市': (34.5739, 135.5967), '藤井寺': (34.5739, 135.5967),
    '泉南市': (34.3639, 135.2858), '泉南': (34.3639, 135.2858),
    '四條畷市': (34.7369, 135.6419), '四條畷': (34.7369, 135.6419),
    '交野市': (34.7878, 135.6886), '交野': (34.7878, 135.6886),
    '大阪狭山市': (34.5036, 135.5539), '大阪狭山': (34.5036, 135.5539),
    '阪南市': (34.3583, 135.2389), '阪南': (34.3583, 135.2389),
    '島本町': (34.8869, 135.6639), '島本': (34.8869, 135.6639),
    '豊能町': (34.9383, 135.4739), '豊能': (34.9383, 135.4739),
    '能勢町': (34.9669, 135.3989), '能勢': (34.9669, 135.3989),
    '忠岡町': (34.4869, 135.3919), '忠岡': (34.4869, 135.3919),
    '熊取町': (34.4036, 135.3539), '熊取': (34.4036, 135.3539),
    '田尻町': (34.3983, 135.2919), '田尻': (34.3983, 135.2919),
    '岬町': (34.3214, 135.1539), '太子町': (34.5222, 135.6489),
    '河南町': (34.4883, 135.6358), '千早赤阪村': (34.4569, 135.6189),
    '千早赤阪': (34.4569, 135.6189), '尼崎市': (34.7083, 135.4083),
    '尼崎': (34.7083, 135.4083), '神戸市': (34.6901, 135.1955),
    '神戸': (34.6901, 135.1955), '西宮市': (34.7376, 135.3414),
    '西宮': (34.7376, 135.3414), '京都市': (34.9858, 135.7583),
    '京都': (34.9858, 135.7583), '堺市': (34.5731, 135.4831),
    '堺': (34.5731, 135.4831), '大阪市': (34.6853, 135.5208),
    '大阪': (34.6853, 135.5208)
}

def geocode_address_precise(address_str):
    if not address_str:
        return (34.6853, 135.5208)
    for kw, coords in sorted(FULL_CITY_COORDS_LOOKUP.items(), key=lambda x: len(x[0]), reverse=True):
        if kw in address_str:
            return coords
    return (34.6853, 135.5208)

def calc_haversine_distance(lat1, lon1, lat2, lon2):
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return round(R * c * 1.2, 1)

# 新規営業所の登録状態をパターン別に完全分離管理（custom_offices_p1 〜 custom_offices_p4）
for p_idx in [1, 2, 3, 4]:
    key_name = f"custom_offices_p{p_idx}"
    if key_name not in st.session_state:
        st.session_state[key_name] = []

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

# 4. 距離・トンキロ算出関数
def compute_distance_and_tonkm(city_name, bulk_comp, bulk_off, j_indiv_off, t_indiv_off, wt_j_kg, wt_t_kg, office_info, municipal_coords):
    c_lat, c_lon = municipal_coords.get(city_name, (34.6853, 135.5208))
    
    # 一括担当営業所が指定されている場合は単一座標から同一距離を算出
    if bulk_off != '-' and bulk_off is not None:
        c_prefix = bulk_comp if bulk_comp in ['J社', 'T社'] else 'J社'
        coords = office_info.get((c_prefix, bulk_off), {}).get('coords')
        if not coords:
            other_prefix = 'T社' if c_prefix == 'J社' else 'J社'
            coords = office_info.get((other_prefix, bulk_off), {}).get('coords')
            
        if coords:
            dist = calc_haversine_distance(coords[0], coords[1], c_lat, c_lon)
            dist_j = dist
            dist_t = dist
        else:
            dist_j = 0.0
            dist_t = 0.0
    else:
        # J社個別指定（未指定なら 0.0 km）
        if j_indiv_off != '-' and j_indiv_off is not None:
            j_coords = office_info.get(('J社', j_indiv_off), {}).get('coords')
            if not j_coords:
                j_coords = office_info.get(('T社', j_indiv_off), {}).get('coords')
            dist_j = calc_haversine_distance(j_coords[0], j_coords[1], c_lat, c_lon) if j_coords else 0.0
        else:
            dist_j = 0.0
            
        # T社個別指定（未指定なら 0.0 km）
        if t_indiv_off != '-' and t_indiv_off is not None:
            t_coords = office_info.get(('T社', t_indiv_off), {}).get('coords')
            if not t_coords:
                t_coords = office_info.get(('J社', t_indiv_off), {}).get('coords')
            dist_t = calc_haversine_distance(t_coords[0], t_coords[1], c_lat, c_lon) if t_coords else 0.0
        else:
            dist_t = 0.0
            
    tk_j = round(dist_j * (wt_j_kg / 1000.0), 2)
    tk_t = round(dist_t * (wt_t_kg / 1000.0), 2)
    
    return dist_j, tk_j, dist_t, tk_t

# 初期データ生成関数（全パターン用完全に独立したディープコピー初期データフレーム生成）
def build_initial_scenario_df(office_info_dict):
    rows = []
    for i in range(len(df_j)):
        code = df_j.loc[i, '市区町村コード']
        city = df_j.loc[i, '市区町村名\n（漢字）']
        wt_j_kg = df_j.loc[i, '配達重量\n(日当たり)']
        wt_t_kg = df_t.loc[i, '配達重量\n(日当たり)']
        off_j_def, off_t_def = df_j.loc[i, 'J社営業所名'], df_t.loc[i, 'T社営業所名']
        sub_j = str(df_j.loc[i, 'J社外部委託地域']).strip() == '◯'
        sub_t = str(df_t.loc[i, 'T社外部委託地域']).strip() == '◯'
        
        dist_j, tk_j, dist_t, tk_t = compute_distance_and_tonkm(city, 'なし', '-', off_j_def, off_t_def, wt_j_kg, wt_t_kg, office_info_dict, MUNICIPAL_HALL_COORDS)
        init_mode = '委託配達' if (sub_j or sub_t) else '自社配達'
        
        rows.append({
            '市町村コード': code, '市町村名': city, '配達方式': init_mode,
            '市町村一括担当': 'なし', '市町村一括担当営業所': '-',
            'J社個別': off_j_def, 'T社個別': off_t_def,
            'J社 距離(km)': dist_j, 'J社 配達重量(kg)': wt_j_kg, 'J社 トンキロ(t・km)': tk_j,
            'T社 距離(km)': dist_t, 'T社 配達重量(kg)': wt_t_kg, 'T社 トンキロ(t・km)': tk_t,
            '_j_def_off': off_j_def, '_t_def_off': off_t_def, '_sub_j': sub_j, '_sub_t': sub_t
        })
    return pd.DataFrame(rows)

# ② 受持選択（編集）用 4つの複製パターン（案1〜案4）の独立したディープコピー初期化
for p_idx in [1, 2, 3, 4]:
    key_name = f"table_data_p{p_idx}"
    if key_name not in st.session_state:
        st.session_state[key_name] = build_initial_scenario_df(EXISTING_OFFICES_INFO).copy(deep=True)

# 5. タブUI構成
tab_names = [
    "📊 全体サマリー＆現状比較",
    "🏢 会社別・営業所別集計",
    "📝 案1：受持選択（編集）",
    "📝 案2：受持選択（編集）",
    "📝 案3：受持選択（編集）",
    "📝 案4：受持選択（編集）"
]
tabs = st.tabs(tab_names)

pattern_active_records = {1: [], 2: [], 3: [], 4: []}
pattern_map_status = {1: {}, 2: {}, 3: {}, 4: {}}
pattern_map_images = {1: None, 2: None, 3: None, 4: None} # 白地図画像保持

display_cols = [
    '市町村コード', '市町村名', '配達方式', '市町村一括担当', '市町村一括担当営業所',
    'J社個別', 'T社個別',
    'J社 距離(km)', 'J社 配達重量(kg)', 'J社 トンキロ(t・km)',
    'T社 距離(km)', 'T社 配達重量(kg)', 'T社 トンキロ(t・km)'
]

# タブ3〜6: 案1〜案4 個別編集画面処理
for p_idx in [1, 2, 3, 4]:
    tab_obj = tabs[p_idx + 1] # 0:サマリー, 1:会社営業所集計, 2:案1, 3:案2, 4:案3, 5:案4
    with tab_obj:
        st.subheader(f"📝 案{p_idx}：市町村別・受持選択テーブル")
        st.caption(f"※ 案{p_idx} の設定を行います。新規営業所の登録や担当選択の変更後に「配達支店変更を反映」を押すと距離・トンキロが自動計算されます。")
        
        # パターン別（案1〜案4独立）の新規営業所管理・座標生成
        p_custom_offices = st.session_state[f"custom_offices_p{p_idx}"]
        
        j_color_map_p, t_color_map_p, j_offices_p, t_offices_p = generate_office_colors(df_j, df_t, p_custom_offices)
        all_offices_p = sorted(list(set(j_offices_p + t_offices_p)))
        
        p_office_info = dict(EXISTING_OFFICES_INFO)
        for cust in p_custom_offices:
            p_office_info[(cust['company'], cust['name'])] = {'address': cust['address'], 'coords': cust['coords']}
        
        # 案1〜案4それぞれのタブで独立した新規営業所登録フォーム
        st.markdown(f"##### ➕ 案{p_idx} 専用 新規営業所の登録（J社 / T社）")
        col_add_j, col_add_t = st.columns(2)
        with col_add_j:
            j_new_name = st.text_input("J社 営業所名", key=f"j_off_input_p{p_idx}", placeholder="例: 茨木西")
            j_new_addr = st.text_input("J社 所在地住所", key=f"j_off_addr_p{p_idx}", placeholder="例: 大阪府茨木市駅前1-1-1")
            if st.button("J社 営業所を追加", key=f"btn_add_j_p{p_idx}"):
                if j_new_name and j_new_addr:
                    detected_coords = geocode_address_precise(j_new_addr)
                    st.session_state[f"custom_offices_p{p_idx}"].append({
                        'company': 'J社', 'name': j_new_name, 'address': j_new_addr, 'coords': detected_coords
                    })
                    st.success(f"案{p_idx} に J社新規営業所「{j_new_name}」（住所: {j_new_addr}）を追加しました！")
                    st.rerun()

        with col_add_t:
            t_new_name = st.text_input("T社 営業所名", key=f"t_off_input_p{p_idx}", placeholder="例: 堺中央")
            t_new_addr = st.text_input("T社 所在地住所", key=f"t_off_addr_p{p_idx}", placeholder="例: 大阪府堺市堺区南瓦町3-1")
            if st.button("T社 営業所を追加", key=f"btn_add_t_p{p_idx}"):
                if t_new_name and t_new_addr:
                    detected_coords = geocode_address_precise(t_new_addr)
                    st.session_state[f"custom_offices_p{p_idx}"].append({
                        'company': 'T社', 'name': t_new_name, 'address': t_new_addr, 'coords': detected_coords
                    })
                    st.success(f"案{p_idx} に T社新規営業所「{t_new_name}」（住所: {t_new_addr}）を追加しました！")
                    st.rerun()
        st.markdown("---")

        curr_p_df = st.session_state[f"table_data_p{p_idx}"]
        
        recalc_clicked = st.button(f"🔄 案{p_idx} の配達支店変更を反映", type="primary", key=f"btn_recalc_p{p_idx}", use_container_width=True)

        edited_p_df = st.data_editor(
            curr_p_df[display_cols],
            column_config={
                "配達方式": st.column_config.SelectboxColumn("配達方式", options=["自社配達", "委託配達"], required=True),
                "市町村一括担当": st.column_config.SelectboxColumn("市町村一括担当", options=["なし", "J社", "T社"], required=True),
                "市町村一括担当営業所": st.column_config.SelectboxColumn("市町村一括担当営業所", options=all_offices_p + ["-"], required=True),
                "J社個別": st.column_config.SelectboxColumn("J社個別", options=j_offices_p + ["-"], required=True),
                "T社個別": st.column_config.SelectboxColumn("T社個別", options=t_offices_p + ["-"], required=True),
                "J社 距離(km)": st.column_config.NumberColumn("J社 距離 (km)", format="%.1f"),
                "J社 配達重量(kg)": st.column_config.NumberColumn("J社 配達重量 (kg)", format="%d"),
                "J社 トンキロ(t・km)": st.column_config.NumberColumn("J社 トンキロ (t・km)", format="%.2f"),
                "T社 距離(km)": st.column_config.NumberColumn("T社 距離 (km)", format="%.1f"),
                "T社 配達重量(kg)": st.column_config.NumberColumn("T社 配達重量 (kg)", format="%d"),
                "T社 トンキロ(t・km)": st.column_config.NumberColumn("T社 トンキロ (t・km)", format="%.2f"),
            },
            disabled=['市町村コード', '市町村名', 'J社 距離(km)', 'J社 配達重量(kg)', 'J社 トンキロ(t・km)', 'T社 距離(km)', 'T社 配達重量(kg)', 'T社 トンキロ(t・km)'],
            use_container_width=True,
            hide_index=True,
            key=f"main_editor_p{p_idx}"
        )

        if recalc_clicked:
            for idx in range(len(edited_p_df)):
                r = edited_p_df.iloc[idx]
                c_name = r['市町村名']
                b_comp = r['市町村一括担当']
                b_off = r['市町村一括担当営業所']
                j_ind = r['J社個別']
                t_ind = r['T社個別']
                wt_j = r['J社 配達重量(kg)']
                wt_t = r['T社 配達重量(kg)']
                
                dj, tkj, dt, tkt = compute_distance_and_tonkm(c_name, b_comp, b_off, j_ind, t_ind, wt_j, wt_t, p_office_info, MUNICIPAL_HALL_COORDS)
                
                curr_p_df.at[idx, '配達方式'] = r['配達方式']
                curr_p_df.at[idx, '市町村一括担当'] = b_comp
                curr_p_df.at[idx, '市町村一括担当営業所'] = b_off
                curr_p_df.at[idx, 'J社個別'] = j_ind
                curr_p_df.at[idx, 'T社個別'] = t_ind
                curr_p_df.at[idx, 'J社 距離(km)'] = dj
                curr_p_df.at[idx, 'J社 トンキロ(t・km)'] = tkj
                curr_p_df.at[idx, 'T社 距離(km)'] = dt
                curr_p_df.at[idx, 'T社 トンキロ(t・km)'] = tkt
                
            st.session_state[f"table_data_p{p_idx}"] = curr_p_df
            st.success(f"案{p_idx} の距離およびトンキロを最新表示に再計算しました！")
            st.rerun()

        # 集計＆マップ保持（全案の最新データの編集内容を確実に同期反映）
        p_active = []
        p_map_dict = {}
        for idx, r in edited_p_df.iterrows():
            orig = curr_p_df.loc[idx]
            city_name = orig['市町村名']
            city_code = orig['市町村コード']
            is_sub_mode = (r['配達方式'] == '委託配達')
            
            bulk_comp = r['市町村一括担当']
            bulk_off = r['市町村一括担当営業所']
            j_indiv_off = r['J社個別']
            t_indiv_off = r['T社個別']
            
            # 最新演算距離・トンキロ（即時同期待避）
            dj_now, tkj_now, dt_now, tkt_now = compute_distance_and_tonkm(city_name, bulk_comp, bulk_off, j_indiv_off, t_indiv_off, orig['J社 配達重量(kg)'], orig['T社 配達重量(kg)'], p_office_info, MUNICIPAL_HALL_COORDS)
            
            curr_p_df.at[idx, '市町村一括担当'] = bulk_comp
            curr_p_df.at[idx, '市町村一括担当営業所'] = bulk_off
            curr_p_df.at[idx, 'J社個別'] = j_indiv_off
            curr_p_df.at[idx, 'T社個別'] = t_indiv_off
            curr_p_df.at[idx, 'J社 距離(km)'] = dj_now
            curr_p_df.at[idx, 'J社 トンキロ(t・km)'] = tkj_now
            curr_p_df.at[idx, 'T社 距離(km)'] = dt_now
            curr_p_df.at[idx, 'T社 トンキロ(t・km)'] = tkt_now

            dist_j = dj_now
            dist_t = dt_now
            tk_j = tkj_now
            tk_t = tkt_now
            
            wt_j_kg = orig['J社 配達重量(kg)']
            wt_t_kg = orig['T社 配達重量(kg)']
            wt_j_t = wt_j_kg / 1000.0
            wt_t_t = wt_t_kg / 1000.0
            
            p_map_dict[city_name] = {
                '一括担当': bulk_comp,
                '一括営業所': bulk_off if bulk_off != '-' else (j_indiv_off if bulk_comp == 'J社' and j_indiv_off != '-' else (t_indiv_off if bulk_comp == 'T社' and t_indiv_off != '-' else orig['_j_def_off'] if bulk_comp == 'J社' else orig['_t_def_off'])),
                'j_indiv_off': j_indiv_off,
                't_indiv_off': t_indiv_off
            }
            
            has_indiv = (j_indiv_off != '-') or (t_indiv_off != '-')
            if has_indiv:
                if j_indiv_off != '-':
                    sub = is_sub_mode or orig['_sub_j']
                    p_active.append({
                        '市区町村コード': city_code, '市区町村名': city_name,
                        '担当会社': 'J社', '会社表示': 'J社（委託）' if sub else 'J社',
                        '担当営業所': j_indiv_off, '営業所表示': j_indiv_off + ('（委託）' if sub else ''),
                        '重量_t': wt_j_t, '距離_km': dist_j, 'トンキロ': tk_j,
                    })
                if t_indiv_off != '-':
                    sub = is_sub_mode or orig['_sub_t']
                    p_active.append({
                        '市区町村コード': city_code, '市区町村名': city_name,
                        '担当会社': 'T社', '会社表示': 'T社（委託）' if sub else 'T社',
                        '担当営業所': t_indiv_off, '営業所表示': t_indiv_off + ('（委託）' if sub else ''),
                        '重量_t': wt_t_t, '距離_km': dist_t, 'トンキロ': tk_t,
                    })
            else:
                if bulk_comp in ['J社', 'T社']:
                    assigned_off = bulk_off if bulk_off != '-' else (orig['_j_def_off'] if bulk_comp == 'J社' else orig['_t_def_off'])
                    c_code = bulk_comp[0]
                    sub = is_sub_mode or (orig['_sub_j'] if c_code == 'J' else orig['_sub_t'])
                    
                    total_wt = wt_j_t + wt_t_t
                    avg_dist = dist_j if c_code == 'J' else dist_t
                    total_tk = round(avg_dist * total_wt, 2)
                    
                    p_active.append({
                        '市区町村コード': city_code, '市区町村名': city_name,
                        '担当会社': bulk_comp, '会社表示': bulk_comp + ('（委託）' if sub else ''),
                        '担当営業所': assigned_off, '営業所表示': assigned_off + ('（委託）' if sub else ''),
                        '重量_t': total_wt, '距離_km': avg_dist, 'トンキロ': total_tk,
                    })

        st.session_state[f"table_data_p{p_idx}"] = curr_p_df
        pattern_active_records[p_idx] = p_active
        pattern_map_status[p_idx] = p_map_dict

        st.markdown("---")
        st.subheader(f"🗺️ 案{p_idx} 白地図エリアマップ")
        
        # カラー凡例のカラーブロックアイコン（拡大）
        st.markdown("##### 📌 営業所別 カラー凡例（拡大カラーブロック表示）")
        leg_cols = st.columns(2)
        with leg_cols[0]:
            st.write("**🔴 J社 営業所**")
            for off in j_offices_p:
                c = j_color_map_p.get(off, (239, 68, 68))
                hex_c = f"#{c[0]:02x}{c[1]:02x}{c[2]:02x}"
                st.markdown(f'<span style="color:{hex_c}; font-size:48px; vertical-align:middle; line-height:1;">■</span> <span style="font-size:18px; font-weight:bold; vertical-align:middle;">{off}</span>', unsafe_allow_html=True)
                
        with leg_cols[1]:
            st.write("**🔵 T社 営業所**")
            for off in t_offices_p:
                c = t_color_map_p.get(off, (59, 130, 246))
                hex_c = f"#{c[0]:02x}{c[1]:02x}{c[2]:02x}"
                st.markdown(f'<span style="color:{hex_c}; font-size:48px; vertical-align:middle; line-height:1;">■</span> <span style="font-size:18px; font-weight:bold; vertical-align:middle;">{off}</span>', unsafe_allow_html=True)

        st.write("")

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
                info = p_map_dict.get(c_name, {'一括担当': 'なし', '一括営業所': '-', 'j_indiv_off': '-', 't_indiv_off': '-'})
                bulk = info['一括担当']
                bulk_off = info['一括営業所']
                j_indiv_off = info.get('j_indiv_off', '-')
                t_indiv_off = info.get('t_indiv_off', '-')
                
                # ③ 地図色塗りルールの元に戻し処理（一括＝営業所色, 個別選択あり＝灰色, 未設定＝白色）
                if bulk == 'J社' and bulk_off in j_color_map_p:
                    fill_rgb = j_color_map_p[bulk_off]
                elif bulk == 'T社' and bulk_off in t_color_map_p:
                    fill_rgb = t_color_map_p[bulk_off]
                elif bulk in ['なし', '-'] and ((j_indiv_off != '-') or (t_indiv_off != '-')):
                    fill_rgb = (200, 200, 200) # 個別選択あり＝灰色
                else:
                    fill_rgb = (255, 255, 255) # 未設定（一括なし＆個別なし）は白色

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

            pattern_map_images[p_idx] = cropped_img

            map_col1, map_col2, map_col3 = st.columns([1, 4, 1])
            with map_col2:
                st.image(cropped_img, width=600)

# タブ1: 全体サマリー ＆ 現状 vs 4つの変更案比較（①J社＋T社「合計」比較追加 ＆ ②エクセル「白地図エリアマップ」に比較表追加）
with tabs[0]:
    st.subheader("📈 会社毎 ＆ 全体合計の「現状」vs「4つの変更案（案1〜案4）」一括比較サマリー")
    st.caption("J社・T社個別の数値に加え、運送2社全体の合計トンキロ・削減率を一括比較します。")

    # ② 現状基準データフレーム（完全同一算出基準による初期集計）
    base_0_df = build_initial_scenario_df(EXISTING_OFFICES_INFO)
    
    cur_j_wt_total = base_0_df['J社 配達重量(kg)'].sum() / 1000.0
    cur_t_wt_total = base_0_df['T社 配達重量(kg)'].sum() / 1000.0
    cur_j_tk_total = base_0_df['J社 トンキロ(t・km)'].sum()
    cur_t_tk_total = base_0_df['T社 トンキロ(t・km)'].sum()

    summary_rows = [
        {'パターン': '現状（基本）', '会社': 'J社', '担当自治体数': len(base_0_df), '配達重量_t': round(cur_j_wt_total, 2), 'トンキロ': round(cur_j_tk_total, 2), 'トンキロ削減量': 0.0, '削減率(%)': 0.0},
        {'パターン': '現状（基本）', '会社': 'T社', '担当自治体数': len(base_0_df), '配達重量_t': round(cur_t_wt_total, 2), 'トンキロ': round(cur_t_tk_total, 2), 'トンキロ削減量': 0.0, '削減率(%)': 0.0}
    ]

    for p_idx in [1, 2, 3, 4]:
        p_df = st.session_state[f"table_data_p{p_idx}"]
        
        j_tk_sum = 0.0; t_tk_sum = 0.0
        j_wt_sum = 0.0; t_wt_sum = 0.0
        j_cnt = 0; t_cnt = 0
        
        for _, r in p_df.iterrows():
            bulk = str(r['市町村一括担当']).strip()
            j_tk = float(r['J社 トンキロ(t・km)'])
            t_tk = float(r['T社 トンキロ(t・km)'])
            j_wt = float(r['J社 配達重量(kg)']) / 1000.0
            t_wt = float(r['T社 配達重量(kg)']) / 1000.0
            
            if bulk == 'J社':
                j_tk_sum += (j_tk + t_tk)
                j_wt_sum += (j_wt + t_wt)
                j_cnt += 1
            elif bulk == 'T社':
                t_tk_sum += (j_tk + t_tk)
                t_wt_sum += (j_wt + t_wt)
                t_cnt += 1
            else: # 'なし' または '-'
                j_tk_sum += j_tk
                j_wt_sum += j_wt
                if j_tk > 0 or str(r['J社個別']).strip() != '-':
                    j_cnt += 1
                    
                t_tk_sum += t_tk
                t_wt_sum += t_wt
                if t_tk > 0 or str(r['T社個別']).strip() != '-':
                    t_cnt += 1

        j_diff = cur_j_tk_total - j_tk_sum
        j_pct = round((j_diff / cur_j_tk_total * 100), 1) if cur_j_tk_total > 0 else 0.0
        summary_rows.append({'パターン': f'案{p_idx}', '会社': 'J社', '担当自治体数': j_cnt, '配達重量_t': round(j_wt_sum, 2), 'トンキロ': round(j_tk_sum, 2), 'トンキロ削減量': round(j_diff, 2), '削減率(%)': j_pct})

        t_diff = cur_t_tk_total - t_tk_sum
        t_pct = round((t_diff / cur_t_tk_total * 100), 1) if cur_t_tk_total > 0 else 0.0
        summary_rows.append({'パターン': f'案{p_idx}', '会社': 'T社', '担当自治体数': t_cnt, '配達重量_t': round(t_wt_sum, 2), 'トンキロ': round(t_tk_sum, 2), 'トンキロ削減量': round(t_diff, 2), '削減率(%)': t_pct})

    summary_df = pd.DataFrame(summary_rows)

    # ① J社＋T社 合計集計行の動的生成
    cur_comb_tk_total = cur_j_tk_total + cur_t_tk_total
    comb_summary_rows = []
    
    # 現状（基本）の合計
    comb_summary_rows.append({
        'パターン': '現状（基本）',
        'J社_トンキロ': round(cur_j_tk_total, 2),
        'T社_トンキロ': round(cur_t_tk_total, 2),
        '合計_トンキロ': round(cur_comb_tk_total, 2),
        '合計_削減量': 0.0,
        '合計_削減率(%)': 0.0
    })

    for p_idx in [1, 2, 3, 4]:
        j_row = summary_df[(summary_df['パターン'] == f'案{p_idx}') & (summary_df['会社'] == 'J社')].iloc[0]
        t_row = summary_df[(summary_df['パターン'] == f'案{p_idx}') & (summary_df['会社'] == 'T社')].iloc[0]
        
        j_tk_p = j_row['トンキロ']
        t_tk_p = t_row['トンキロ']
        comb_tk_p = j_tk_p + t_tk_p
        comb_diff = cur_comb_tk_total - comb_tk_p
        comb_pct = round((comb_diff / cur_comb_tk_total * 100), 1) if cur_comb_tk_total > 0 else 0.0
        
        comb_summary_rows.append({
            'パターン': f'案{p_idx}',
            'J社_トンキロ': j_tk_p,
            'T社_トンキロ': t_tk_p,
            '合計_トンキロ': round(comb_tk_p, 2),
            '合計_削減量': round(comb_diff, 2),
            '合計_削減率(%)': comb_pct
        })

    comb_summary_df = pd.DataFrame(comb_summary_rows)

    # 全体サマリーへの一括Excel出力ボタン（②シート「白地図エリアマップ」に比較表追加）
    excel_full_bytes = io.BytesIO()
    with pd.ExcelWriter(excel_full_bytes, engine='openpyxl') as writer:
        all_cust_rows = []
        for p_i in [1, 2, 3, 4]:
            p_custs = st.session_state[f"custom_offices_p{p_i}"]
            for c_item in p_custs:
                all_cust_rows.append({
                    '対象パターン': f'案{p_i}',
                    '担当会社': c_item['company'],
                    '新規営業所名': c_item['name'],
                    '所在地住所': c_item['address']
                })
        if all_cust_rows:
            df_cust_all = pd.DataFrame(all_cust_rows)
        else:
            df_cust_all = pd.DataFrame(columns=['対象パターン', '担当会社', '新規営業所名', '所在地住所'])
        df_cust_all.to_excel(writer, sheet_name='新規営業所一覧', index=False)

        for p_i in [1, 2, 3, 4]:
            p_df_export = st.session_state[f"table_data_p{p_i}"][display_cols].copy()
            p_df_export.to_excel(writer, sheet_name=f'案{p_i}', index=False)

        wb = writer.book
        ws_map = wb.create_sheet(title="白地図エリアマップ")
        ws_map['A1'] = "🗺️ 大阪府 市区町村別受持選択 白地図エリアマップ（案1〜案4 比較一覧）"
        
        # ② シート「白地図エリアマップ」に「J社・T社・合計トンキロの比較表」を追加出力（A3〜F9セル）
        ws_map['A3'] = "【全体トンキロ比較対比表（現状 vs 案1〜案4）】"
        map_headers = ['パターン', 'J社 トンキロ(t・km)', 'T社 トンキロ(t・km)', '合計 トンキロ(t・km)', '合計 削減量', '合計 削減率(%)']
        for col_idx, h_text in enumerate(map_headers, 1):
            ws_map.cell(row=4, column=col_idx, value=h_text)
            
        for r_idx, c_row in enumerate(comb_summary_rows, 5):
            ws_map.cell(row=r_idx, column=1, value=c_row['パターン'])
            ws_map.cell(row=r_idx, column=2, value=c_row['J社_トンキロ'])
            ws_map.cell(row=r_idx, column=3, value=c_row['T社_トンキロ'])
            ws_map.cell(row=r_idx, column=4, value=c_row['合計_トンキロ'])
            ws_map.cell(row=r_idx, column=5, value=c_row['合計_削減量'])
            ws_map.cell(row=r_idx, column=6, value=f"{c_row['合計_削減率(%)']:+.1f}%")

        # マップ画像のセル配置（比較表の下部、12行目からマップ画像を配置）
        map_cell_positions = {
            1: ('B', 12, 13),   # 案1: B12タイトル, B13画像
            2: ('K', 12, 13),   # 案2: K12タイトル, K13画像
            3: ('T', 12, 13),   # 案3: T12タイトル, T13画像
            4: ('AC', 12, 13)   # 案4: AC12タイトル, AC13画像
        }
        
        for p_i in [1, 2, 3, 4]:
            col_let, t_row, img_row = map_cell_positions[p_i]
            ws_map[f"{col_let}{t_row}"] = f"■ 案{p_i} マップ"
            
            m_img = pattern_map_images[p_i]
            if m_img is not None:
                b_rgb = cv2.cvtColor(m_img, cv2.COLOR_RGB2BGR)
                is_ok, buffer = cv2.imencode(".png", b_rgb)
                if is_ok:
                    img_stream = io.BytesIO(buffer)
                    xl_img = OpenPyxlImage(img_stream)
                    xl_img.width = 380
                    xl_img.height = 520
                    ws_map.add_image(xl_img, f"{col_let}{img_row}")

        for col_idx in range(1, 35):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            ws_map.column_dimensions[col_letter].width = 6.0

    st.download_button(
        label="📥 全体シミュレーション＆4案一括データ（白地図比較表・新規営業所併記）Excel出力",
        data=excel_full_bytes.getvalue(),
        file_name="大阪府配達エリア最適化シミュレーション_全体比較結果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )

    st.markdown("---")
    # ① J社＋T社「合計」の現状 vs 4案比較メトリックの表示
    st.markdown("### 📊 J社 ＋ T社 全体合計：現状 vs 4案の比較")
    cols_comb = st.columns(5)
    for idx, col in enumerate(cols_comb):
        row_c = comb_summary_df.iloc[idx]
        with col:
            st.markdown(f"#### {row_c['パターン']}")
            st.metric("合計トンキロ", "{:,.1f} ton-km".format(row_c['合計_トンキロ']), delta=f"削減: {row_c['合計_削減量']:,.1f} ({row_c['合計_削減率(%)']:+.1f}%)" if idx > 0 else None)
            st.caption(f"J社: {row_c['J社_トンキロ']:,.1f} ton-km / T社: {row_c['T社_トンキロ']:,.1f} ton-km")

    st.markdown("---")
    st.markdown("### 🏢 J社：現状 vs 4案の比較")
    cols_j = st.columns(5)
    j_sum = summary_df[summary_df['会社'] == 'J社']
    for idx, col in enumerate(cols_j):
        row_d = j_sum.iloc[idx]
        with col:
            st.markdown(f"#### {row_d['パターン']}")
            st.metric("トンキロ", "{:,.1f} ton-km".format(row_d['トンキロ']), delta=f"削減: {row_d['トンキロ削減量']:,.1f} ({row_d['削減率(%)']:+.1f}%)" if idx > 0 else None)
            st.caption(f"件数: {int(row_d['担当自治体数'])} 件 / 重量: {row_d['配達重量_t']:,.1f} t")

    st.markdown("---")
    st.markdown("### 🏢 T社：現状 vs 4案の比較")
    cols_t = st.columns(5)
    t_sum = summary_df[summary_df['会社'] == 'T社']
    for idx, col in enumerate(cols_t):
        row_d = t_sum.iloc[idx]
        with col:
            st.markdown(f"#### {row_d['パターン']}")
            st.metric("トンキロ", "{:,.1f} ton-km".format(row_d['トンキロ']), delta=f"削減: {row_d['トンキロ削減量']:,.1f} ({row_d['削減率(%)']:+.1f}%)" if idx > 0 else None)
            st.caption(f"件数: {int(row_d['担当自治体数'])} 件 / 重量: {row_d['配達重量_t']:,.1f} t")

    st.markdown("---")
    st.write("### 📊 全パターン比較対比表（会社別）")
    st.dataframe(
        summary_df.style.format({
            '配達重量_t': '{:,.2f}', 'トンキロ': '{:,.2f}',
            'トンキロ削減量': '{:,.2f}', '削減率(%)': '{:+.1f}%'
        }),
        use_container_width=True, hide_index=True
    )

# タブ2: 会社別 ＆ 営業所別 詳細集計（案1〜案4から選択可能）
with tabs[1]:
    st.subheader("🏢 会社別 ＆ 🏬 営業所別 詳細集計")
    selected_p = st.selectbox("集計表示する変更案を選択", [1, 2, 3, 4], format_func=lambda x: f"案{x}")
    
    p_act = pattern_active_records[selected_p]
    p_act_df = pd.DataFrame(p_act) if len(p_act) > 0 else pd.DataFrame()
    
    if not p_act_df.empty:
        st.markdown(f"#### 🏢 案{selected_p} 会社別 集計（自社配達 vs 外部委託）")
        comp_sub_summary = p_act_df.groupby(['会社表示']).agg(
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
        
        st.markdown("---")
        st.markdown(f"#### 🏬 案{selected_p} 営業所別 集計（自社配達 vs 外部委託）")
        off_sub_summary = p_act_df.groupby(['担当会社', '営業所表示']).agg(
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
        st.info(f"案{selected_p} の割り当てデータがありません。")