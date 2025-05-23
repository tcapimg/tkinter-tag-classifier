import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import pandas as pd
import os
import re
import random
import uuid # UUIDを生成するために追加

# --- データの保存先ファイル名 ---
DATA_FILE = 'tag_dictionary.json'

# --- アプリケーションのグローバル状態管理 ---
app_state = {
    'dictionary': {"categories": []},
    'unclassified_df': pd.DataFrame(columns=["英語タグ名", "日本語説明", "カテゴリ"]),
    'edited_dict_df': pd.DataFrame(columns=["英語タグ名", "日本語説明", "カテゴリ", "_category_id"]),
    'selected_generating_tags': [],
    'random_generated_tags': []
}

# グローバルなソート方向を保持する辞書
# キー: (tree_widget_name, column_id), 値: True (降順) または False (昇順)
sort_reverse_flags = {}

# --- ヘルパー関数 ---

def load_dictionary():
    """辞書データをファイルから読み込む関数"""
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                app_state['dictionary'] = json.load(f)
            # 読み込み時に全てのタグの日本語説明をstripする
            for category in app_state['dictionary'].get('categories', []):
                for tag in category.get('tags', []):
                    if 'ja' in tag and tag['ja'] is not None:
                        tag['ja'] = tag['ja'].strip()
        except json.JSONDecodeError:
            messagebox.showerror("エラー", "辞書ファイルが破損しているようです。新しい辞書を作成します。")
            app_state['dictionary'] = {"categories": []}
    else:
        app_state['dictionary'] = {"categories": []}
    
    app_state['edited_dict_df'] = pd.DataFrame(columns=["英語タグ名", "日本語説明", "カテゴリ", "_category_id"])
    update_category_dropdowns() # all_category_options をここで更新

def save_dictionary():
    """辞書データをファイルに保存する関数"""
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(app_state['dictionary'], f, ensure_ascii=False, indent=4)

def get_category_path(category_id):
    """カテゴリIDからカテゴリパス（例: 服装 / 女性 / トップス）を取得する関数"""
    path = []
    current_id = category_id
    # カテゴリマップを一度構築して効率化
    all_categories_map = {cat['id']: cat for cat in app_state['dictionary']['categories']}
    while current_id:
        category = all_categories_map.get(current_id)
        if category:
            path.insert(0, category['name'])
            current_id = category.get('parent_id')
        else:
            break
    return " / ".join(path) if path else ""

def find_category_by_id(category_id, categories=None):
    """カテゴリIDからカテゴリ情報を検索する関数"""
    if categories is None:
        categories = app_state['dictionary']['categories']
    for category in categories:
        if category['id'] == category_id:
            return category
    return None

def get_category_id_from_path(path_string):
    """カテゴリパス（例: 服装 / 女性）からカテゴリIDを取得する関数"""
    path_parts = path_string.split(' / ')
    current_categories = [cat for cat in app_state['dictionary']['categories'] if cat.get('parent_id') is None] # トップレベルカテゴリから開始
    current_category = None
    
    for i, name_part in enumerate(path_parts):
        found = None
        for cat in current_categories:
            if cat['name'] == name_part:
                found = cat
                break
        if found:
            current_category = found
            # 次のレベルの子カテゴリを準備
            current_categories = [cat for cat in app_state['dictionary']['categories'] if cat.get('parent_id') == current_category['id']]
        else:
            return None # パスが見つからない
    return current_category['id'] if current_category else None


def add_tag_to_dictionary(tag_en, tag_ja, category_id):
    """タグを辞書に追加する関数"""
    category = find_category_by_id(category_id)
    if category:
        # 既存のタグをチェック (英語タグ名で大文字小文字を区別せずチェック)
        existing_tag = next((t for t in category.setdefault('tags', []) if t['en'].lower() == tag_en.lower()), None)
        if existing_tag:
            # 既存のタグが見つかった場合、日本語説明を更新 (stripを適用)
            existing_tag['ja'] = tag_ja.strip()
            return True, f"タグ '{tag_en}' の日本語説明をカテゴリ '{category['name']}' で更新しました。"
        else:
            # 新規タグとして追加 (stripを適用)
            category.setdefault('tags', []).append({"en": tag_en, "ja": tag_ja.strip()})
            return True, f"タグ '{tag_en}' をカテゴリ '{category['name']}' に追加しました。"
    else:
        return False, "指定されたカテゴリが見つかりません。"

def generate_initial_dictionary():
    """デモ用の初期辞書を生成する関数"""
    initial_data = {
      "categories": [
        { "id": "person", "name": "人物", "parent_id": None, "tags": [{"en": "1girl", "ja": "少女/女性一人"}, {"en": "2girls", "ja": "少女/女性二人"}, {"en": "multiple girls", "ja": "少女/女性複数"}, {"en": "boy", "ja": "少年"}, {"en": "shota", "ja": "ショタ"}, {"en": "loli", "ja": "ロリ"}, {"en": "age difference", "ja": "年齢差"}, {"en": "teacher", "ja": "先生"}, {"en": "student", "ja": "生徒"}, {"en": "male", "ja": "男性"}, {"en": "female", "ja": "女性"}, {"en": "asian", "ja": "アジア人"}, {"en": "child", "ja": "子供"}, {"en": "adult", "ja": "大人"}, {"en": "elderly", "ja": "高齢者"}] },
        { "id": "body", "name": "身体", "parent_id": None, "tags": [{"en": "small breasts", "ja": "貧乳"}, {"en": "medium breasts", "ja": "普通サイズの胸"}, {"en": "large breasts", "ja": "巨乳"}, {"en": "huge breasts", "ja": "爆乳"}, {"en": "hanging breasts", "ja": "垂れ乳"}, {"en": "ass", "ja": "尻"}, {"en": "pussy", "ja": "女性器"}, {"en": "testicles", "ja": "睾丸"}, {"en": "penis", "ja": "ペニス"}, {"en": "lips", "ja": "唇"}, {"en": "legs", "ja": "脚"}, {"en": "bare legs", "ja": "生脚"}, {"en": "thighs", "ja": "太もも"}, {"en": "feet", "ja": "足（足首から下）"}, {"en": "hands", "ja": "手"}, {"en": "fingers", "ja": "指"}, {"en": "pubic hair", "ja": "陰毛"}, {"en": "female pubic hair", "ja": "女性の陰毛"}, {"en": "armpits", "ja": "脇"}, {"en": "navel", "ja": "へそ"}, {"en": "nipples", "ja": "乳首"}, {"en": "areolae", "ja": "乳輪"}, {"en": "clitoris", "ja": "クリトリス"}, {"en": "anus", "ja": "アナル"}] },
         { "id": "hair", "name": "髪", "parent_id": None, "tags": [{"en": "ponytail", "ja": "ポニーテール"}, {"en": "short hair", "ja": "ショートヘア"}, {"en": "long hair", "ja": "ロングヘア"}, {"en": "black hair", "ja": "黒髪"}, {"en": "blonde hair", "ja": "金髪"}, {"en": "brown hair", "ja": "茶髪"}, {"en": "blue hair", "ja": "青髪"}, {"en": "pink hair", "ja": "ピンク髪"}, {"en": "red hair", "ja": "赤髪"}, {"en": "green hair", "ja": "緑髪"}, {"en": "purple hair", "ja": "紫髪"}, {"en": "twintails", "ja": "ツインテール"}, {"en": "braids", "ja": "三つ編み"}, {"en": "bun", "ja": "お団子ヘア"}, {"en": "ahoge", "ja": "アホ毛"}, {"en": "bangs", "ja": "前髪"}, {"en": "sideburns", "ja": "もみあげ"}] },
         { "id": "eyes", "name": "瞳", "parent_id": None, "tags": [{"en": "blue eyes", "ja": "青い瞳"}, {"en": "red eyes", "ja": "赤い瞳"}, {"en": "green eyes", "ja": "緑の瞳"}, {"en": "brown eyes", "ja": "茶色い瞳"}, {"en": "purple eyes", "ja": "紫の瞳"}, {"en": "yellow eyes", "ja": "黄色い瞳"}, {"en": "heterochromia", "ja": "オッドアイ"}, {"en": "closed eyes", "ja": "閉じた目"}, {"en": "eyelashes", "ja": "まつげ"}, {"en": "eyebrows", "ja": "眉毛"}] },
        { "id": "clothes", "name": "服装", "parent_id": None, "tags": [{"en": "shirt", "ja": "シャツ"}, {"en": "t-shirt", "ja": "Tシャツ"}, {"en": "dress", "ja": "ワンピース"}, {"en": "skirt", "ja": "スカート"}, {"en": "pencil skirt", "ja": "タイトスカート"}, {"en": "pleated skirt", "ja": "プリーツスカート"}, {"en": "pants", "ja": "パンツ"}, {"en": "jeans", "ja": "ジーンズ"}, {"en": "shorts", "ja": "ショートパンツ"}, {"en": "uniform", "ja": "ユニフォーム"}, {"en": "school uniform", "ja": "制服"}, {"en": "sailor uniform", "ja": "セーラー服"}, {"en": "suit", "ja": "スーツ"}, {"en": "jacket", "ja": "ジャケット"}, {"en": "cardigan", "ja": "カーディガン"}, {"en": "sweater", "ja": "セーター"}, {"en": "vest", "ja": "ベスト"}, {"en": "hoodie", "ja": "パーカー"}, {"en": "coat", "ja": "コート"}, {"en": "lingerie", "ja": "ランジェリー"}, {"en": "bra", "ja": "ブラジャー"}, {"en": "black bra", "ja": "黒ブラジャー"}, {"en": "pink bra", "ja": "ピンクブラジャー"}, {"en": "yellow bra", "ja": "黄色ブラジャー"}, {"en": "bow bra", "ja": "リボン付きブラジャー"}, {"en": "panties", "ja": "パンティー"}, {"en": "black panties", "ja": "黒パンティー"}, {"en": "pink panties", "ja": "ピンクパンティー"}, {"en": "bow panties", "ja": "リボン付きパンティー"}, {"en": "stockings", "ja": "ストッキング"}, {"en": "thighhighs", "ja": "サイハイソックス"}, {"en": "pantyhose", "ja": "パンティストッキング"}, {"en": "socks", "ja": "靴下"}, {"en": "footwear", "ja": "履物"}, {"en": "shoes", "ja": "靴"}, {"en": "sneakers", "ja": "スニーカー"}, {"en": "boots", "ja": "ブーツ"}, {"en": "sandals", "ja": "サンダル"},
            {"en": "heels", "ja": "ヒール"}, {"en": "necktie", "ja": "ネクタイ"}, {"en": "bowtie", "ja": "蝶ネクタイ"},
            {"en": "bow", "ja": "リボン"}, {"en": "ribbon", "ja": "リボン"}, {"en": "scarf", "ja": "スカーフ"},
            {"en": "gloves", "ja": "手袋"}, {"en": "hat", "ja": "帽子"}, {"en": "cap", "ja": "キャップ"},
            {"en": "glasses", "ja": "眼鏡"}, {"en": "sunglasses", "ja": "サングラス"},
            {"en": "swimsuit", "ja": "水着"}, {"en": "bikini", "ja": "ビキニ"},
            {"en": "school swimsuit", "ja": "スクール水着"}, {"en": "leotard", "ja": "レオタード"},
            {"en": "apron", "ja": "エプロン"},
            {"en": "bunny suit", "ja": "バニースーツ"}, {"en": "nurse uniform", "ja": "ナース服"}, {"en": "police uniform", "ja": "警察官の制服"}, {"en": "military uniform", "ja": "軍服"}, {"en": "yukata", "ja": "浴衣"}, {"en": "kimono", "ja": "着物"}, {"en": "cheongsam", "ja": "チャイナドレス"},
            {"en": "garter straps", "ja": "ガーターストラップ"}, {"en": "suspenders", "ja": "サスペンダー"},
            {"en": "belt", "ja": "ベルト"}, {"en": "zipper", "ja": "ジッパー"},
            {"en": "buttons", "ja": "ボタン"}
          ]
        },
         {
          "id": "clothes_state_action",
          "name": "服装の状態/アクション",
          "parent_id": "clothes",
          "tags": [
            {"en": "open clothes", "ja": "服が開いている"},
            {"en": "open shirt", "ja": "シャツが開いている"},
            {"en": "clothes pull", "ja": "服を引っ張る"},
            {"en": "panty pull", "ja": "パンティーを引っ張る"},
            {"en": "pants pull", "ja": "パンツを引っ張る"},
            {"en": "skirt lift", "ja": "スカートを持ち上げる"},
            {"en": "clothes lift", "ja": "服を持ち上げる"},
            {"en": "undressing", "ja": "脱衣中"},
            {"en": "unworn panties", "ja": "着用していないパンティー"},
            {"en": "no panties", "ja": "ノーパン"},
            {"en": "no socks", "ja": "靴下なし"},
            {"en": "no legwear", "ja": "レッグウェアなし"},
            {"en": "one breast out", "ja": "片乳出し"},
            {"en": "erection under clothes", "ja": "服の上から勃起"},
            {"en": "bulge", "ja": "膨らみ"},
            {"en": "sleeves rolled up", "ja": "袖まくり"},
            {"en": "torn clothes", "ja": "破れた服"},
            {"en": "wet clothes", "ja": "濡れた服"}
          ]
        },
        {
          "id": "pose",
          "name": "ポーズ",
          "parent_id": None,
          "tags": [
            {"en": "standing", "ja": "立つ"},
            {"en": "standing on one leg", "ja": "片足立ち"},
            {"en": "sitting", "ja": "座る"},
            {"en": "kneeling", "ja": "跪く"},
            {"en": "on one knee", "ja": "片膝立ち"},
            {"en": "squatting", "ja": "しゃがむ"},
            {"en": "bent over", "ja": "かがむ"},
            {"en": "lying", "ja": "寝そべる"},
            {"en": "on back", "ja": "仰向け"},
            {"en": "on stomach", "ja": "うつ伏せ"},
            {"en": "on side", "ja": "横向き"},
            {"en": "leg lift", "ja": "足を上げる"},
            {"en": "leg up", "ja": "片足を上げる"},
            {"en": "spread legs", "ja": "開脚"},
            {"en": "tiptoes", "ja": "つま先立ち"},
            {"en": "hands on own knees", "ja": "自分の膝に手を置く"},
            {"en": "arms up", "ja": "腕を上げる"},
            {"en": "arms crossed", "ja": "腕組み"},
            {"en": "hands on hips", "ja": "腰に手を当てる"},
            {"en": "fingers in mouth", "ja": "指を口に入れる"},
            {"en": "licking lips", "ja": "唇を舐める"},
            {"en": "looking at viewer", "ja": "こちらを見ている"},
            {"en": "looking away", "ja": "視線を外す"},
            {"en": "profile", "ja": "横顔"},
            {"en": "back", "ja": "後ろ姿"},
            {"en": "from below", "ja": "下から見るアングル"},
            {"en": "from above", "ja": "上から見るアングル"}
          ]
        },
         {
          "id": "pose_action",
          "name": "ポーズのアクション",
          "parent_id": "pose",
          "tags": [
            {"en": "grabbing", "ja": "掴む"},
            {"en": "ass grab", "ja": "尻を掴む"},
            {"en": "grabbing another's ass", "ja": "他人の尻を掴む"},
            {"en": "grabbing another's breast", "ja": "他人の胸を掴む"},
            {"en": "groping", "ja": "揉む"},
            {"en": "biting", "ja": "噛む"},
            {"en": "kissing", "ja": "キス"}
          ]
        },
        {
          "id": "scene_context",
          "name": "情景/状況",
          "parent_id": None,
          "tags": [
            {"en": "indoors", "ja": "屋内"},
            {"en": "outdoors", "ja": "屋外"},
            {"en": "office", "ja": "オフィス"},
            {"en": "school", "ja": "学校"},
            {"en": "bedroom", "ja": "寝室"},
            {"en": "bathroom", "ja": "浴室"},
            {"en": "kitchen", "ja": "キッチン"},
            {"en": "living room", "ja": "リビング"},
            {"en": "street", "ja": "通り"},
            {"en": "park", "ja": "公園"},
            {"en": "beach", "ja": "ビーチ"},
            {"en": "pool", "ja": "プール"},
            {"en": "against wall", "ja": "壁際"},
            {"en": "on bed", "ja": "ベッドの上"},
            {"en": "on floor", "ja": "床の上"},
            {"en": "in water", "ja": "水の中"},
            {"en": "formal", "ja": "フォーマル"},
            {"en": "casual", "ja": "カジュアル"},
            {"en": "cosplay", "ja": "コスプレ"},
            {"en": "raining", "ja": "雨"},
            {"en": "snowing", "ja": "雪"},
            {"en": "sunny", "ja": "晴れ"},
            {"en": "night", "ja": "夜"},
            {"en": "day", "ja": "昼"},
            {"en": "sunset", "ja": "夕焼け"},
            {"en": "sunrise", "ja": "日の出"},
            {"en": "indoors", "ja": "屋内"},
            {"en": "outdoors", "ja": "屋外"}
          ]
        },
         {
          "id": "action_relationship",
          "name": "行為/関係性",
          "parent_id": None,
          "tags": [
            {"en": "sex", "ja": "セックス"},
            {"en": "doggystyle", "ja": "バック体位"},
            {"en": "standing sex", "ja": "立ちセックス"},
            {"en": "clothed sex", "ja": "服を着たままセックス"},
            {"en": "fellatio", "ja": "フェラチオ"},
            {"en": "cunnilingus", "ja": "クンニリングス"},
            {"en": "anal", "ja": "アナル"},
            {"en": "vaginal", "ja": "膣内セックス"},
            {"en": "handjob", "ja": "手コキ"},
            {"en": "footjob", "ja": "足コキ"},
            {"en": "titjob", "ja": "パイズリ"},
            {"en": "rape", "ja": "レイプ"},
            {"en": "gangbang", "ja": "輪姦"},
            {"en": "threesome", "ja": "三人組"},
            {"en": "mmf threesome", "ja": "男男女性三人組"},
            {"en": "ffm threesome", "ja": "女女男性三人組"},
            {"en": "foursome", "ja": "四人組"},
            {"en": "group sex", "ja": "集団セックス"},
            {"en": "onee-shota", "ja": "お姉さんショタ"},
            {"en": "yaoi", "ja": "ヤオイ"},
            {"en": "yuri", "ja": "百合"},
            {"en": "incest", "ja": "近親相姦"},
            {"en": "bdsm", "ja": "BDSM"},
            {"en": " bondage", "ja": "緊縛"},
            {"en": "spanking", "ja": "尻叩き"},
            {"en": "fisting", "ja": "フィストファック"}
          ]
        },
         {
          "id": "accessories",
          "name": "アクセサリー",
          "parent_id": None,
          "tags": [
            {"en": "necklace", "ja": "ネックレス"},
            {"en": "earrings", "ja": "イヤリング"},
            {"en": "bracelet", "ja": "ブレスレット"},
            {"en": "ring", "ja": "指輪"},
            {"en": "choker", "ja": "チョーカー"},
            {"en": "hair ornament", "ja": "髪飾り"},
            {"en": "headband", "ja": "カチューシャ"},
            {"en": "hairpin", "ja": "ヘアピン"},
            {"en": "piercing", "ja": "ピアス"},
            {"en": "tattoo", "ja": "タトゥー"},
            {"en": "mask", "ja": "マスク"},
            {"en": "eyepatch", "ja": "眼帯"}
          ]
        },
         {
          "id": "expression",
          "name": "表情",
          "parent_id": None,
          "tags": [
            {"en": "smile", "ja": "笑顔"},
            {"en": "laughing", "ja": "笑っている"},
            {"en": "blush", "ja": "赤面"},
            {"en": "embarrassed", "ja": "恥ずかしがっている"},
            {"en": "angry", "ja": "怒っている"},
            {"en": "sad", "ja": "悲しい"},
            {"en": "crying", "ja": "泣いている"},
            {"en": "shocked", "ja": "驚いている"},
            {"en": "confused", "ja": "困惑している"},
            {"en": "ahegao", "ja": "アヘ顔"},
            {"en": "ecstasy", "ja": "絶頂"},
            {"en": "pain", "ja": "苦痛"},
            {"en": "pleasure", "ja": "快感"}
          ]
        },
        {
          "id": "camera_angle",
          "name": "カメラアングル",
          "parent_id": None,
          "tags": [
            {"en": "low angle", "ja": "ローアングル"},
            {"en": "high angle", "ja": "ハイアングル"},
            {"en": "eye level", "ja": "アイレベル"},
            {"en": "dutch angle", "ja": "ダッチアングル"}
          ]
        },
        {
          "id": "composition",
          "name": "構図",
          "parent_id": None,
          "tags": [
            {"en": "full body", "ja": "全身構図"},
            {"en": "upper body", "ja": "上半身構図"},
            {"en": "close-up", "ja": "クローズアップ"},
            {"en": "rule of thirds", "ja": "三分割法"},
            {"en": "leading lines", "ja": "誘導線"}
          ]
        },
        {
          "id": "time",
          "name": "時間",
          "parent_id": None,
          "tags": [
            {"en": "morning", "ja": "朝"},
            {"en": "noon", "ja": "昼"},
            {"en": "evening", "ja": "夕方"},
            {"en": "night", "ja": "夜"},
            {"en": "spring", "ja": "春"},
            {"en": "summer", "ja": "夏"},
            {"en": "autumn", "ja": "秋"},
            {"en": "winter", "ja": "冬"}
          ]
        }
      ]
    }
    app_state['dictionary'] = initial_data
    save_dictionary()
    messagebox.showinfo("情報", "初期辞書を生成しました。")
    populate_dict_treeview() # 辞書管理タブのTreeviewを更新
    update_available_tags_treeview() # タグセット生成タブも更新
    # populate_category_hierarchy_treeview(category_hierarchy_tree_manage) # 辞書管理タブの階層Treeviewを更新
    # populate_category_hierarchy_treeview(category_hierarchy_tree_classify) # 分類タブの階層Treeviewを更新


def get_classification_hint(tag_en):
    """タグの自動分類ヒントを生成する関数"""
    hints = []
    tag_en_lower = tag_en.lower()

    for category in app_state['dictionary'].get('categories', []):
        for dict_tag in category.get('tags', []):
            dict_tag_en_lower = dict_tag['en'].lower()
            if tag_en_lower == dict_tag_en_lower:
                hints.append({
                    'type': '完全一致',
                    'category_id': category['id'],
                    'category_path': get_category_path(category['id']),
                    'tag_en': dict_tag['en'],
                    'tag_ja': dict_tag.get('ja', '説明なし'),
                    'confidence': 1.0
                })
                return hints

            if dict_tag_en_lower in tag_en_lower:
                 hints.append({
                     'type': '部分一致 (含む)',
                     'category_id': category['id'],
                     'category_path': get_category_path(category['id']),
                     'tag_en': dict_tag['en'],
                     'tag_ja': dict_tag.get('ja', '説明なし'),
                     'confidence': 0.8
                 })

            if tag_en_lower in dict_tag_en_lower:
                 hints.append({
                     'type': '部分一致 (含まれる)',
                     'category_id': category['id'],
                     'category_path': get_category_path(category['id']),
                     'tag_en': dict_tag['en'],
                     'tag_ja': dict_tag.get('ja', '説明なし'),
                     'confidence': 0.7
                 })

    words = re.split(r'[ _-]', tag_en_lower)
    for word in words:
        if not word: continue
        for category in app_state['dictionary'].get('categories', []):
            if word in category['name'].lower():
                 hints.append({
                     'type': 'カテゴリ名に単語一致',
                     'category_id': category['id'],
                     'category_path': get_category_path(category['id']),
                     'tag_en': None,
                     'tag_ja': None,
                     'confidence': 0.6
                 })
            for dict_tag in category.get('tags', []):
                 dict_words = re.split(r'[ _-]', dict_tag['en'].lower())
                 if word in dict_words:
                      hints.append({
                          'type': '辞書タグの単語に一致',
                          'category_id': category['id'],
                          'category_path': get_category_path(category['id']),
                          'tag_en': dict_tag['en'],
                          'tag_ja': dict_tag.get('ja', '説明なし'),
                          'confidence': 0.5
                      })

    hints = sorted(hints, key=lambda x: x['confidence'], reverse=True)
    seen_hints = set()
    unique_hints = []
    for hint in hints:
        hint_identifier = (hint['category_id'], hint.get('tag_en'))
        if hint_identifier not in seen_hints:
            seen_hints.add(hint_identifier)
            unique_hints.append(hint)

    return unique_hints

# --- UI更新ヘルパー ---

def update_treeview(tree_widget, df_data):
    """Treeviewの内容を更新する汎用関数"""
    for item in tree_widget.get_children():
        tree_widget.delete(item)
    
    for index, row in df_data.iterrows():
        # Ensure all values are strings before inserting into Treeview
        values_as_strings = [str(val) if val is not None else '' for val in list(row)]
        tree_widget.insert("", "end", iid=index, values=values_as_strings)
    
    tree_widget.selection_remove(tree_widget.selection())

def update_category_dropdowns():
    """全てのカテゴリドロップダウンのオプションを更新する関数"""
    category_options_list = ["--カテゴリを選択--"]
    category_path_to_id_map = {"--カテゴリを選択--": None}

    def flatten_categories_for_dropdown(categories, parent_path=""):
        for cat in categories:
            current_path = f"{parent_path} / {cat['name']}" if parent_path else cat['name']
            category_options_list.append(current_path)
            category_path_to_id_map[current_path] = cat['id']
            # 子カテゴリも再帰的に追加
            children = [c for c in app_state['dictionary']['categories'] if c.get('parent_id') == cat['id']]
            flatten_categories_for_dropdown(children, current_path)


    top_level_categories = [cat for cat in app_state['dictionary'].get('categories', []) if cat.get('parent_id') is None]
    flatten_categories_for_dropdown(top_level_categories)

    global all_category_options, all_category_path_to_id
    all_category_options = category_options_list
    all_category_path_to_id = category_path_to_id_map

    # 各コンボボックスの値を更新
    # ここでグローバル変数がNoneでないことを確認してから更新
    if 'unclassified_category_combobox' in globals() and unclassified_category_combobox is not None:
        unclassified_category_combobox['values'] = all_category_options
    if 'dict_category_combobox' in globals() and dict_category_combobox is not None:
        dict_category_combobox['values'] = all_category_options
    if 'tag_gen_filter_combobox' in globals() and tag_gen_filter_combobox is not None:
        tag_gen_filter_combobox['values'] = ["--全てのカテゴリ--"] + all_category_options[1:]
    if 'dict_filter_combobox' in globals() and dict_filter_combobox is not None: # 新しいフィルタ
        dict_filter_combobox['values'] = ["--全てのカテゴリ--"] + all_category_options[1:]
    
    # ここで add_tag_category_combobox の値を更新
    if 'new_category_parent_combobox' in globals() and new_category_parent_combobox is not None:
        new_category_parent_combobox['values'] = ["--カテゴリを選択--"] + all_category_options[1:]
    if 'new_category_parent_combobox_classify' in globals() and new_category_parent_combobox_classify is not None:
        new_category_parent_combobox_classify['values'] = ["--カテゴリを選択--"] + all_category_options[1:]
    if 'add_tag_category_combobox' in globals() and add_tag_category_combobox is not None: # 新しいタグ追加用
        add_tag_category_combobox['values'] = all_category_options


def treeview_sort_column(tree_widget, col_name, reverse):
    """Treeviewの指定された列でソートする関数"""
    # ツリー表示の場合、#0列はソートしない
    if tree_widget.cget("show") == "tree" and col_name == "#0":
        return

    # Treeviewが階層表示の場合、アイテムのテキスト（#0列）でソートするか、valuesでソートするかを区別
    if tree_widget.cget("show") == "tree" or tree_widget.cget("show") == "tree headings":
        # #0列のソート
        if col_name == "#0":
            l = [(tree_widget.item(k, 'text'), k) for k in tree_widget.get_children('')]
        else:
            # その他の列のソート
            l = [(tree_widget.set(k, col_name), k) for k in tree_widget.get_children('')]
    else:
        # 通常のTreeview（headingsのみ）のソート
        l = [(tree_widget.set(k, col_name), k) for k in tree_widget.get_children('')]
    
    # ソートキーの型を考慮したソート（ここでは文字列としてソート）
    l.sort(key=lambda t: t[0], reverse=reverse)

    # ソートされた順序でアイテムを再配置
    for index, (val, k) in enumerate(l):
        tree_widget.move(k, '', index)

    # 次回クリック時のソート方向を反転
    tree_widget_name = str(tree_widget)
    sort_key = (tree_widget_name, col_name)
    sort_reverse_flags[sort_key] = not reverse

    # ヘッディングのテキストにソート方向を示す矢印を追加/更新
    current_heading_text = tree_widget.heading(col_name, "text")
    # 既存の矢印を削除
    clean_heading_text = re.sub(r' \u25b2| \u25bc', '', current_heading_text)
    # 新しい矢印を追加
    arrow = ' \u25b2' if not reverse else ' \u25bc' # True (降順) なら下矢印、False (昇順) なら上矢印
    tree_widget.heading(col_name, text=clean_heading_text + arrow, 
                        command=lambda _col_name=col_name: treeview_sort_column(tree_widget, _col_name, sort_reverse_flags.get((tree_widget_name, _col_name), False)))


# --- カテゴリ追加機能 ---
def add_new_category(name_entry, parent_combobox, target_notebook_tab_index=None):
    new_name = name_entry.get().strip()
    parent_path = parent_combobox.get()

    if not new_name:
        messagebox.showwarning("警告", "新しいカテゴリ名を入力してください。")
        return

    # Check for duplicate category names at the same level (simple check)
    parent_id = all_category_path_to_id.get(parent_path) if parent_path != "--カテゴリを選択--" else None

    for cat in app_state['dictionary']['categories']:
        if cat['name'].lower() == new_name.lower() and cat.get('parent_id') == parent_id:
            messagebox.showwarning("警告", f"カテゴリ '{new_name}' は既に存在します。")
            return

    # Generate a simple unique ID
    # UUIDの使用を推奨
    new_id = str(uuid.uuid4()) # ユニークなIDを生成
    
    app_state['dictionary']['categories'].append({
        "id": new_id,
        "name": new_name,
        "parent_id": parent_id,
        "tags": []
    })
    save_dictionary()
    messagebox.showinfo("情報", f"カテゴリ '{new_name}' を追加しました。")
    name_entry.delete(0, tk.END)
    parent_combobox.set("--カテゴリを選択--")
    update_category_dropdowns() # Update all dropdowns including the new category parent combobox
    populate_dict_treeview() # Refresh dictionary treeview
    populate_category_hierarchy_treeview(category_hierarchy_tree_manage) # 辞書管理タブの階層Treeviewを更新
    populate_category_hierarchy_treeview(category_hierarchy_tree_classify) # 分類タブの階層Treeviewを更新
    populate_available_categories_treeview() # タグセット生成タブのカテゴリツリーを更新
    
    # カテゴリ追加後、元のタブに戻る
    if target_notebook_tab_index is not None:
        notebook.select(target_notebook_tab_index)

# --- タグ直接追加機能 ---
def add_direct_tag(english_entry, japanese_entry, category_combobox):
    tag_en = english_entry.get().strip()
    tag_ja = japanese_entry.get().strip() # stripを適用
    category_path = category_combobox.get()

    if not tag_en:
        messagebox.showwarning("警告", "英語タグ名を入力してください。")
        return
    if category_path == "--カテゴリを選択--":
        messagebox.showwarning("警告", "カテゴリを選択してください。")
        return

    category_id = all_category_path_to_id.get(category_path)
    if category_id is None:
        messagebox.showwarning("エラー", "選択されたカテゴリが無効です。")
        return

    success, message = add_tag_to_dictionary(tag_en, tag_ja, category_id)
    if success:
        save_dictionary()
        messagebox.showinfo("情報", message)
        english_entry.delete(0, tk.END)
        japanese_entry.delete(0, tk.END)
        category_combobox.set("--カテゴリを選択--")
        populate_dict_treeview(dict_search_entry.get(), dict_filter_var.get()) # フィルタを維持して更新
        update_category_dropdowns()
        update_available_tags_treeview()
        populate_category_hierarchy_treeview(category_hierarchy_tree_manage)
        populate_category_hierarchy_treeview(category_hierarchy_tree_classify)
    else:
        messagebox.showwarning("警告", message)


# --- カテゴリ階層Treeviewのヘルパー関数 ---
def get_all_categories_flat_map():
    """全カテゴリをIDをキーとする辞書として返すヘルパー関数"""
    return {cat['id']: cat for cat in app_state['dictionary']['categories']}

def populate_category_hierarchy_treeview(tree_widget):
    """カテゴリ階層Treeviewにデータをロードする関数"""
    # TreeviewがNoneでないことを確認
    if tree_widget is None:
        return

    # Treeviewをクリア
    for item in tree_widget.get_children():
        tree_widget.delete(item)

    all_categories_map = get_all_categories_flat_map()

    def insert_category_into_tree(category_info, parent_iid=""):
        # カテゴリに含まれるタグの総数を計算
        total_tags = len(category_info.get('tags', []))
        
        # Treeviewにカテゴリを挿入
        iid = tree_widget.insert(parent_iid, "end", text=f"📂 {category_info['name']} ({total_tags}タグ)", open=False, values=(category_info['id'], category_info['name']))
        
        # このカテゴリに直接属するタグを子として挿入
        for tag in category_info.get('tags', []):
            tree_widget.insert(iid, "end", text=f"  - {tag['en']} ({tag.get('ja', '説明なし')})", values=("tag",))

        # このカテゴリの子カテゴリを検索し、再帰的に挿入
        children_categories = [cat for cat in app_state['dictionary']['categories'] if cat.get('parent_id') == category_info['id']]
        for child_cat in children_categories:
            insert_category_into_tree(child_cat, iid)

    # トップレベルカテゴリ（parent_idがNoneのカテゴリ）を挿入
    top_level_categories = [cat for cat in app_state['dictionary']['categories'] if cat.get('parent_id') is None]
    for category in top_level_categories:
        insert_category_into_tree(category)

def show_category_tree_context_menu(event, tree_widget, name_entry_widget, parent_combobox_widget, notebook_widget, target_tab_frame):
    """カテゴリ階層Treeviewの右クリックメニューを表示する"""
    item_id = tree_widget.identify_row(event.y)
    if not item_id:
        return

    # クリックされたアイテムの値をチェック
    item_values = tree_widget.item(item_id, 'values')
    if not item_values or item_values[0] == "tag": # タグの場合はメニューを表示しない
        return

    category_id = item_values[0] # ここで正確なカテゴリIDを取得
    category_name = item_values[1] # カテゴリ名を取得 (表示用)

    context_menu = tk.Menu(tree_widget, tearoff=0)
    context_menu.add_command(label=f"'{category_name}' の子カテゴリを追加", 
                              command=lambda: set_parent_category_and_switch_tab(category_id, name_entry_widget, parent_combobox_widget, notebook_widget, target_tab_frame))
    context_menu.add_command(label=f"'{category_name}' を削除", 
                             command=lambda: delete_category(category_id, category_name))
    
    try:
        context_menu.tk_popup(event.x_root, event.y_root)
    finally:
        context_menu.grab_release()

def delete_category(category_id, category_name):
    """カテゴリを削除する関数"""
    category_to_delete = find_category_by_id(category_id)

    if not category_to_delete:
        messagebox.showerror("エラー", "指定されたカテゴリが見つかりません。")
        return

    # 子カテゴリの存在チェック
    children_categories = [cat for cat in app_state['dictionary']['categories'] if cat.get('parent_id') == category_id]
    if children_categories:
        messagebox.showwarning("警告", f"カテゴリ '{category_name}' には子カテゴリが存在するため削除できません。\n先に子カテゴリを削除してください。")
        return

    # タグの存在チェック
    if category_to_delete.get('tags') and len(category_to_delete['tags']) > 0:
        messagebox.showwarning("警告", f"カテゴリ '{category_name}' にはタグが存在するため削除できません。\n先にタグを別のカテゴリに移動するか、削除してください。")
        return

    if messagebox.askyesno("確認", f"カテゴリ '{category_name}' を本当に削除しますか？\nこの操作は元に戻せません。"):
        app_state['dictionary']['categories'] = [
            cat for cat in app_state['dictionary']['categories'] if cat['id'] != category_id
        ]
        save_dictionary()
        messagebox.showinfo("情報", f"カテゴリ '{category_name}' を削除しました。")
        update_category_dropdowns()
        populate_dict_treeview(dict_search_entry.get(), dict_filter_var.get())
        populate_category_hierarchy_treeview(category_hierarchy_tree_manage)
        populate_category_hierarchy_treeview(category_hierarchy_tree_classify)
        populate_available_categories_treeview() # タグセット生成タブのカテゴリツリーを更新


def set_parent_category_and_switch_tab(category_id, name_entry_widget, parent_combobox_widget, notebook_widget, target_tab_frame):
    """親カテゴリを設定し、指定されたタブに切り替える"""
    # category_id を直接使用してパスを取得
    category_path_to_set = get_category_path(category_id)

    parent_combobox_widget.set(category_path_to_set)
    notebook_widget.select(notebook_widget.index(target_tab_frame)) # 指定されたタブに切り替え


# --- タブごとのUI構築関数 ---

def create_file_management_tab(notebook_frame):
    """ファイル管理タブのUIを構築する関数"""
    tab_frame = ttk.Frame(notebook_frame, padding="10")

    # 全てのファイル操作をまとめる新しいフレーム
    file_operations_group_frame = ttk.LabelFrame(tab_frame, text="辞書ファイルの入出力", padding="10")
    file_operations_group_frame.pack(fill=tk.X, pady=10, padx=10)

    # JSONファイル操作
    file_ops_json_frame = ttk.LabelFrame(file_operations_group_frame, text="JSONファイル操作", padding="10")
    file_ops_json_frame.pack(fill=tk.X, pady=5)
    
    inner_json_frame = tk.Frame(file_ops_json_frame)
    inner_json_frame.pack(fill=tk.BOTH, expand=True)

    ttk.Button(inner_json_frame, text="既存の辞書JSONをアップロード", command=upload_dictionary_file).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
    ttk.Button(inner_json_frame, text="現在の辞書JSONをダウンロード", command=download_dictionary_file).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
    ttk.Button(inner_json_frame, text="デモ用初期辞書を生成", command=generate_initial_dictionary).grid(row=0, column=2, padx=5, pady=5, sticky="ew")
    ttk.Button(inner_json_frame, text="追加辞書JSONをインポート (マージ)", command=import_additional_dictionary_json).grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky="ew")

    # 日本語説明の一括追加・更新 (CSV関連)
    bulk_ja_frame = ttk.LabelFrame(file_operations_group_frame, text="日本語説明の一括追加・更新 (CSV)", padding="10")
    bulk_ja_frame.pack(fill=tk.X, pady=5)
    
    inner_bulk_ja_frame = tk.Frame(bulk_ja_frame)
    inner_bulk_ja_frame.pack(fill=tk.BOTH, expand=True)

    ttk.Button(inner_bulk_ja_frame, text="日本語説明がないタグをエクスポート (.txt)", command=export_tags_without_ja).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
    ttk.Button(inner_bulk_ja_frame, text="翻訳済みタグをインポート (.csv)", command=import_translated_tags).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
    ttk.Button(inner_bulk_ja_frame, text="全タグをCSVでエクスポート (編集用)", command=export_all_tags_to_csv).grid(row=0, column=2, padx=5, pady=5, sticky="ew")

    return tab_frame


def create_manage_dictionary_tab(notebook_frame):
    """カテゴリ・辞書管理タブのUIを構築する関数"""
    tab_frame = ttk.Frame(notebook_frame, padding="10")

    # PanedWindowで左右に分割
    paned_window = ttk.PanedWindow(tab_frame, orient=tk.HORIZONTAL)
    paned_window.pack(fill=tk.BOTH, expand=True)

    # 左側のフレーム (カテゴリ階層の閲覧)
    left_frame = ttk.Frame(paned_window, padding="10")
    paned_window.add(left_frame, weight=1) # 左側を伸縮可能に

    category_view_frame = ttk.LabelFrame(left_frame, text="カテゴリ階層の閲覧", padding="10")
    category_view_frame.pack(fill=tk.BOTH, expand=True)

    global category_hierarchy_tree_manage # 辞書管理タブ用のTreeview
    hierarchy_scrollbar_manage = ttk.Scrollbar(category_view_frame, orient="vertical")
    hierarchy_scrollbar_manage.pack(side="right", fill="y")
    category_hierarchy_tree_manage = ttk.Treeview(category_view_frame, show="tree", selectmode="browse", yscrollcommand=hierarchy_scrollbar_manage.set)
    category_hierarchy_tree_manage.pack(side="left", fill=tk.BOTH, expand=True)
    hierarchy_scrollbar_manage.config(command=category_hierarchy_tree_manage.yview)

    # 右クリックメニューのバインド
    # ここで tab_frame を明示的にキャプチャ
    category_hierarchy_tree_manage.bind("<Button-3>", 
                                        lambda e, tf=tab_frame: show_category_tree_context_menu(e, category_hierarchy_tree_manage, 
                                                                                new_category_name_entry, new_category_parent_combobox, 
                                                                                notebook, tf))
    
    # 初期ロードはmain関数で行う

    # 右側のフレーム (カテゴリ追加、タグ編集、一括更新)
    right_frame = ttk.Frame(paned_window, padding="10")
    paned_window.add(right_frame, weight=2) # 右側をより広く伸縮可能に

    # カテゴリとタグの追加機能をまとめる新しいフレーム
    add_category_and_tag_frame = ttk.LabelFrame(right_frame, text="カテゴリとタグの追加", padding="10")
    add_category_and_tag_frame.pack(fill=tk.X, pady=5)

    # カテゴリ追加機能とタグ直接追加機能を横並びにするためのフレーム
    input_sections_frame = ttk.Frame(add_category_and_tag_frame)
    input_sections_frame.pack(fill=tk.BOTH, expand=True)

    # カテゴリ追加セクション
    category_add_section = ttk.LabelFrame(input_sections_frame, text="新しいカテゴリを追加", padding="10")
    category_add_section.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
    input_sections_frame.grid_columnconfigure(0, weight=1) # カラムの伸縮設定

    ttk.Label(category_add_section, text="新しいカテゴリ名:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    global new_category_name_entry
    new_category_name_entry = ttk.Entry(category_add_section)
    new_category_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

    ttk.Label(category_add_section, text="親カテゴリ:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
    global new_category_parent_combobox
    new_category_parent_combobox = ttk.Combobox(category_add_section, state="readonly")
    new_category_parent_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
    new_category_parent_combobox.set("--カテゴリを選択--")

    ttk.Button(category_add_section, text="カテゴリを追加", command=lambda: add_new_category(new_category_name_entry, new_category_parent_combobox, notebook.index(tab_frame))).grid(row=2, column=0, columnspan=2, pady=10)
    category_add_section.grid_columnconfigure(1, weight=1) # カラムの伸縮設定

    # タグ直接追加セクション
    tag_add_section = ttk.LabelFrame(input_sections_frame, text="タグを直接追加", padding="10")
    tag_add_section.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
    input_sections_frame.grid_columnconfigure(1, weight=1) # カラムの伸縮設定

    ttk.Label(tag_add_section, text="英語タグ名:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    global add_tag_english_entry
    add_tag_english_entry = ttk.Entry(tag_add_section)
    add_tag_english_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

    ttk.Label(tag_add_section, text="日本語説明:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
    global add_tag_japanese_entry
    add_tag_japanese_entry = ttk.Entry(tag_add_section)
    add_tag_japanese_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

    ttk.Label(tag_add_section, text="カテゴリ:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
    global add_tag_category_combobox
    add_tag_category_combobox = ttk.Combobox(tag_add_section, state="readonly")
    add_tag_category_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
    add_tag_category_combobox.set("--カテゴリを選択--")

    ttk.Button(tag_add_section, text="タグを追加", command=lambda: add_direct_tag(add_tag_english_entry, add_tag_japanese_entry, add_tag_category_combobox)).grid(row=3, column=0, columnspan=2, pady=10)
    tag_add_section.grid_columnconfigure(1, weight=1) # カラムの伸縮設定


    edit_frame = ttk.LabelFrame(right_frame, text="カテゴリとタグの編集", padding="10")
    edit_frame.pack(fill=tk.BOTH, expand=True, pady=5)

    # 検索・フィルタ機能
    filter_search_frame_dict = ttk.Frame(edit_frame, padding="5")
    filter_search_frame_dict.pack(fill=tk.X, pady=5)
    
    ttk.Label(filter_search_frame_dict, text="カテゴリで絞り込み:").pack(side="left", padx=5)
    global dict_filter_var, dict_filter_combobox
    dict_filter_var = tk.StringVar(root)
    dict_filter_combobox = ttk.Combobox(filter_search_frame_dict, textvariable=dict_filter_var, state="readonly", values=["--全てのカテゴリ--"] + all_category_options[1:])
    dict_filter_combobox.set("--全てのカテゴリ--")
    dict_filter_combobox.bind("<<ComboboxSelected>>", lambda e: populate_dict_treeview(dict_search_entry.get(), dict_filter_var.get()))
    dict_filter_combobox.pack(side="left", padx=5, expand=True, fill=tk.X)

    ttk.Label(filter_search_frame_dict, text="タグを検索:").pack(side="left", padx=5)
    global dict_search_entry
    dict_search_entry = ttk.Entry(filter_search_frame_dict)
    dict_search_entry.bind("<KeyRelease>", lambda e: populate_dict_treeview(dict_search_entry.get(), dict_filter_var.get()))
    dict_search_entry.pack(side="left", padx=5, expand=True, fill=tk.X)


    columns = ("英語タグ名", "日本語説明", "カテゴリ")
    global dict_tree
    dict_tree = ttk.Treeview(edit_frame, columns=columns, show="headings", selectmode="extended")

    # ヘッディングとソート機能のバインド
    for col_name in columns:
        dict_tree.heading(col_name, text=col_name, command=lambda _col_name=col_name: treeview_sort_column(dict_tree, _col_name, sort_reverse_flags.get((str(dict_tree), _col_name), False)))
        dict_tree.column(col_name, width=200, anchor="w")

    scrollbar = ttk.Scrollbar(edit_frame, orient="vertical", command=dict_tree.yview)
    dict_tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    dict_tree.pack(fill=tk.BOTH, expand=True)

    def on_dict_tree_double_click(event):
        item_id = dict_tree.focus()
        if not item_id: return
        column_id = dict_tree.identify_column(event.x)
        column_index = int(column_id[1:]) - 1

        if column_index == 0: return

        row_index = int(item_id)
        current_value = dict_tree.item(item_id, 'values')[column_index]

        x, y, width, height = dict_tree.bbox(item_id, column_id)
        
        if column_index == 2:
            editor = ttk.Combobox(dict_tree, values=all_category_options, state="readonly")
            editor.set(current_value)
            def on_combobox_select(event):
                new_value = editor.get()
                app_state['edited_dict_df'].loc[row_index, columns[column_index]] = new_value
                dict_tree.item(item_id, values=list(app_state['edited_dict_df'].loc[row_index][["英語タグ名", "日本語説明", "カテゴリ"]]))
                editor.destroy()
                # コンボボックス選択後にフォーカスをTreeviewに戻す
                dict_tree.focus_set()
            editor.bind("<<ComboboxSelected>>", on_combobox_select)
            editor.bind("<FocusOut>", lambda e: editor.destroy())
        else:
            editor = ttk.Entry(dict_tree)
            editor.insert(0, current_value)
            def on_entry_return(event):
                new_value = editor.get().strip() # stripを適用
                app_state['edited_dict_df'].loc[row_index, columns[column_index]] = new_value
                dict_tree.item(item_id, values=list(app_state['edited_dict_df'].loc[row_index][["英語タグ名", "日本語説明", "カテゴリ"]]))
                editor.destroy()
                # エンターキー押下後にフォーカスをTreeviewに戻す
                dict_tree.focus_set()
            editor.bind("<Return>", on_entry_return)
            editor.bind("<FocusOut>", lambda e: editor.destroy())
        
        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()

    dict_tree.bind("<Double-1>", on_dict_tree_double_click)

    def show_dict_tree_tag_context_menu(event):
        item_id = dict_tree.identify_row(event.y)
        if not item_id: return
        
        # 選択されたアイテムの英語タグを取得
        english_tag = dict_tree.item(item_id, 'values')[0]

        context_menu = tk.Menu(dict_tree, tearoff=0)
        context_menu.add_command(label=f"'{english_tag}' をコピー", command=lambda: copy_to_clipboard(english_tag))
        
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

    def copy_to_clipboard(text):
        root.clipboard_clear()
        root.clipboard_append(text)
        messagebox.showinfo("情報", f"'{text}' をクリップボードにコピーしました。")

    dict_tree.bind("<Button-3>", show_dict_tree_tag_context_menu) # 右クリックでメニュー表示

    bulk_apply_frame = ttk.Frame(edit_frame, padding="10")
    bulk_apply_frame.pack(fill=tk.X, pady=5)

    ttk.Label(bulk_apply_frame, text="選択したタグにまとめてカテゴリを適用:").pack(side="left", padx=5)
    global dict_category_var, dict_category_combobox
    dict_category_var = tk.StringVar(root)
    dict_category_combobox = ttk.Combobox(bulk_apply_frame, textvariable=dict_category_var, state="readonly", values=all_category_options)
    dict_category_combobox.set(all_category_options[0])
    dict_category_combobox.pack(side="left", padx=5, expand=True, fill=tk.X)
    ttk.Button(bulk_apply_frame, text="選択したタグに適用", command=apply_selected_category_dict_tab).pack(side="left", padx=5)

    # 「タグの変更を保存」と「選択したタグを削除」ボタンを横並びにするためのフレーム
    save_delete_buttons_frame = ttk.Frame(edit_frame, padding="5")
    save_delete_buttons_frame.pack(fill=tk.X, pady=5)

    ttk.Button(save_delete_buttons_frame, text="タグの変更を保存", command=save_dict_changes).pack(side="left", expand=True, padx=5)
    ttk.Button(save_delete_buttons_frame, text="選択したタグを削除", command=delete_selected_tags).pack(side="left", expand=True, padx=5)
    
    # 初期ロードはmain関数で行う

    return tab_frame

def delete_selected_tags():
    """辞書管理タブで選択したタグを辞書から削除する関数"""
    selected_items = dict_tree.selection()
    if not selected_items:
        messagebox.showwarning("警告", "削除するタグを選択してください。")
        return

    if not messagebox.askyesno("確認", f"{len(selected_items)}件のタグを本当に削除しますか？\nこの操作は元に戻せません。"):
        return

    deleted_count = 0
    tags_to_delete_en = {dict_tree.item(item, 'values')[0] for item in selected_items} # 選択されたアイテムの英語タグ名を取得

    # 新しい辞書構造を構築し、削除対象のタグを含めない
    new_dictionary_structure = {"categories": []}
    for category in app_state['dictionary'].get('categories', []):
        new_category = {
            "id": category['id'],
            "name": category['name'],
            "parent_id": category.get('parent_id'),
            "tags": []
        }
        for tag in category.get('tags', []):
            if tag['en'] not in tags_to_delete_en:
                new_category['tags'].append(tag)
            else:
                deleted_count += 1
        new_dictionary_structure['categories'].append(new_category)

    app_state['dictionary'] = new_dictionary_structure
    save_dictionary()
    messagebox.showinfo("情報", f"{deleted_count}件のタグを辞書から削除しました。")
    populate_dict_treeview(dict_search_entry.get(), dict_filter_var.get()) # Treeviewを更新
    update_category_dropdowns() # ドロップダウンリストを更新
    update_available_tags_treeview() # タグセット生成タブのリストを更新
    populate_category_hierarchy_treeview(category_hierarchy_tree_manage) # 辞書管理タブの階層Treeviewを更新
    populate_category_hierarchy_treeview(category_hierarchy_tree_classify) # 分類タブの階層Treeviewも更新


def export_all_tags_to_csv():
    """辞書内の全てのタグを英語タグ名と日本語説明のCSVでエクスポートする"""
    all_tags_for_export = []
    for category in app_state['dictionary'].get('categories', []):
        for tag in category.get('tags', []):
            all_tags_for_export.append({
                "English Tag": tag.get('en', ''),
                "日本語説明": tag.get('ja', '')
            })
    
    if not all_tags_for_export:
        messagebox.showinfo("情報", "エクスポートするタグがありません。")
        return

    export_df = pd.DataFrame(all_tags_for_export)

    filepath = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSVファイル", "*.csv")],
        title="全タグをCSVでエクスポート"
    )
    if filepath:
        try:
            export_df.to_csv(filepath, index=False, encoding='utf-8-sig') # Excelで開けるようにutf-8-sig
            messagebox.showinfo("情報", f"全タグ ({len(all_tags_for_export)}件) をCSVでエクスポートしました。")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの保存中にエラーが発生しました: {e}")


def populate_dict_treeview(search_query="", filter_category_path="--全てのカテゴリ--"):
    """辞書管理タブのTreeviewにデータをロードする (検索・フィルタ機能付き)"""
    all_tags_data = []
    search_query_lower = search_query.lower()
    filter_category_id = all_category_path_to_id.get(filter_category_path)

    for category in app_state['dictionary'].get('categories', []):
        category_path = get_category_path(category['id'])
        
        is_under_filter = False
        if filter_category_id is None or filter_category_path == "--全てのカテゴリ--":
            is_under_filter = True
        else:
            # 現在のカテゴリがフィルタカテゴリの子孫であるか、またはフィルタカテゴリ自体であるかをチェック
            current_cat_id = category['id']
            while current_cat_id:
                if current_cat_id == filter_category_id:
                    is_under_filter = True
                    break
                parent_cat = find_category_by_id(current_cat_id)
                current_cat_id = parent_cat.get('parent_id') if parent_cat else None
        
        if is_under_filter:
            for tag in category.get('tags', []):
                tag_en = tag.get('en', '')
                tag_ja = tag.get('ja', '')
                
                # 検索クエリでフィルタリング
                if search_query_lower in tag_en.lower() or search_query_lower in tag_ja.lower():
                    all_tags_data.append({
                        "英語タグ名": tag_en,
                        "日本語説明": tag_ja,
                        "カテゴリ": category_path,
                        "_category_id": tag.get('id', '')
                    })
    app_state['edited_dict_df'] = pd.DataFrame(all_tags_data, columns=["英語タグ名", "日本語説明", "カテゴリ", "_category_id"])
    app_state['edited_dict_df'] = app_state['edited_dict_df'].fillna('')
    # dict_treeがNoneでないことを確認
    if dict_tree is not None:
        update_treeview(dict_tree, app_state['edited_dict_df'][["英語タグ名", "日本語説明", "カテゴリ"]])

def apply_selected_category_dict_tab():
    """辞書管理タブで選択したタグにカテゴリを一括適用する"""
    selected_category_path = dict_category_var.get()
    if selected_category_path == "--カテゴリを選択--":
        messagebox.showwarning("警告", "適用するカテゴリを選択してください。")
        return

    selected_items = dict_tree.selection()
    if not selected_items:
        messagebox.showwarning("警告", "カテゴリを適用するタグを選択してください。")
        return

    selected_indices = [int(item) for item in selected_items]
    app_state['edited_dict_df'].loc[selected_indices, 'カテゴリ'] = selected_category_path
    
    for index in selected_indices:
        dict_tree.item(index, values=list(app_state['edited_dict_df'].loc[index][["英語タグ名", "日本語説明", "カテゴリ"]]))
    
    messagebox.showinfo("情報", f"{len(selected_indices)}件のタグにカテゴリ '{selected_category_path}' を適用しました。変更を保存するには「タグの変更を保存」ボタンを押してください。")

def save_dict_changes():
    """辞書管理タブでの変更を辞書データに反映し保存する"""
    if app_state['edited_dict_df'].empty and not app_state['dictionary'].get('categories'):
        messagebox.showwarning("警告", "保存する変更がありません。")
        return

    # 現在の辞書データをディープコピーして、タグ情報を操作するためのフラットなマップを作成
    # {english_tag_lower: {'en': '...', 'ja': '...', 'category_id': '...'}}
    all_tags_flat = {}
    for category in app_state['dictionary'].get('categories', []):
        for tag in category.get('tags', []):
            all_tags_flat[tag['en'].lower()] = {
                'en': tag['en'],
                'ja': tag.get('ja', ''), # ここは読み込んだまま
                'category_id': category['id']
            }

    updated_count = 0
    added_count = 0

    # edited_dict_df の内容で all_tags_flat を更新
    # edited_dict_df にあるタグは、元の辞書から更新されたもの、または新規追加されたもの
    for index, row in app_state['edited_dict_df'].iterrows():
        tag_en = row["英語タグ名"]
        tag_ja = row["日本語説明"].strip() # stripを適用
        category_path = row["カテゴリ"]
        category_id = get_category_id_from_path(category_path)

        if category_id is None:
            messagebox.showwarning("警告", f"タグ '{tag_en}': 無効なカテゴリパス '{category_path}' です。このタグは保存されません。")
            continue

        if tag_en.lower() in all_tags_flat:
            # 既存のタグを更新
            existing_tag_info = all_tags_flat[tag_en.lower()]
            # 変更があった場合のみ更新カウント (比較時もstripを適用)
            if existing_tag_info['ja'].strip() != tag_ja: # 比較時もstripを適用
                existing_tag_info['ja'] = tag_ja
                updated_count += 1
        else:
            # 新規タグを追加
            all_tags_flat[tag_en.lower()] = {
                'en': tag_en,
                'ja': tag_ja,
                'category_id': category_id
            }
            added_count += 1

    # 新しい辞書構造を再構築
    # カテゴリ構造は既存のものを維持し、タグを再割り当て
    new_dictionary_structure = {"categories": []}
    for category in app_state['dictionary'].get('categories', []):
        new_dictionary_structure['categories'].append({
            "id": category['id'],
            "name": category['name'],
            "parent_id": category.get('parent_id'),
            "tags": [] # ここは空にして、後でフラットなリストから追加
        })

    # フラットなタグリストから、正しいカテゴリにタグを割り当てる
    for tag_key, tag_info in all_tags_flat.items():
        target_category_id = tag_info['category_id']
        target_category_in_new_dict = find_category_by_id(target_category_id, new_dictionary_structure['categories'])
        if target_category_in_new_dict:
            target_category_in_new_dict['tags'].append({"en": tag_info['en'], "ja": tag_info['ja']})
        else:
            # これが発生する場合、カテゴリが削除されたか、無効なカテゴリIDがタグに割り当てられた
            # このケースは本来発生しないはずだが、念のため警告
            messagebox.showwarning("警告", f"タグ '{tag_info['en']}' のカテゴリID '{target_category_id}' が見つかりませんでした。このタグは保存されません。")


    app_state['dictionary'] = new_dictionary_structure
    save_dictionary()
    messagebox.showinfo("情報", f"タグの変更を保存しました。更新されたタグ: {updated_count}件, 新規追加タグ: {added_count}件")
    populate_dict_treeview(dict_search_entry.get(), dict_filter_var.get()) # フィルタを維持して更新
    update_category_dropdowns()
    update_available_tags_treeview() # タグセット生成タブも更新
    populate_category_hierarchy_treeview(category_hierarchy_tree_manage) # 辞書管理タブの階層Treeviewを更新
    populate_category_hierarchy_treeview(category_hierarchy_tree_classify) # 分類タブの階層Treeviewも更新

def upload_dictionary_file():
    """辞書JSONファイルをアップロードする"""
    filepath = filedialog.askopenfilename(title="辞書JSONファイルを選択", filetypes=[("JSONファイル", "*.json")])
    if filepath:
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                app_state['dictionary'] = json.load(f)
            # 読み込み時に全てのタグの日本語説明をstripする
            for category in app_state['dictionary'].get('categories', []):
                for tag in category.get('tags', []):
                    if 'ja' in tag and tag['ja'] is not None:
                        tag['ja'] = tag['ja'].strip()
            save_dictionary()
            messagebox.showinfo("情報", "辞書ファイルを読み込みました。")
            populate_dict_treeview()
            update_category_dropdowns()
            update_available_tags_treeview() # タグセット生成タブも更新
            populate_category_hierarchy_treeview(category_hierarchy_tree_manage) # 辞書管理タブの階層Treeviewを更新
            populate_category_hierarchy_treeview(category_hierarchy_tree_classify) # 分類タブの階層Treeviewも更新
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの読み込み中にエラーが発生しました: {e}")

def download_dictionary_file():
    """現在の辞書データをJSONでダウンロードする"""
    if not app_state['dictionary'].get('categories'):
        messagebox.showwarning("警告", "ダウンロードする辞書データがありません。")
        return
    
    filepath = filedialog.asksaveasfilename(
        defaultextension=".json",
        filetypes=[("JSONファイル", "*.json")],
        title="辞書データを保存"
    )
    if filepath:
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(app_state['dictionary'], f, ensure_ascii=False, indent=4)
            messagebox.showinfo("情報", "辞書データをダウンロードしました。")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの保存中にエラーが発生しました: {e}")

def export_tags_without_ja():
    """日本語説明がないタグをエクスポートする"""
    tags_without_ja = [
        tag['en']
        for category in app_state['dictionary'].get('categories', [])
        for tag in category.get('tags', [])
        if not tag.get('ja') or tag.get('ja') == '説明なし'
    ]

    if not tags_without_ja:
        messagebox.showinfo("情報", "日本語説明がないタグはありません。")
        return

    filepath = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("テキストファイル", "*.txt")],
        title="日本語説明がないタグをエクスポート"
    )
    if filepath:
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write("\n".join(tags_without_ja))
            messagebox.showinfo("情報", f"日本語説明がないタグ ({len(tags_without_ja)}件) をエクスポートしました。")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの保存中にエラーが発生しました: {e}")

def import_translated_tags():
    """
    翻訳済みタグをインポートする関数。
    CSVファイルから読み込んだ日本語説明で、辞書内の既存の日本語説明を上書きします。
    """
    filepath = filedialog.askopenfilename(title="翻訳済みタグリストCSVファイルを選択", filetypes=[("CSVファイル", "*.csv")])
    if not filepath:
        return

    try:
        translated_df = pd.read_csv(filepath)
        if 'English Tag' not in translated_df.columns or '日本語説明' not in translated_df.columns:
            messagebox.showerror("エラー", "CSVファイルには 'English Tag' と '日本語説明' の列が必要です。")
            return
        
        update_count = 0
        not_found_count = 0

        # 現在の辞書のタグを効率的にルックアップできるように、英語タグ名をキーとする辞書を作成
        current_tags_map = {}
        for category in app_state['dictionary'].get('categories', []):
            for tag in category.get('tags', []):
                current_tags_map[tag['en'].lower()] = tag # 小文字化した英語タグ名をキーに

        for index, row in translated_df.iterrows():
            english_tag = str(row['English Tag']).strip()
            japanese_description = str(row['日本語説明']).strip() # stripを適用

            if english_tag.lower() in current_tags_map:
                # 既存のタグが見つかった場合、日本語説明を上書き (stripを適用)
                existing_tag_obj = current_tags_map[english_tag.lower()]
                if existing_tag_obj['ja'].strip() != japanese_description: # 比較時もstripを適用
                    existing_tag_obj['ja'] = japanese_description
                    update_count += 1
            else:
                not_found_count += 1
                print(f"辞書にタグ '{english_tag}' が見つかりませんでした。このタグの日本語説明は更新されません。")
        
        save_dictionary()
        messagebox.showinfo("情報", 
                            f"翻訳済みタグのインポートが完了しました。\n"
                            f"更新: {update_count}件\n"
                            f"辞書に見つからなかったタグ: {not_found_count}件")
        populate_dict_treeview(dict_search_entry.get(), dict_filter_var.get()) # フィルタを維持して更新
        update_category_dropdowns()
        update_available_tags_treeview() # タグセット生成タブも更新
        populate_category_hierarchy_treeview(category_hierarchy_tree_manage) # 辞書管理タブの階層Treeviewを更新
        populate_category_hierarchy_treeview(category_hierarchy_tree_classify) # 分類タブの階層Treeviewも更新

    except Exception as e:
        messagebox.showerror("エラー", f"ファイルの読み込みまたは処理中にエラーが発生しました: {e}")

def import_additional_dictionary_json():
    """追加辞書JSONファイルをインポートし、現在の辞書とマージする関数"""
    filepath = filedialog.askopenfilename(title="追加辞書JSONファイルを選択", filetypes=[("JSONファイル", "*.json")])
    if not filepath:
        return

    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            imported_data = json.load(f)
        
        if 'categories' not in imported_data:
            messagebox.showerror("エラー", "インポートするJSONファイルは 'categories' キーを持つ必要があります。")
            return

        added_categories_count = 0
        added_tags_count = 0
        updated_tags_count = 0

        # --- 現在の辞書の状態を効率的にルックアップできるように準備 ---
        # (カテゴリ名.lower(), parent_id) -> カテゴリオブジェクト のマップ
        current_category_name_parent_map = {} 
        # 英語タグ名(小文字) -> タグオブジェクト のマップ
        current_tag_en_lower_to_obj = {} 

        # Populate current maps
        for cat in app_state['dictionary']['categories']:
            parent_id = cat.get('parent_id')
            current_category_name_parent_map[(cat['name'].lower(), parent_id)] = cat
            for tag in cat.get('tags', []):
                current_tag_en_lower_to_obj[tag['en'].lower()] = tag

        # --- インポートされる辞書内のカテゴリIDを、マージ後の辞書のIDにマッピングするための準備 ---
        # imported_cat_id -> current_cat_id (または新しく生成されたID)
        imported_id_to_final_id = {} 
        # imported_cat_id -> imported_cat_object
        imported_categories_by_id = {c['id']: c for c in imported_data['categories']}
        
        # 処理済みのインポートカテゴリIDを追跡
        processed_imported_category_ids = set()
        # 未処理のインポートカテゴリリストを初期化
        unprocessed_imported_categories = list(imported_data['categories'])

        # --- カテゴリのマージ (親カテゴリが先に処理されるように反復処理) ---
        # 最大ループ回数を設定し、無限ループを防ぐ
        max_iterations = len(imported_data['categories']) * 2 # Safety break
        iteration_count = 0

        while len(processed_imported_category_ids) < len(imported_data['categories']) and iteration_count < max_iterations:
            iteration_count += 1
            categories_processed_in_this_iteration = []

            # 未処理のカテゴリリストのコピーをイテレート
            for imported_cat in list(unprocessed_imported_categories):
                if imported_cat['id'] in processed_imported_category_ids:
                    continue # 既に処理済み

                imported_parent_id = imported_cat.get('parent_id')
                imported_cat_name = imported_cat['name']

                resolved_parent_id_in_main_dict = None
                can_process_category = False

                if imported_parent_id is None: # Top-level category in imported file
                    can_process_category = True
                    resolved_parent_id_in_main_dict = None
                elif imported_parent_id == "general": # Special "general" parent
                    # Try to find "general" category in current dictionary
                    general_cat_obj = current_category_name_parent_map.get(("general", None))
                    if general_cat_obj:
                        resolved_parent_id_in_main_dict = general_cat_obj['id']
                        can_process_category = True
                    else:
                        # If "general" doesn't exist, treat this imported category as a new top-level
                        resolved_parent_id_in_main_dict = None
                        can_process_category = True
                elif imported_parent_id in imported_id_to_final_id: # Parent was processed in this import session
                    resolved_parent_id_in_main_dict = imported_id_to_final_id[imported_parent_id]
                    can_process_category = True
                else:
                    # Check if parent exists in the main dictionary already
                    existing_parent_in_main_dict = find_category_by_id(imported_parent_id, app_state['dictionary']['categories'])
                    if existing_parent_in_main_dict:
                        resolved_parent_id_in_main_dict = existing_parent_in_main_dict['id']
                        can_process_category = True
                    else:
                        # Parent not found in main dictionary or in current import session.
                        # Treat this imported category as a new top-level category and issue a warning.
                        messagebox.showwarning("警告", f"インポートされたカテゴリ '{imported_cat_name}' の親カテゴリID '{imported_parent_id}' が見つかりませんでした。このカテゴリはトップレベルカテゴリとして追加されます。")
                        resolved_parent_id_in_main_dict = None
                        can_process_category = True # Can process now, as it's top-level

                if can_process_category:
                    # Check if this category (by name and resolved parent) already exists in the main dictionary
                    existing_cat_obj = current_category_name_parent_map.get((imported_cat_name.lower(), resolved_parent_id_in_main_dict))

                    if existing_cat_obj:
                        # Category already exists, reuse it
                        final_cat_id = existing_cat_obj['id']
                        imported_id_to_final_id[imported_cat['id']] = final_cat_id
                    else:
                        # New category, add it
                        final_cat_id = str(uuid.uuid4())
                        new_category_obj = {
                            "id": final_cat_id,
                            "name": imported_cat_name,
                            "parent_id": resolved_parent_id_in_main_dict,
                            "tags": [] # Tags will be added later
                        }
                        app_state['dictionary']['categories'].append(new_category_obj)
                        current_category_name_parent_map[(imported_cat_name.lower(), resolved_parent_id_in_main_dict)] = new_category_obj
                        imported_id_to_final_id[imported_cat['id']] = final_cat_id
                        added_categories_count += 1
                    
                    categories_processed_in_this_iteration.append(imported_cat)
                    processed_imported_category_ids.add(imported_cat['id']) # ここで処理済みとしてマーク

            # このイテレーションで処理されたカテゴリを未処理リストから削除
            unprocessed_imported_categories = [cat for cat in unprocessed_imported_categories if cat['id'] not in processed_imported_category_ids]

            if not categories_processed_in_this_iteration and unprocessed_imported_categories:
                # このイテレーションで何も処理されなかったが、まだ未処理のカテゴリが残っている場合
                # これは解決できない依存関係（例：循環参照）を示唆している可能性が高い
                messagebox.showwarning("警告", "インポートするカテゴリ階層に解決できない依存関係があるか、親カテゴリが見つかりません。一部のカテゴリがインポートされない可能性があります。")
                break # 無限ループを防ぐためにループを抜ける

        # --- タグのマージ (全てのカテゴリがマッピングされた後) ---
        for imported_cat in imported_data['categories']:
            # Get the corresponding category object in the main dictionary
            target_current_cat_id = imported_id_to_final_id.get(imported_cat['id'])
            if not target_current_cat_id:
                # This imported category was not successfully processed, skip its tags
                continue 
            
            target_category_obj = None
            # Find the actual category object in the main dictionary using its final ID
            for cat in app_state['dictionary']['categories']:
                if cat['id'] == target_current_cat_id:
                    target_category_obj = cat
                    break
            
            if not target_category_obj:
                # Should not happen if target_current_cat_id exists in map
                continue

            for imported_tag in imported_cat.get('tags', []):
                imported_tag_en_lower = imported_tag['en'].lower()
                imported_tag_ja = imported_tag.get('ja', '').strip() # stripを適用

                if imported_tag_en_lower in current_tag_en_lower_to_obj:
                    # 既存のタグが見つかった場合、日本語説明を更新 (比較時もstripを適用)
                    existing_tag_obj = current_tag_en_lower_to_obj[imported_tag_en_lower]
                    if existing_tag_obj['ja'].strip() != imported_tag_ja: # 比較時もstripを適用
                        existing_tag_obj['ja'] = imported_tag_ja
                        updated_tags_count += 1
                else:
                    # 新規タグとして追加 (stripを適用)
                    new_tag = {"en": imported_tag['en'], "ja": imported_tag_ja}
                    target_category_obj.setdefault('tags', []).append(new_tag)
                    current_tag_en_lower_to_obj[imported_tag_en_lower] = new_tag # Add to map for future checks
                    added_tags_count += 1

        save_dictionary()
        messagebox.showinfo("インポート完了", 
                            f"追加辞書JSONのインポートが完了しました。\n"
                            f"追加されたカテゴリ: {added_categories_count}件\n"
                            f"追加されたタグ: {added_tags_count}件\n"
                            f"更新されたタグ: {updated_tags_count}件")
        
        # UIを更新
        update_category_dropdowns()
        populate_dict_treeview(dict_search_entry.get(), dict_filter_var.get())
        update_available_tags_treeview()
        populate_category_hierarchy_treeview(category_hierarchy_tree_manage)
        populate_category_hierarchy_treeview(category_hierarchy_tree_classify)

    except json.JSONDecodeError:
        messagebox.showerror("エラー", "選択されたファイルは有効なJSON形式ではありません。")
    except Exception as e:
        messagebox.showerror("エラー", f"ファイルの読み込みまたは処理中にエラーが発生しました: {e}")


def create_classify_tags_tab(notebook_frame):
    """タグ分類作業タブのUIを構築する関数"""
    tab_frame = ttk.Frame(notebook_frame, padding="10")

    # PanedWindowで左右に分割
    paned_window = ttk.PanedWindow(tab_frame, orient=tk.HORIZONTAL)
    paned_window.pack(fill=tk.BOTH, expand=True)

    # 左側のフレーム (カテゴリ階層の閲覧とカテゴリ追加)
    left_frame = ttk.Frame(paned_window, padding="10")
    paned_window.add(left_frame, weight=1) # 左側を伸縮可能に

    category_view_frame = ttk.LabelFrame(left_frame, text="カテゴリ階層の閲覧", padding="10")
    category_view_frame.pack(fill=tk.BOTH, expand=True)

    global category_hierarchy_tree_classify # 分類タブ用のTreeview
    hierarchy_scrollbar_classify = ttk.Scrollbar(category_view_frame, orient="vertical")
    hierarchy_scrollbar_classify.pack(side="right", fill="y")
    category_hierarchy_tree_classify = ttk.Treeview(category_view_frame, show="tree", selectmode="browse", yscrollcommand=hierarchy_scrollbar_classify.set)
    category_hierarchy_tree_classify.pack(side="left", fill=tk.BOTH, expand=True)
    hierarchy_scrollbar_classify.config(command=category_hierarchy_tree_classify.yview)

    # カテゴリ追加機能 (分類タブ内)
    add_category_frame_classify = ttk.LabelFrame(left_frame, text="新しいカテゴリの追加", padding="10")
    add_category_frame_classify.pack(fill=tk.X, pady=5)

    inner_add_category_frame_classify = tk.Frame(add_category_frame_classify)
    inner_add_category_frame_classify.pack(fill=tk.BOTH, expand=True)

    ttk.Label(inner_add_category_frame_classify, text="新しいカテゴリ名:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    global new_category_name_entry_classify
    new_category_name_entry_classify = ttk.Entry(inner_add_category_frame_classify)
    new_category_name_entry_classify.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

    ttk.Label(inner_add_category_frame_classify, text="親カテゴリ:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
    global new_category_parent_combobox_classify
    new_category_parent_combobox_classify = ttk.Combobox(inner_add_category_frame_classify, state="readonly")
    new_category_parent_combobox_classify.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
    new_category_parent_combobox_classify.set("--カテゴリを選択--")

    ttk.Button(inner_add_category_frame_classify, text="カテゴリを追加", command=lambda: add_new_category(new_category_name_entry_classify, new_category_parent_combobox_classify, notebook.index(tab_frame))).grid(row=2, column=0, columnspan=2, pady=10)

    # 右クリックメニューのバインド (分類タブ用)
    # ここで tab_frame を明示的にキャプチャ
    category_hierarchy_tree_classify.bind("<Button-3>", 
                                          lambda e, tf=tab_frame: show_category_tree_context_menu(e, category_hierarchy_tree_classify, 
                                                                                  new_category_name_entry_classify, new_category_parent_combobox_classify, 
                                                                                  notebook, tf))
    
    # 初期ロードはmain関数で行う


    # 右側のフレーム (未分類タグの読み込みと分類作業)
    right_frame = ttk.Frame(paned_window, padding="10")
    paned_window.add(right_frame, weight=2) # 右側をより広く伸縮可能に

    load_unclassified_frame = ttk.LabelFrame(right_frame, text="未分類タグの読み込み", padding="10")
    load_unclassified_frame.pack(fill=tk.X, pady=5)
    ttk.Button(load_unclassified_frame, text="未分類タグリストをファイルから読み込む", command=load_unclassified_tags_from_file_classify_tab).pack(pady=5)
    
    # タグ文字列直接入力欄
    ttk.Label(load_unclassified_frame, text="または、タグ文字列を直接貼り付け:").pack(pady=5)
    global unclassified_paste_text_area
    unclassified_paste_text_area = tk.Text(load_unclassified_frame, height=5, wrap="word")
    unclassified_paste_text_area.pack(fill=tk.X, pady=5)
    ttk.Button(load_unclassified_frame, text="貼り付けたタグを読み込む", command=load_unclassified_tags_from_paste).pack(pady=5)

    global unclassified_status_label
    unclassified_status_label = ttk.Label(load_unclassified_frame, text="未分類タグリストを読み込んでください。")
    unclassified_status_label.pack(pady=5)

    classify_tags_frame = ttk.LabelFrame(right_frame, text="未分類タグの分類", padding="10")
    classify_tags_frame.pack(fill=tk.BOTH, expand=True, pady=5)

    columns = ("英語タグ名", "日本語説明", "カテゴリ")
    global unclassified_tree
    unclassified_tree = ttk.Treeview(classify_tags_frame, columns=columns, show="headings", selectmode="extended")

    for col_name in columns:
        unclassified_tree.heading(col_name, text=col_name, command=lambda _col_name=col_name: treeview_sort_column(unclassified_tree, _col_name, sort_reverse_flags.get((str(unclassified_tree), _col_name), False)))
        unclassified_tree.column(col_name, width=200, anchor="w")

    scrollbar = ttk.Scrollbar(classify_tags_frame, orient="vertical", command=unclassified_tree.yview)
    unclassified_tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    unclassified_tree.pack(fill=tk.BOTH, expand=True)

    def on_unclassified_tree_double_click(event):
        item_id = unclassified_tree.focus()
        if not item_id: return
        column_id = unclassified_tree.identify_column(event.x)
        column_index = int(column_id[1:]) - 1

        if column_index == 0: return

        row_index = int(item_id)
        current_value = unclassified_tree.item(item_id, 'values')[column_index]

        x, y, width, height = unclassified_tree.bbox(item_id, column_id)
        
        if column_index == 2:
            editor = ttk.Combobox(unclassified_tree, values=all_category_options, state="readonly")
            editor.set(current_value)
            def on_combobox_select(event):
                new_value = editor.get()
                app_state['unclassified_df'].loc[row_index, columns[column_index]] = new_value
                update_treeview(unclassified_tree, app_state['unclassified_df']) # DataFrame全体を更新してTreeviewを再描画
                editor.destroy()
                # コンボボックス選択後にフォーカスをTreeviewに戻す
                unclassified_tree.focus_set()
            editor.bind("<<ComboboxSelected>>", on_combobox_select)
            editor.bind("<FocusOut>", lambda e: editor.destroy())
        else:
            editor = ttk.Entry(unclassified_tree)
            editor.insert(0, current_value)
            def on_entry_return(event):
                new_value = editor.get().strip() # stripを適用
                app_state['unclassified_df'].loc[row_index, columns[column_index]] = new_value
                update_treeview(unclassified_tree, app_state['unclassified_df']) # DataFrame全体を更新してTreeviewを再描画
                editor.destroy()
                # エンターキー押下後にフォーカスをTreeviewに戻す
                unclassified_tree.focus_set()
            editor.bind("<Return>", on_entry_return)
            editor.bind("<FocusOut>", lambda e: editor.destroy())
        
        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()

    unclassified_tree.bind("<Double-1>", on_unclassified_tree_double_click)

    bulk_apply_unclassified_frame = ttk.Frame(classify_tags_frame, padding="10")
    bulk_apply_unclassified_frame.pack(fill=tk.X, pady=5)

    ttk.Label(bulk_apply_unclassified_frame, text="選択したタグにまとめてカテゴリを適用:").pack(side="left", padx=5)
    global unclassified_category_var, unclassified_category_combobox
    unclassified_category_var = tk.StringVar(root)
    unclassified_category_combobox = ttk.Combobox(bulk_apply_unclassified_frame, textvariable=unclassified_category_var, state="readonly", values=all_category_options)
    unclassified_category_combobox.set(all_category_options[0])
    unclassified_category_combobox.pack(side="left", padx=5, expand=True, fill=tk.X)
    ttk.Button(bulk_apply_unclassified_frame, text="選択したタグに適用", command=apply_selected_category_unclassified_tab).pack(side="left", padx=5)

    process_buttons_frame = ttk.Frame(classify_tags_frame, padding="10")
    process_buttons_frame.pack(fill=tk.X, pady=5)
    ttk.Button(process_buttons_frame, text="テーブルで分類したタグを辞書に追加", command=add_classified_tags_to_dictionary_unclassified_tab).pack(side="left", padx=5)
    ttk.Button(process_buttons_frame, text="未分類タグリストをクリア", command=clear_unclassified_tags_classify_tab).pack(side="right", padx=5)

    return tab_frame

def process_unclassified_tags(tags_list_cleaned):
    """未分類タグリストを処理し、DataFrameを更新する共通関数"""
    newly_unclassified = []
    # 辞書内のすべてのタグを効率的にルックアップできるように、英語タグ名をキーとする辞書を作成
    all_dict_tags_en_map = {t['en'].lower(): t for cat in app_state['dictionary'].get('categories', []) for t in cat.get('tags', [])}

    for tag in tags_list_cleaned:
        if tag.lower() in all_dict_tags_en_map:
            # 既存のタグが見つかった場合、日本語説明を更新
            existing_tag_obj = all_dict_tags_en_map[tag.lower()]
            # ここでは既存のタグの日本語説明を更新するロジックは含めない（分類タブの役割ではないため）
            print(f"タグ '{tag}' は既に辞書に存在します。スキップします。")
        else:
            newly_unclassified.append(tag)

    unclassified_tags_data = []
    for tag_en in newly_unclassified:
        hints = get_classification_hint(tag_en)
        suggested_ja = ""
        suggested_cat_path = "--カテゴリを選択--"

        if hints:
            top_hint = hints[0]
            # stripを適用
            suggested_ja = (top_hint.get('tag_ja', '') if top_hint.get('tag_ja') != '説明なし' else '').strip()
            suggested_cat_path = top_hint.get('category_path', '')

        unclassified_tags_data.append({
            "英語タグ名": tag_en,
            "日本語説明": suggested_ja,
            "カテゴリ": suggested_cat_path
        })
    
    app_state['unclassified_df'] = pd.DataFrame(unclassified_tags_data, columns=["英語タグ名", "日本語説明", "カテゴリ"])
    app_state['unclassified_df'] = app_state['unclassified_df'].fillna('')
    update_treeview(unclassified_tree, app_state['unclassified_df'])
    unclassified_status_label.config(text=f"未分類タグ ({len(app_state['unclassified_df'])}件):")
    messagebox.showinfo("情報", f"{len(app_state['unclassified_df'])} 個の新しい未分類タグを読み込みました。")


def load_unclassified_tags_from_file_classify_tab():
    """未分類タグリストファイルを読み込む（分類タブ用）"""
    filepath = filedialog.askopenfilename(
        title="未分類タグリストファイルを選択",
        filetypes=[("テキストファイル", "*.txt"), ("CSVファイル", "*.csv")]
    )
    if not filepath:
        return

    try:
        tags_list_cleaned = []
        if filepath.endswith('.csv'):
            df_uploaded = pd.read_csv(filepath)
            if not df_uploaded.empty:
                tags_list_cleaned = df_uploaded.iloc[:, 0].astype(str).tolist()
            else:
                messagebox.showwarning("警告", "アップロードされたCSVファイルは空です。")
                return
        else:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
            tags_list_raw = re.split(r'[,\n]+', content)
            tags_list_cleaned = [tag.strip() for tag in tags_list_raw if tag.strip()]
        
        process_unclassified_tags(tags_list_cleaned)

    except Exception as e:
        messagebox.showerror("エラー", f"ファイルの読み込み中にエラーが発生しました: {e}")

def load_unclassified_tags_from_paste():
    """未分類タグ文字列を直接貼り付けて読み込む（分類タブ用）"""
    if unclassified_paste_text_area is None: return
    content = unclassified_paste_text_area.get(1.0, tk.END).strip()
    if not content:
        messagebox.showwarning("警告", "貼り付けるタグ文字列がありません。")
        return
    
    tags_list_raw = re.split(r'[,\n]+', content)
    tags_list_cleaned = [tag.strip() for tag in tags_list_raw if tag.strip()]

    if not tags_list_cleaned:
        messagebox.showwarning("警告", "有効なタグが見つかりませんでした。")
        return

    process_unclassified_tags(tags_list_cleaned)
    unclassified_paste_text_area.delete(1.0, tk.END) # 読み込み後、テキストエリアをクリア

def apply_selected_category_unclassified_tab():
    """未分類タグタブで選択したタグにカテゴリを一括適用する"""
    selected_category_path = unclassified_category_var.get()
    if selected_category_path == "--カテゴリを選択--":
        messagebox.showwarning("警告", "適用するカテゴリを選択してください。")
        return

    selected_items = unclassified_tree.selection()
    if not selected_items:
        messagebox.showwarning("警告", "カテゴリを適用するタグを選択してください。")
        return

    selected_indices = [int(item) for item in selected_items]

    app_state['unclassified_df'].loc[selected_indices, 'カテゴリ'] = selected_category_path
    
    update_treeview(unclassified_tree, app_state['unclassified_df'])
    
    messagebox.showinfo("情報", f"{len(selected_indices)}件のタグにカテゴリ '{selected_category_path}' を適用しました。")

def add_classified_tags_to_dictionary_unclassified_tab():
    """未分類タグタブで分類したタグを辞書に追加する"""
    added_count = 0
    updated_count = 0 # 更新されたタグのカウントを追加
    unclassified_after_add = []
    
    # 辞書内のすべてのタグを効率的にルックアップできるように、英語タグ名をキーとする辞書を作成
    all_dict_tags_en_map = {}
    for category in app_state['dictionary'].get('categories', []):
        for tag in category.get('tags', []):
            all_dict_tags_en_map[tag['en'].lower()] = tag

    for index, row in app_state['unclassified_df'].iterrows():
        tag_en = row["英語タグ名"]
        tag_ja = row["日本語説明"].strip() # stripを適用
        category_path = row["カテゴリ"]

        if category_path and category_path != "--カテゴリを選択--":
            category_id = get_category_id_from_path(category_path)
            if category_id is not None:
                if tag_en.lower() in all_dict_tags_en_map:
                    # 既存のタグが見つかった場合、日本語説明を更新
                    existing_tag_obj = all_dict_tags_en_map[tag_en.lower()]
                    # カテゴリも更新できるように修正 (ただし、カテゴリ移動は慎重に)
                    # ここでは、同じ英語タグ名であれば日本語説明を更新するのみとする
                    if existing_tag_obj['ja'].strip() != tag_ja: # 比較時もstripを適用
                        existing_tag_obj['ja'] = tag_ja
                        updated_count += 1
                else:
                    # 新規タグとして追加
                    success, message = add_tag_to_dictionary(tag_en, tag_ja, category_id)
                    if success:
                        added_count += 1
                        # 新しく追加されたタグもマップに反映
                        target_category = find_category_by_id(category_id)
                        if target_category:
                            new_tag_obj = next((t for t in target_category['tags'] if t['en'].lower() == tag_en.lower()), None)
                            if new_tag_obj:
                                all_dict_tags_en_map[tag_en.lower()] = new_tag_obj
                    else:
                        messagebox.showwarning("警告", f"タグ '{tag_en}' の追加に失敗しました: {message}")
                        unclassified_after_add.append(tag_en)
            else:
                messagebox.showwarning("警告", f"タグ '{tag_en}': 無効なカテゴリパス '{category_path}' です。スキップしました。")
                unclassified_after_add.append(tag_en)
        else:
            unclassified_after_add.append(tag_en)

    # 未分類のまま残ったタグを再処理
    unclassified_tags_data = []
    for tag_en in unclassified_after_add:
        hints = get_classification_hint(tag_en)
        suggested_ja = ""
        suggested_cat_path = "--カテゴリを選択--"
        if hints:
            top_hint = hints[0]
            suggested_ja = (top_hint.get('tag_ja', '') if top_hint.get('tag_ja') != '説明なし' else '').strip() # stripを適用
            suggested_cat_path = top_hint.get('category_path', '')
        unclassified_tags_data.append({
            "英語タグ名": tag_en,
            "日本語説明": suggested_ja,
            "カテゴリ": suggested_cat_path
        })
    app_state['unclassified_df'] = pd.DataFrame(unclassified_tags_data, columns=["英語タグ名", "日本語説明", "カテゴリ"])
    app_state['unclassified_df'] = app_state['unclassified_df'].fillna('')
    update_treeview(unclassified_tree, app_state['unclassified_df'])
    save_dictionary()
    update_category_dropdowns()
    populate_dict_treeview(dict_search_entry.get(), dict_filter_var.get()) # フィルタを維持して更新
    update_available_tags_treeview() # タグセット生成タブも更新
    populate_category_hierarchy_treeview(category_hierarchy_tree_manage) # 辞書管理タブの階層Treeviewを更新
    populate_category_hierarchy_treeview(category_hierarchy_tree_classify) # 分類タブの階層Treeviewも更新
    unclassified_status_label.config(text=f"未分類タグ ({len(app_state['unclassified_df'])}件):")
    messagebox.showinfo("情報", f"{added_count}件のタグを辞書に追加し、{updated_count}件のタグを更新しました。辞書ファイルも更新されました。")


def clear_unclassified_tags_classify_tab():
    """未分類タグリストをクリアする（分類タブ用）"""
    app_state['unclassified_df'] = pd.DataFrame(columns=["英語タグ名", "日本語説明", "カテゴリ"])
    update_treeview(unclassified_tree, app_state['unclassified_df'])
    unclassified_status_label.config(text="未分類タグリストを読み込んでください。")
    messagebox.showinfo("情報", "未分類タグリストをクリアしました。")

# 新しく追加する関数
def copy_generated_text():
    """生成されたタグテキストをクリップボードにコピーする関数"""
    if generated_text_area is None: return
    text_to_copy = generated_text_area.get(1.0, tk.END).strip()
    if text_to_copy:
        root.clipboard_clear()
        root.clipboard_append(text_to_copy)
        messagebox.showinfo("情報", "生成されたタグテキストをクリップボードにコピーしました！")
    else:
        messagebox.showwarning("警告", "コピーするテキストがありません。")


def create_generate_tags_tab(notebook_frame):
    """タグセット生成タブのUIを構築する関数"""
    tab_frame = ttk.Frame(notebook_frame, padding="10")

    # PanedWindowで左右に分割
    paned_window = ttk.PanedWindow(tab_frame, orient=tk.HORIZONTAL)
    paned_window.pack(fill=tk.BOTH, expand=True)

    # 左側のフレーム (カテゴリ階層の閲覧)
    left_gen_frame = ttk.Frame(paned_window, padding="10")
    paned_window.add(left_gen_frame, weight=1)

    category_view_gen_frame = ttk.LabelFrame(left_gen_frame, text="カテゴリ選択", padding="10")
    category_view_gen_frame.pack(fill=tk.BOTH, expand=True)

    global available_categories_tree
    gen_cat_scrollbar = ttk.Scrollbar(category_view_gen_frame, orient="vertical")
    gen_cat_scrollbar.pack(side="right", fill="y")
    available_categories_tree = ttk.Treeview(category_view_gen_frame, show="tree", selectmode="browse", yscrollcommand=gen_cat_scrollbar.set)
    available_categories_tree.pack(side="left", fill=tk.BOTH, expand=True)
    gen_cat_scrollbar.config(command=available_categories_tree.yview)

    # Bind selection event
    available_categories_tree.bind("<<TreeviewSelect>>", on_available_category_select)
    # Search entry for category tree (optional, but consistent with other trees)
    global tag_gen_search_entry # Re-use the existing global search entry
    tag_gen_search_entry_label = ttk.Label(category_view_gen_frame, text="カテゴリ検索:")
    tag_gen_search_entry_label.pack(pady=5)
    tag_gen_search_entry = ttk.Entry(category_view_gen_frame)
    tag_gen_search_entry.bind("<KeyRelease>", lambda e: populate_available_categories_treeview()) # Update category tree on search
    tag_gen_search_entry.pack(fill=tk.X, padx=5, pady=5)


    # 右側のフレーム (タグリストと検索/フィルタ)
    right_gen_frame = ttk.Frame(paned_window, padding="10")
    paned_window.add(right_gen_frame, weight=2)

    available_tags_frame = ttk.LabelFrame(right_gen_frame, text="利用可能タグ (選択カテゴリ内)", padding="10")
    available_tags_frame.pack(fill=tk.BOTH, expand=True, pady=5)

    filter_search_frame = ttk.Frame(available_tags_frame)
    filter_search_frame.pack(fill=tk.X, pady=5)
    
    # カテゴリフィルタは左のツリーで代替されるため、ここでは削除または非表示にする
    #ttk.Label(filter_search_frame, text="カテゴリで絞り込み:").pack(side="left", padx=5)
    #global tag_gen_filter_var, tag_gen_filter_combobox
    #tag_gen_filter_var = tk.StringVar(root)
    #tag_gen_filter_combobox = ttk.Combobox(filter_search_frame, textvariable=tag_gen_filter_var, state="readonly", values=["--全てのカテゴリ--"] + all_category_options[1:])
    #tag_gen_filter_combobox.set("--全てのカテゴリ--")
    #tag_gen_filter_combobox.bind("<<ComboboxSelected>>", lambda e: update_available_tags_treeview())
    #tag_gen_filter_combobox.pack(side="left", padx=5, expand=True, fill=tk.X)

    ttk.Label(filter_search_frame, text="タグ名または説明で検索:").pack(side="left", padx=5)
    # tag_gen_search_entry は左のカテゴリツリー検索に再利用されるため、ここでは別のEntryを使うか、タグリストの検索専用にする
    # ここでは、タグリストの検索専用のEntryを新しく作成する
    global tag_list_search_entry
    tag_list_search_entry = ttk.Entry(filter_search_frame)
    tag_list_search_entry.bind("<KeyRelease>", lambda e: on_available_category_select(None)) # Re-trigger update based on current selection
    tag_list_search_entry.pack(side="left", padx=5, expand=True, fill=tk.X)


    columns = ("英語タグ名", "日本語説明", "カテゴリ") # タグの列
    global available_tags_tree # このTreeviewは右側のタグリストになる
    available_tags_tree = ttk.Treeview(available_tags_frame, columns=columns, show="headings", selectmode="browse")

    # ヘッディングの設定 (英語タグ名、日本語説明、フルパスカテゴリ)
    for col_name in columns:
        available_tags_tree.heading(col_name, text=col_name, command=lambda _col_name=col_name: treeview_sort_column(available_tags_tree, _col_name, sort_reverse_flags.get((str(available_tags_tree), _col_name), False)))
        available_tags_tree.column(col_name, width=150, anchor="w")

    scrollbar = ttk.Scrollbar(available_tags_frame, orient="vertical", command=available_tags_tree.yview)
    available_tags_tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    available_tags_tree.pack(fill=tk.BOTH, expand=True) # expandをTrueに

    ttk.Button(available_tags_frame, text="選択したタグを追加", command=add_selected_tag_to_generating_list).pack(pady=5)

    selected_tags_frame = ttk.LabelFrame(tab_frame, text="選択済みタグ", padding="10")
    selected_tags_frame.pack(fill=tk.X, pady=5) # expandをFalseに

    columns = ("英語タグ名", "日本語説明", "カテゴリ")
    global selected_generating_tree
    # Treeviewの高さ固定
    selected_generating_tree = ttk.Treeview(selected_tags_frame, columns=columns, show="headings", selectmode="browse", height=8)

    for col_name in columns:
        selected_generating_tree.heading(col_name, text=col_name, command=lambda _col_name=col_name: treeview_sort_column(selected_generating_tree, _col_name, sort_reverse_flags.get((str(selected_generating_tree), _col_name), False)))
        selected_generating_tree.column(col_name, width=150, anchor="w")

    scrollbar = ttk.Scrollbar(selected_tags_frame, orient="vertical", command=selected_generating_tree.yview)
    selected_generating_tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    selected_generating_tree.pack(fill=tk.BOTH, expand=False) # expandをFalseに

    selected_tags_buttons_frame = ttk.Frame(selected_tags_frame, padding="5")
    selected_tags_buttons_frame.pack(fill=tk.X, pady=5)
    ttk.Button(selected_tags_buttons_frame, text="削除", command=remove_selected_generating_tag).pack(side="left", padx=5)
    ttk.Button(selected_tags_buttons_frame, text="上に移動", command=move_selected_generating_tag_up).pack(side="left", padx=5)
    ttk.Button(selected_tags_buttons_frame, text="下に移動", command=move_selected_generating_tag_down).pack(side="left", padx=5)
    # 新しく追加するクリアボタン
    ttk.Button(selected_tags_buttons_frame, text="選択済みタグをクリア", command=clear_selected_generating_tags).pack(side="right", padx=5)


    generated_text_frame = ttk.LabelFrame(tab_frame, text="生成されたタグテキスト", padding="10")
    generated_text_frame.pack(fill=tk.X, pady=5)

    ttk.Label(generated_text_frame, text="区切り文字:").pack(side="left", padx=5)
    global delimiter_var
    delimiter_var = tk.StringVar(root)
    delimiter_combobox = ttk.Combobox(generated_text_frame, textvariable=delimiter_var, values=[", ", "_", "|", " "], state="readonly")
    delimiter_combobox.set(", ")
    delimiter_combobox.bind("<<ComboboxSelected>>", lambda e: update_generated_text())
    delimiter_combobox.pack(side="left", padx=5)

    global generated_text_area
    generated_text_area = tk.Text(generated_text_frame, height=5, wrap="word")
    generated_text_area.pack(fill=tk.BOTH, expand=True, pady=5)
    ttk.Button(generated_text_frame, text="クリップボードにコピー", command=copy_generated_text).pack(pady=5)

    # 初期ロードはmain関数で行う

    return tab_frame

def create_random_tag_gen_tab(notebook_frame):
    """ランダムタグセット生成タブのUIを構築する関数"""
    tab_frame = ttk.Frame(notebook_frame, padding="10")

    random_gen_frame = ttk.LabelFrame(tab_frame, text="ランダムタグセット生成", padding="10")
    random_gen_frame.pack(fill=tk.BOTH, expand=True, pady=5)

    ttk.Label(random_gen_frame, text="各最終カテゴリ（子カテゴリを持たないカテゴリ）からランダムに1つずつタグを選んでタグセットを生成します。").pack(pady=5)
    ttk.Button(random_gen_frame, text="ランダムタグセットを生成", command=generate_random_tag_set).pack(pady=5)
    
    # 生成されたタグを表示するTextウィジェットとスクロールバー
    text_area_frame = ttk.Frame(random_gen_frame)
    text_area_frame.pack(fill=tk.BOTH, expand=True, pady=5)

    global random_generated_label # LabelからTextに変更
    random_generated_label = tk.Text(text_area_frame, height=10, wrap="word") # heightを設定
    random_generated_label.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    text_scrollbar = ttk.Scrollbar(text_area_frame, orient="vertical", command=random_generated_label.yview)
    random_generated_label.config(yscrollcommand=text_scrollbar.set)
    text_scrollbar.pack(side=tk.RIGHT, fill="y")

    ttk.Button(random_gen_frame, text="このランダムタグセットを選択済みに追加", command=add_random_tags_to_selected).pack(pady=5)

    return tab_frame

def get_leaf_categories(categories_list):
    """子カテゴリを持たない最終カテゴリのリストを返す"""
    # カテゴリIDをキーとするマップを作成
    category_id_map = {cat['id']: cat for cat in categories_list}
    
    # 全てのカテゴリのparent_idをセットに収集
    all_parent_ids = {cat.get('parent_id') for cat in categories_list if cat.get('parent_id') is not None}
    
    leaf_categories = []
    for category in categories_list:
        # 自身のIDが他のカテゴリのparent_idとして存在しない場合、それはリーフカテゴリ
        if category['id'] not in all_parent_ids:
            leaf_categories.append(category)
    return leaf_categories


def check_if_category_or_descendant_matches_search(category_info, search_query_lower, all_categories_map):
    """カテゴリまたはその子孫が検索クエリに一致するか再帰的にチェックする"""
    # 現在のカテゴリ名が検索クエリに一致するか
    if search_query_lower in category_info['name'].lower():
        return True
    # 現在のカテゴリのタグが検索クエリに一致するか
    for tag in category_info.get('tags', []):
        if search_query_lower in tag.get('en', '').lower() or \
           search_query_lower in tag.get('ja', '').lower():
            return True
    # 子カテゴリを再帰的にチェック
    children_categories = [cat for cat in app_state['dictionary']['categories'] if cat.get('parent_id') == category_info['id']]
    for child_cat in children_categories:
        if check_if_category_or_descendant_matches_search(child_cat, search_query_lower, all_categories_map):
            return True
    return False


def populate_available_categories_treeview():
    """タグセット生成タブの左側カテゴリツリーにデータをロードする関数"""
    if available_categories_tree is None: return
    for item in available_categories_tree.get_children():
        available_categories_tree.delete(item)

    all_categories_map = get_all_categories_flat_map()
    search_query_lower = tag_gen_search_entry.get().lower() if tag_gen_search_entry is not None else ""

    def insert_category_node(category_info, parent_iid=""):
        # 検索クエリがある場合、このカテゴリまたは子孫が検索にヒットしない場合はスキップ
        if search_query_lower and not check_if_category_or_descendant_matches_search(category_info, search_query_lower, all_categories_map):
            return

        iid = available_categories_tree.insert(parent_iid, "end", text=f"📂 {category_info['name']}", open=False, values=(category_info['id'],))
        
        children_categories = [cat for cat in app_state['dictionary']['categories'] if cat.get('parent_id') == category_info['id']]
        for child_cat in children_categories:
            insert_category_node(child_cat, iid)

    top_level_categories = [cat for cat in app_state['dictionary'].get('categories', []) if cat.get('parent_id') is None]
    for category in top_level_categories:
        insert_category_node(category)

def populate_available_tags_list_treeview(selected_category_id=None):
    """タグセット生成タブの右側タグリストTreeviewにデータをロードする関数"""
    if available_tags_tree is None: return
    for item in available_tags_tree.get_children():
        available_tags_tree.delete(item)

    all_tags_to_display = []
    # tag_list_search_entry から検索クエリを取得
    search_query_lower = tag_list_search_entry.get().lower() if tag_list_search_entry is not None else ""

    # 選択されたカテゴリとその子孫カテゴリのタグを再帰的に取得
    def get_tags_recursively(cat_id):
        current_category = find_category_by_id(cat_id)
        if not current_category: return []
        
        tags = []
        for tag in current_category.get('tags', []):
            tags.append({
                'en': tag['en'],
                'ja': tag.get('ja', ''),
                'category_path': get_category_path(current_category['id'])
            })
        
        children_categories = [cat for cat in app_state['dictionary']['categories'] if cat.get('parent_id') == cat_id]
        for child_cat in children_categories:
            tags.extend(get_tags_recursively(child_cat['id']))
        return tags
    
    if selected_category_id:
        all_tags_to_display = get_tags_recursively(selected_category_id)
    else: # カテゴリが選択されていない場合、全てのタグを表示
        for category in app_state['dictionary'].get('categories', []):
            for tag in category.get('tags', []):
                all_tags_to_display.append({
                    'en': tag['en'],
                    'ja': tag.get('ja', ''),
                    'category_path': get_category_path(category['id'])
                })

    # 検索クエリでフィルタリング
    filtered_tags = []
    for tag_info in all_tags_to_display:
        if search_query_lower in tag_info['en'].lower() or \
           search_query_lower in tag_info['ja'].lower() or \
           search_query_lower in tag_info['category_path'].lower():
            filtered_tags.append(tag_info)

    # フィルタリングされたタグをTreeviewに挿入
    for index, tag_info in enumerate(filtered_tags):
        available_tags_tree.insert("", "end", iid=index, values=(tag_info['en'], tag_info['ja'], tag_info['category_path']))


def on_available_category_select(event):
    """タグセット生成タブの左側カテゴリツリーでカテゴリが選択されたときのイベントハンドラ"""
    selected_item_id = available_categories_tree.focus()
    if selected_item_id:
        # 選択されたアイテムがカテゴリノードであることを確認
        item_values = available_categories_tree.item(selected_item_id, 'values')
        if item_values and item_values[0] != "tag": # "tag"はタグノードの識別子
            category_id = item_values[0] # Get category ID from values
            populate_available_tags_list_treeview(category_id)
        else:
            # タグノードが選択された場合や無効な選択の場合、タグリストをクリア
            populate_available_tags_list_treeview(None) # 全てのタグを表示するか、クリアするか
    else:
        populate_available_tags_list_treeview(None) # 選択が解除された場合も全てのタグを表示


def update_available_tags_treeview():
    """タグセット生成タブのTreeviewを更新する（カテゴリツリーとタグリストの両方）"""
    populate_available_categories_treeview() # 左側のカテゴリツリーを更新
    
    # 左側のカテゴリツリーで現在選択されているカテゴリに基づいて右側のタグリストを更新
    selected_item_id = available_categories_tree.focus()
    if selected_item_id:
        item_values = available_categories_tree.item(selected_item_id, 'values')
        if item_values and item_values[0] != "tag":
            category_id = item_values[0]
            populate_available_tags_list_treeview(category_id)
        else:
            populate_available_tags_list_treeview(None) # タグノードが選択されている場合は全てのタグを表示
    else:
        populate_available_tags_list_treeview(None) # 何も選択されていない場合は全てのタグを表示


def add_selected_tag_to_generating_list():
    """選択したタグを選択済みリストに追加する (階層型Treeviewに対応)"""
    if available_tags_tree is None: return
    selected_item_id = available_tags_tree.focus()
    if not selected_item_id:
        messagebox.showwarning("警告", "追加するタグを選択してください。")
        return
    
    item_values = available_tags_tree.item(selected_item_id, 'values')
    
    # available_tags_tree はもはや階層表示ではないため、item_values[-1] != "tag_node" のチェックは不要
    # valuesの長さが3であることを確認（英語、日本語、カテゴリパス）
    if not item_values or len(item_values) < 3:
        messagebox.showwarning("警告", "タグを選択してください。（カテゴリは追加できません）")
        return

    tag_en = item_values[0]
    tag_ja = item_values[1]
    category_path = item_values[2]

    if tag_en not in [t['en'] for t in app_state['selected_generating_tags']]:
        app_state['selected_generating_tags'].append({'en': tag_en, 'ja': tag_ja, 'category_path': category_path})
        update_selected_generating_treeview()
        update_generated_text()
        messagebox.showinfo("情報", f"タグ '{tag_en}' を追加しました。")
    else:
        messagebox.showwarning("警告", "そのタグは既に追加されています。")

def update_selected_generating_treeview():
    """選択済みタグTreeviewを更新する"""
    processed_selected_tags = []
    for tag_info in app_state['selected_generating_tags']:
        processed_selected_tags.append({
            "英語タグ名": tag_info.get('en', ''),
            "日本語説明": tag_info.get('ja', ''),
            "カテゴリ": tag_info.get('category_path', '')
        })
    
    selected_df = pd.DataFrame(processed_selected_tags, columns=["英語タグ名", "日本語説明", "カテゴリ"])
    selected_df = selected_df.fillna('')
    # selected_generating_treeがNoneでないことを確認
    if selected_generating_tree is not None:
        # Treeviewのクリアはupdate_treeview内で処理される
        update_treeview(selected_generating_tree, selected_df)

def remove_selected_generating_tag():
    """選択済みタグリストからタグを削除する"""
    if selected_generating_tree is None: return
    selected_item = selected_generating_tree.focus()
    if not selected_item:
        messagebox.showwarning("警告", "削除するタグを選択してください。")
        return
    
    # Treeviewのitem_idはDataFrameのindexと一致するようにしている
    row_index_to_remove = int(selected_item)
    
    if 0 <= row_index_to_remove < len(app_state['selected_generating_tags']):
        tag_en_to_remove = app_state['selected_generating_tags'][row_index_to_remove]['en']
        del app_state['selected_generating_tags'][row_index_to_remove]
        update_selected_generating_treeview()
        update_generated_text()
        messagebox.showinfo("情報", f"タグ '{tag_en_to_remove}' を削除しました。")
    else:
        messagebox.showwarning("エラー", "選択されたタグが見つかりませんでした。")


def move_selected_generating_tag_up():
    """選択済みタグを上に移動する"""
    if selected_generating_tree is None: return
    selected_item = selected_generating_tree.focus()
    if not selected_item: return
    
    current_index = int(selected_generating_tree.index(selected_item))
    if current_index > 0:
        tag_info = app_state['selected_generating_tags'].pop(current_index)
        app_state['selected_generating_tags'].insert(current_index - 1, tag_info)
        update_selected_generating_treeview()
        # 移動後も選択状態を維持
        new_item_id = str(current_index - 1) # 新しいindexに対応するiid
        selected_generating_tree.focus(new_item_id)
        selected_generating_tree.selection_set(new_item_id)
        update_generated_text()

def move_selected_generating_tag_down():
    """選択済みタグを下に移動する"""
    if selected_generating_tree is None: return
    selected_item = selected_generating_tree.focus()
    if not selected_item: return
    
    current_index = int(selected_generating_tree.index(selected_item))
    if current_index < len(app_state['selected_generating_tags']) - 1:
        tag_info = app_state['selected_generating_tags'].pop(current_index)
        app_state['selected_generating_tags'].insert(current_index + 1, tag_info)
        update_selected_generating_treeview()
        # 移動後も選択状態を維持
        new_item_id = str(current_index + 1) # 新しいindexに対応するiid
        selected_generating_tree.focus(new_item_id)
        selected_generating_tree.selection_set(new_item_id)
        update_generated_text()

def update_generated_text():
    """生成されたタグテキストを更新する"""
    # delimiter_varとgenerated_text_areaがNoneでないことを確認
    if delimiter_var is None or generated_text_area is None: return
    delimiter = delimiter_var.get()
    generated_text = delimiter.join([tag_info['en'] for tag_info in app_state['selected_generating_tags']])
    generated_text_area.delete(1.0, tk.END)
    generated_text_area.insert(tk.END, generated_text)

def clear_selected_generating_tags():
    """選択済みタグリストをクリアする関数"""
    if not app_state['selected_generating_tags']:
        messagebox.showinfo("情報", "選択済みタグリストはすでに空です。")
        return

    if messagebox.askyesno("確認", "選択済みタグリストをすべてクリアしますか？"):
        app_state['selected_generating_tags'] = []
        update_selected_generating_treeview()
        update_generated_text()
        messagebox.showinfo("情報", "選択済みタグリストをクリアしました。")


def generate_random_tag_set():
    """ランダムタグセットを生成する（最終カテゴリから1つずつ）"""
    random_tags = []
    
    # 最終カテゴリ（子カテゴリを持たないカテゴリ）を抽出
    leaf_categories = get_leaf_categories(app_state['dictionary'].get('categories', []))
    
    # タグが全く存在しない最終カテゴリをフィルタリング
    leaf_categories_with_tags = [cat for cat in leaf_categories if cat.get('tags')]

    if not leaf_categories_with_tags:
        messagebox.showwarning("警告", "ランダムタグセットを生成できませんでした。\n辞書にタグが登録されていないか、初期辞書が生成されていません。\n「カテゴリ・辞書管理」タブでデモ用初期辞書を生成するか、辞書をアップロードしてください。")
        # random_generated_labelがNoneでないことを確認
        if random_generated_label is not None:
            random_generated_label.delete(1.0, tk.END) # Textウィジェットなのでdelete
            random_generated_label.insert(tk.END, "ランダムタグセットを生成できませんでした。辞書にタグがありません。")
        return

    for category in leaf_categories_with_tags:
        random_tag = random.choice(category['tags'])
        random_tags.append({
            'en': random_tag['en'],
            'ja': random_tag.get('ja', ''),
            'category_path': get_category_path(category['id'])
        })
    app_state['random_generated_tags'] = random_tags
    
    if random_tags:
        display_text = "生成されたランダムタグセット:\n" + "\n".join([f"- {t['en']} ({t['ja']}) [カテゴリ: {t['category_path']}]" for t in random_tags])
    else:
        # このケースは、leaf_categories_with_tagsが空でなければ通常到達しないはず
        display_text = "ランダムタグセットを生成できませんでした。辞書にタグがありません。"
    # random_generated_labelがNoneでないことを確認
    if random_generated_label is not None:
        random_generated_label.delete(1.0, tk.END) # Textウィジェットなのでdelete
        random_generated_label.insert(tk.END, display_text)
    messagebox.showinfo("情報", "ランダムタグセットを生成しました。")

def add_random_tags_to_selected():
    """ランダム生成されたタグを選択済みに追加する (既存の選択済みタグをクリアしてから追加)"""
    # 既存の選択済みタグをクリア
    app_state['selected_generating_tags'] = []
    
    added_count = 0
    for tag_info in app_state['random_generated_tags']:
        # 重複チェックは不要になるが、念のため残す
        if tag_info['en'] not in [t['en'] for t in app_state['selected_generating_tags']]:
            app_state['selected_generating_tags'].append(tag_info)
            added_count += 1
    
    app_state['random_generated_tags'] = []
    # random_generated_labelがNoneでないことを確認
    if random_generated_label is not None:
        random_generated_label.delete(1.0, tk.END) # Textウィジェットなのでdelete
    update_selected_generating_treeview()
    update_generated_text()
    messagebox.showinfo("情報", f"{added_count}件のランダムタグを選択済みタグに追加しました。")


# --- メインアプリケーションのセットアップ ---
def main():
    global root, notebook, all_category_options, all_category_path_to_id

    root = tk.Tk()
    root.title("タグ分類・生成アプリ (Tkinter)")
    root.geometry("1000x800")

    # ここでまず辞書をロードし、all_category_options を初期化する
    load_dictionary()

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # 各タブのフレームを先に作成し、notebookに追加
    # この時点で、各 create_tab 関数内でグローバル変数にウィジェットが割り当てられる
    file_management_tab = create_file_management_tab(notebook) # 新しいファイル管理タブ
    manage_dict_tab = create_manage_dictionary_tab(notebook)
    classify_tags_tab = create_classify_tags_tab(notebook)
    generate_tags_tab = create_generate_tags_tab(notebook)
    random_tag_gen_tab = create_random_tag_gen_tab(notebook)

    notebook.add(file_management_tab, text="ファイル管理") # ファイル管理タブを最初に追加
    notebook.add(manage_dict_tab, text="カテゴリ・辞書管理")
    notebook.add(classify_tags_tab, text="タグ分類作業")
    notebook.add(generate_tags_tab, text="タグセット生成")
    notebook.add(random_tag_gen_tab, text="ランダムタグ生成")

    # 全てのUI要素が作成されてから、残りのUI更新を行う
    # populate_dict_treeview() は引数を持つようになったため、初期呼び出しも引数を渡す
    populate_dict_treeview(dict_search_entry.get(), dict_filter_var.get())
    update_available_tags_treeview() # ここで階層型Treeviewを初期化
    populate_category_hierarchy_treeview(category_hierarchy_tree_manage)
    populate_category_hierarchy_treeview(category_hierarchy_tree_classify)

    # アプリケーション起動時にドロップダウンリストの値を更新
    update_category_dropdowns()

    root.mainloop()

if __name__ == "__main__":
    # グローバル変数を初期化
    category_hierarchy_tree_manage = None
    category_hierarchy_tree_classify = None
    new_category_name_entry = None
    new_category_parent_combobox = None
    new_category_name_entry_classify = None
    new_category_parent_combobox_classify = None
    notebook = None # notebookもグローバルでアクセスできるように

    # 新しく追加されたグローバルUI要素の初期化
    add_tag_english_entry = None
    add_tag_japanese_entry = None
    add_tag_category_combobox = None
    unclassified_paste_text_area = None

    # その他のグローバルUI要素もNoneで初期化しておくことで、
    # update_category_dropdowns()などの関数が安全にアクセスチェックできるようにする
    dict_tree = None
    unclassified_tree = None
    available_tags_tree = None # タグセット生成タブの右側タグリスト用
    available_categories_tree = None # タグセット生成タブの左側カテゴリツリー用
    selected_generating_tree = None
    generated_text_area = None
    random_generated_label = None
    tag_gen_search_entry = None # 左側カテゴリツリーの検索用
    tag_list_search_entry = None # 右側タグリストの検索用
    tag_gen_filter_var = None # これは使われなくなるが、初期化は残す
    delimiter_var = None
    dict_search_entry = None
    dict_filter_var = None
    dict_filter_combobox = None
    
    try:
        main()
    except Exception as e:
        # エラーメッセージをコンソールにも出力
        import traceback
        traceback.print_exc()
        # エラーメッセージをメッセージボックスで表示
        messagebox.showerror("アプリケーションエラー", f"予期せぬエラーが発生しました:\n{e}\n\n詳細については、コンソール出力を確認してください。")

