"""HS コード自動判定モジュール

商品タイトル（英語/日本語）から、戸井さん指定の HS コードを推定する。
不明な場合はデフォルト値 (9999999000) を返す。

使い方:
    from hs_code_lookup import lookup_hs_code

    code = lookup_hs_code("Ultraman Kaiju Encyclopedia New Edition 2008")
    # → "4901990093" (本/漫画)

    code = lookup_hs_code("Ensky Puzzle Frame One Piece")
    # → "9504500000" (ゲームソフト→パズルも近い扱い) or 9503000000 (玩具)
"""

import re


DEFAULT_HS_CODE = "9999999000"


# === キーワード → HS コード マッピング ===
# 優先順位の高い物（より具体的）を上に配置。
# 順番に走査して、最初にマッチしたコードを返す。
HS_CODE_RULES = [
    # (HSコード, [マッチするキーワードリスト], 説明)

    # === 楽器系 ===
    ("9205900000", ["clarinet", "クラリネット", "flute", "フルート",
                    "saxophone", "サックス", "trumpet", "トランペット",
                    "trombone", "wind instrument"], "クラリネット、フルート"),
    ("9207100000", ["electronic keyboard", "synthesizer", "synth",
                    "電子楽器", "キーボード", "midi keyboard"], "電子楽器、キーボード"),

    # === 時計 ===
    ("9102210000", ["wristwatch", "腕時計", "wrist watch",
                    "mechanical watch", "自動巻"], "腕時計（機械式）"),
    ("9102110000", ["quartz watch", "クオーツ"], "腕時計（クオーツ）"),

    # === メディア ===
    ("8523801000", ["vinyl record", "lp record", "レコード",
                    "vinyl ", "33rpm", "45rpm"], "レコード"),
    ("8523809000", ["blu-ray", "blue-ray", "bluray", "ブルーレイ"], "ブルーレイディスク"),
    ("8523492000", ["compact disc", "音楽cd", "music cd",
                    "audio cd"], "CD"),
    ("9504500000", ["game software", "video game", "ゲームソフト",
                    "nintendo switch", "ps4 game", "ps5 game",
                    "game disc", "game cartridge"], "ゲームソフト"),

    # === エレクトロニクス ===
    ("8517120000", ["smartphone", "iphone", "android phone",
                    "mobile phone", "携帯", "スマホ", "スマートフォン"], "携帯、スマホ"),
    ("8471300000", ["laptop", "notebook computer", "macbook",
                    "thinkpad", "パソコン", "ノートパソコン",
                    "デスクトップパソコン", "tablet computer"], "パソコン"),
    ("8527190000", ["radio cassette", "boombox", "ラジカセ",
                    "portable radio"], "ラジカセ"),
    ("8518220000", ["speaker", "スピーカ", "loudspeaker",
                    "bookshelf speaker"], "スピーカ"),

    # === 家電 ===
    ("8516320000", ["hair iron", "flat iron", "curling iron",
                    "アイロン", "美容アイロン", "ストレートアイロン"], "美容アイロン"),
    ("8516310000", ["hair dryer", "ドライヤー", "blow dryer"], "ドライヤー"),
    ("8438500000", ["food slicer", "slicer", "食品スライサー"], "食品スライサー"),
    ("8419810000", ["cooker", "rice cooker", "炊飯器",
                    "調理用機器", "induction cooker"], "調理用機器"),
    ("9405100000", ["fluorescent light", "蛍光灯", "lighting fixture",
                    "ceiling light", "照明器具"], "蛍光灯器具"),

    # === 玩具 / ぬいぐるみ ===
    ("9503000000", ["figure", "フィギュア", "plush", "ぬいぐるみ",
                    "doll", "toy", "玩具", "stuffed", "action figure",
                    "model toy", "diecast", "プラモデル", "puzzle"], "フィギュア、ぬいぐるみ"),

    # === 服飾 ===
    ("6211320000", ["yukata", "浴衣", "kimono", "着物"], "浴衣、着物"),
    ("6110200000", ["sweater", "セーター", "knit shirt",
                    "polo shirt", "tシャツ", "t-shirt",
                    "シャツ", "服 ", "cotton shirt"], "服 (一般)"),

    # === 履物・防具 ===
    ("6404110000", ["sneaker", "athletic shoe", "sports shoe",
                    "スニーカー", "running shoe"], "スニーカー"),
    ("6506100000", ["helmet", "ヘルメット", "safety helmet",
                    "bike helmet", "motorcycle helmet"], "ヘルメット"),

    # === 工具 ===
    ("8201500000", ["bonsai scissors", "bonsai shears", "盆栽鋏",
                    "garden shears", "pruning shears", "剪定鋏"], "盆栽鋏、園芸鋏"),

    # === バッグ類（"Book Bag" などが book に誤マッチするのを防ぐため本より先）===
    # ※戸井さん要確認: HSコード 4202929900 (バッグ・カバン類の一般値)
    ("4202929900", ["shoulder bag", "tote bag", "handbag", "backpack",
                    "school bag", "book bag", "messenger bag", "duffel",
                    "briefcase", "バッグ", "鞄", "ハンドバッグ"], "バッグ類 [要確認]"),

    # === 翻訳機 ===
    # ※戸井さん要確認: HSコード 8517620000
    ("8517620000", ["pocketalk", "translator", "翻訳機"], "翻訳機 [要確認]"),

    # === 置時計・掛時計 (腕時計と別) ===
    # ※戸井さん要確認: HSコード 9105210000
    ("9105210000", ["flip clock", "alarm clock", "wall clock", "desk clock",
                    "table clock", "置時計", "目覚まし時計", "掛時計"], "時計（置/掛）[要確認]"),

    # === 本 / 漫画（範囲広いので最後の方に置く）===
    ("4901990093", ["book", "encyclopedia", "manga", "magazine",
                    "本", "漫画", "雑誌", "guide book", "art book",
                    "photo book", "kaiju encyclopedia", "guide from japan",
                    "shogakukan", "kodansha", "novel", "comic"], "本、漫画、雑誌"),
]


def normalize(text):
    """テキストを小文字化、全角→半角、空白整形"""
    if not text:
        return ""
    t = str(text).lower()
    # 全角英数を半角に
    t = t.translate(str.maketrans(
        "０１２３456789ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ"
        "ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ",
        "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    ))
    return t


def lookup_hs_code(item_title, ebay_category_id=None):
    """商品タイトル（and/or eBay カテゴリID）からHSコードを判定

    Args:
        item_title: 商品名（英語/日本語）
        ebay_category_id: eBay Category ID（optional、より正確）

    Returns:
        str: 10桁のHSコード（ドットなし）
    """
    title = normalize(item_title or "")

    # キーワードマッチ
    for code, keywords, _desc in HS_CODE_RULES:
        for kw in keywords:
            if kw.lower() in title:
                return code

    return DEFAULT_HS_CODE


def lookup_with_description(item_title, ebay_category_id=None):
    """HSコード + 該当した分類説明を返す（デバッグ用）"""
    title = normalize(item_title or "")
    for code, keywords, desc in HS_CODE_RULES:
        for kw in keywords:
            if kw.lower() in title:
                return (code, desc, kw)
    return (DEFAULT_HS_CODE, "不明（デフォルト）", "")


if __name__ == "__main__":
    # テスト
    samples = [
        "Ultraman Kaiju Encyclopedia New Edition 2008 Shogakukan Tsuburaya Kaiju Guide From Japan",
        "POCKETALK S PTSGW Translator Global Communication White Multilingual SOURCENEXT",
        "National Panasonic Flip Clock TG02 Orange Alarm 70s Japan Young suyasuya Tested",
        "Ensky Puzzle Frame One Piece Ultimate Frame Metal 50x75cm Japan",
        "MARC JACOBS Shoulder Bag Book Bag M0017047 Cotton Canvas Black From Japan",
        "Yamaha Clarinet YCL-450",
        "Casio G-Shock Wristwatch GA-2100",
        "Vintage Vinyl Record Pink Floyd Dark Side of the Moon",
        "Nintendo Switch Game: Zelda Breath of the Wild",
        "ぬいぐるみ ピカチュウ",
    ]

    print("=" * 70)
    print("HSコード判定テスト")
    print("=" * 70)
    for title in samples:
        code, desc, kw = lookup_with_description(title)
        print()
        print("Title : " + title[:60])
        print("HSCode: " + code + " (" + desc + ", キーワード=" + kw + ")")
