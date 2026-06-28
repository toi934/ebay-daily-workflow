"""寸法・重量 自動推定モジュール

商品タイトルから、典型的な寸法 (長さ × 幅 × 高さ cm) と重量 (kg) を推定する。
CPaSS の編集画面に自動入力する値の決定に使う。

使い方:
    from dimension_weight_lookup import lookup_dimensions_weight

    dims = lookup_dimensions_weight("Ultraman Kaiju Encyclopedia 2008")
    # → {"length_cm": 26, "width_cm": 19, "height_cm": 2, "weight_kg": 0.5, "category": "図録/百科"}

漫画・本のセット販売対応:
- "Vol.1-30" などのパターンを検出 → 巻数 N を抽出
- 寸法: 13 × 19 × (1.6 × N) cm
- 重量: 0.2 × N kg

注意:
- これは平均的な推定値です。実際の商品によって差があります。
- DHLの送料計算は容積重量で算出されるので、過小推定するとリスク有。
  → 不明な時は **大きめ**に取る。
"""

import re


# (キーワードリスト, 長さcm, 幅cm, 高さcm, 重量kg, 説明)
DIMENSION_RULES = [
    # ============ 楽器 ============
    (["clarinet", "クラリネット"], 70, 25, 18, 1.5, "クラリネット (ケース込)"),
    (["flute", "フルート"], 50, 15, 10, 1.2, "フルート (ケース込)"),
    (["saxophone", "サックス"], 70, 35, 25, 4.5, "サックス (ケース込)"),
    (["trumpet", "トランペット"], 55, 25, 15, 2.5, "トランペット"),
    (["electronic keyboard", "synthesizer", "synth", "電子楽器",
      "midi keyboard"], 100, 30, 15, 7, "電子キーボード"),

    # ============ 時計 ============
    (["wristwatch", "wrist watch", "腕時計", "g-shock",
      "g shock", "rolex", "seiko watch", "casio watch"], 15, 15, 10, 0.6, "腕時計 (箱込)"),

    # ============ メディア (CD/レコード/BD/ゲーム) ============
    (["vinyl record", "lp record", "レコード", "vinyl ", "33rpm"],
     33, 33, 1, 0.3, "レコード（LP）"),
    (["blu-ray", "blue-ray", "bluray", "ブルーレイ"],
     19, 14, 1.5, 0.15, "ブルーレイディスク"),
    (["compact disc", "audio cd", "music cd"],
     14, 13, 1, 0.15, "音楽CD"),
    (["nintendo switch game", "switch game", "ps4 game", "ps5 game",
      "video game", "game software", "ゲームソフト", "game cartridge"],
     17, 11, 1.5, 0.15, "ゲームソフト"),

    # ============ 電子機器 ============
    (["iphone", "smartphone", "android phone", "スマホ", "スマートフォン",
      "携帯電話", "mobile phone"], 17, 9, 5, 0.4, "スマホ (箱込)"),
    (["macbook", "laptop", "notebook computer", "thinkpad", "ノートパソコン"],
     38, 28, 5, 2.5, "ノートパソコン"),
    (["tablet computer", "ipad"], 30, 22, 3, 1.0, "タブレット"),
    (["radio cassette", "boombox", "ラジカセ"],
     45, 20, 18, 3.5, "ラジカセ"),
    (["loudspeaker", "bookshelf speaker", "スピーカ"],
     35, 25, 30, 5, "スピーカー (1台)"),
    (["camera", "カメラ", "lens", "レンズ"], 20, 15, 15, 1.2, "カメラ/レンズ"),
    (["pocketalk", "translator", "翻訳機"], 15, 9, 5, 0.3, "翻訳機"),

    # ============ 家電 ============
    (["hair iron", "flat iron", "curling iron", "アイロン",
      "美容アイロン", "ストレートアイロン"], 35, 15, 8, 0.7, "美容アイロン"),
    (["hair dryer", "ドライヤー", "blow dryer"], 30, 25, 15, 0.9, "ヘアドライヤー"),
    (["food slicer", "slicer", "食品スライサー"], 40, 30, 25, 5, "食品スライサー"),
    (["rice cooker", "炊飯器"], 35, 30, 30, 5, "炊飯器"),
    (["induction cooker", "ih"], 35, 30, 7, 3, "IH調理器"),
    (["fluorescent light", "蛍光灯", "ceiling light", "照明器具"],
     65, 30, 15, 3, "照明器具"),

    # ============ 玩具 ============
    (["puzzle", "jigsaw"], 50, 35, 5, 1.5, "ジグソーパズル/フレーム"),
    (["plush", "stuffed", "ぬいぐるみ"], 40, 30, 30, 0.6, "ぬいぐるみ"),
    (["figure", "フィギュア", "diecast", "プラモデル", "action figure",
      "doll", "model toy"], 25, 18, 28, 0.7, "フィギュア"),

    # ============ 服飾 ============
    (["yukata", "kimono", "浴衣", "着物"], 35, 30, 6, 1.2, "浴衣・着物"),
    (["sweater", "セーター"], 35, 28, 4, 0.5, "セーター"),
    (["polo shirt", "t-shirt", "tシャツ", "シャツ", "knit shirt"],
     30, 25, 3, 0.3, "シャツ類"),

    # ============ 履物・防具 ============
    (["sneaker", "athletic shoe", "running shoe", "スニーカー"],
     33, 22, 13, 1.1, "スニーカー (箱込)"),
    (["motorcycle helmet", "bike helmet", "safety helmet", "ヘルメット"],
     32, 27, 27, 1.8, "ヘルメット"),

    # ============ 工具 ============
    (["bonsai scissors", "bonsai shears", "盆栽鋏", "garden shears",
      "pruning shears", "剪定鋏"], 25, 10, 3, 0.4, "盆栽鋏"),

    # ============ バッグ ============
    (["shoulder bag", "tote bag", "handbag", "バッグ", "鞄"],
     35, 30, 15, 1.0, "バッグ"),

    # ============ 本/雑誌 (最後の方に置く) ============
    (["encyclopedia", "art book", "photo book", "guide book",
      "shogakukan", "kodansha"], 27, 21, 2, 0.7, "図録/百科"),
    (["magazine", "雑誌"], 26, 19, 0.7, 0.35, "雑誌"),
    (["manga", "漫画", "comic"], 18, 11, 2, 0.3, "漫画"),
    (["book", "novel", "本", "guide from japan"], 21, 15, 2, 0.4, "本（一般）"),
]


DEFAULT_DIMS = {
    "length_cm": 30,
    "width_cm": 20,
    "height_cm": 10,
    "weight_kg": 0.5,
    "category": "不明（デフォルト）",
    "matched_keyword": "",
}


def normalize(text):
    if not text:
        return ""
    return str(text).lower()


def detect_manga_set_volumes(title):
    """漫画・本のセット販売の場合、巻数 N を返す。なければ None。

    検出パターン例:
        "Vol.1-30" → 30
        "Vol 1-15" → 15
        "Volume 1-50 Complete Set" → 50
        "1-20 Vols Complete" → 20
        "5 Volume Set" → 5
        "Complete 30 Volume Manga Set" → 30
    """
    t = normalize(title)

    # Pattern 1: vol.X-Y or vol X-Y or volume X-Y
    m = re.search(r'vol(?:ume)?\.?\s*(\d+)\s*[-〜~–]\s*(\d+)', t)
    if m:
        start = int(m.group(1))
        end = int(m.group(2))
        if end > start and end - start <= 200:
            return end - start + 1

    # Pattern 2: X-Y vols/volumes
    m = re.search(r'(\d+)\s*[-〜~–]\s*(\d+)\s*(?:vols?|volumes?)', t)
    if m:
        start = int(m.group(1))
        end = int(m.group(2))
        if end > start and end - start <= 200:
            return end - start + 1

    # Pattern 3: "N volume set" / "N volumes set" / "complete N volume"
    if "complete set" in t or "full set" in t or "complete" in t:
        m = re.search(r'(\d+)\s*(?:vols?|volumes?|books?|冊)', t)
        if m:
            n = int(m.group(1))
            if 2 <= n <= 200:
                return n

    return None


def lookup_dimensions_weight(item_title):
    """商品タイトルから 寸法 + 重量 を推定

    Returns:
        dict: {length_cm, width_cm, height_cm, weight_kg, category, matched_keyword}
    """
    title = normalize(item_title or "")

    # === 漫画・本のセット販売を優先的に検出 ===
    volumes = detect_manga_set_volumes(title)
    if volumes:
        # 漫画/本キーワードがあれば確実にセット
        is_book_related = any(kw in title for kw in [
            "manga", "comic", "book", "novel", "vol", "volume",
            "漫画", "本", "巻", "全巻"
        ])
        if is_book_related:
            return {
                "length_cm": 19.0,
                "width_cm": 13.0,
                "height_cm": float(min(60.0, 1.6 * volumes)),  # 60cm cap (CPaSS制限想定)
                "weight_kg": round(0.2 * volumes, 2),
                "category": "漫画/本セット (" + str(volumes) + "冊)",
                "matched_keyword": "Vol.1-" + str(volumes),
            }

    # === 通常の単品検索 ===
    for entry in DIMENSION_RULES:
        keywords, length, width, height, weight, desc = entry
        for kw in keywords:
            if kw.lower() in title:
                return {
                    "length_cm": float(length),
                    "width_cm": float(width),
                    "height_cm": float(height),
                    "weight_kg": float(weight),
                    "category": desc,
                    "matched_keyword": kw,
                }
    return dict(DEFAULT_DIMS)


if __name__ == "__main__":
    # テスト
    samples = [
        "Ultraman Kaiju Encyclopedia New Edition 2008 Shogakukan Tsuburaya Kaiju Guide From Japan",
        "POCKETALK S PTSGW Translator Global Communication White",
        "National Panasonic Flip Clock TG02 Orange Alarm",
        "Ensky Puzzle Frame One Piece Ultimate Frame Metal 50x75cm Japan",
        "MARC JACOBS Shoulder Bag Book Bag M0017047 Cotton Canvas Black",
        "Yamaha Clarinet YCL-450 with Case",
        "Casio G-Shock Wristwatch GA-2100",
        "Pink Floyd Dark Side of the Moon Vinyl Record LP",
        "Nintendo Switch Game: Zelda Breath of the Wild",
        "Pokemon Pikachu Plush Stuffed Toy 30cm",
        "Yukata Cotton Summer Kimono Japan",
        "Sony Boombox CFD-S70 Radio Cassette",
        "Yamaha YPT-260 Electronic Keyboard 61 keys",
        # 漫画セット
        "Arpeggio of Blue Steel Vol.1-30 Complete Full Set Japanese Manga Comics Japan",
        "One Piece Vol.1-105 Complete Set Japanese Manga",
        "Dragon Ball Vol.1-42 Complete Set Manga",
        "Vol 5-15 Manga Series",
        "Naruto 72 Volume Complete Set",
    ]

    print("=" * 80)
    print("寸法・重量 自動推定テスト")
    print("=" * 80)
    for title in samples:
        d = lookup_dimensions_weight(title)
        print()
        print("Title : " + title[:65])
        print("  → " + str(int(d["length_cm"])) + " x " + str(int(d["width_cm"])) +
              " x " + str(int(d["height_cm"])) + " cm, " +
              str(d["weight_kg"]) + " kg  [" + d["category"] +
              " / kw=" + d["matched_keyword"] + "]")
