from flask import Flask, request, render_template, send_file
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
import os
import io
import re
import subprocess
import tempfile

app = Flask(__name__)

# -----------------------------
# 設定（あなたの要望すべて反映）
# -----------------------------

# 1スライドあたりの最大文字数
MAX_CHARS_PER_SLIDE = 150

# テキスト領域設定（右下枠と被らない範囲）
TEXT_LEFT_CM = 0.79
TEXT_TOP_CM = 0.80
TEXT_WIDTH_CM = 25.2   # 枠にかからない右端まで
TEXT_HEIGHT_CM = 15.6

# 右下の枠設定（位置・サイズ）
FRAME_LEFT_CM = 25.87
FRAME_TOP_CM = 14.55
FRAME_WIDTH_CM = 8.0
FRAME_HEIGHT_CM = 4.5

# ページ番号（分割時のみ表示）
PAGE_LEFT_CM = 21.94
PAGE_TOP_CM = 16.93
PAGE_COLOR = RGBColor(0x00, 0x9D, 0xFF)  # #009DFF
PAGE_FONT_SIZE_PT = 32
PAGE_FONT_BOLD = True

# 明示的な話者ごとの色（固定）
NAME_FIXED_COLORS = {
    "仲條": RGBColor(0x00, 0xFD, 0xFF),  # #00FDFF（水色）
    "三村": RGBColor(0xFF, 0xFF, 0xFF),  # #FFFFFF（白）
    "星野": RGBColor(0xFF, 0xFF, 0x00),  # #FFFF00（黄色）
}

# 明示指定以外の話者に自動割り当て（以降固定）
AUTO_COLOR_POOL = [
    RGBColor(0xFF, 0x40, 0xFF),  # ピンク
    RGBColor(0xFF, 0xA5, 0x00),  # オレンジ
    RGBColor(0xFF, 0xFB, 0x00),  # 黄（予備）
]
name_color_map = {}
_auto_color_idx = 0


# =============================
# 補助関数群
# =============================

def get_color_for_name(name: str):
    """話者ごとに色を固定して返す"""
    global _auto_color_idx
    if not name:
        return RGBColor(0xFF, 0xFF, 0xFF)  # デフォルト白

    # 明示指定があればそれを使う
    if name in NAME_FIXED_COLORS:
        return NAME_FIXED_COLORS[name]

    # 既存登録があればそれを再利用
    if name in name_color_map:
        return name_color_map[name]

    # 新規話者 → 自動カラー割り当て
    color = AUTO_COLOR_POOL[_auto_color_idx % len(AUTO_COLOR_POOL)]
    name_color_map[name] = color
    _auto_color_idx += 1
    return color


def clean_text(text: str) -> str:
    """ノート欄から制御文字を除去"""
    return text.replace("\x0b", "").strip()


def parse_notes_into_segments(note_text: str):
    """
    ノートを《名前》単位で抽出 [(名前, テキスト)]。
    連続する同一話者は結合して扱う。
    """
    segments = []
    current_name = None
    buffer = []

    for raw in note_text.splitlines():
        line = raw.strip()
        if not line:
            continue

        match = re.match(r"《(.+?)》", line)
        if match:
            # 直前バッファを保存
            if buffer:
                joined = "".join(buffer).strip()
                if joined:
                    segments.append((current_name, joined))
                buffer = []
            current_name = match.group(1)
            rest = line[match.end():]
            if rest:
                buffer.append(rest)
        else:
            buffer.append(line)

    if buffer:
        joined = "".join(buffer).strip()
        if joined:
            segments.append((current_name, joined))

    # 連続する同名を結合
    merged = []
    for name, text in segments:
        if merged and merged[-1][0] == name:
            merged[-1] = (name, merged[-1][1] + text)
        else:
            merged.append((name, text))
    return merged


def pack_segments_into_chunks(segments, max_len=MAX_CHARS_PER_SLIDE):
    """
    同じスライド内で話者を区切らず150文字単位で分割。
    1チャンク内に収まる文字数が<=max_len。
    """
    chunks = []
    cur = []
    cur_len = 0

    def flush():
        nonlocal cur, cur_len
        if cur:
            chunks.append(cur)
            cur = []
            cur_len = 0

    for name, text in segments:
        i = 0
        while i < len(text):
            remain = max_len - cur_len
            if remain <= 0:
                flush()
                remain = max_len
            take = min(remain, len(text) - i)
            part = text[i:i + take]
            cur.append((name, part))
            cur_len += len(part)
            i += take
    flush()
    return chunks


def add_photo_frame(slide):
    """右下に固定枠を追加"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Cm(FRAME_LEFT_CM),
        Cm(FRAME_TOP_CM),
        Cm(FRAME_WIDTH_CM),
        Cm(FRAME_HEIGHT_CM)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0xF0, 0xF0, 0xF0)
    shape.line.color.rgb = RGBColor(0x64, 0x64, 0x64)
    shape.line.width = Pt(2)
    return shape


def add_page_indicator(slide, index, total):
    """分割時のみページ番号を表示"""
    if total <= 1:
        return
    tb = slide.shapes.add_textbox(Cm(PAGE_LEFT_CM), Cm(PAGE_TOP_CM), Cm(4), Cm(1.5))
    tf = tb.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = f"{index}/{total}"
    p.alignment = PP_ALIGN.LEFT
    font = p.font
    font.name = "メイリオ"
    font.size = Pt(PAGE_FONT_SIZE_PT)
    font.bold = PAGE_FONT_BOLD
    font.color.rgb = PAGE_COLOR


# =============================
# スライド生成
# =============================

def create_script_slides(notes):
    prs = Presentation()
    prs.slide_width = Cm(33.867)
    prs.slide_height = Cm(19.05)

    for note in notes:
        segments = parse_notes_into_segments(note)
        chunks = pack_segments_into_chunks(segments, MAX_CHARS_PER_SLIDE)
        total_parts = len(chunks)

        for idx, chunk in enumerate(chunks, start=1):
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            # 背景を黒に
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(0, 0, 0)

            # テキストボックス
            txBox = slide.shapes.add_textbox(
                Cm(TEXT_LEFT_CM),
                Cm(TEXT_TOP_CM),
                Cm(TEXT_WIDTH_CM),
                Cm(TEXT_HEIGHT_CM)
            )
            tf = txBox.text_frame
            tf.clear()
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.space_before = Pt(0)
            p.space_after = Pt(0)

            # 各話者テキストをrun単位で追加
            for name, part in chunk:
                run = p.add_run()
                prefix = f"《{name}》" if name else ""
                run.text = prefix + part
                f = run.font
                f.name = "メイリオ"
                f.size = Pt(40)
                f.bold = True
                f.color.rgb = get_color_for_name(name)

            # 右下の枠
            add_photo_frame(slide)

            # ページ番号
            add_page_indicator(slide, idx, total_parts)

    return prs


# =============================
# Flaskメイン処理
# =============================

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["pptx_file"]
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = os.path.join(temp_dir, "input.pptx")
            file.save(input_path)

            prs = Presentation(input_path)
            notes = []
            for slide in prs.slides:
                if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                    text = slide.notes_slide.notes_text_frame.text
                    cleaned = clean_text(text)
                    if cleaned:
                        notes.append(cleaned)

            new_prs = create_script_slides(notes)
            buf = io.BytesIO()
            new_prs.save(buf)
            buf.seek(0)
            return send_file(
                buf,
                as_attachment=True,
                download_name="スクリプトスライド_自動生成.pptx",
                mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
