"""script.md とスライド画像から、ナレーション・字幕つきの解説動画を作る。

必要なもの: open_jtalk（naist-jdic）, HTS Voice "Mei", ffmpeg（libass）, Pillow, Noto Sans CJK JP
使い方: python3 build.py <スライドPNGのディレクトリ> <mei_normal.htsvoice> <作業ディレクトリ>
"""
import re
import subprocess
import sys
import wave
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).parent
SLIDES, VOICE, WORK = (Path(p) for p in sys.argv[1:4])
OUT = HERE / "output"
DIC = "/var/lib/mecab/dic/open-jtalk/naist-jdic"
FONT = "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"

RATE = 48000
LEAD, GAP, TAIL = 0.4, 0.35, 0.9   # 秒：スライド冒頭・文間・スライド末尾の間
COUNTDOWN_SLIDE, COUNTDOWN_SEC = 2, 5
CREDIT_SEC = 5
MAX_CUE = 32                        # 字幕1枚あたりの最大文字数

# 字幕はそのまま、読み上げだけ置き換える語
READINGS = [
    ("1日", "いちにち"), ("3,500", "3500"), ("35,000", "35000"),
    ("未来①", "未来いち"), ("未来②", "未来に"),
    ("未来A", "未来エー"), ("未来B", "未来ビー"), ("FA選択", "エフエー選択"),
    ("1・2限", "1、2限"), ("3・4限", "3、4限"), ("5・6限", "5、6限"),
    ("Word", "ワード"), ("Excel", "エクセル"), ("「IT」", "「アイティー」"),
    ("「P.E.」", "「ピーイー」"), ("先の自分", "さきの自分"), ("「先」", "「さき」"),
    ("創", "つく"),
]


def reading(text):
    for src, dst in READINGS:
        text = text.replace(src, dst)
    return text


def parse_script():
    """### N. 見出し の下の本文を {スライド番号: [文, ...]} にする。【演出】行は除く。"""
    slides, cur = {}, None
    for line in (HERE / "script.md").read_text(encoding="utf-8").splitlines():
        m = re.match(r"### (\d+)\.", line)
        if m:
            cur = int(m.group(1))
            slides[cur] = []
        elif cur and line.strip() and not line.startswith("【"):
            slides[cur] += [s for s in re.findall(r"[^。！？]+[。！？]?」?", line.strip()) if s.strip()]
    return slides


def split_cue(sentence):
    """長い文を読点の位置で MAX_CUE 文字以下の字幕に分ける。"""
    if len(sentence) <= MAX_CUE:
        return [sentence]
    parts, buf = [], ""
    for piece in re.findall(r"[^、]+、?", sentence):
        if buf and len(buf) + len(piece) > MAX_CUE:
            parts.append(buf)
            buf = ""
        buf += piece
    return parts + [buf] if buf else parts


def synth(text, path):
    subprocess.run(["open_jtalk", "-x", DIC, "-m", str(VOICE), "-ow", str(path)],
                   input=reading(text).encode(), check=True)
    with wave.open(str(path)) as w:
        assert w.getframerate() == RATE and w.getsampwidth() == 2 and w.getnchannels() == 1
        return w.readframes(w.getnframes())


def silence(sec):
    return b"\x00\x00" * int(RATE * sec)


def ass_time(t):
    cs = round(t * 100)
    return f"{cs // 360000}:{cs // 6000 % 60:02}:{cs // 100 % 60:02}.{cs % 100:02}"


def srt_time(t):
    ms = round(t * 1000)
    return f"{ms // 3600000:02}:{ms // 60000 % 60:02}:{ms // 1000 % 60:02},{ms % 1000:03}"


def make_credit(path):
    img = Image.new("RGB", (1920, 1080), "#FAFAFA")
    d = ImageDraw.Draw(img)
    big, small = ImageFont.truetype(FONT, 64), ImageFont.truetype(FONT, 36)
    d.text((960, 430), "京都翔英高等学校　未来開拓講座 履修登録ガイダンス", font=big, fill="#2B2B2B", anchor="mm")
    d.text((960, 600), "音声合成：Open JTalk ／ HTS Voice \"Mei\"", font=small, fill="#556189", anchor="mm")
    d.text((960, 660), "(c) Nagoya Institute of Technology　CC BY 3.0", font=small, fill="#556189", anchor="mm")
    img.save(path)


def main():
    WORK.mkdir(parents=True, exist_ok=True)
    OUT.mkdir(exist_ok=True)
    script = parse_script()
    audio, cues, timeline = bytearray(), [], []   # timeline: (画像, 秒)
    t = 0.0
    countdown_at = None

    for n in sorted(script):
        start = t
        audio += silence(LEAD)
        t += LEAD
        for i, sent in enumerate(script[n]):
            pcm = synth(sent, WORK / f"{n:02}_{i:02}.wav")
            dur = len(pcm) / 2 / RATE
            parts = split_cue(sent)
            total = sum(len(p) for p in parts)
            ct = t
            for p in parts:
                d = dur * len(p) / total
                cues.append((ct, ct + d, p))
                ct += d
            audio += pcm + silence(GAP)
            t += dur + GAP
        if n == COUNTDOWN_SLIDE:
            countdown_at = t
            audio += silence(COUNTDOWN_SEC)
            t += COUNTDOWN_SEC
        audio += silence(TAIL - GAP)
        t += TAIL - GAP
        timeline.append((SLIDES / f"s-{n:02}.png", t - start))

    credit = WORK / "credit.png"
    make_credit(credit)
    audio += silence(CREDIT_SEC)
    t += CREDIT_SEC
    timeline.append((credit, CREDIT_SEC))

    with wave.open(str(WORK / "narration.wav"), "wb") as w:
        w.setnchannels(1), w.setsampwidth(2), w.setframerate(RATE)
        w.writeframes(bytes(audio))

    concat = WORK / "slides.txt"
    lines = [f"file '{img}'\nduration {d:.3f}" for img, d in timeline]
    concat.write_text("\n".join(lines) + f"\nfile '{timeline[-1][0]}'\n", encoding="utf-8")

    # 字幕：焼き込み用 ASS と、配信先で切り替えられる SRT
    header = (
        "[Script Info]\nScriptType: v4.00+\nPlayResX: 1920\nPlayResY: 1080\n\n"
        "[V4+ Styles]\nFormat: Name, Fontname, Fontsize, PrimaryColour, OutlineColour, BackColour, "
        "Bold, BorderStyle, Outline, Shadow, Alignment, MarginV\n"
        "Style: Default,Noto Sans CJK JP,52,&H00FFFFFF,&H00000000,&H66402B2B,1,3,14,0,2,40\n\n"
        "[Events]\nFormat: Layer, Start, End, Style, Text\n"
    )
    ass = header + "".join(f"Dialogue: 0,{ass_time(a)},{ass_time(b)},Default,{s}\n" for a, b, s in cues)
    (WORK / "subs.ass").write_text(ass, encoding="utf-8")
    srt = "\n".join(f"{i}\n{srt_time(a)} --> {srt_time(b)}\n{s}\n" for i, (a, b, s) in enumerate(cues, 1))
    (OUT / "未来開拓講座_履修登録ガイダンス.srt").write_text(srt, encoding="utf-8")

    c0, c1 = countdown_at, countdown_at + COUNTDOWN_SEC
    vf = (
        f"fps=30,subtitles={WORK / 'subs.ass'},"
        f"drawtext=fontfile={FONT}:text='%{{eif\\:{COUNTDOWN_SEC}-floor(t-{c0:.3f})\\:d}}':"
        f"fontsize=180:fontcolor=0x556189:box=1:boxcolor=0xFFFFFF@0.85:boxborderw=40:"
        f"x=w-text_w-170:y=150:enable='between(t,{c0:.3f},{c1 - 0.01:.3f})',"
        f"fade=t=in:d=0.5,fade=t=out:st={t - 1:.3f}:d=1,format=yuv420p"
    )
    subprocess.run([
        "ffmpeg", "-y", "-loglevel", "error",
        "-f", "concat", "-safe", "0", "-i", str(concat), "-i", str(WORK / "narration.wav"),
        "-vf", vf, "-c:v", "libx264", "-preset", "medium", "-crf", "24", "-tune", "stillimage",
        "-af", "loudnorm=I=-16:TP=-1.5", "-c:a", "aac", "-b:a", "128k", "-ar", "48000",
        "-movflags", "+faststart", "-shortest", str(OUT / "未来開拓講座_履修登録ガイダンス.mp4"),
    ], check=True)
    print(f"total {t:.1f}s, cues {len(cues)}, countdown {c0:.1f}s")


if __name__ == "__main__":
    main()
