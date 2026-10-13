#!/usr/bin/env python3
"""
normalize.py — اصلاحات مکانیکیِ «بی‌خطر» در ویرایش فنی متن فارسی
(بر پایه‌ی فصل‌های شیوه‌ی خط و نشانه‌گذاری کتاب ذوالفقاری).

فقط کارهایی را انجام می‌دهد که تقریباً هیچ‌وقت معنا را عوض نمی‌کنند:
  ۱. ي/ك عربی ← ی/ک فارسی
  ۲. رقم عربی (٠-٩) ← رقم فارسی؛ رقم لاتین ← فارسی فقط اگر --fa-digits داده شود
  ۳. «می / نمی» + فاصله ← نیم‌فاصله  (می روم ← می‌روم)
  ۴. واژه + فاصله + «ها/های/هایی…» ← نیم‌فاصله
  ۵. واژه + فاصله + «تر/ترین» ← نیم‌فاصله (فقط اگر واژه‌ی قبل دست‌کم ۲ حرف باشد)
  ۶. حذف فاصله پیش از ، ؛ : . ! ؟  و افزودن یک فاصله پس از آن‌ها (پیش از حرف فارسی)
  ۷. , ; ? در بافت فارسی ← ، ؛ ؟
  ۸. فاصله‌های تکراری ← یک فاصله؛ نیم‌فاصله‌های تکراری/بی‌جا حذف
  ۹. بیش از سه نقطه ← «…»

هرچه نیاز به داوری دارد (املا، همزه، «را»، حشو، جمع‌ها…) این‌جا نیست؛ آن‌ها کار خودِ ویراستار است.

کاربرد:
    python3 normalize.py in.txt -o out.txt [--log changes.tsv] [--fa-digits]
    python3 normalize.py in.docx -o out.docx          # متن پاراگراف‌ها و جدول‌ها (قالب run اول حفظ می‌شود)
"""
import re, sys, argparse
sys.stdout.reconfigure(encoding="utf-8"); sys.stderr.reconfigure(encoding="utf-8")

ZW = "‌"
FA = "آ-یء-يپچژکگی"
W = f"[{FA}{ZW}]"

RULES = [
    ("ya-kaf", r"ي", "ی"),
    ("ya-kaf", r"ك", "ک"),
    ("arabic-digits", r"[٠-٩]", lambda m: chr(ord(m.group(0)) - 0x0660 + 0x06F0)),
    ("mi", rf"(?<![{FA}{ZW}])(ن?می) (?={W})", lambda m: m.group(1) + ZW),
    ("mi-joined", rf"(?<![{FA}{ZW}])(ن?می)(?=(?:توان|تواند|توانند|توانم|توانی|توانیم|توانید|شود|شوند|شوم|شویم|کند|کنند|کنم|کنی|کنیم|کنید|کرد|کردند|کردم|رود|روند|روم|رفت|دهد|دهند|دهم|داد|دادند|گوید|گویند|گویم|گفت|گفتند|داند|دانند|دانم|دانست|خواهد|خواهند|خواهم|خواست|آید|آیند|آمد|آمدند|گیرد|گیرند|گرفت|بیند|بینند|بینم|دید|دیدند|نویسد|نویسند|ماند|مانند|گردد|گردند|یابد|یابند|رسد|رسند|رسید|سازد|سازند|پرسد|خورد|خورند|شد|شدند|بود|بودند|کشد|کشند|زند|زنند|زد|پذیرد|گذارد|گذرد|افتد|اندازد|نشیند|پردازد|کوشد|خواند|خوانند)(?![{FA}]))", lambda m: m.group(1) + ZW),
    ("ha", rf"(?<={W}) (ها(?:ی|یی|یم|یت|یش|یمان|یتان|یشان)?)(?![{FA}{ZW}])", lambda m: ZW + m.group(1)),
    ("tar", rf"(?<=[{FA}]{{2}}) (تر|ترین)(?![{FA}{ZW}])", lambda m: ZW + m.group(1)),
    ("latin-punct", rf"(?<=[{FA}])\s*,", "،"),
    ("latin-punct", rf"(?<=[{FA}])\s*;", "؛"),
    ("latin-punct", rf"(?<=[{FA}])\s*\?", "؟"),
    ("space-before-punct", r"(?<=\S)[ \t]+([،؛:.!؟])(?!\.)", lambda m: m.group(1)),
    ("space-after-punct", rf"([،؛؟!])(?=[{FA}])", lambda m: m.group(1) + " "),
    ("dots", r"\.{4,}", "…"),
    ("double-space", r"(?<=\S)[ ]{2,}(?=\S)", " "),
    ("zwnj-cleanup", rf"{ZW}{{2,}}", ZW),
    ("zwnj-cleanup", rf"{ZW}(?=\s)|(?<=\s){ZW}", ""),
]
LATIN_DIGITS = ("fa-digits", rf"(?<=[{FA}\s])[0-9]+(?=[\s{FA}،.؛])", lambda m: m.group(0).translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")))


def normalize(text, fa_digits=False, log=None):
    rules = RULES + ([LATIN_DIGITS] if fa_digits else [])
    for name, pat, rep in rules:
        def _r(m, name=name, rep=rep):
            new = rep(m) if callable(rep) else rep
            if log is not None and new != m.group(0):
                log.append((name, m.group(0), new))
            return new
        text = re.sub(pat, _r, text)
    return text


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path"); ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--log"); ap.add_argument("--fa-digits", action="store_true")
    a = ap.parse_args()
    log = []
    if a.path.lower().endswith(".docx"):
        import docx
        d = docx.Document(a.path)
        paras = list(d.paragraphs) + [p for t in d.tables for r in t.rows for c in r.cells for p in c.paragraphs]
        for p in paras:
            old = p.text
            new = normalize(old, a.fa_digits, log)
            if new != old and p.runs:
                p.runs[0].text = new
                for r in p.runs[1:]:
                    r.text = ""
        d.save(a.out)
    else:
        with open(a.path, encoding="utf-8") as f:
            t = f.read()
        with open(a.out, "w", encoding="utf-8") as f:
            f.write(normalize(t, a.fa_digits, log))
    if a.log:
        with open(a.log, "w", encoding="utf-8") as f:
            f.write("rule\tbefore\tafter\n")
            for r in log:
                f.write("\t".join(x.replace("\t", " ").replace(ZW, "<ZWNJ>") for x in r) + "\n")
    from collections import Counter
    c = Counter(r[0] for r in log)
    print("اصلاحات مکانیکی:", ", ".join(f"{k}={v}" for k, v in c.most_common()) or "هیچ")


if __name__ == "__main__":
    main()
