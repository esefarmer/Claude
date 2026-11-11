#!/bin/bash
# 全文書を生成する。usage: bash build_all.sh  （_src ディレクトリで実行）
set -e
OUT=..
b() { node build.js "content/$1" "$OUT/$2"; zip -dq "$OUT/$2" '*/' 2>/dev/null || true; }
b 00_gaiyou.txt            "00_整備概要・要確認事項.docx"
b 01b_shinkyu.txt          "01b_就業規則_新旧対照表.docx"
b 02_taisho_kitei.txt      "02_児童対象性暴力等対処規程.docx"
b 03_hanzai_joho_kitei.txt "03_犯罪事実確認・情報管理規程.docx"
b 04_soudan_youryou.txt    "04_相談窓口設置要領.docx"
b 05_taiou_manual.txt      "05_性暴力事案発生時の対応マニュアル.docx"
b 06_kenshu_seiyaku.txt    "06_教職員向け研修資料・誓約書.docx"
b 07_youshiki.txt          "07_様式集（相談受付票・事実確認記録票・対応記録票ほか）.docx"
b 08_saiyou_shoshiki.txt   "08_採用時の確認・誓約に関する書式.docx"
b 09_shuuchi.txt           "09_保護者・生徒への周知文書.docx"
