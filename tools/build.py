# -*- coding: utf-8 -*-
"""spec_*.json (extract.py の出力) + テンプレート -> ruikei-a/b/c.html と index.html を生成する。

使い方:
    python tools/extract.py   # Excel -> spec_*.json
    python tools/build.py     # spec_*.json -> HTML
設問文・選択肢・分岐条件はすべてExcelが正。ここで手を入れるのは「Excel固有の言い回しのWeb向け置換」
「大きなブロックの見出し(GROUP_HEADINGS)」「規則の概要(SUMMARY)」「表紙の文言」のみ。
"""
import json, os, re, sys, html

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS = os.path.join(ROOT, "tools")
BRAND = "つるい未来補助金"

CONFIG = {"returnTo": "", "deadline": ""}
cfg_path = os.path.join(TOOLS, "config.json")
if os.path.exists(cfg_path):
    CONFIG.update(json.load(open(cfg_path, encoding="utf-8")))

TYPES = {
    "a": {
        "file": "ruikei-a.html",
        "typeName": "類型A　新規開業・新分野進出支援事業",
        "short": "類型A",
        "lead": "村内で新しく事業を始める方（新規開業）、または村内で営業中の事業者が新しい分野に進出する方向け。",
        "summary": [
            ["対象事業", "新規開業支援事業（起業に必要な施設の整備・改修等）／新分野進出支援事業（既存の村内事業者が、日本標準産業分類の中分類を超えて新たに開始する事業）"],
            ["補助率", "補助対象経費の1/2以内"],
            ["補助限度額", "500万円。次に当てはまる場合は限度額に加算：村長が指定する区域内で事業を開始（＋100万円）／村長が指定する空き家等を活用（＋50万円）／申請日から起算して鶴居村に15年以上居住した経験のある45歳以下の方（＋200万円）"],
            ["補助金の下限額", "50万円"],
            ["対象経費", "工事請負費（事務所・店舗等の建設費・改修費等）、委託費（調査・設計・開発費等）、備品購入費（設備・機械装置・測量器具等）、そのほか村長が必要かつ適当と認めた費用"],
            ["ご注意", "住居兼店舗事務所は事業専有部分に係る経費のみが対象です（住居等の他用途部分と物理的に明確な区分が必要）。施設整備等の上限単価は24.8万円／㎡です。"],
        ],
    },
    "b": {
        "file": "ruikei-b.html",
        "typeName": "類型B　事業継続支援事業",
        "short": "類型B",
        "lead": "村内で事業を営む方が、安定的な事業継続のために設備の更新・改修などに取り組む場合向け。",
        "summary": [
            ["対象事業", "安定的な事業継続を図るために行う事業"],
            ["補助率", "補助対象経費の1/2以内"],
            ["補助限度額", "250万円"],
            ["補助金の下限額", "50万円"],
            ["対象経費", "報償費（専門家等への謝礼金など）、広告宣伝費、工事請負費（事務所・店舗等の建設費・改修費等）、委託費（調査・設計・開発費等）、備品購入費（設備・機械装置等）、そのほか村長が必要かつ適当と認めた費用"],
            ["ご注意", "施設整備等の上限単価は24.8万円／㎡です。"],
        ],
    },
    "c": {
        "file": "ruikei-c.html",
        "typeName": "類型C　観光推進事業",
        "short": "類型C",
        "lead": "旅館業法に基づく宿泊施設を整備・改修し、村内観光の推進に取り組む方向け。",
        "summary": [
            ["対象事業", "村内観光業の推進に資する、旅館業法に基づく宿泊施設の整備及び改修等を行う事業"],
            ["補助率", "補助対象経費の1/2以内"],
            ["補助限度額", "500万円"],
            ["補助金の下限額", "50万円"],
            ["対象経費", "工事請負費（事務所・店舗等の建設費・改修費等）、委託費（調査・設計・開発費等）、備品購入費（設備・機械装置等）、そのほか村長が必要かつ適当と認めた費用"],
            ["ご注意", "施設整備等の上限単価は24.8万円／㎡です。"],
        ],
    },
}
SUMMARY_CAVEAT = "※つるい未来へつなぐ商工観光経済活性化支援事業補助規則（別表1）に基づく概要です。募集期間や詳細な要件は、村が別に定める公募案内等で必ずご確認ください。"

# 設問が多いブロックだけ、開始ID -> 見出し で小見出しを付ける（Excelには見出しがないため）
GROUP_HEADINGS = {
    "a": {
        "A6-01": "コンセプト・商品・サービス", "A6-09": "営業日・営業時間・開業時期", "A6-14": "許認可・運営体制・提携先",
        "A7-01": "事業を行う場所・建物", "A7-05": "工事・設備①", "A7-10": "工事・設備②", "A7-15": "工事・設備③",
        "A7-20": "その他の費用・スケジュール",
    },
    "b": {
        "B1-01": "創業・沿革", "B1-10": "代表者・後継者", "B1-13": "地域とのつながり",
        "B3-01": "従業員", "B3-04": "営業の状況", "B3-08": "お客様",
        "B4-01": "最大の課題", "B4-06": "対象設備の状況", "B4-08": "お客様の要望",
        "B5-01": "取り組みの概要", "B5-02": "工事・設備①", "B5-08": "工事・設備②", "B5-14": "実施場所・スケジュール", "B5-19": "取り組みの効果",
        "B5-21": "報償費（専門家等への謝礼）", "B5-25": "広告宣伝費",
    },
    "c": {
        "C6-01": "宿のコンセプト・プラン", "C6-08": "営業の概要", "C6-11": "運営人数・雇用",
        "C7-01": "施設の場所・建物", "C7-05": "工事・設備①", "C7-10": "工事・設備②", "C7-15": "工事・設備③",
        "C7-20": "その他の費用", "C7-21": "客室数・定員", "C7-25": "スケジュール・手続き",
    },
}

# 打ち合わせ(2026-09-29)での確認事項に基づく、Excelからの差分。Excel本体は未改訂のためここで吸収する。
# 類型C: 観光推進事業は別表1に加算がないため、上乗せ確認の4問(指定区域・空き家・居住歴・年齢)を削除
REMOVE_IDS = {"c": {"C1-06", "C1-07", "C1-08", "C1-09"}}
# 類型B: 補助対象経費のうち報償費・広告宣伝費の設問を追加（既存IDを動かさないよう末尾の連番で追加）
_YOMIKOMI = ["あり", "依頼中", "まだ"]
def _q(id, label, prompt, example, options=None, dep=None):
    return {"id": id, "label": label, "prompt": prompt, "example": example, "options": ",".join(options) if options else None,
            "deps": [{"ref": dep, "op": "=", "val": "なし", "kind": "対象外"}] if dep else [], "row": 0}
EXTRA_QUESTIONS = {
    "b": {"05_補助金でやりたいこと": [
        _q("B5-21", "報償費（専門家等への謝礼）", "補助金を使って、専門家（コンサルタント・デザイナー・講師など）へ謝礼を支払う予定はありますか。ある場合は、依頼する内容と依頼先を教えてください。ない場合は「なし」（続く設問は記入不要になります）。", "例）商品開発のための料理研究家への監修依頼"),
        _q("B5-22", "報償費の目的・効果", "その依頼で、何を目的とし、今と比べて何がどのくらい良くなりますか。", "例）冷凍商品の味と歩留まりを改善し、新商品を2品開発する", dep="B5-21"),
        _q("B5-23", "報償費の金額", "報償費のおおよその金額（税抜）を教えてください。", "例）30万円", dep="B5-21"),
        _q("B5-24", "報償費の見積書", "報償費の見積書または依頼内容がわかる書類はありますか。（あり／依頼中／まだ） ▼リストから選んでください", "例）あり", _YOMIKOMI, dep="B5-21"),
        _q("B5-25", "広告宣伝費", "補助金を使って、広告・宣伝（チラシ、パンフレット、看板、Web広告、ホームページ制作など）を行う予定はありますか。ある場合は、その内容を教えてください。ない場合は「なし」（続く設問は記入不要になります）。", "例）新商品を紹介するパンフレットの制作と、SNS広告", dep=None),
        _q("B5-26", "広告宣伝費の目的・効果", "その広告・宣伝で、誰に何を伝え、どのような効果を見込みますか。", "例）道内の旅行客に新商品を知ってもらい、通販の注文を月20件増やす", dep="B5-25"),
        _q("B5-27", "広告宣伝費の金額", "広告宣伝費のおおよその金額（税抜）を教えてください。", "例）45万円", dep="B5-25"),
        _q("B5-28", "広告宣伝費の見積書", "広告宣伝費の見積書はありますか。（あり／依頼中／まだ） ▼リストから選んでください", "例）依頼中", _YOMIKOMI, dep="B5-25"),
    ]},
}

SKIP_SHEETS = ("はじめに", "提出書類チェック", "回答一覧（支援者用）")


def clean_prompt(p):
    p = re.sub(r"\s*▼リストから選んでください", "", p or "")
    return p.replace("\n", " ").strip()


def clean_example(e):
    e = (e or "").strip()
    return e[2:] if e.startswith("例）") else e


def clean_note(n):
    if not n:
        return ""
    n = n.replace("あなたが記入しなくてよい設問が自動で灰色（対象外）になります", "あなたが記入しなくてよい設問は自動的に表示されなくなります")
    n = n.replace("は、このシートは飛ばして次へお進みください。", "は、このブロックは自動的にスキップされます。")
    n = n.replace("このシート", "このブロック")
    n = n.replace("指定区域〜年齢要件は補助金の上乗せに当てはまるかを確認する質問です。分からなければ「わからない」を選んでください。", "")
    n = n.replace("次にお進みください。", "")
    return n


def build_form(key, spec):
    t = TYPES[key]
    blocks = []
    all_ids = set()
    for sh in spec["sheets"]:
        if sh["title"] in SKIP_SHEETS:
            continue
        m = re.match(r"^■\s*(\d+)\.\s*(.+)$", sh["head"])
        n, title = int(m.group(1)), m.group(2).strip()
        heads = GROUP_HEADINGS[key]
        groups = []
        qlist = [q for q in sh["qs"] if q["id"] not in REMOVE_IDS.get(key, ())] + EXTRA_QUESTIONS.get(key, {}).get(sh["title"], [])
        for q in qlist:
            all_ids.add(q["id"])
            opts = [o.strip() for o in q["options"].split(",")] if q["options"] else None
            item = {
                "id": q["id"], "label": q["label"], "prompt": clean_prompt(q["prompt"]),
                "example": clean_example(q["example"]),
                "type": "select" if opts else ("text" if n == 0 else "textarea"),
            }
            if opts:
                item["options"] = opts
            if n == 0:
                if "電話" in q["label"]:
                    item["inputmode"] = "tel"
                elif "メール" in q["label"]:
                    item["inputmode"] = "email"
            if item["type"] == "textarea" and len(item["example"]) <= 14:
                item["short"] = True
            deps = [d for d in q["deps"] if d["kind"] == "対象外"]
            if deps:
                item["dependsOn"] = [{"id": d["ref"], "hideIf": "eq" if d["op"] == "=" else "ne", "val": d["val"]} for d in deps]
            if q["id"] in heads or not groups:
                groups.append({"heading": heads.get(q["id"], ""), "questions": []})
            groups[-1]["questions"].append(item)
        blocks.append({
            "id": "b%d" % n, "no": n,
            "eyebrow": "基本情報" if n == 0 else "第%dブロック" % n,
            "title": title, "railLabel": "%d．%s" % (n, title), "sheetLabel": "%d. %s" % (n, title),
            "note": clean_note(sh["note"]) + (" 専門家への謝礼（報償費）や広告宣伝費は、続く設問でお伺いします。" if key == "b" and n == 5 else ""), "groups": groups,
        })
    # 依存先IDの存在確認
    for b in blocks:
        for g in b["groups"]:
            for q in g["questions"]:
                for d in q.get("dependsOn", []):
                    assert d["id"] in all_ids, (q["id"], d)

    docs = []
    doc_states = []
    for sh in spec["sheets"]:
        if sh["title"] == "提出書類チェック":
            docs = sh["docs"]
            doc_states = sh["docStates"]
    total_q = sum(len(g["questions"]) for b in blocks for g in b["groups"])
    has_deps = any("dependsOn" in q for b in blocks for g in b["groups"] for q in g["questions"])

    usage = [
        ["記入のしかた", "難しく考えず、思いつくことを短い言葉や箇条書きでご記入ください。1つの欄には1つのことをお書きください。"],
        ["「なし」", "当てはまらない・特にない場合は「なし」とご記入ください（空欄のままだと未回答として扱われます）。"],
        ["「わからない」", "わからない・まだ決まっていない場合は「わからない」「未定」とご記入のうえ、先へ進んでください。後ほど追加のヒアリングでお伺いします。"],
        ["数字・記入例", "金額・人数・割合などはおおよその値で構いません。記入例は実際の事例をもとにしたイメージです。"],
    ]
    if has_deps:
        usage.append(["設問の切替", "前の設問の回答によって、記入が不要な設問は自動的に表示されなくなります。上から順にお答えください。"])
    usage += [
        ["所要時間", "設問は最大で約%d問（回答内容によって減ります）、2〜3時間ほどです。1日で終わらなくて大丈夫です。数日に分けてご記入ください。" % total_q],
        ["保存", "入力内容はこの端末に自動的に保存されます。途中でやめても続きから再開できます。最後の確認画面でファイル保存・コピーができます。"],
        ["お困りのとき", "入力が難しい場合は、担当の支援者にご相談ください。お電話でのヒアリングも承ります。"],
    ]
    info = []
    if CONFIG.get("returnTo"):
        info.append(["返送先", CONFIG["returnTo"]])
    if CONFIG.get("deadline"):
        info.append(["締切", CONFIG["deadline"]])
    if info:
        info.append(["返送の方法", "最後の「確認・保存」画面で保存したファイル（CSVまたはテキスト）、またはコピーした全文を、返送先へお送りください。メールアドレスや住所などの連絡先は、事業者様へ個別にお知らせします。"])

    return {
        "key": key, "brand": BRAND, "typeName": t["typeName"], "short": t["short"],
        "blocks": blocks, "docs": docs, "docStates": doc_states,
        "docDefault": doc_states.index("まだ") if "まだ" in doc_states else 0,
        "cover": {"usage": usage, "info": info, "summary": t["summary"], "summaryCaveat": SUMMARY_CAVEAT},
        "_total": total_q,
    }


def main():
    head = open(os.path.join(TOOLS, "template_head.html"), encoding="utf-8").read()
    script = open(os.path.join(TOOLS, "template_script.js"), encoding="utf-8").read()
    counts = {}
    for key, t in TYPES.items():
        spec = json.load(open(os.path.join(TOOLS, "spec_%s.json" % key), encoding="utf-8"))
        form = build_form(key, spec)
        counts[key] = (form["_total"], len(form["blocks"]))
        del form["_total"]
        data = json.dumps(form, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
        page = head.replace("__PAGE_TITLE__", "%s ヒアリングフォーム（%s）" % (BRAND, t["short"]))
        page += "<script>\n" + script.replace("__FORM_JSON__", data) + "</script>\n</body>\n</html>\n"
        open(os.path.join(ROOT, t["file"]), "w", encoding="utf-8", newline="\n").write(page)
    open(os.path.join(ROOT, "index.html"), "w", encoding="utf-8", newline="\n").write(build_index(counts))
    print(counts)


def build_index(counts):
    idx = open(os.path.join(TOOLS, "template_index.html"), encoding="utf-8").read()
    cards = ""
    for key, t in TYPES.items():
        total, nblocks = counts[key]
        cards += (
            '    <a class="card" href="%s">\n      <div class="eyebrow">%s</div>\n      <h2>%s</h2>\n'
            '      <p>%s（全%dブロック・最大約%d問）</p>\n      <div class="go">回答をはじめる →</div>\n    </a>\n'
            % (t["file"], html.escape(t["short"]), html.escape(t["typeName"].split("　", 1)[1]), html.escape(t["lead"]), nblocks + 1, total)
        )
    return idx.replace("__CARDS__", cards)


if __name__ == "__main__":
    main()
