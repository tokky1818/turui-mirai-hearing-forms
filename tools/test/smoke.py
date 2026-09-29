# -*- coding: utf-8 -*-
"""ブラウザ(Chrome)で類型A/B/Cを通しで操作するスモークテスト。 python tools/test/smoke.py"""
import os, sys, json
from playwright.sync_api import sync_playwright
sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def run(pw, fname, first_choice, other_choice_hidden_prefix):
    b = pw.chromium.launch(channel="chrome", headless=True)
    ctx = b.new_context(viewport={"width": 1100, "height": 800}, accept_downloads=True)
    page = ctx.new_page()
    errs = []
    page.on("pageerror", lambda e: errs.append(str(e)))
    page.on("console", lambda m: errs.append(m.text) if m.type == "error" and "fonts" not in m.text and "cdnjs" not in m.text and "ERR_" not in m.text else None)
    page.goto("file:///" + ROOT.replace(os.sep, "/") + "/" + fname)
    page.on("dialog", lambda d: d.accept())
    page.click("#btnStart")
    # block 0: fill all
    for i, el in enumerate(page.query_selector_all("[data-qid]")):
        el.fill("テスト%d" % i)
    page.click("#btnNext")
    # block 1: choose first select
    sel = page.query_selector("select[data-qid]")
    qid = sel.get_attribute("data-qid") if sel else None
    print(fname, "first select", qid)
    # before selection dependents hidden
    hidden_before = page.eval_on_selector_all(".qcard[hidden]", "els=>els.length")
    if qid: page.select_option("select[data-qid='%s']" % qid, first_choice)
    hidden_after = page.eval_on_selector_all(".qcard[hidden]", "els=>els.length")
    print("  hidden cards before/after choosing:", hidden_before, hidden_after)
    # fill visible questions in this block
    for el in page.query_selector_all(".qcard:not([hidden]) [data-qid]"):
        if el.get_attribute("data-qid") == qid: continue
        tag = el.evaluate("e=>e.tagName")
        if tag == "SELECT":
            el.select_option(index=1)
        else:
            el.fill("回答")
    page.click("#btnNext")
    step_label = page.inner_text(".footer-progress-text span:first-child")
    print("  after next ->", step_label.strip())
    # go through all remaining steps, fill visible, count NA skips
    labels = []
    for _ in range(30):
        lab = page.inner_text(".footer-progress-text span:first-child").split("最終保存")[0].strip()
        labels.append(lab)
        if lab == "確認・保存": break
        for el in page.query_selector_all(".qcard:not([hidden]) [data-qid]"):
            tag = el.evaluate("e=>e.tagName")
            if tag == "SELECT": el.select_option(index=1)
            else:
                # 'なし' first for chain questions
                el.fill("なし" if el.get_attribute("data-qid")[-2:] in ("15","10") else "回答")
        page.click("#btnNext")
    print("  visited:", labels)
    txt = page.input_value("#copyArea")
    print("  review stats:", page.inner_text(".review-stats").replace("\n", " "))
    page.click("details.save-section summary")
    with page.expect_download() as d:
        page.click("#btnSaveCsv")
    path = d.value.path()
    csv = open(path, encoding="utf-8-sig").read().splitlines()
    print("  csv lines:", len(csv), "|", csv[0])
    st = {}
    for l in csv[2:]:
        s = l.rsplit(",", 1)[-1]
        st[s] = st.get(s, 0) + 1
    print("  status counts:", st)
    print("  errors:", errs)
    # mobile viewport render check
    page.set_viewport_size({"width": 390, "height": 800})
    page.screenshot(path=os.path.join(os.environ.get("SHOT_DIR", "."), fname.replace(".html", "_review_m.png")))
    b.close()

with sync_playwright() as pw:
    run(pw, "ruikei-a.html", "新規開業", None)
    run(pw, "ruikei-b.html", None, None)
    run(pw, "ruikei-c.html", "営んでいる", None)
