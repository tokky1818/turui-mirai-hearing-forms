# -*- coding: utf-8 -*-
"""第1回レビュー対応の回帰テスト: 進捗カウンタ更新 / 選択ボタン / 未記入確認 / 再開リンク / 横スクロールなし"""
import os, sys
from playwright.sync_api import sync_playwright
sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))).replace(os.sep, "/")
with sync_playwright() as pw:
    b = pw.chromium.launch(channel="chrome", headless=True)
    ctx = b.new_context(viewport={"width": 390, "height": 800})
    p = ctx.new_page(); errs = []
    p.on("pageerror", lambda e: errs.append(str(e)))
    p.goto("file:///" + ROOT + "/ruikei-a.html")
    print("cover overflow px:", p.evaluate("document.documentElement.scrollWidth-innerWidth"))
    p.click("#btnStart")
    p.fill("[data-qid='00-01']", "テスト事業者")
    print("progress:", p.inner_text("#progressText"))
    p.on("dialog", lambda d: (print("dialog:", d.message.split(chr(10))[0]), d.accept()))
    p.click("#btnNext")
    print("step:", p.inner_text(".footer-label"))
    print("radio chips:", p.eval_on_selector_all(".opt-chip", "e=>e.length"))
    p.click(".opt-chip[data-val='新規開業']")
    print("A1-01 selected ->", p.evaluate("state.answers['A1-01']"), "| hidden A1-02:", p.evaluate("document.getElementById('card_A1-02').hidden"))
    p.click("#btnNext"); print("after next:", p.inner_text(".footer-label"), "scrollY", p.evaluate("scrollY"))
    p.reload(); p.click("#btnStart") if p.query_selector("#btnStart") else None
    print("resume step:", p.inner_text(".footer-label"))
    print("errors:", errs)
    b.close()
