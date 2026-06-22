# -*- coding: utf-8 -*-
"""
AI-Powered Contact Form Bot (Zevahit - SEO Agency Outreach)
- Gemini: form analyze karta hai (gemini-3.1-flash-lite)
- 2captcha: captcha automatically solve karta hai
- Google Sheets: real-time status update
- GitHub Actions: scheduled cloud run with strict anti-hang timeouts

FIXES vs previous version:
  FIX 1: JS fallback now uses page.query_selector() (element handle) not locator
  FIX 2: Submit confirmation: 20+ phrases, 24s wait, CSS class check, form-hide check
  FIX 3: Click handler no longer JS-clicks when page is navigating (execution context fix)
  FIX 4: AI prompt now warns against dynamic/numeric IDs, gives better GravityForms selectors
  FIX 5: Iframe form detection + frame context switching
  FIX 6: Label-based fallback fill for unlabelled inputs
  FIX 7: WPForms / CF7 / GravityForms / Elementor specific success class detection
  FIX 8: contact page nav wait 0.5s -> 2s + networkidle wait
  FIX 9: form load wait before get_page_html (wait_for_selector + networkidle)
  FIX 10: debug screenshot + HTML dump before AI call
  FIX 11: Gemini prompt gets generic visible input fallback selectors rule
"""
import os
import json
import time
import logging
import sys
from datetime import datetime

import google.generativeai as genai
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright
import twocaptcha

# ------------------------------------------
#  CONFIGURATION
# ------------------------------------------

GEMINI_API_KEY    = os.environ.get("GEMINI_API_KEY", "")
CAPTCHA_API_KEY   = os.environ.get("CAPTCHA_API_KEY", "")
GOOGLE_SHEET_ID   = os.environ.get("GOOGLE_SHEET_ID", "")
GOOGLE_CREDS_JSON = os.environ.get("GOOGLE_CREDS_JSON", "{}")

genai.configure(api_key=GEMINI_API_KEY)
gemini_model = genai.GenerativeModel("gemini-3.1-flash-lite")

FIRST_NAME = "Ray"
LAST_NAME  = ""
FULL_NAME  = "Ray"
COMPANY    = "Zevahit"
EMAIL      = "sales@zevahit.com"
PHONE      = "+18005550199"

SUBJECT_TEMPLATE = "Quick question regarding white-label link inventory"

MESSAGE_TEMPLATE = """Hi team,

{intro}Quick question—are you currently taking on new SEO clients that require genuine, high-traffic editorial placements rather than the usual spammed-out guest post networks?

We handle direct-to-editorial outreach strictly for agencies. For instance, we secure placements on live, highly vetted platforms like The Boss Magazine (https://thebossmagazine.com) - real brands with massive organic traffic, strict editorial standards, and zero footprint.

If you're looking to scale your link-building fulfillment with inventory that easily passes manual client approval, just reply to this with your primary client niche. I'll send over 3 live examples with pricing.

Warm Regards,
Ray
Zevahit.com
Client reviews: https://clutch.co/profile/zevahit#reviews"""

PROCESS_LIMIT = None

# FIX 10: Debug mode — set False in production to skip HTML/screenshot dumps
DEBUG_MODE = os.environ.get("DEBUG_MODE", "false").lower() == "true"

CONTACT_KEYWORDS = [
    "contact", "contact-us", "contactus", "contact-form", "get-in-touch",
    "getintouch", "reach-us", "reachus", "reach-out", "write-to-us",
    "get-started", "getstarted", "start-here", "enquiry", "enquire",
    "enquiries", "inquiry", "inquire", "lets-talk", "let-s-talk", "lets-connect",
    "work-with-us", "hire-us", "hire", "start-project", "start-a-project",
    "request-quote", "request-a-quote", "get-a-quote", "get-quote", "quote",
    "book-a-call", "book-call", "book-a-consultation", "book-consultation",
    "free-consultation", "free-audit", "free-quote", "schedule", "schedule-a-call",
    "consultation", "talk-to-us", "connect", "connect-with-us", "say-hello",
    "hello", "support", "help", "get-in-touch-with-us", "contact-sales"
]

# ------------------------------------------
#  SUCCESS DETECTION CONSTANTS
# ------------------------------------------

SUCCESS_PHRASES = [
    "thank you", "thanks", "thank-you",
    "message sent", "message received", "message has been sent",
    "message has been received", "your message",
    "we'll be in touch", "we will be in touch",
    "we'll get back", "we will get back", "get back to you",
    "submitted successfully", "submission received", "form submitted",
    "successfully submitted", "successfully sent",
    "received your", "we got your", "got your message",
    "appreciate you", "appreciate your", "appreciate your message",
    "contact you soon", "be in touch soon", "in touch shortly",
    "sent successfully", "inquiry received", "enquiry received",
    "we have received", "has been received",
]

SUCCESS_CSS_CLASSES = [
    "wpforms-confirmation", "wpforms-confirmation-container",
    "wpcf7-mail-sent-ok",
    "gform_confirmation_wrapper", "gform_confirmation_message",
    "elementor-message-success",
    "nf-response-msg",
    "frm_message",
    "submitted-message",
    "alert-success", "form-success", "message-success",
    "success-message", "sent-message", "contact-success",
    "form-sent", "is-success", "success-box",
    "confirmation-message", "form-confirmation",
    "submission-success", "form_success",
]

# ------------------------------------------
#  LOGGING
# ------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
log = logging.getLogger(__name__)

# ------------------------------------------
#  GOOGLE SHEETS SETUP
# ------------------------------------------

def init_sheets():
    creds_dict = json.loads(GOOGLE_CREDS_JSON)
    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=[
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    try:
        ws = sh.worksheet("websites")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet("websites", rows=1000, cols=7)
        ws.update("A1:G1", [["website", "city", "status", "submitted_at",
                              "notes", "fields_filled", "ai_actions"]])
    return ws


def update_sheet_row(ws, row_num, status, notes="", fields_filled="", ai_actions=""):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    excel_row = row_num + 1
    headers = ws.row_values(1)
    try:
        status_idx = headers.index("status")
        start_col = chr(65 + status_idx)
        end_col   = chr(65 + status_idx + 4)
        ws.update("{}{}:{}{}".format(start_col, excel_row, end_col, excel_row),
                  [[status, now, notes, fields_filled, ai_actions]])
    except ValueError:
        ws.update("C{}:G{}".format(excel_row, excel_row),
                  [[status, now, notes, fields_filled, ai_actions]])
    log.info("  [Sheets] Row {} -> {}".format(excel_row, status))


def get_pending_rows(ws):
    rows = ws.get_all_records()
    pending = []
    for i, row in enumerate(rows):
        url    = str(row.get("website", "")).strip()
        status = str(row.get("status", "")).strip().lower()
        if url and status not in ("submitted",):
            pending.append((i + 1, row))
    return pending

# ------------------------------------------
#  URL HELPERS
# ------------------------------------------

def normalise_url(url):
    url = str(url).strip()
    if not url.startswith("http"):
        url = "https://" + url
    return url.rstrip("/")

# ------------------------------------------
#  COOKIE BANNER DISMISS
# ------------------------------------------

def dismiss_cookie_banner(page):
    accept_texts = [
        "accept all", "accept all cookies", "accept cookies", "accept",
        "i agree", "agree", "agree & continue", "got it", "allow all",
        "allow cookies", "allow", "ok", "okay", "i accept", "accept & close",
        "continue", "i understand", "understand", "consent", "yes, i agree",
        "close", "dismiss", "no problem", "sounds good"
    ]
    selectors = ("button, a, input[type='button'], input[type='submit'], "
                 "[role='button'], div[onclick], span[onclick], div, span")
    try:
        buttons = page.locator(selectors).all()
        for btn in buttons[:80]:
            try:
                txt = (btn.inner_text(timeout=300) or "").strip().lower()
            except Exception:
                continue
            if not txt or len(txt) > 20:
                continue
            if any(t == txt for t in accept_texts):
                try:
                    if btn.is_visible(timeout=500):
                        btn.click(timeout=2000)
                        log.info("  [Cookie] dismissed: {}".format(txt[:25]))
                        time.sleep(1)
                        return True
                except Exception:
                    pass
    except Exception:
        pass
    return False

# ------------------------------------------
#  CONTACT PAGE FINDER
# ------------------------------------------

def find_contact_page(page, base_url):
    current_url = page.url
    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except Exception:
        pass

    try:
        links = page.locator("a").all()
        for link in links:
            try:
                href      = link.get_attribute("href") or ""
                link_text = ""
                try:
                    link_text = (link.inner_text(timeout=500) or "").lower()
                except Exception:
                    pass
                if any(kw in href.lower() for kw in CONTACT_KEYWORDS) or \
                   any(kw.replace("-", " ") in link_text for kw in CONTACT_KEYWORDS):
                    if any(kw in current_url.lower() for kw in CONTACT_KEYWORDS):
                        log.info("  Already on contact page: {}".format(current_url))
                        return True
                    log.info("  Contact link: {}".format(href))
                    try:
                        link.click(timeout=5000)
                        page.wait_for_load_state("domcontentloaded", timeout=10000)
                    except Exception:
                        pass

                    # FIX 8: Extended wait after contact page navigation
                    time.sleep(2)
                    try:
                        page.wait_for_load_state("networkidle", timeout=8000)
                    except Exception:
                        pass

                    return True
            except Exception:
                pass
    except Exception:
        pass

    if any(kw in current_url.lower() for kw in CONTACT_KEYWORDS):
        return True

    for kw in CONTACT_KEYWORDS:
        candidate = "{}/{}".format(base_url, kw)
        try:
            resp  = page.goto(candidate, timeout=10000, wait_until="domcontentloaded")
            title = page.title().lower()
            if resp and resp.status < 400 and "404" not in title and "not found" not in title:
                log.info("  Contact page: {}".format(candidate))
                # FIX 8: Extended wait here too
                time.sleep(2)
                try:
                    page.wait_for_load_state("networkidle", timeout=8000)
                except Exception:
                    pass
                return True
        except Exception:
            pass
    return False

# ------------------------------------------
#  CAPTCHA SOLVER
# ------------------------------------------

def solve_captcha(page, website):
    solver = twocaptcha.TwoCaptcha(CAPTCHA_API_KEY)
    try:
        frame = page.locator('iframe[src*="recaptcha"]').first
        if frame.is_visible(timeout=1000):
            src     = frame.get_attribute("src") or ""
            sitekey = ""
            for part in src.split("&"):
                if "k=" in part:
                    sitekey = part.split("k=")[1].split("&")[0]
                    break
            if not sitekey:
                div     = page.locator('.g-recaptcha').first
                sitekey = div.get_attribute("data-sitekey") or ""

            if sitekey:
                log.info("  [CAPTCHA] reCAPTCHA detected, solving via 2captcha...")
                result = solver.recaptcha(sitekey=sitekey, url=website)
                token  = result["code"]
                page.evaluate("""(token) => {
                    document.getElementById('g-recaptcha-response').innerHTML = token;
                    if (typeof ___grecaptcha_cfg !== 'undefined') {
                        Object.entries(___grecaptcha_cfg.clients).forEach(([key, client]) => {
                            Object.entries(client).forEach(([k, v]) => {
                                if (typeof v === 'object' && v !== null && 'callback' in v) {
                                    try { v.callback(token); } catch(e) {}
                                }
                            });
                        });
                    }
                }""", token)
                log.info("  [CAPTCHA] reCAPTCHA solved!")
                return True
    except Exception:
        pass
    return False

# ------------------------------------------
#  SUCCESS DETECTION
# ------------------------------------------

def check_submission_success(page, url_before):
    try:
        current_url = page.url

        if current_url != url_before:
            return True

        if "#" in current_url and current_url.split("#")[0] == url_before.split("#")[0]:
            frag = current_url.split("#")[-1].lower()
            if any(w in frag for w in ["thank", "success", "confirm", "sent", "done"]):
                return True

        body_text = ""
        try:
            body_text = page.inner_text("body", timeout=3000).lower()
        except Exception:
            pass

        if any(phrase in body_text for phrase in SUCCESS_PHRASES):
            return True

        for css_class in SUCCESS_CSS_CLASSES:
            try:
                el = page.locator(".{}".format(css_class)).first
                if el.is_visible(timeout=300):
                    log.info("  [OK] success class detected: .{}".format(css_class))
                    return True
            except Exception:
                pass

        try:
            forms = page.locator("form").all()
            if forms:
                all_hidden = all(
                    not f.is_visible(timeout=300) for f in forms
                )
                if all_hidden:
                    return True
        except Exception:
            pass

    except Exception:
        pass

    return False

# ------------------------------------------
#  IFRAME FORM DETECTION
# ------------------------------------------

def try_iframe_form(page, actions):
    iframe_selectors = [
        'iframe[src*="jotform"]',
        'iframe[src*="typeform"]',
        'iframe[src*="cognitoforms"]',
        'iframe[src*="123formbuilder"]',
        'iframe[src*="formstack"]',
        'iframe[src*="wufoo"]',
        'iframe[src*="paperform"]',
    ]
    for sel in iframe_selectors:
        try:
            iframe_el = page.locator(sel).first
            if not iframe_el.is_visible(timeout=500):
                continue
            frame = iframe_el.content_frame()
            if not frame:
                continue
            log.info("  [Iframe] Detected embedded form: {}".format(sel))
            filled = []
            submitted = False
            for action in actions:
                act      = action.get("action", "").lower()
                selector = action.get("selector", "")
                value    = action.get("value", "")
                if not selector:
                    continue
                try:
                    if act == "fill":
                        frame.locator(selector).first.fill(value, timeout=3000)
                        filled.append(selector[:30])
                    elif act == "click":
                        url_before = page.url
                        frame.locator(selector).first.click(timeout=5000)
                        for _ in range(8):
                            time.sleep(3)
                            if check_submission_success(page, url_before):
                                submitted = True
                                break
                        if submitted:
                            log.info("  [OK] iframe submit confirmed.")
                        else:
                            log.warning("  [??] iframe clicked but NO confirmation.")
                except Exception as e:
                    log.warning("  [Iframe][--] {}: {} -> {}".format(
                        act, selector[:40], str(e)[:30]))
            return filled, submitted
        except Exception:
            pass
    return [], False

# ------------------------------------------
#  LABEL-BASED FALLBACK FILL
# ------------------------------------------

def label_based_fill(page, field_type, value):
    label_map = {
        "name":    ["name", "full name", "your name", "contact name", "first name"],
        "email":   ["email", "e-mail", "email address", "your email"],
        "phone":   ["phone", "telephone", "mobile", "phone number", "tel"],
        "company": ["company", "organisation", "organization", "business", "agency"],
        "subject": ["subject", "topic", "regarding", "title"],
        "message": ["message", "comments", "enquiry", "inquiry", "how can we help",
                    "tell us", "your message", "details"],
    }
    keywords = label_map.get(field_type, [])
    if not keywords:
        return False

    try:
        labels = page.locator("label").all()
        for label in labels:
            try:
                label_text = (label.inner_text(timeout=300) or "").strip().lower()
                if not any(kw in label_text for kw in keywords):
                    continue
                for_attr = label.get_attribute("for") or ""
                if for_attr:
                    el = page.query_selector("#{}".format(for_attr))
                    if el:
                        page.locator("#{}".format(for_attr)).fill(value, timeout=2000)
                        return True
            except Exception:
                pass
    except Exception:
        pass
    return False

# ------------------------------------------
#  SMART JS FILL
# ------------------------------------------

def smart_js_fill(page, selector, value):
    try:
        el = page.query_selector(selector)
        if not el:
            return False
        page.evaluate("""(el, val) => {
            el.focus();
            el.value = val;
            el.dispatchEvent(new Event('focus',  {bubbles: true}));
            el.dispatchEvent(new Event('input',  {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
            el.dispatchEvent(new Event('blur',   {bubbles: true}));
        }""", el, value)
        filled_val = page.evaluate("el => el.value", el)
        return bool(filled_val)
    except Exception:
        return False

# ------------------------------------------
#  PAGE TEXT & HTML EXTRACTION
# ------------------------------------------

def get_page_text(page):
    try:
        txt = page.evaluate("""() => {
            const isVisible = (el) => {
                const s = window.getComputedStyle(el);
                return s && s.display !== 'none' && s.visibility !== 'hidden';
            };
            let out = '';
            document.querySelectorAll('h1,h2,h3,h4,p,li,span,a,.tagline,title').forEach(el => {
                if (el.children.length === 0 && isVisible(el) && el.innerText) {
                    const t = el.innerText.trim();
                    if (t.length > 2) out += t + ' | ';
                }
            });
            return out;
        }""")
        return (txt or "")[:4000]
    except Exception:
        return ""


def get_page_html(page):
    # FIX 9: Wait for form elements to be present before extracting HTML
    try:
        page.wait_for_selector(
            "form, input[type='text'], input[type='email'], textarea",
            timeout=8000
        )
    except Exception:
        log.warning("  [HTML] No form/input found within timeout — extracting anyway")

    try:
        page.wait_for_load_state("networkidle", timeout=6000)
    except Exception:
        pass

    try:
        return page.evaluate("""() => {
            const els = document.querySelectorAll(
                'input, textarea, button, select, label, form, [class*="form"], [id*="form"]'
            );
            return Array.from(els).map(el => {
                let extra = '';
                if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
                    if (el.id) {
                        const lbl = document.querySelector('label[for="' + el.id + '"]');
                        if (lbl) extra = ' data-label="' + lbl.innerText.trim() + '"';
                    }
                    if (!extra && el.getAttribute('aria-label')) {
                        extra = ' data-label="' + el.getAttribute('aria-label') + '"';
                    }
                    if (!extra && el.placeholder) {
                        extra = ' data-label="' + el.placeholder + '"';
                    }
                }
                return el.outerHTML.replace('>', extra + '>');
            }).join('\\n');
        }""")[:16000]
    except Exception:
        return ""

# ------------------------------------------
#  DEBUG DUMP  (FIX 10)
# ------------------------------------------

def debug_dump(page, row_idx):
    """Save screenshot + HTML for debugging no_form_found cases."""
    if not DEBUG_MODE:
        return
    try:
        os.makedirs("screenshots/debug", exist_ok=True)
        pg_url = page.url
        pg_html = get_page_html(page)
        shot_path = "screenshots/debug/before_ai_{}.png".format(row_idx)
        html_path = "screenshots/debug/html_{}.txt".format(row_idx)
        page.screenshot(path=shot_path)
        with open(html_path, "w", encoding="utf-8") as f:
            f.write("URL: {}\n\n".format(pg_url))
            f.write(pg_html)
        log.info("  [Debug] Saved: {} | {}".format(shot_path, html_path))
    except Exception as e:
        log.warning("  [Debug] dump failed: {}".format(str(e)[:60]))

# ------------------------------------------
#  AI: MERGED PERSONALIZE + FORM FILL
# ------------------------------------------

def ask_claude(page, website, subject, message_template, homepage_text=""):
    page_html  = get_page_html(page)
    site_text  = (homepage_text or "").strip()
    if len(site_text) < 40:
        site_text = get_page_text(page)
    site_text = site_text[:2500]

    prompt = """You are a web automation expert AND a cold-outreach copywriter. Do BOTH tasks below for: {website}

=== TASK 1: Personalized opening line ===
Homepage text:
---
{site_text}
---
Write ONE short, specific, genuine opening line (max 22 words).
Rules:
- Mention something REAL and specific about what this agency does.
- Sound human and sharp. No "I hope this finds you well".
- Do NOT mention SEO, links, AI search, rankings, or any offer.
- End with a comma or dash so the next sentence flows naturally.
- If text is too thin, use empty string "" for the line.

=== TASK 2: Fill the contact form ===
MESSAGE TEMPLATE (replace {{INTRO}} with your line from Task 1 + two newlines, or remove if empty):
{message_template}

FORM HTML (inputs have data-label attributes showing their purpose):
{html}

Fill details:
- Full Name / Name: {full_name}
- First Name: {first_name}
- Last Name: {last_name}
- Company: {company}
- Email: {email}
- Phone: {phone}
- Subject/Title: {subject}
- Message: final message with {{INTRO}} replaced. Copy EXACTLY with all line breaks.

=== CRITICAL SELECTOR RULES ===
1. NEVER use pure numeric IDs like #1761201320 or #1088408602. These are dynamic and break.
2. For Gravity Forms: use .gfield input[type="text"], .gfield input[type="email"],
   .gfield textarea, .gfield input[type="tel"] — NOT #input_X_Y IDs.
3. For WPForms: use #wpforms-X-field_Y or input[name="wpforms[fields][N]"].
4. For Contact Form 7: use input[name="your-name"], input[name="your-email"], etc.
5. Prefer name= attributes over id= when id looks numeric or random.
6. For submit button: prefer input[type="submit"], button[type="submit"],
   button:contains("Send"), button:contains("Submit") over dynamic IDs.
7. Skip any field whose data-label contains: date, time, day, month, year, appointment, calendar.
8. Skip checkbox/radio/select fields unless they look like consent or service choice.
9. If you see NO proper form fields, return an empty actions array [].
10. If labeled selectors fail, use these visible generic fallbacks IN ORDER:
    - Name: input[type="text"]:visible:first-of-type
    - Email: input[type="email"]:visible
    - Phone: input[type="tel"]:visible
    - Message: textarea:visible
    - Submit: button[type="submit"]:visible, input[type="submit"]:visible

=== OUTPUT FORMAT ===
Return ONLY valid JSON (no markdown, no extra text):
{{
  "intro_line": "the line from Task 1 (or empty string)",
  "actions": [
    {{"action": "fill"|"check"|"click"|"select", "selector": "CSS selector", "value": "value"}}
  ]
}}""".format(
        website=website, site_text=site_text, html=page_html,
        message_template=message_template.replace("{intro}", "{INTRO}"),
        full_name=FULL_NAME, first_name=FIRST_NAME, last_name=LAST_NAME,
        company=COMPANY, email=EMAIL, phone=PHONE, subject=subject
    )

    raw   = None
    waits = [15, 30, 60]
    for attempt in range(3):
        try:
            resp = gemini_model.generate_content(prompt)
            raw  = resp.text.strip()
            break
        except Exception as e:
            msg = str(e)
            if any(c in msg for c in ("429", "quota", "rate", "exceeded", "503", "overloaded")):
                w = waits[attempt]
                log.warning("  [AI] rate limit, retry in {}s...".format(w))
                time.sleep(w)
                continue
            raise Exception("Gemini API generation failed: {}".format(e))

    if raw is None:
        raise Exception("Gemini API failed after retries (daily quota exhausted)")

    if "```" in raw:
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    data = json.loads(raw)
    if isinstance(data, list):
        return data, ""

    actions    = data.get("actions", [])
    intro_line = (data.get("intro_line") or "").strip()
    if intro_line:
        log.info("  [Personalize] {}".format(intro_line[:80]))
    return actions, intro_line

# ------------------------------------------
#  EXECUTE ACTIONS
# ------------------------------------------

def execute_actions(page, actions):
    filled    = []
    submitted = False

    for action in actions:
        act      = action.get("action", "").lower()
        selector = action.get("selector", "")
        value    = action.get("value", "")

        if not selector:
            continue

        if act == "fill":
            try:
                locator = page.locator(selector).first
                try:
                    locator.scroll_into_view_if_needed(timeout=2000)
                except Exception:
                    pass

                filled_ok = False
                try:
                    locator.fill(value, timeout=3000)
                    filled_ok = True
                except Exception:
                    pass

                if not filled_ok:
                    filled_ok = smart_js_fill(page, selector, value)

                if not filled_ok:
                    try:
                        locator.click(timeout=2000)
                        locator.type(value, delay=30)
                        filled_ok = True
                    except Exception:
                        pass

                if filled_ok:
                    filled.append(selector[:30])
                else:
                    log.warning("  [--] fill: {} -> all methods failed".format(selector[:50]))

            except Exception as e:
                log.warning("  [--] fill: {} -> {}".format(selector[:50], str(e)[:40]))

        elif act == "click":
            url_before = page.url
            click_ok   = False

            try:
                locator = page.locator(selector).first
                try:
                    locator.scroll_into_view_if_needed(timeout=2000)
                except Exception:
                    pass
                locator.click(timeout=5000)
                click_ok = True
            except Exception as e1:
                log.warning("  [--] click: {} -> {}".format(selector[:50], str(e1)[:40]))

            if not click_ok:
                try:
                    if page.url == url_before:
                        el = page.query_selector(selector)
                        if el:
                            page.evaluate("el => el.click()", el)
                            click_ok = True
                except Exception as e2:
                    log.warning("  [--] js-click: {} -> {}".format(selector[:50], str(e2)[:40]))

            if not click_ok:
                try:
                    locator = page.locator(selector).first
                    locator.focus(timeout=2000)
                    page.keyboard.press("Enter")
                    click_ok = True
                except Exception:
                    pass

            for poll in range(8):
                time.sleep(3)
                if check_submission_success(page, url_before):
                    submitted = True
                    log.info("  [OK] submit confirmed (poll {}).".format(poll + 1))
                    break

            if not submitted:
                log.warning("  [??] clicked but NO confirmation.")

        elif act == "check":
            try:
                locator = page.locator(selector).first
                locator.check(timeout=3000)
            except Exception as e:
                log.warning("  [--] check: {} -> {}".format(selector[:50], str(e)[:30]))

        elif act == "select":
            try:
                locator = page.locator(selector).first
                locator.select_option(value, timeout=3000)
            except Exception as e:
                log.warning("  [--] select: {} -> {}".format(selector[:50], str(e)[:30]))

        if submitted:
            break

    return filled, submitted

# ------------------------------------------
#  MAIN
# ------------------------------------------

def main():
    log.info("Connecting to Google Sheets...")
    ws = init_sheets()

    pending = get_pending_rows(ws)
    log.info("Pending sites: {}".format(len(pending)))
    if not pending:
        return

    to_process = pending[:PROCESS_LIMIT]

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-gpu"]
        )
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            )
        )
        pg = context.new_page()
        pg.set_default_timeout(20000)
        pg.set_default_navigation_timeout(30000)
        pg.route("**/*", lambda route: route.abort()
                 if route.request.resource_type in ("image", "media")
                 else route.continue_())

        for row_idx, row_data in to_process:
            website_raw     = row_data.get("website", "")
            website         = normalise_url(website_raw)
            current_subject = SUBJECT_TEMPLATE

            log.info("\nOpening: {}".format(website))

            try:
                pg.goto(website, timeout=30000, wait_until="domcontentloaded")
                time.sleep(2)
                dismiss_cookie_banner(pg)

                homepage_text = get_page_text(pg)

                find_contact_page(pg, website)
                time.sleep(1)
                dismiss_cookie_banner(pg)
                solve_captcha(pg, website)

                # FIX 10: Debug dump before AI call
                debug_dump(pg, row_idx)

                try:
                    actions, intro_line = ask_claude(
                        pg, website, current_subject, MESSAGE_TEMPLATE, homepage_text
                    )
                except Exception as e:
                    update_sheet_row(ws, row_idx, "error",
                                     "AI error: {}".format(str(e)[:80]))
                    time.sleep(10)
                    continue

                filled, submitted = execute_actions(pg, actions)

                if not filled and not submitted:
                    log.info("  [Iframe] Trying iframe form fallback...")
                    filled, submitted = try_iframe_form(pg, actions)

                if submitted:
                    status, note_text = "submitted", "OK"
                elif not filled:
                    status, note_text = "no_form_found", "No fillable form found"
                else:
                    status, note_text = "filled_not_submitted", "Submit failed"

                update_sheet_row(
                    ws, row_idx, status,
                    notes=note_text,
                    fields_filled=", ".join(filled),
                    ai_actions=str(len(actions))
                )

                log.info("  Waiting 15s to avoid Gemini rate limit...")
                time.sleep(15)

            except Exception as e:
                log.error("  ERROR: {}".format(str(e)[:100]))
                try:
                    os.makedirs("screenshots/errors", exist_ok=True)
                    pg.screenshot(path="screenshots/errors/crash_{}.png".format(row_idx))
                except Exception:
                    pass
                update_sheet_row(ws, row_idx, "error", str(e)[:100])
                time.sleep(10)

        browser.close()
    log.info("\nRun complete!")


if __name__ == "__main__":
    main()
