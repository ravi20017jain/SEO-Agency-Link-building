# -*- coding: utf-8 -*-
"""
AI-Powered Contact Form Bot
- Claude Vision API: form analyze karta hai
- 2captcha: captcha automatically solve karta hai
- Google Sheets: real-time status update (Dynamic City + Static Niche)
- GitHub Actions: scheduled cloud run
"""
import os
import json
import base64
import time
import logging
import sys
from datetime import datetime

import warnings
warnings.filterwarnings("ignore", category=FutureWarning)
import google.generativeai as genai
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright
import twocaptcha

# ------------------------------------------
#  CONFIGURATION - GitHub Secrets se aata hai
# ------------------------------------------

GEMINI_API_KEY      = os.environ["GEMINI_API_KEY"]       # Google AI Studio se free key
CAPTCHA_API_KEY     = os.environ["CAPTCHA_API_KEY"]
GOOGLE_SHEET_ID     = os.environ["GOOGLE_SHEET_ID"]       # Sheet URL se ID
GOOGLE_CREDS_JSON   = os.environ["GOOGLE_CREDS_JSON"]     # Service account JSON

# Gemini setup - 3.1 Flash Lite (500 req/day free tier)
genai.configure(api_key=GEMINI_API_KEY)
gemini_model = genai.GenerativeModel("gemini-3.1-flash-lite")

FIRST_NAME  = "Salman"
LAST_NAME   = "Khan"
FULL_NAME   = "Salman Khan"
COMPANY     = "LocalTuneUp"
EMAIL       = "salman@localtuneup.com"
PHONE       = "+918889652586"

# LocalTuneup ka message fixed hai (city placeholder ki zarurat nahi)
SUBJECT_TEMPLATE = "Free 14-Day Trial - Google Business Profile Management"

MESSAGE_TEMPLATE = "Hi,\n\nWhen patients search \"dentist near me,\" your Google Business Profile decides whether they pick you or a competitor. Most clinics don't manage it actively.\n\nLocalTuneUp handles it for you - reviews, posts, photos, local ranking - so you rank higher and book more patients.\n\nFree 14-day trial, no card needed. Reply YES and we'll set it up.\n\n- Team LocalTuneUp\nlocaltuneup.com"

PROCESS_LIMIT = None  # None = sab sites ek hi run mein

CONTACT_KEYWORDS = ["contact", "contact-us", "contactus", "contact-form", "get-in-touch",
                    "getintouch", "reach-us", "reachus", "reach-out", "write-to-us",
                    "get-started", "getstarted", "start-here", "enquiry", "enquire",
                    "enquiries", "inquiry", "inquire", "lets-talk", "let-s-talk", "lets-connect",
                    "work-with-us", "hire-us", "hire", "start-project", "start-a-project",
                    "request-quote", "request-a-quote", "get-a-quote", "get-quote", "quote",
                    "book-a-call", "book-call", "book-a-consultation", "book-consultation",
                    "free-consultation", "free-audit", "free-quote", "schedule", "schedule-a-call",
                    "consultation", "talk-to-us", "connect", "connect-with-us", "say-hello",
                    "hello", "support", "help", "get-in-touch-with-us", "contact-sales"]

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
    """Google Sheets connection initialize karo aur city column set karo."""
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

    # Websites sheet check karo ya banao (7 Columns ke saath)
    try:
        ws = sh.worksheet("websites")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet("websites", rows=1000, cols=7)
        ws.update("A1:G1", [["website", "city", "status", "submitted_at", "notes", "fields_filled", "ai_actions"]])

    return ws


def get_all_rows(ws):
    """Saari rows fetch karo."""
    return ws.get_all_records()


def update_sheet_row(ws, row_num, status, notes="", fields_filled="", ai_actions=""):
    """Headers scan karke bina data mix kiye sahi columns update karega."""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    excel_row = row_num + 1
    
    headers = ws.row_values(1)
    try:
        # Dynamic check taaki city column ki wajah se shift na ho data
        status_idx = headers.index("status")
        start_col = chr(65 + status_idx)  # Column letter automatically find karega (e.g. 'C')
        end_col = chr(65 + status_idx + 4)
        ws.update("{}{}:{}{}".format(start_col, excel_row, end_col, excel_row),
                  [[status, now, notes, fields_filled, ai_actions]])
    except ValueError:
        # Fallback agar auto match na ho
        ws.update("C{}:G{}".format(excel_row, excel_row),
                  [[status, now, notes, fields_filled, ai_actions]])
        
    log.info("  [Sheets] Row {} -> {}".format(excel_row, status))


def get_pending_rows(ws):
    """Poori row data return karega taaki loop mein city access ho sake."""
    rows = ws.get_all_records()
    pending = []
    for i, row in enumerate(rows):
        url     = str(row.get("website", "")).strip()
        status  = str(row.get("status", "")).strip().lower()
        if url and status not in ("submitted",):
            pending.append((i + 1, row))   # Full row dict pass ho rahi hai
    return pending

# ------------------------------------------
#  URL HELPERS
# ------------------------------------------

def normalise_url(url):
    url = str(url).strip()
    if not url.startswith("http"):
        url = "https://" + url
    return url.rstrip("/")


def dismiss_cookie_banner(page):
    accept_texts = ["accept all", "accept all cookies", "accept cookies", "accept",
                    "i agree", "agree", "agree & continue", "got it", "allow all",
                    "allow cookies", "allow", "ok", "okay", "i accept", "accept & close",
                    "continue", "i understand", "understand", "consent", "yes, i agree",
                    "close", "dismiss", "no problem", "sounds good"]
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


def find_contact_page(page, base_url):
    current_url = page.url

    try:
        page.wait_for_load_state("networkidle", timeout=6000)
    except Exception:
        pass
    try:
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(1)
        page.evaluate("window.scrollTo(0, 0)")
        time.sleep(0.5)
    except Exception:
        pass

    try:
        links = page.locator("a").all()
        for link in links:
            try:
                href = link.get_attribute("href") or ""
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
                        link.click()
                        page.wait_for_load_state("domcontentloaded", timeout=10000)
                    except Exception:
                        pass
                    time.sleep(0.5)
                    return True
            except Exception:
                pass
    except Exception:
        pass

    if any(kw in current_url.lower() for kw in CONTACT_KEYWORDS):
        log.info("  Already on contact page: {}".format(current_url))
        return True

    for kw in CONTACT_KEYWORDS:
        candidate = "{}/{}".format(base_url, kw)
        try:
            resp = page.goto(candidate, timeout=10000, wait_until="domcontentloaded")
            title = page.title().lower()
            if resp and resp.status < 400 and "404" not in title and "not found" not in title:
                log.info("  Contact page: {}".format(candidate))
                return True
        except Exception:
            pass
    return False

# ------------------------------------------
#  CAPTCHA SOLVER (2captcha)
# ------------------------------------------

def solve_captcha(page, website):
    solver = twocaptcha.TwoCaptcha(CAPTCHA_API_KEY)

    try:
        frame = page.locator('iframe[src*="recaptcha"]').first
        if frame.is_visible(timeout=1000):
            src = frame.get_attribute("src") or ""
            sitekey = ""
            for part in src.split("&"):
                if "k=" in part:
                    sitekey = part.split("k=")[1].split("&")[0]
                    break
            if not sitekey:
                div = page.locator('.g-recaptcha').first
                sitekey = div.get_attribute("data-sitekey") or ""

            if sitekey:
                log.info("  [CAPTCHA] reCAPTCHA detected, solving via 2captcha...")
                result = solver.recaptcha(sitekey=sitekey, url=website)
                token = result["code"]
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
    except Exception as e:
        log.debug("  reCAPTCHA solve attempt: {}".format(e))

    try:
        frame = page.locator('iframe[src*="hcaptcha.com"]').first
        if frame.is_visible(timeout=1000):
            div = page.locator('.h-captcha').first
            sitekey = div.get_attribute("data-sitekey") or ""
            if sitekey:
                log.info("  [CAPTCHA] hCaptcha detected, solving...")
                result = solver.hcaptcha(sitekey=sitekey, url=website)
                token = result["code"]
                page.evaluate("""(token) => {
                    document.querySelector('[name="h-captcha-response"]').value = token;
                    document.querySelector('[name="g-recaptcha-response"]') &&
                        (document.querySelector('[name="g-recaptcha-response"]').value = token);
                }""", token)
                log.info("  [CAPTCHA] hCaptcha solved!")
                return True
    except Exception as e:
        log.debug("  hCaptcha solve attempt: {}".format(e))

    try:
        div = page.locator('.cf-turnstile').first
        if div.is_visible(timeout=1000):
            sitekey = div.get_attribute("data-sitekey") or ""
            if sitekey:
                log.info("  [CAPTCHA] Cloudflare Turnstile detected, solving...")
                result = solver.turnstile(sitekey=sitekey, url=website)
                token = result["code"]
                page.evaluate("""(token) => {
                    document.querySelector('[name="cf-turnstile-response"]').value = token;
                }""", token)
                log.info("  [CAPTCHA] Turnstile solved!")
                return True
    except Exception as e:
        log.debug("  Turnstile solve attempt: {}".format(e))

    return False

# ------------------------------------------
#  AI FORM ANALYSIS (Claude Vision)
# ------------------------------------------

def get_page_html(page):
    def grab(frame):
        try:
            return frame.evaluate("""() => {
                const els = document.querySelectorAll(
                    'input, textarea, button, select, label, form'
                );
                return Array.from(els).map(el => el.outerHTML).join('\\n');
            }""")
        except Exception:
            return ""
    parts = []
    try:
        parts.append(grab(page))
    except Exception:
        pass
    try:
        for fr in page.frames:
            if fr == page.main_frame:
                continue
            h = grab(fr)
            if h and ("input" in h or "form" in h):
                parts.append(h)
    except Exception:
        pass
    return "\n".join(p for p in parts if p)[:18000]


def ask_claude(page, website, subject, message):
    """Claude ko dynamic elements (subject/message) ke saath call karein."""
    try:
        page.wait_for_load_state("networkidle", timeout=8000)
    except Exception:
        pass
    for _ in range(4):
        try:
            page.wait_for_selector("input, textarea, select", timeout=3000)
            break
        except Exception:
            try:
                page.evaluate("window.scrollTo(0, document.body.scrollHeight/2)")
            except Exception:
                pass
            time.sleep(1)

    page_html = get_page_html(page)
    if len(page_html) > 50000:
        page_html = page_html[:50000]

    prompt = """You are a web automation expert. Fill this contact form on: {website}

Form HTML:
{html}

Details to fill:
- Full Name: {full_name}
- First Name: {first_name}
- Last Name: {last_name}
- Company: {company}
- Email: {email}
- Phone: {phone}
- Subject/Title: {subject}
- Message (copy EXACTLY, keep all line breaks):
{message}

IMPORTANT: Fill the message field with the COMPLETE text above. Do not truncate or summarize.

Return ONLY a JSON array of actions. Each action:
  "action": "fill" | "check" | "click" | "select"
  "selector": CSS selector (prefer name/id/type attributes)
  "value": value to use

Rules:
- Only include fields that exist in the HTML
- IMPORTANT: Only fill an ACTUAL CONTACT/ENQUIRY form. Do NOT fill search boxes (input name="s", role="search"), login forms (name="log"/"pwd"/"username"/"password"), or newsletter-only email boxes. If there is no real contact form, return an empty array [].
- HUMAN-CHECK QUESTIONS: If the form has a simple text question to prove you're human (e.g. "which is bigger, 2 or 8?", "what is 3+4?", "type the word yes", "what color is the sky?"), SOLVE it and fill the answer in that field. Answer with the simplest correct value (e.g. "8", "7", "yes", "blue").
- For checkboxes (terms/agree/consent/privacy) use "check".
- REQUIRED checkbox groups (marked with * like "Services", "Interested in", "Budget"): you MUST select at least one option, else the form won't submit. Prefer an SEO / digital-marketing / "Google SEO" / "search" related option if available; otherwise pick the first reasonable option. Use "check" for it.
- For the submit button use "click" - include it LAST. Pick the form's actual submit button (type="submit" inside the contact form), not a search or login button.
- COMMON FIELDS: fill phone/mobile with the phone, website/url with our site, subject/topic with a short subject like "Partnership enquiry". For dropdowns/select (subject, service, "how did you hear", country), use "select" and pick the most relevant option (e.g. SEO/marketing/general enquiry); if unsure pick the first non-empty option.
- SKIP appointment-booking fields: do NOT fill Date, Time, date pickers, calendar fields, age, or appointment-slot fields. Leave them empty. Fill only name, email, phone, company, and message. Put all the outreach text in the message/comment field.
- Message field: use the FULL message text provided
- Return ONLY JSON, no markdown, no explanation""".format(
        website=website,
        html=page_html,
        full_name=FULL_NAME,
        first_name=FIRST_NAME,
        last_name=LAST_NAME,
        company=COMPANY,
        email=EMAIL,
        phone=PHONE,
        subject=subject,
        message=message
    )

    raw = None
    waits = [5, 20, 40, 60]
    for attempt in range(4):
        try:
            resp = gemini_model.generate_content(prompt)
            raw = resp.text.strip()
            break
        except Exception as e:
            msg = str(e)
            if any(code in msg for code in ("429", "500", "503", "overloaded", "quota", "timeout", "rate")):
                w = waits[attempt]
                log.warning("  [AI] Gemini busy ({}), retry in {}s...".format(msg[:40], w))
                time.sleep(w)
                continue
            raise
    if raw is None:
        raise Exception("Gemini API failed after 4 retries")

    if "```" in raw:
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())

# ------------------------------------------
#  EXECUTE ACTIONS
# ------------------------------------------

def scroll_to(page, locator):
    try:
        locator.scroll_into_view_if_needed(timeout=2000)
        time.sleep(0.2)
    except Exception:
        pass


def execute_actions(page, actions):
    filled = []
    submitted = False

    for action in actions:
        act      = action.get("action", "").lower()
        selector = action.get("selector", "")
        value    = action.get("value", "")

        if not selector:
            continue

        try:
            locator = page.locator(selector).first
            scroll_to(page, locator)

            if act == "fill":
                if locator.is_visible(timeout=1000):
                    locator.fill(value)
                    log.info("  [OK] fill: {}".format(selector[:50]))
                    filled.append(selector[:30])

            elif act == "check":
                if locator.is_visible(timeout=1000) and not locator.is_checked():
                    locator.check()
                    log.info("  [OK] check: {}".format(selector[:50]))

            elif act == "select":
                if locator.is_visible(timeout=1000):
                    locator.select_option(value)
                    log.info("  [OK] select: {}".format(selector[:50]))

            elif act == "click":
                if locator.is_visible(timeout=1000):
                    url_before = page.url
                    try:
                        locator.scroll_into_view_if_needed(timeout=2000)
                    except Exception:
                        pass
                    try:
                        locator.click(timeout=5000)
                    except Exception:
                        try:
                            locator.evaluate("el => el.click()")
                        except Exception:
                            pass
                    success_words = ["thank you", "thanks", "message sent", "we'll be in touch",
                                     "we have received", "submitted successfully", "your message",
                                     "successfully sent", "received your", "get back to you",
                                     "contacting us", "be in touch", "form submitted", "sent successfully",
                                     "we'll get back", "message has been sent", "successfully submitted",
                                     "your submission", "appreciate you", "has been received",
                                     "will respond", "soon as possible", "form was submitted",
                                     "message received", "we received", "submission received",
                                     "we'll reach out", "reach out to you", "talk to you soon",
                                     "we will contact", "request received", "got your message",
                                     "ticket has been", "enquiry received", "inquiry received",
                                     "we'll respond", "in touch shortly", "received and"]
                    confirmed = False
                    captcha_done = False
                    retried_click = False
                    for i in range(20):
                        time.sleep(3)
                        if not captcha_done:
                            try:
                                if solve_captcha(page, page.url):
                                    captcha_done = True
                                    try:
                                        locator.click(timeout=2000)
                                    except Exception:
                                        pass
                                    time.sleep(2)
                            except Exception:
                                pass
                        page_text = ""
                        try:
                            page_text = page.inner_text("body", timeout=3000).lower()
                        except Exception:
                            pass
                        url_changed = page.url != url_before
                        if any(w in page_text for w in success_words) or url_changed:
                            confirmed = True
                            break
                        if i == 3 and not retried_click:
                            retried_click = True
                            try:
                                if locator.is_visible(timeout=1000):
                                    locator.evaluate("el => el.click()")
                                    time.sleep(1)
                                    locator.press("Enter")
                            except Exception:
                                pass
                    if confirmed:
                        submitted = True
                        log.info("  [OK] submit confirmed: {}".format(selector[:50]))
                    else:
                        log.warning("  [??] clicked but NO confirmation: {}".format(selector[:50]))

        except Exception as e:
            log.warning("  [--] {}: {} -> {}".format(act, selector[:50], e))

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
        log.info("No pending sites. Done!")
        return

    to_process = pending[:PROCESS_LIMIT]

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
            ]
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
            website_raw = row_data.get("website", "")
            website = normalise_url(website_raw)
            
            # Sheet se unique row ki city read karega, khali hone par fallback lagayega
            city = str(row_data.get("city", "Somers Point")).strip() or "Somers Point"
            
            # Dynamic text allocation
            current_subject = SUBJECT_TEMPLATE.format(city=city)
            current_message = MESSAGE_TEMPLATE.format(city=city)

            log.info("\nOpening: {} [City: {}]".format(website, city))

            try:
                pg.goto(website, timeout=30000, wait_until="domcontentloaded")
                time.sleep(2)
                dismiss_cookie_banner(pg)

                contact_found = find_contact_page(pg, website)
                if not contact_found:
                    log.info("  No separate contact page - checking current page for form")

                time.sleep(1)
                dismiss_cookie_banner(pg)

                try:
                    pg.wait_for_load_state("networkidle", timeout=6000)
                except Exception:
                    pass
                try:
                    pg.wait_for_selector("form, input[type='email'], input[type='text'], textarea",
                                         timeout=8000)
                except Exception:
                    pass
                try:
                    pg.evaluate("window.scrollTo(0, document.body.scrollHeight/2)")
                    time.sleep(1.5)
                except Exception:
                    pass

                solve_captcha(pg, website)

                # Claude processes dynamic templates
                try:
                    actions = ask_claude(pg, website, current_subject, current_message)
                    log.info("  [AI] {} actions".format(len(actions)))
                except Exception as e:
                    log.error("  [AI] Error: {}".format(e))
                    update_sheet_row(ws, row_idx, "error", "AI error: {}".format(str(e)[:80]))
                    continue

                filled, submitted = execute_actions(pg, actions)
                time.sleep(1)

                try:
                    import re, os
                    safe_name = re.sub(r'[^a-zA-Z0-9]', '_', website)[:50]
                    os.makedirs("screenshots/before_submit", exist_ok=True)
                    screenshot_path = "screenshots/before_submit/{}.png".format(safe_name)
                    pg.screenshot(path=screenshot_path, full_page=False)
                    log.info("  [Screenshot] Before submit saved: {}".format(screenshot_path))
                except Exception as e:
                    log.warning("  [Screenshot] Failed: {}".format(e))

                if submitted:
                    status = "submitted"
                elif not filled:
                    status = "no_form_found"
                else:
                    status = "filled_not_submitted"

                try:
                    import re, os
                    try:
                        pg.wait_for_load_state("networkidle", timeout=8000)
                    except Exception:
                        pass
                    time.sleep(2)
                    safe_name = re.sub(r'[^a-zA-Z0-9]', '_', website)[:50]
                    os.makedirs("screenshots/after_submit", exist_ok=True)
                    screenshot_path = "screenshots/after_submit/{}.png".format(safe_name)
                    pg.screenshot(path=screenshot_path, full_page=False)
                    log.info("  [Screenshot] After submit saved: {}".format(screenshot_path))
                except Exception as e:
                    log.warning("  [Screenshot] Failed: {}".format(e))

                if submitted:
                    note_text = "OK"
                elif not filled:
                    note_text = "No form on page (manual not needed)"
                else:
                    note_text = "Submit failed - try manually"
                    
                update_sheet_row(
                    ws, row_idx, status,
                    notes=note_text,
                    fields_filled=", ".join(filled),
                    ai_actions=str(len(actions))
                )

                log.info("  Status: {}".format(status))
                time.sleep(1)

            except Exception as e:
                log.error("  ERROR: {}".format(e))
                update_sheet_row(ws, row_idx, "error", str(e)[:100])

        browser.close()

    log.info("\nRun complete!")


if __name__ == "__main__":
    main()
