# -*- coding: utf-8 -*-
"""
AI-Powered Contact Form Bot (Zevahit - SEO Agency Outreach)
- Claude Vision API / Gemini: form analyze karta hai (gemini-3.1-flash-lite)
- 2captcha: captcha automatically solve karta hai
- Google Sheets: real-time status update 
- GitHub Actions: scheduled cloud run with strict anti-hang timeouts
"""
import os
import json
import base64
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

GEMINI_API_KEY      = os.environ.get("GEMINI_API_KEY", "")
CAPTCHA_API_KEY     = os.environ.get("CAPTCHA_API_KEY", "")
GOOGLE_SHEET_ID     = os.environ.get("GOOGLE_SHEET_ID", "")
GOOGLE_CREDS_JSON   = os.environ.get("GOOGLE_CREDS_JSON", "{}")

genai.configure(api_key=GEMINI_API_KEY)
# Gemini Model updated to gemini-3.1-flash-lite as requested
gemini_model = genai.GenerativeModel("gemini-3.1-flash-lite")

FIRST_NAME  = "Ray"
LAST_NAME   = ""
FULL_NAME   = "Ray"
COMPANY     = "Zevahit"
EMAIL       = "sales@zevahit.com"
PHONE       = "+18005550199" # Agar zaroorat ho toh apna real phone number dalein

SUBJECT_TEMPLATE = "Quick question regarding white-label link inventory"

# Option 1 template with Zevahit signature
MESSAGE_TEMPLATE = """Hi team,

{intro}Quick question—are you currently taking on new SEO clients that require genuine, high-traffic editorial placements rather than the usual spammed-out guest post networks?

We handle direct-to-editorial outreach strictly for agencies. For instance, we secure placements on live, highly vetted platforms like The Boss Magazine (https://thebossmagazine.com) - real brands with massive organic traffic, strict editorial standards, and zero footprint.

If you're looking to scale your link-building fulfillment with inventory that easily passes manual client approval, just reply to this with your primary client niche. I'll send over 3 live examples with pricing.

Warm Regards,
Ray
Zevahit.com
Client reviews: https://clutch.co/profile/zevahit#reviews"""

PROCESS_LIMIT = None 

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
        ws.update("A1:G1", [["website", "city", "status", "submitted_at", "notes", "fields_filled", "ai_actions"]])

    return ws

def update_sheet_row(ws, row_num, status, notes="", fields_filled="", ai_actions=""):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    excel_row = row_num + 1
    headers = ws.row_values(1)
    try:
        status_idx = headers.index("status")
        start_col = chr(65 + status_idx) 
        end_col = chr(65 + status_idx + 4)
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
        url     = str(row.get("website", "")).strip()
        status  = str(row.get("status", "")).strip().lower()
        if url and status not in ("submitted",):
            pending.append((i + 1, row))   
    return pending

# ------------------------------------------
#  URL HELPERS & BROWSER ACTIONS
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
        page.wait_for_load_state("networkidle", timeout=5000)
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
                        link.click(timeout=5000)
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
#  CAPTCHA SOLVER
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
        pass
    return False

# ------------------------------------------
#  AI PERSONALIZATION & FORM ANALYSIS
# ------------------------------------------

def get_page_text(page):
    try:
        txt = page.evaluate(
            """() => {
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
            }"""
        )
        return (txt or "")[:4000]
    except Exception:
        return ""

def generate_personalized_line(page, website):
    """
    Specifically looks for SEO Agency context (niches, case studies, awards).
    """
    site_text = get_page_text(page)
    if len(site_text.strip()) < 40:
        return "" 

    prompt = """You are writing the FIRST sentence of a cold outreach message to an SEO or Digital Marketing Agency.

Here is text scraped from their website ({website}):
---
{site_text}
---

Write ONE short, specific, genuine opening line (max 22 words) that shows we actually looked at their site.
Rules:
- Mention something specific about their agency: a specific industry/niche they serve (e.g. SaaS, Lawyers), an impressive case study, an award, or their specific approach to digital marketing.
- Sound human and warm, NOT salesy or generic. No "I hope this finds you well".
- Do NOT mention anything about selling them links yet.
- End with a comma or dash so the next sentence flows naturally.
- Return ONLY the line itself. No quotes, no explanation.

Example good output: Saw you've been doing some impressive SEO work for e-commerce brands recently -"""

    prompt = prompt.format(website=website, site_text=site_text)
    raw = None

    waits = [15, 30, 60]
    for attempt in range(3):
        try:
            resp = gemini_model.generate_content(prompt)
            raw = (resp.text or "").strip()
            break
        except Exception as e:
            msg = str(e)
            if any(c in msg for c in ("429", "quota", "rate", "exceeded", "503", "overloaded")):
                w = waits[attempt]
                log.warning("  [Personalize] rate limit, retry in {}s...".format(w))
                time.sleep(w)
                continue
            log.warning("  [Personalize] failed: {}".format(msg[:60]))
            return ""
    if not raw:
        return ""

    line = raw.replace("```", "").strip().strip('"').strip("'").strip()
    line = line.split("\n")[0].strip()
    if len(line) > 200 or len(line.split()) > 30:
        return ""
    
    log.info("  [Personalize] {}".format(line[:80]))
    return line

def get_page_html(page):
    try:
        return page.evaluate("""() => {
            const els = document.querySelectorAll('input, textarea, button, select, label, form');
            return Array.from(els).map(el => el.outerHTML).join('\\n');
        }""")
    except Exception:
        return ""

def ask_claude(page, website, subject, message_template, homepage_text=""):
    """
    MERGED single Gemini call:
    - homepage_text (homepage se capture kiya) se personalized opening line banata hai
    - Us line ko message_template ke {intro} me daalta hai
    - Phir contact page ke form HTML se form-fill actions banata hai
    Sab kuch EK hi Gemini call me -> quota aadhi.
    Personalization homepage text se (strong), form-fill contact page se (sahi).

    Returns: (actions_list, intro_line)
    """
    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except Exception:
        pass

    page_html = get_page_html(page)[:16000]
    # Personalization homepage text se. Agar homepage text na mile to current page se fallback.
    site_text = (homepage_text or "").strip()
    if len(site_text) < 40:
        site_text = get_page_text(page)
    site_text = site_text[:2500]

    prompt = """You are a web automation expert AND a cold-outreach copywriter. Do BOTH tasks below for the site: {website}

=== TASK 1: Write a personalized opening line ===
Here is visible text scraped from their HOMEPAGE (this describes what the business actually does):
---
{site_text}
---
Write ONE short, specific, genuine opening line (max 22 words) that shows we actually looked at their site.
Rules for the line:
- Mention something REAL and specific about what this B2B brand does: their product category, who they serve, or a specific value prop you can see in the text.
- Sound human and sharp, NOT salesy or generic. No "I hope this finds you well".
- Do NOT mention SEO, links, AI search, rankings, or any offer.
- End with a comma or dash so the next sentence flows naturally.
- If the text is too thin to say anything specific, use an empty string "" for the line.

=== TASK 2: Fill the contact form ===
Take this MESSAGE TEMPLATE and replace the literal token {{INTRO}} with your line from Task 1 (followed by two newlines). If your line is empty, just remove the {{INTRO}} token.

MESSAGE TEMPLATE:
{message_template}

Now fill this contact form. Form HTML:
{html}

Details to fill:
- Full Name: {full_name}
- First Name: {first_name}
- Last Name: {last_name}
- Company: {company}
- Email: {email}
- Phone: {phone}
- Subject/Title: {subject}
- Message: use the FINAL message (template with {{INTRO}} replaced). Copy it EXACTLY, keep all line breaks.

=== OUTPUT FORMAT ===
Return ONLY a valid JSON object (no markdown, no extra text) with EXACTLY these two keys:
{{
  "intro_line": "the line you wrote in Task 1 (or empty string)",
  "actions": [
    {{"action": "fill"|"check"|"click"|"select", "selector": "CSS selector", "value": "value to use"}}
  ]
}}
""".format(
        website=website, site_text=site_text, html=page_html,
        message_template=message_template.replace("{intro}", "{INTRO}"),
        full_name=FULL_NAME, first_name=FIRST_NAME, last_name=LAST_NAME,
        company=COMPANY, email=EMAIL, phone=PHONE, subject=subject
    )

    raw = None
    waits = [15, 30, 60]
    for attempt in range(3):
        try:
            resp = gemini_model.generate_content(prompt)
            raw = resp.text.strip()
            break
        except Exception as e:
            msg = str(e)
            if any(c in msg for c in ("429", "quota", "rate", "exceeded", "503", "overloaded")):
                w = waits[attempt]
                log.warning("  [AI] rate limit, retry in {}s... ({})".format(w, msg[:40]))
                time.sleep(w)
                continue
            raise Exception(f"Gemini API generation failed: {e}")
    if raw is None:
        raise Exception("Gemini API failed after retries (likely daily quota exhausted)")

    # Clean markdown fences
    if "```" in raw:
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    data = json.loads(raw)

    # Backward-safe: agar kabhi AI sirf array bhej de to use actions maan lo
    if isinstance(data, list):
        return data, ""

    actions = data.get("actions", [])
    intro_line = (data.get("intro_line") or "").strip()
    if intro_line:
        log.info("  [Personalize] {}".format(intro_line[:80]))
    return actions, intro_line



# ------------------------------------------
#  EXECUTE ACTIONS
# ------------------------------------------

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
            try:
                locator.scroll_into_view_if_needed(timeout=2000)
            except Exception:
                pass

            if act == "fill":
                try:
                    locator.fill(value, timeout=3000)
                    filled.append(selector[:30])
                except Exception:
                    # Fallback JS fill
                    page.evaluate("""(el, val) => {
                        el.value = val;
                        el.dispatchEvent(new Event('input', {bubbles:true}));
                        el.dispatchEvent(new Event('change', {bubbles:true}));
                    }""", locator, value)
                    filled.append(selector[:30])

            elif act == "click":
                url_before = page.url
                try:
                    locator.click(timeout=5000)
                except Exception:
                    page.evaluate("el => el.click()", locator)
                
                # Strict anti-hang verification
                success_words = ["thank you", "message sent", "we'll be in touch", "submitted successfully"]
                for _ in range(5): 
                    time.sleep(3)
                    try:
                        page_text = page.inner_text("body", timeout=3000).lower()
                        if any(w in page_text for w in success_words) or (page.url != url_before):
                            submitted = True
                            break
                    except Exception:
                        pass
                    
                if submitted:
                    log.info("  [OK] submit confirmed.")
                else:
                    log.warning("  [??] clicked but NO confirmation.")

        except Exception as e:
            log.warning("  [--] {}: {} -> {}".format(act, selector[:50], str(e)[:30]))

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
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-gpu"])
        context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
        
        pg = context.new_page()
        pg.set_default_timeout(20000)
        pg.set_default_navigation_timeout(30000)
        
        pg.route("**/*", lambda route: route.abort() if route.request.resource_type in ("image", "media") else route.continue_())

        for row_idx, row_data in to_process:
            website_raw = row_data.get("website", "")
            website = normalise_url(website_raw)
            current_subject = SUBJECT_TEMPLATE

            log.info("\nOpening: {}".format(website))

            try:
                pg.goto(website, timeout=30000, wait_until="domcontentloaded")
                time.sleep(2)
                dismiss_cookie_banner(pg)

                # Homepage ka business text ABHI capture karo (contact page jaane se PEHLE).
                # Yahi text personalization ke liye strong hai - business kya karta hai
                # ye homepage pe hota hai, contact page pe nahi.
                homepage_text = get_page_text(pg)

                find_contact_page(pg, website)
                time.sleep(1)
                dismiss_cookie_banner(pg)
                solve_captcha(pg, website)

                # MERGED single Gemini call: personalize line (homepage text se)
                # + form actions (contact page HTML se) ek saath. Sirf 1 call = quota aadhi.
                try:
                    actions, intro_line = ask_claude(
                        pg, website, current_subject, MESSAGE_TEMPLATE, homepage_text
                    )
                except Exception as e:
                    update_sheet_row(ws, row_idx, "error", "AI error: {}".format(str(e)[:80]))
                    continue

                filled, submitted = execute_actions(pg, actions)
                
                if submitted:
                    status, note_text = "submitted", "OK"
                elif not filled:
                    status, note_text = "no_form_found", "No form on page"
                else:
                    status, note_text = "filled_not_submitted", "Submit failed"
                    
                update_sheet_row(
                    ws, row_idx, status,
                    notes=note_text,
                    fields_filled=", ".join(filled),
                    ai_actions=str(len(actions))
                )

                # ---- Gemini 429 Rate Limit fix ----
                # Har site pe 2 Gemini call jaati hain (personalize + form-fill).
                # Free tier 15 req/min ko cross na karne ke liye har site ke baad wait.
                log.info("  Waiting 15s to avoid Gemini rate limit...")
                time.sleep(15)

            except Exception as e:
                log.error("  ERROR: {}".format(str(e)[:100]))
                
                try:
                    os.makedirs("screenshots/errors", exist_ok=True)
                    pg.screenshot(path=f"screenshots/errors/crash_{row_idx}.png")
                except:
                    pass
                
                update_sheet_row(ws, row_idx, "error", str(e)[:100])
                # Error aane par bhi thoda zyada wait (kabhi-kabhi 429 hi crash karta hai)
                time.sleep(10)

        browser.close()
    log.info("\nRun complete!")

if __name__ == "__main__":
    main()
