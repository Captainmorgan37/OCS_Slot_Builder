import json
import os
import re
import PySimpleGUI as sg
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

# -----------------------------
# URLs
# -----------------------------
OCS_LOGIN = "https://online-coordination.com/frontend/#/login"
OCS_ADD_FLIGHTS = "https://online-coordination.com/frontend/#/addFlightsGaba"

# -----------------------------
# Cred storage helpers
# -----------------------------
CRED_FILE = "ocs_creds.json"
LOG_FILE = "ocs_autofill_debug.log"


def log_debug(msg: str):
    """Mirror debug output to stdout and append to a local log file."""

    try:
        print(msg)
    finally:
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(f"{msg}\n")
        except Exception:
            # Debug logging should never break the flow
            pass

def load_saved_creds():
    if os.path.exists(CRED_FILE):
        try:
            with open(CRED_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_creds(username: str, password: str, passphrase: str):
    try:
        with open(CRED_FILE, "w") as f:
            json.dump(
                {
                    "username": username,
                    "password": password,
                    "passphrase": passphrase,
                },
                f,
            )
    except Exception:
        # Fail silently; app can still run without saving
        pass

# -----------------------------
# Helpers
# -----------------------------
def parse_feas_json(raw: str):
    try:
        data = json.loads(raw)
        return data if isinstance(data, dict) else None
    except Exception:
        return None

def popup_passphrase_chars(request_text: str):
    """
    Old helper (no longer used) kept for reference.
    """
    layout = [
        [sg.Text("OCS needs 2 passphrase characters.")],
        [sg.Text(request_text, text_color="yellow")],
        [sg.Text("Character #1:"), sg.Input(key="c1", size=(5,1))],
        [sg.Text("Character #2:"), sg.Input(key="c2", size=(5,1))],
        [sg.Button("Continue"), sg.Button("Cancel")]
    ]
    win = sg.Window("OCS Passphrase", layout, modal=True)
    event, values = win.read()
    win.close()
    if event == "Continue":
        return values["c1"], values["c2"]
    return None, None

def _dump_react_select_debug(section, row_label):
    try:
        html = section.evaluate("el => el.innerHTML")
        log_debug(f"[DEBUG {row_label}] section HTML:\n{html}\n")
    except Exception as e:
        log_debug(f"[DEBUG {row_label}] Unable to dump HTML: {e}")

    try:
        controls = section.locator(".ocs__control").all()
        log_debug(
            f"[DEBUG {row_label}] .ocs__control count: {len(controls)}"
        )
        for idx, ctl in enumerate(controls):
            cls = ctl.evaluate("el => el.className")
            log_debug(f"  control[{idx}] class: {cls}")
    except Exception as e:
        log_debug(f"[DEBUG {row_label}] Control introspection failed: {e}")

    try:
        inputs = section.locator(
            "xpath=.//input[contains(@id,'react-select') and contains(@id,'-input')]"
        ).all()
        log_debug(f"[DEBUG {row_label}] react-select inputs: {len(inputs)}")
        for idx, inp in enumerate(inputs):
            rid = inp.get_attribute("id")
            log_debug(f"  input[{idx}] id: {rid}")
    except Exception as e:
        log_debug(f"[DEBUG {row_label}] Input introspection failed: {e}")


def select_row_react_select(page, row_label, value_text, timeout=6000):
    """
    Select a react-select dropdown within the first editable flight row that sits
    immediately after a label containing ``row_label`` (e.g., 'A/P', 'STC').

    This uses explicit XPath selectors (not chained CSS) so Playwright never tries
    to interpret the XPath as CSS, avoiding the "Unexpected token '/'" failure the
    inspector reported.
    """

    row_xpath = "//div[contains(@class,'ocs-transaction-flight-fields') and contains(@class,'first-flight')]"
    target_section = None

    try:
        # 1) Grab the first row container explicitly via XPath to avoid selector mixing
        row = page.locator(f"xpath={row_xpath}").first
        row.wait_for(state="visible", timeout=timeout)

        # 2) Identify the header section that holds the label (the top row with column headings)
        header = row.locator(
            "xpath=.//div[contains(@class,'trans-fields-headings') or contains(@class,'trans-fields-headings-local-time')]"
        ).first
        header.wait_for(state="visible", timeout=timeout)

        label_section = header.locator(
            f"xpath=.//section[.//label[normalize-space()='{row_label}']]"
        ).first
        label_section.wait_for(state="visible", timeout=timeout)

        # 3) Determine the column index of the label inside the header row
        column_index = label_section.evaluate(
            "el => Array.from(el.parentElement.children).indexOf(el)"
        )
        if column_index < 0:
            raise Exception(f"Could not compute column index for label '{row_label}'")

        # 4) Grab the next d-flex row that holds the actual inputs (sections align by position)
        field_row = header.locator(
            "xpath=following::div[contains(@class,'d-flex')][1]"
        ).first
        field_row.wait_for(state="visible", timeout=timeout)

        target_section = field_row.locator(
            f"xpath=.//section[contains(@class,'ocs-field-item')][{column_index + 1}]"
        ).first
        target_section.wait_for(state="visible", timeout=timeout)

        # Prefer an explicit react-select input if present (IDs tend to change
        # across deployments), then fall back to the generic control wrapper.
        control = target_section.locator(
            "xpath=.//input[contains(@id,'react-select') and contains(@id,'-input')]"
        ).first

        if control.count() == 0:
            control = target_section.locator(".ocs__control").first

        # Even when the control is present, the library sometimes holds it in a
        # collapsed/hidden state until a click happens on the parent section.
        # Avoid waiting on visibility until after we've nudged the container.
        control.wait_for(state="attached", timeout=timeout)
        target_section.scroll_into_view_if_needed()
        target_section.click(force=True)
        page.wait_for_timeout(150)

        # If we still don't see a visible control, try clicking the control
        # itself before enforcing visibility to coax React-Select to render.
        if not control.is_visible():
            control.click(force=True)
            page.wait_for_timeout(200)

        if not control.is_visible():
            log_debug(
                f"[DEBUG {row_label}] control still hidden before wait_for – writing {LOG_FILE}"
            )
            _dump_react_select_debug(target_section, row_label)
        control.wait_for(state="visible", timeout=timeout)

        # 5) Open dropdown (double click fallback)
        for _ in range(2):
            control.click(force=True)
            page.wait_for_timeout(200)
            menu = page.locator(".ocs__menu")
            if menu.is_visible():
                break
        else:
            raise Exception(f"Dropdown never opened for label '{row_label}'")

        # 6) Pick the option
        option = page.get_by_role("option", name=value_text)
        option.wait_for(timeout=timeout)
        option.click(force=True)

        page.wait_for_timeout(150)
        return True

    except Exception as e:
        log_debug(f"[ERROR selecting '{value_text}' in row '{row_label}']: {e}")
        if target_section is not None:
            try:
                _dump_react_select_debug(target_section, row_label)
            except Exception:
                pass

        # Fallback: locate the label anywhere inside the first-flight row and
        # pick the nearest react-select control (useful when the header/column
        # alignment shifts or STC renders slightly differently).
        try:
            row = page.locator(
                "div.ocs-transaction-flight-fields.first-flight"
            ).first
            row.wait_for(state="visible", timeout=timeout)

            label = row.locator(
                f"xpath=.//*[normalize-space()='{row_label}']"
            ).first
            label.wait_for(state="visible", timeout=timeout)

            # Prefer the same section, otherwise step to the next field item.
            section_locator = label.locator(
                "xpath=ancestor::section[contains(@class,'ocs-field-item')][1]"
            ).first
            if section_locator.count() == 0:
                section_locator = label.locator(
                    "xpath=following::section[contains(@class,'ocs-field-item')][1]"
                ).first

            control = section_locator.locator(
                "xpath=.//input[contains(@id,'react-select') and contains(@id,'-input')]"
            ).first
            if control.count() == 0:
                control = section_locator.locator(".ocs__control").first

            control.wait_for(state="visible", timeout=timeout)
            control.scroll_into_view_if_needed()

            for _ in range(2):
                control.click(force=True)
                page.wait_for_timeout(200)
                menu = page.locator(".ocs__menu")
                if menu.is_visible():
                    break
            else:
                raise Exception(f"Fallback dropdown never opened for '{row_label}'")

            option = page.get_by_role("option", name=value_text)
            option.wait_for(timeout=timeout)
            option.click(force=True)
            page.wait_for_timeout(150)
            return True
        except Exception as fallback_error:
            log_debug(
                f"[FALLBACK ERROR selecting '{value_text}' in row '{row_label}']: {fallback_error}"
            )
            return False






def select_ap_dropdown(page, airport_code):
    try:
        # Locate the A/P label
        label = page.get_by_text("A/P", exact=True)
        label.wait_for()

        # Locate the react-select control immediately after it
        control = label.locator("xpath=following::div[contains(@class,'ocs__control')][1]")
        control.wait_for()
        control.scroll_into_view_if_needed()
        page.wait_for_timeout(150)

        # Click twice to ensure dropdown opens
        for _ in range(2):
            control.click(force=True)
            page.wait_for_timeout(200)
            menu = page.locator(".ocs__menu")
            if menu.is_visible():
                break
        else:
            raise Exception("Dropdown failed to open")

        # Select the airport
        option = page.get_by_role("option", name=airport_code)
        option.wait_for()
        option.click()

        page.wait_for_timeout(150)
        return True

    except Exception as e:
        print(f"[AP ERROR] {e}")
        return False


        
def open_react_select(page, index):
    control = page.locator(".ocs__control").nth(index)

    # scroll into view just in case
    control.scroll_into_view_if_needed()

    # First click
    control.click(force=True)
    page.wait_for_timeout(150)

    # If menu didn't open, click again
    menu = page.locator(".ocs__menu")
    if not menu.is_visible():
        control.click(force=True)
        page.wait_for_timeout(150)


def select_react_select(page, index, value_text, timeout=6000):
    """Open a react-select control by index and pick the desired option."""

    control = page.locator(".ocs__control").nth(index)
    control.wait_for(timeout=timeout)
    control.scroll_into_view_if_needed()

    # Open the dropdown (double click as a fallback)
    for _ in range(2):
        control.click(force=True)
        page.wait_for_timeout(200)
        menu = page.locator(".ocs__menu")
        if menu.is_visible():
            break
    else:
        raise Exception(f"Dropdown at index {index} did not open")

    option = page.get_by_role("option", name=value_text)
    option.wait_for(timeout=timeout)
    option.click()

    page.wait_for_timeout(150)


def select_dropdown_value(page, label_text, value_text, timeout=5000):
    """
    Select a value from an OCS react-select dropdown based on the left-hand label.
    Works for A/P, STC, ParkLoc and all other react-select controls in the slot table.
    """

    try:
        # 1. Find the <td> with the label
        label_td = page.locator(f"//td[normalize-space()='{label_text}']").first
        label_td.wait_for(timeout=timeout)

        # 2. Move to the dropdown's cell (the next TD)
        dropdown_cell = label_td.locator("xpath=following-sibling::td[1]")

        # 3. Get the react-select control
        control = dropdown_cell.locator(".ocs__control").first
        control.wait_for(state="visible", timeout=timeout)

        # 4. Click to open
        control.click(force=True)

        # 5. Wait for menu to appear
        menu = page.locator("//div[contains(@class,'ocs__menu')]")
        menu.wait_for(state="visible", timeout=timeout)

        # 6. Select the value
        option = menu.get_by_role("option", name=value_text)
        option.wait_for(timeout=timeout)
        option.click()

        # 7. Confirm dropdown closed
        page.wait_for_timeout(300)

        return True

    except Exception as e:
        print(f"Dropdown selection error for {label_text}: {e}")
        return False


def fill_text_cell(page, label_text, value):
    """Fill a text field that sits immediately to the right of ``label_text``.

    We scope to the first editable flight row using explicit XPath selectors to
    avoid CSS parsing of the XPath snippet (which previously threw an inspector
    error). This keeps adjacent input targeting reliable even when nested markup
    or whitespace varies.
    """

    # Prefer stable IDs/`name` attributes when the markup provides them; fall back
    # to the positional label-based lookup if anything goes wrong.
    selector_map = {
        "A/C Reg": "#aircraftRegistration",
        "Date": "input[name='startDate']",
        "Seats": "#numSeats",
        "A/C Type": "#aircraftType",
        "Time": "#clearedTimeDep",
        "Dest": "#destinationStation",
        "Orig": "#originStation",
    }

    direct_selector = selector_map.get(label_text)
    if direct_selector:
        try:
            fill_field_by_selector(page, direct_selector, value, timeout=12000)
            return
        except Exception as e:
            print(
                f"Direct selector failed for {label_text} ({direct_selector}); falling back: {e}"
            )

    row_xpath = "//div[contains(@class,'ocs-transaction-flight-fields') and contains(@class,'first-flight')]"

    def _locate_label():
        row = page.locator(f"xpath={row_xpath}").first
        row.wait_for(state="visible", timeout=15000)
        return page.locator(
            f"xpath={row_xpath}//*[normalize-space()='{label_text}']"
        ).first

    try:
        label = _locate_label()
        label.wait_for(state="visible", timeout=10000)
    except Exception:
        # Fallback: search anywhere on the page for the label if the new row
        # is missing the expected "first-flight" class.
        label = page.locator(f"//*[normalize-space()='{label_text}']").first
        label.wait_for(state="visible", timeout=10000)

    container = label.locator(
        "xpath=ancestor::section[1]/following-sibling::section[1]"
    )
    container.scroll_into_view_if_needed()

    field = container.locator(
        "xpath=.//input[not(@type='checkbox') and not(@type='radio')]"
    ).first
    field.wait_for(state="visible", timeout=10000)
    field.scroll_into_view_if_needed()

    # 5) Fill the input
    field.fill(str(value))


def fill_field_by_selector(page, selector: str, value, timeout=10000):
    """Wait for a specific input selector and fill it."""

    field = page.locator(selector)
    field.wait_for(state="visible", timeout=timeout)
    field.fill(str(value))


def fill_slot_form(page, slot, operation, parkloc):
    """Fill the entire slot form using react-select + text inputs."""

    row = page.locator("div.ocs-transaction-flight-fields.first-flight").first
    row.wait_for(state="visible", timeout=15000)

    if slot.get("airport"):
        control = row.locator(".ocs__control").first
        control.wait_for(state="visible", timeout=8000)
        control.scroll_into_view_if_needed()

        for _ in range(2):
            control.click(force=True)
            page.wait_for_timeout(200)
            menu = page.locator(".ocs__menu")
            if menu.is_visible():
                break
        else:
            raise Exception("A/P dropdown did not open")

        option = page.get_by_role("option", name=slot["airport"])
        option.wait_for(timeout=8000)
        option.click()
        page.wait_for_timeout(200)

    if slot.get("acreg"):
        fill_field_by_selector(page, "#aircraftRegistration", slot["acreg"], timeout=8000)

    if slot.get("date"):
        fill_field_by_selector(page, "input[name='startDate']", slot["date"], timeout=8000)
        page.keyboard.press("Escape")
        page.wait_for_timeout(150)

    if slot.get("time"):
        fill_field_by_selector(page, "#clearedTimeDep", slot["time"], timeout=8000)

    if slot.get("other_airport"):
        dest_selector = "#destinationStation" if operation == "departure" else "#originStation"
        try:
            fill_field_by_selector(page, dest_selector, slot["other_airport"], timeout=8000)
        except Exception:
            # fallback to label-based lookup if the direct selector isn't present
            if operation == "departure":
                fill_text_cell(page, "Dest", slot["other_airport"])
            else:
                fill_text_cell(page, "Orig", slot["other_airport"])

    select_row_react_select(page, "STC", "D")

    if slot.get("airport") == "CYYZ":
        try:
            svc_control = row.locator(".trans-field-w-service-type .ocs__control").first
            svc_control.wait_for(state="visible", timeout=8000)
            svc_control.click(force=True)
            page.wait_for_timeout(150)
            page.get_by_role("option", name="D General Aviation").click()
        except Exception:
            select_row_react_select(page, "ParkLoc", parkloc)


    return True



def click_add_slot_button(page, operation):
    """
    Click the correct 'Add Flight No' or 'Add A/C Reg' inside the proper block
    (Departure / Arrival / Turnaround / Out and Back).
    """

    # Map operation to human labels and the SECOND SPAN label
    op_map = {
        "dep-flightno": ("Departure",     "Flight No"),
        "dep-reg":      ("Departure",     "A/C Reg"),
        "arr-flightno": ("Arrival",       "Flight No"),
        "arr-reg":      ("Arrival",       "A/C Reg"),
        "turn-flightno":("Turnaround",    "Flight No"),
        "turn-reg":     ("Turnaround",    "A/C Reg"),
        "out-flightno": ("Out and Back",  "Flight No"),
        "out-reg":      ("Out and Back",  "A/C Reg"),
    }

    block_label, second_span = op_map[operation]

    # 1) Find the section header (<p>Departure</p>, <p>Arrival</p>, etc.)
    header = page.locator(f"//p[normalize-space()='{block_label}']").first
    header.wait_for(state="visible", timeout=15000)

    # 2) Move up to the container that holds the two Add buttons
    container = header.locator(
        "xpath=ancestor::div[contains(@class,'ocs-double-btn')]"
    ).first
    container.wait_for(state="visible", timeout=15000)

    # 3) Locate the correct button using *two separate inner spans*
    button_xpath = f".//button[.//span[normalize-space()='Add'] and .//span[normalize-space()='{second_span}']]"

    btn = container.locator(f"xpath={button_xpath}").first
    btn.wait_for(state="visible", timeout=8000)

    # 4) Click it
    btn.click()

def select_dropdown_by_label(page, label_text, option_text):
    # Scope to the editable row, not the header
    row = page.locator("div.ocs-transaction-flight-fields.first-flight").first

    # Find the label inside THIS row
    label = row.get_by_text(label_text, exact=True)

    # Move to the next <section> which contains the dropdown
    container = label.locator("xpath=ancestor::section[1]/following-sibling::section[1]")

    # React-Select control
    box = container.locator(".ocs__control").first

    # Open dropdown
    box.click()

    # Wait for menu
    page.wait_for_selector(".ocs__menu")

    # Select option
    page.get_by_role("option", name=option_text).click()




def run_ocs_autofill(slot: dict, creds: dict):
    """
    slot keys (manual or FEAS):
      operation: 'departure' | 'arrival'
      airport: CYYZ/CYYC/CYUL/CYVR
      acreg: CFASY / tail
      date: 27NOV or 2025-11-27 (OCS accepts both in practice)
      time: 0800 (UTC)
      other_airport: KTEB (Dest for departure, Orig for arrival)
      parkloc: defaults to SKYCHARTER
    """
    operation = slot.get("operation", "departure").lower()
    parkloc = slot.get("parkloc") or "SKYCHARTER"

    # Default aircraft details so the seat and type fields are always populated.
    ac_type = slot.get("ac_type") or "E545"
    slot["ac_type"] = ac_type

    if not slot.get("seats"):
        if ac_type == "E545":
            slot["seats"] = "9"
        elif ac_type in ("C25A", "C25B"):
            slot["seats"] = "7"

    # reset debug log per run so the latest dump is easy to spot
    try:
        with open(LOG_FILE, "w", encoding="utf-8") as f:
            f.write("OCS autofill debug log\n")
    except Exception:
        pass

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(ignore_https_errors=True)
        page = context.new_page()

        # ---------------- LOGIN STEP A ----------------
        # Load base URL and wait for initialization scripts
        page.goto("https://www.online-coordination.com/", wait_until="networkidle")

        # Sometimes the Login button appears, sometimes the login page loads directly
        try:
            page.wait_for_selector("text=Login", timeout=8000)
            page.click("text=Login")
        except Exception:
            pass

        # --- BRANCH A: Native OCS login ---
        try:
            page.wait_for_selector("input[name='username']", timeout=5000)
            print("Detected native OCS login page")

            page.locator("input[name='username']").fill(creds["username"])
            page.locator("input[name='password']").fill(creds["password"])
            page.locator("button:has-text('Next'), input[type='submit']").click()

        except Exception:
            # Not the native login → move on and try Azure
            print("Native OCS login not found, trying Azure…")

        # --- BRANCH B: Azure login ---
        try:
            page.wait_for_selector("input[type='email'], input[name='loginfmt']", timeout=8000)
            print("Detected Azure login page")

            # Username
            page.locator("input[type='email'], input[name='loginfmt']").fill(creds["username"])
            page.locator("button:has-text('Next'), input[type='submit']").click()

            # Password
            page.wait_for_selector("input[type='password']", timeout=10000)
            page.locator("input[type='password']").fill(creds["password"])
            page.locator("button:has-text('Sign in'), input[type='submit']").click()

            # Stay signed in? page
            try:
                page.wait_for_selector("#idSIButton9", timeout=5000)
                page.click("#idSIButton9")
            except PWTimeoutError:
                pass

        except Exception as e:
            print("Azure login not detected:", e)

        # --- HANDLE "USER ALREADY LOGGED IN" SCREEN ---
        try:
            page.wait_for_selector("text=already logged in", timeout=3000)
            print("Detected 'User already logged in' warning")

            logout_button = page.locator("button:has-text('Log out other user')")
            if logout_button.count() > 0:
                logout_button.click()
                print("Clicked 'Log out other user'")

                # Wait for login page again and re-enter creds
                page.wait_for_selector("input[name='username']", timeout=8000)
                page.locator("input[name='username']").fill(creds["username"])
                page.locator("input[name='password']").fill(creds["password"])
                page.locator("button:has-text('Next'), input[type='submit']").click()
        except PWTimeoutError:
            pass

        # Final wait for dashboard (or passphrase page)
        page.wait_for_load_state("networkidle")
        print("Login complete (native or Azure, pre-passphrase)")

        # ---------------- LOGIN STEP B (AUTO PASSPHRASE) ----------------
        try:
            page.wait_for_selector("text=Please type the", timeout=8000)
            print("Detected passphrase challenge")

            instruction = page.locator("text=Please type the").inner_text()
            # Example: "Please type the fourth and seventh character in your pass phrase"
            match = re.search(r"the\s+(\w+)\s+and\s+(\w+)\s+character", instruction, re.IGNORECASE)
            if not match:
                raise Exception(f"Could not parse passphrase instructions: {instruction}")

            word1, word2 = match.groups()
            ordinal_map = {
                "first": 1, "second": 2, "third": 3, "fourth": 4,
                "fifth": 5, "sixth": 6, "seventh": 7, "eighth": 8,
                "ninth": 9, "tenth": 10
            }

            if word1.lower() not in ordinal_map or word2.lower() not in ordinal_map:
                raise Exception(f"Unknown ordinals in passphrase prompt: {word1}, {word2}")

            idx1 = ordinal_map[word1.lower()]
            idx2 = ordinal_map[word2.lower()]

            phrase = creds.get("passphrase", "") or ""
            if len(phrase) < max(idx1, idx2):
                sg.popup_error(
                    f"Stored passphrase is too short for positions {idx1} and {idx2}.\n"
                    f"Length is {len(phrase)}. Please update it in the app."
                )
                browser.close()
                return

            char1 = phrase[idx1 - 1]
            char2 = phrase[idx2 - 1]

            print(f"Passphrase autofill: {idx1}='{char1}', {idx2}='{char2}'")

            inputs = page.locator("input[type='password'], input[type='text']")
            inputs.nth(0).fill(char1)
            inputs.nth(1).fill(char2)
            page.locator("form button:has-text('Login')").first.click()

        except PWTimeoutError:
            print("No passphrase challenge detected.")

        page.wait_for_load_state("networkidle")

        # ---------------- NAV TO ADD FLIGHTS (UI-driven) ----------------
        # Wait for dashboard to fully load after login
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1500)  # give Angular some time

        # DEBUG PAUSE: inspect Chromium BEFORE submenu clicks
        print("Pausing for manual inspection BEFORE navigating to Add Flights")
        page.pause()
        # ---------------------------------------------------------------

        # Click "GA/BA Flights" in the top menu
        # Attempt 1: Standard ARIA locator
        try:
            page.get_by_role("link", name="GA/BA Flights").click(timeout=3000)
        except:
            # Fallback: simple text-based click (works reliably on OCS)
            print("[INFO] Falling back to text locator for GA/BA Flights...")
            page.get_by_text("GA/BA Flights").click()


        # Wait for submenu items to appear
        page.wait_for_selector("text=Add Flights", timeout=8000)

        # Click the "Add Flights" submenu
        page.click("text=Add Flights")

        # ---------------- Ensure Add Flights page is fully ready ----------------
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1500)  # allow Angular to finish animations

        page.wait_for_selector("//p[normalize-space()='Departure']", timeout=20000)
        page.wait_for_selector("//p[normalize-space()='Arrival']", timeout=20000)

        page.wait_for_timeout(1000)

        # ---------------- ADD SLOT ROW ----------------
        if operation == "departure":
            opkey = "dep-reg"
        elif operation == "arrival":
            opkey = "arr-reg"
        else:
            raise Exception(f"Unknown operation: {operation}")

        click_add_slot_button(page, opkey)
        page.wait_for_timeout(800)


        # ---------------- FILL REQUIRED FIELDS ----------------

        # A/P dropdown (react-select)
        if slot.get("airport"):
            ok = select_row_react_select(page, "A/P", slot["airport"])
            if not ok:
                sg.popup_error("Couldn't select A/P dropdown. We'll need to tweak selector.")
                page.pause()


        # 2) A/C Reg
        if slot.get("acreg"):
            fill_text_cell(page, "A/C Reg", slot["acreg"])

        # 3) Date
        if slot.get("date"):
            fill_text_cell(page, "Date", slot["date"])
            page.keyboard.press("Escape")   # Close any date pop-up
            page.wait_for_timeout(150)

        # 4) Seats (optional but available)
        if slot.get("seats"):
            fill_text_cell(page, "Seats", slot["seats"])

        # 5) A/C Type (optional but available)
        if slot.get("ac_type"):
            fill_text_cell(page, "A/C Type", slot["ac_type"])

        # 6) Time  (clearedTimeDep or clearedTimeArr handled automatically via label)
        if slot.get("time"):
            fill_text_cell(page, "Time", slot["time"])

        # 7) Dest or Orig depending on operation
        if slot.get("other_airport"):
            if operation == "departure":
                fill_text_cell(page, "Dest", slot["other_airport"])
            else:
                fill_text_cell(page, "Orig", slot["other_airport"])

        # STC dropdown
        select_row_react_select(page, "STC", "D")

        # ParkLoc dropdown only applies to CYYZ slots
        if slot.get("airport") == "CYYZ":
            select_row_react_select(page, "ParkLoc", parkloc)




        # ---------------- STOP SHORT OF SEND ALL ----------------
        sg.popup("Autofill complete.\nReview in browser, then click Send All manually.")
        page.pause()
        browser.close()

def main():
    sg.theme("SystemDefault")  # now safe with the updated PySimpleGUI install

    saved_creds = load_saved_creds()

    layout = [
        [sg.Text("OCS Slot Autofill Tool (FEAS-ready)", font=("Segoe UI", 14, "bold"))],

        [sg.Frame("Optional: paste FEAS JSON here",
                  [[sg.Multiline(size=(80,6), key="feas_json")],
                   [sg.Button("Parse FEAS JSON")]])],

        [sg.Frame("Operation Type",
                  [[sg.Radio("Departure", "op", default=True, key="op_dep"),
                    sg.Radio("Arrival", "op", default=False, key="op_arr")]])],

        [sg.Frame("Slot Details (manual or FEAS-filled)",
                  [
                      [sg.Text("A/P (Airport):"), sg.Input(key="airport", size=(8,1)),
                       sg.Text("A/C Reg:"), sg.Input(key="acreg", size=(12,1))],
                      [sg.Text("Date:"), sg.Input(key="date", size=(10,1)),
                       sg.Text("Time (UTC):"), sg.Input(key="time", size=(6,1))],
                      [sg.Text("Dest/Orig airport:"), sg.Input(key="other_airport", size=(10,1)),
                       sg.Text("ParkLoc:"), sg.Input(key="parkloc", size=(14,1))]
                  ])],

        [sg.Frame("OCS Credentials",
                  [
                      [sg.Text("Username:"),
                       sg.Input(key="username", size=(20,1),
                                default_text=saved_creds.get("username", ""))],
                      [sg.Text("Password:"),
                       sg.Input(key="password", password_char="*",
                                size=(20,1),
                                default_text=saved_creds.get("password", ""))],
                      [sg.Text("Passphrase:"),
                       sg.Input(key="passphrase", password_char="*",
                                size=(20,1),
                                default_text=saved_creds.get("passphrase", ""))]
                  ])],

        [sg.Button("Launch Autofill", button_color=("white", "#0B84F3")),
         sg.Button("Exit")]
    ]

    window = sg.Window("OCS Slot Autofill v1.1", layout)

    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, "Exit"):
            break

        if event == "Parse FEAS JSON":
            data = parse_feas_json(values["feas_json"])
            if not data:
                sg.popup_error("Invalid JSON.")
                continue

            op = (data.get("operation") or data.get("type") or "departure").lower()
            window["op_dep"].update(value=(op == "departure"))
            window["op_arr"].update(value=(op == "arrival"))

            window["airport"].update(data.get("airport",""))
            window["acreg"].update(data.get("acreg","") or data.get("tail",""))
            window["date"].update(data.get("date",""))
            window["time"].update(data.get("time",""))
            window["other_airport"].update(
                data.get("other_airport","") or data.get("dest","") or data.get("orig","")
            )
            window["parkloc"].update(data.get("parkloc","SKYCHARTER"))

        if event == "Launch Autofill":
            operation = "arrival" if values["op_arr"] else "departure"

            slot = {
                "operation": operation,
                "airport": values["airport"].strip(),
                "acreg": values["acreg"].strip(),
                "date": values["date"].strip(),
                "time": values["time"].strip(),
                "other_airport": values["other_airport"].strip(),
                "parkloc": values["parkloc"].strip() or "SKYCHARTER",
            }
            creds = {
                "username": values["username"].strip(),
                "password": values["password"],
                "passphrase": values["passphrase"],
            }
            if not creds["username"] or not creds["password"]:
                sg.popup_error("Enter OCS credentials.")
                continue

            # Save creds so they auto-populate next time
            save_creds(creds["username"], creds["password"], creds["passphrase"])

            run_ocs_autofill(slot, creds)

    window.close()

if __name__ == "__main__":
    main()
