import json
import os
import re
import tkinter as tk
from tkinter import messagebox, scrolledtext, simpledialog, ttk
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
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("OCS Passphrase", "OCS needs 2 passphrase characters.\n" + request_text)
    c1 = simpledialog.askstring("Passphrase", "Character #1:", parent=root)
    c2 = simpledialog.askstring("Passphrase", "Character #2:", parent=root)
    root.destroy()
    return c1, c2

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
            f"xpath=(.//section[contains(@class,'ocs-field-item')])[{column_index + 1}]"
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

            # If the label lives in the header, compute its sibling index and
            # pick the same position in the first input row. This avoids
            # depending on positional predicates that can drift when markup
            # shifts or additional columns are inserted.
            section_locator = label.locator(
                "xpath=ancestor::section[contains(@class,'ocs-field-item')][1]"
            )
            if section_locator.count() > 0:
                try:
                    section_index = section_locator.evaluate(
                        "el => Array.from(el.parentElement.children).indexOf(el)"
                    )
                except Exception:
                    section_index = -1
            else:
                section_index = -1

            if section_index >= 0:
                field_row = row.locator(
                    "xpath=.//div[contains(@class,'d-flex')][1]"
                ).first
                section_locator = field_row.locator(
                    f"xpath=(.//section[contains(@class,'ocs-field-item')])[{section_index + 1}]"
                ).first
            else:
                # Fallback to the next field-item if we cannot compute an index
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


def select_stc(page, value, timeout=8000):
    """Select the STC dropdown within the first flight row by targeting its own section."""

    row = page.locator("div.ocs-transaction-flight-fields.first-flight").first
    row.wait_for(state="visible", timeout=timeout)

    # Prefer the inline section with an STC label; fall back to the service-type
    # composite control if the inline one isn't present/visible yet.
    stc_section = row.locator(
        "xpath=.//section[.//label[normalize-space()='STC']]"
    ).first

    control = None
    try:
        stc_section.wait_for(state="visible", timeout=timeout)
        inline_control = stc_section.locator(".ocs__control").first
        inline_control.wait_for(state="visible", timeout=timeout)
        control = inline_control
    except Exception:
        # Inline control may not be visible on some layouts; use the combined
        # service-type control instead.
        fallback_control = row.locator(
            ".trans-field-w-service-type .ocs__control"
        ).first
        fallback_control.wait_for(state="visible", timeout=timeout)
        control = fallback_control

    control.scroll_into_view_if_needed()

    for _ in range(2):
        control.click(force=True)
        page.wait_for_timeout(200)
        menu = page.locator(".ocs__menu")
        if menu.is_visible():
            break
    else:
        raise Exception("STC dropdown did not open")

    option = page.get_by_role("option", name=value)
    option.wait_for(timeout=timeout)
    option.click()
    page.wait_for_timeout(150)


def select_parkloc(page, value, timeout=8000):
    """Select ParkLoc as the LAST react-select dropdown in the first-flight row."""

    row = page.locator(".ocs-transaction-flight-fields.first-flight").first
    row.wait_for(state="visible", timeout=timeout)

    # Grab ALL react-select controls in the row
    controls = row.locator(".ocs__control")
    count = controls.count()

    if count == 0:
        raise Exception("No react-select controls found in first-flight row")

    # ParkLoc = last dropdown (STC is second-last)
    control = controls.nth(count - 1)
    control.wait_for(state="visible", timeout=timeout)
    control.scroll_into_view_if_needed()

    # Open dropdown reliably
    for _ in range(2):
        control.click(force=True)
        page.wait_for_timeout(200)
        if page.locator(".ocs__menu").is_visible():
            break
    else:
        raise Exception("ParkLoc dropdown did not open")

    # Select the requested option
    option = page.get_by_role("option", name=value)
    option.wait_for(timeout=timeout)
    option.click()

    page.wait_for_timeout(150)


def click_send_all(page, timeout=12000):
    """Click the Send All button and wait for the confirmation view to load."""

    try:
        button = page.locator("#sendAllBtn").first
        button.wait_for(state="visible", timeout=timeout)
        button.scroll_into_view_if_needed()
        button.click()
        page.wait_for_load_state("networkidle")
        log_debug("Clicked Send All button")
        return True
    except Exception as e:
        log_debug(f"[ERROR clicking Send All]: {e}")
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
        # Arrival uses clearedTimeArr while departure uses clearedTimeDep.
        "Time": ["#clearedTimeDep", "#clearedTimeArr"],
        "Dest": "#destinationStation",
        "Orig": "#originStation",
    }

    direct_selector = selector_map.get(label_text)
    if direct_selector:
        direct_selectors = (
            direct_selector if isinstance(direct_selector, (list, tuple)) else [direct_selector]
        )
        last_error = None

        for selector in direct_selectors:
            try:
                fill_field_by_selector(page, selector, value, timeout=12000)
                return
            except Exception as e:
                last_error = e
                continue

        print(
            f"Direct selector failed for {label_text} ({direct_selectors}); falling back: {last_error}"
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
        time_selectors = ["#clearedTimeDep", "#clearedTimeArr"]
        last_error = None
        for selector in time_selectors:
            try:
                fill_field_by_selector(page, selector, slot["time"], timeout=8000)
                break
            except Exception as e:
                last_error = e
        else:
            raise last_error if last_error else Exception("No time input found")

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

    select_stc(page, "D")

    if slot.get("airport") == "CYYZ":
        try:
            svc_control = row.locator(".trans-field-w-service-type .ocs__control").first
            svc_control.wait_for(state="visible", timeout=8000)
            svc_control.click(force=True)
            page.wait_for_timeout(150)
            page.get_by_role("option", name="D General Aviation").click()
        except Exception:
            select_parkloc(page, parkloc)


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




class OCSAutomationSession:
    def __init__(self):
        self._pw = None
        self.browser = None
        self.context = None
        self.page = None

    @property
    def is_active(self):
        return self.page is not None

    def _reset_debug_log(self):
        try:
            with open(LOG_FILE, "w", encoding="utf-8") as f:
                f.write("OCS autofill debug log\n")
        except Exception:
            pass

    def start(self, creds: dict):
        if self.page:
            return

        self._pw = sync_playwright().start()
        self.browser = self._pw.chromium.launch(headless=False)
        self.context = self.browser.new_context(ignore_https_errors=True)
        self.page = self.context.new_page()

        self._login(creds)
        self._nav_to_add_flights()

    def _login(self, creds: dict):
        page = self.page

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
                "first": 1,
                "second": 2,
                "third": 3,
                "fourth": 4,
                "fifth": 5,
                "sixth": 6,
                "seventh": 7,
                "eighth": 8,
                "ninth": 9,
                "tenth": 10,
            }

            if word1.lower() not in ordinal_map or word2.lower() not in ordinal_map:
                raise Exception(f"Unknown ordinals in passphrase prompt: {word1}, {word2}")

            idx1 = ordinal_map[word1.lower()]
            idx2 = ordinal_map[word2.lower()]

            phrase = creds.get("passphrase", "") or ""
            if len(phrase) < max(idx1, idx2):
                messagebox.showerror(
                    "OCS Slot Autofill",
                    (
                        f"Stored passphrase is too short for positions {idx1} and {idx2}.\n"
                        f"Length is {len(phrase)}. Please update it in the app."
                    ),
                )
                self.close()
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

    def _nav_to_add_flights(self):
        page = self.page
        # ---------------- NAV TO ADD FLIGHTS (UI-driven) ----------------
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1500)  # give Angular some time

        # Prefer the sidebar/menu control (button/anchor) instead of a loose
        # text match to avoid strict-mode collisions with the page header.
        try:
            nav_button = page.locator(
                "button:has(span:has-text('GA/BA Flights')), "
                "a:has(span:has-text('GA/BA Flights'))"
            ).first
            nav_button.click(timeout=3000)
        except Exception:
            try:
                page.get_by_role("link", name="GA/BA Flights").first.click(timeout=3000)
            except Exception:
                print("[INFO] Falling back to explicit text locator for GA/BA Flights...")
                page.locator("button:has-text('GA/BA Flights'), a:has-text('GA/BA Flights')").first.click()

        # The Add Flights control is also duplicated as a page header; target a
        # clickable element (button/anchor) to avoid strict mode errors.
        page.locator("button:has-text('Add Flights'), a:has-text('Add Flights')").first.click()

        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1500)  # allow Angular to finish animations
        page.wait_for_selector("//p[normalize-space()='Departure']", timeout=20000)
        page.wait_for_selector("//p[normalize-space()='Arrival']", timeout=20000)
        page.wait_for_timeout(1000)

    def ensure_add_flights_page(self):
        if not self.page:
            raise RuntimeError("Session is not active")

        # If we're not on the Add Flights page anymore, navigate via the UI
        # instead of hitting the route directly (direct deep links can 404
        # when the session hasn't initialized routing state).
        # If we're already on the slot-entry form, skip re-navigation; otherwise
        # walk through the sidebar path to avoid deep-link 404s.
        try:
            if not self.page.locator("//p[normalize-space()='Departure']").is_visible():
                self._nav_to_add_flights()
        except Exception:
            self._nav_to_add_flights()

        self.page.wait_for_selector("//p[normalize-space()='Departure']", timeout=20000)
        self.page.wait_for_selector("//p[normalize-space()='Arrival']", timeout=20000)
        self.page.wait_for_timeout(500)

    def _apply_slot_defaults(self, slot: dict):
        operation = slot.get("operation", "departure").lower()
        parkloc = slot.get("parkloc") or "SKYCHARTER"

        # Default aircraft details so the seat and type fields are always populated.
        ac_type = slot.get("ac_type") or "E545"
        slot["ac_type"] = ac_type

        # Default tail registrations when the user leaves A/C Reg blank.
        default_regs = {"E545": "CGASL", "C25A": "CFASP", "C25B": "CFASY"}
        if not slot.get("acreg") and ac_type in default_regs:
            slot["acreg"] = default_regs[ac_type]

        if not slot.get("seats"):
            if ac_type == "E545":
                slot["seats"] = "9"
            elif ac_type in ("C25A", "C25B"):
                slot["seats"] = "7"

        return operation, parkloc

    def book_slot(self, slot: dict):
        if not self.page:
            raise RuntimeError("Session has not been started. Call start() first.")

        self._reset_debug_log()
        operation, parkloc = self._apply_slot_defaults(slot)
        page = self.page

        self.ensure_add_flights_page()

        if operation == "departure":
            opkey = "dep-reg"
        elif operation == "arrival":
            opkey = "arr-reg"
        else:
            raise Exception(f"Unknown operation: {operation}")

        click_add_slot_button(page, opkey)
        page.wait_for_timeout(800)

        if slot.get("airport"):
            ok = select_row_react_select(page, "A/P", slot["airport"])
            if not ok:
                messagebox.showerror(
                    "OCS Slot Autofill",
                    "Couldn't select A/P dropdown. We'll need to tweak selector.",
                )

        if slot.get("acreg"):
            fill_text_cell(page, "A/C Reg", slot["acreg"])

        if slot.get("date"):
            fill_text_cell(page, "Date", slot["date"])
            page.keyboard.press("Escape")
            page.wait_for_timeout(150)

        if slot.get("seats"):
            fill_text_cell(page, "Seats", slot["seats"])

        if slot.get("ac_type"):
            fill_text_cell(page, "A/C Type", slot["ac_type"])

        if slot.get("time"):
            fill_text_cell(page, "Time", slot["time"])

        if slot.get("other_airport"):
            if operation == "departure":
                fill_text_cell(page, "Dest", slot["other_airport"])
            else:
                fill_text_cell(page, "Orig", slot["other_airport"])

        select_stc(page, "D")

        if slot.get("airport") == "CYYZ":
            select_parkloc(page, parkloc)

        send_clicked = click_send_all(page)

        if send_clicked:
            messagebox.showinfo(
                "OCS Slot Autofill",
                "Send All clicked automatically. Review the confirmation page in the browser.",
            )
        else:
            messagebox.showerror(
                "OCS Slot Autofill",
                "Couldn't click Send All automatically. Please click it manually in the browser.",
            )

    def close(self):
        try:
            if self.context:
                self.context.close()
            if self.browser:
                self.browser.close()
            if self._pw:
                self._pw.stop()
        finally:
            self.page = None
            self.context = None
            self.browser = None
            self._pw = None


def run_ocs_autofill(slot: dict, creds: dict, session: OCSAutomationSession | None = None):
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
    owns_session = session is None
    active_session = session or OCSAutomationSession()

    try:
        active_session.start(creds)
        active_session.book_slot(slot)
    finally:
        if owns_session:
            active_session.close()

def main():
    root = tk.Tk()
    root.title("OCS Slot Autofill v1.1")

    session = OCSAutomationSession()

    saved_creds = load_saved_creds()

    operation_var = tk.StringVar(value="departure")
    airport_var = tk.StringVar()
    acreg_var = tk.StringVar()
    date_var = tk.StringVar()
    time_var = tk.StringVar()
    other_airport_var = tk.StringVar()
    parkloc_var = tk.StringVar(value="SKYCHARTER")
    ac_type_var = tk.StringVar(value="E545")

    def parse_feas():
        data = parse_feas_json(feas_text.get("1.0", tk.END).strip())
        if not data:
            messagebox.showerror("OCS Slot Autofill", "Invalid JSON.")
            return

        op = (data.get("operation") or data.get("type") or "departure").lower()
        operation_var.set("arrival" if op == "arrival" else "departure")

        airport_var.set(data.get("airport", ""))
        acreg_var.set(data.get("acreg", "") or data.get("tail", ""))
        date_var.set(data.get("date", ""))
        time_var.set(data.get("time", ""))
        other_airport_var.set(
            data.get("other_airport", "")
            or data.get("dest", "")
            or data.get("orig", "")
        )
        parkloc_var.set(data.get("parkloc", "SKYCHARTER"))
        ac_type_var.set(data.get("ac_type", "") or data.get("aircraft_type", ""))

    def launch_autofill():
        operation = operation_var.get()
        slot = {
            "operation": operation,
            "airport": airport_var.get().strip(),
            "acreg": acreg_var.get().strip(),
            "date": date_var.get().strip(),
            "time": time_var.get().strip(),
            "other_airport": other_airport_var.get().strip(),
            "parkloc": parkloc_var.get().strip() or "SKYCHARTER",
            "ac_type": ac_type_var.get().strip(),
        }
        creds = {
            "username": username_var.get().strip(),
            "password": password_var.get(),
            "passphrase": passphrase_var.get(),
        }

        if not creds["username"] or not creds["password"]:
            messagebox.showerror("OCS Slot Autofill", "Enter OCS credentials.")
            return

        save_creds(creds["username"], creds["password"], creds["passphrase"])
        try:
            run_ocs_autofill(slot, creds, session=session)
        except Exception as e:
            messagebox.showerror("OCS Slot Autofill", f"Error booking slot: {e}")

    ttk.Label(root, text="OCS Slot Autofill Tool (FEAS-ready)", font=("Segoe UI", 14, "bold")).pack(pady=(10, 5))

    feas_frame = ttk.LabelFrame(root, text="Optional: paste FEAS JSON here")
    feas_frame.pack(fill="x", padx=10, pady=5)
    feas_text = scrolledtext.ScrolledText(feas_frame, width=80, height=6)
    feas_text.pack(fill="both", expand=True, padx=5, pady=5)
    ttk.Button(feas_frame, text="Parse FEAS JSON", command=parse_feas).pack(padx=5, pady=(0, 5), anchor="e")

    op_frame = ttk.LabelFrame(root, text="Operation Type")
    op_frame.pack(fill="x", padx=10, pady=5)
    ttk.Radiobutton(op_frame, text="Departure", variable=operation_var, value="departure").pack(side="left", padx=5, pady=5)
    ttk.Radiobutton(op_frame, text="Arrival", variable=operation_var, value="arrival").pack(side="left", padx=5, pady=5)

    slot_frame = ttk.LabelFrame(root, text="Slot Details (manual or FEAS-filled)")
    slot_frame.pack(fill="x", padx=10, pady=5)

    ttk.Label(slot_frame, text="A/P (Airport):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
    ttk.Combobox(
        slot_frame,
        textvariable=airport_var,
        values=["CYYZ", "CYUL", "CYVR", "CYYC"],
        width=8,
        state="readonly",
    ).grid(row=0, column=1, sticky="w", padx=5, pady=2)
    ttk.Label(slot_frame, text="A/C Reg:").grid(row=0, column=2, sticky="w", padx=5, pady=2)
    ttk.Entry(slot_frame, textvariable=acreg_var, width=16).grid(row=0, column=3, sticky="w", padx=5, pady=2)

    ttk.Label(slot_frame, text="Date:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
    ttk.Entry(slot_frame, textvariable=date_var, width=12).grid(row=1, column=1, sticky="w", padx=5, pady=2)
    ttk.Label(slot_frame, text="Time (Local):").grid(row=1, column=2, sticky="w", padx=5, pady=2)
    ttk.Entry(slot_frame, textvariable=time_var, width=10).grid(row=1, column=3, sticky="w", padx=5, pady=2)

    ttk.Label(slot_frame, text="Dest/Orig airport:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
    ttk.Entry(slot_frame, textvariable=other_airport_var, width=14).grid(row=2, column=1, sticky="w", padx=5, pady=2)
    ttk.Label(slot_frame, text="ParkLoc:").grid(row=2, column=2, sticky="w", padx=5, pady=2)
    ttk.Entry(slot_frame, textvariable=parkloc_var, width=16).grid(row=2, column=3, sticky="w", padx=5, pady=2)

    ttk.Label(slot_frame, text="A/C Type:").grid(row=3, column=0, sticky="w", padx=5, pady=2)
    ttk.Combobox(
        slot_frame,
        textvariable=ac_type_var,
        values=["E545", "C25A", "C25B"],
        width=8,
        state="readonly",
    ).grid(row=3, column=1, sticky="w", padx=5, pady=2)

    creds_frame = ttk.LabelFrame(root, text="OCS Credentials")
    creds_frame.pack(fill="x", padx=10, pady=5)

    username_var = tk.StringVar(value=saved_creds.get("username", ""))
    password_var = tk.StringVar(value=saved_creds.get("password", ""))
    passphrase_var = tk.StringVar(value=saved_creds.get("passphrase", ""))

    ttk.Label(creds_frame, text="Username:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
    ttk.Entry(creds_frame, textvariable=username_var, width=24).grid(row=0, column=1, sticky="w", padx=5, pady=2)
    ttk.Label(creds_frame, text="Password:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
    ttk.Entry(creds_frame, textvariable=password_var, width=24, show="*").grid(row=1, column=1, sticky="w", padx=5, pady=2)
    ttk.Label(creds_frame, text="Passphrase:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
    ttk.Entry(creds_frame, textvariable=passphrase_var, width=24, show="*").grid(row=2, column=1, sticky="w", padx=5, pady=2)

    btn_frame = ttk.Frame(root)
    btn_frame.pack(fill="x", padx=10, pady=10)
    def exit_app():
        session.close()
        root.destroy()

    ttk.Button(btn_frame, text="Launch Autofill", command=launch_autofill).pack(side="left", padx=5)
    ttk.Button(btn_frame, text="Exit", command=exit_app).pack(side="left", padx=5)

    root.mainloop()

if __name__ == "__main__":
    main()

