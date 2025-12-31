"""
Fully Automated County & City Downloader
- Opens CAR.org Power BI page
- Automatically selects each target county in the slicer
- Extracts SharePoint image links
- Downloads county and city PNGs into the work directory

Usage: python final_working_downloader_auto.py

Optional env:
- SALA_COUNTIES="Santa Clara,San Mateo,Alameda,San Francisco"
- SALA_DEBUG=1
"""

import os
import time
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

try:
    from final_working_downloader import (
        setup_work_directory,
        get_sharepoint_links_from_powerbi,
        download_image_file,
    )
except Exception:
    setup_work_directory = None
    get_sharepoint_links_from_powerbi = None
    download_image_file = None


DEFAULT_COUNTIES = [
    "Santa Clara",
    "San Mateo",
    "Alameda",
    "San Francisco",
]

DEBUG = os.getenv("SALA_DEBUG", "").strip().lower() in ("1", "true", "yes")


def load_target_counties():
    raw = os.getenv("SALA_COUNTIES", "").strip()
    if raw:
        names = [name.strip() for name in raw.split(",") if name.strip()]
    else:
        names = DEFAULT_COUNTIES
    return [(name, idx + 1) for idx, name in enumerate(names)]


def _shorten(text, max_len=160):
    if not text:
        return ""
    text = " ".join(str(text).split())
    if len(text) <= max_len:
        return text
    return text[: max_len - 3] + "..."


def debug(msg):
    if DEBUG:
        print(f"   🐛 {msg}")


def summarize_element(el):
    try:
        tag = el.tag_name
        attrs = {
            "id": el.get_attribute("id"),
            "class": el.get_attribute("class"),
            "role": el.get_attribute("role"),
            "aria-label": el.get_attribute("aria-label"),
            "title": el.get_attribute("title"),
        }
        attr_str = ", ".join(f"{k}={_shorten(v, 80)}" for k, v in attrs.items() if v)
        text = _shorten(el.text, 80)
        if text:
            attr_str = f"{attr_str}, text={text}" if attr_str else f"text={text}"
        return f"<{tag} {attr_str}>" if attr_str else f"<{tag}>"
    except Exception as e:
        return f"<element error={e}>"


def log_elements(label, elements, max_items=3):
    debug(f"{label}: {len(elements)} found")
    if DEBUG and elements:
        for idx, el in enumerate(elements[:max_items], 1):
            debug(f"{label} #{idx}: {summarize_element(el)}")


def safe_click(driver, element, label="element"):
    try:
        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                element,
            )
        except Exception:
            pass
        element.click()
        return True
    except Exception as e:
        debug(f"{label} direct click failed: {e}")
    try:
        driver.execute_script("arguments[0].click();", element)
        return True
    except Exception as e:
        debug(f"{label} JS click failed: {e}")
    return False


def debug_element_attrs(element, label="element"):
    if not DEBUG:
        return
    attrs = [
        "id",
        "class",
        "role",
        "aria-label",
        "aria-owns",
        "aria-controls",
        "aria-expanded",
        "data-automation-id",
        "title",
    ]
    try:
        parts = []
        for name in attrs:
            val = element.get_attribute(name)
            if val:
                parts.append(f"{name}={_shorten(val, 120)}")
        if parts:
            debug(f"{label} attrs: " + ", ".join(parts))
    except Exception as e:
        debug(f"{label} attrs error: {e}")


def click_element_or_ancestor(driver, element, label="item", max_depth=3):
    current = element
    for depth in range(max_depth):
        if safe_click(driver, current, label=f"{label} depth={depth}"):
            return True
        try:
            current = current.find_element(By.XPATH, "..")
        except Exception:
            break
    return False


def find_listbox_from_combobox(driver, combobox):
    try:
        owns = combobox.get_attribute("aria-owns") or combobox.get_attribute("aria-controls")
        if not owns:
            return None
        for list_id in owns.split():
            try:
                el = driver.find_element(By.ID, list_id)
                if el and el.is_displayed():
                    debug(f"listbox from combobox id={list_id}")
                    return el
            except Exception:
                continue
    except Exception as e:
        debug(f"find_listbox_from_combobox failed: {e}")
    return None


def debug_global_listbox_counts(driver):
    if not DEBUG:
        return
    try:
        counts = driver.execute_script(
            """
            const all = document.querySelectorAll('[role="listbox"], [role="option"], [role="listitem"]');
            const listboxes = document.querySelectorAll('[role="listbox"]');
            return {
              all: all.length,
              listboxes: listboxes.length,
              options: document.querySelectorAll('[role="option"]').length,
              listitems: document.querySelectorAll('[role="listitem"]').length,
              slicerItems: document.querySelectorAll('[data-automation-id="slicer-item"]').length,
            };
            """
        )
        debug(f"global counts: {counts}")
    except Exception as e:
        debug(f"global counts failed: {e}")


def log_listbox_html(driver, listbox):
    if not DEBUG or not listbox:
        return
    try:
        html = driver.execute_script("return arguments[0].innerHTML;", listbox)
        debug(f"listbox innerHTML: {_shorten(html, 200)}")
    except Exception as e:
        debug(f"listbox html debug failed: {e}")


def find_scrollable_descendant(driver, root):
    if not root:
        return None
    try:
        return driver.execute_script(
            """
            const root = arguments[0];
            const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT);
            while (walker.nextNode()) {
              const el = walker.currentNode;
              if (el.scrollHeight && el.clientHeight && el.scrollHeight > el.clientHeight + 2) {
                return el;
              }
            }
            return null;
            """,
            root,
        )
    except Exception:
        return None


def find_scrollable_ancestor(driver, element):
    if not element:
        return None
    try:
        return driver.execute_script(
            """
            let el = arguments[0];
            while (el && el !== document.body) {
              if (el.scrollHeight && el.clientHeight && el.scrollHeight > el.clientHeight + 2) {
                return el;
              }
              el = el.parentElement;
            }
            return null;
            """,
            element,
        )
    except Exception:
        return None


def click_matching_item_in_listbox(driver, listbox, county_name):
    county_lower = county_name.lower()
    candidate_selectors = [
        "[role='option']",
        "[role='listitem']",
        "[data-automation-id='slicer-item']",
        ".slicerItemContainer",
        ".slicerItem",
    ]
    if not listbox:
        return False
    log_listbox_html(driver, listbox)
    debug_global_listbox_counts(driver)
    scroll_target = find_scrollable_descendant(driver, listbox)
    if not scroll_target:
        scroll_target = find_scrollable_ancestor(driver, listbox)
    if not scroll_target:
        scroll_target = listbox

    last_scroll_top = None
    for _ in range(8):
        for sel in candidate_selectors:
            try:
                elements = listbox.find_elements(By.CSS_SELECTOR, sel)
            except Exception:
                elements = []
            for el in elements:
                try:
                    text = (el.text or "").strip().lower()
                    aria = (el.get_attribute("aria-label") or "").strip().lower()
                    title = (el.get_attribute("title") or "").strip().lower()
                    hay = " ".join(val for val in (text, aria, title) if val)
                    if not hay or county_lower not in hay:
                        continue
                    try:
                        driver.execute_script(
                            "arguments[0].scrollIntoView({block: 'center'});",
                            el,
                        )
                    except Exception:
                        pass
                    if click_element_or_ancestor(driver, el, label="listbox option"):
                        return True
                    try:
                        driver.execute_script(
                            "arguments[0].dispatchEvent(new MouseEvent('click', {bubbles:true}));",
                            el,
                        )
                        return True
                    except Exception as e:
                        debug(f"listbox option JS click failed: {e}")
                except Exception:
                    continue

        try:
            scroll_height = driver.execute_script("return arguments[0].scrollHeight;", scroll_target)
            client_height = driver.execute_script("return arguments[0].clientHeight;", scroll_target)
            scroll_top = driver.execute_script("return arguments[0].scrollTop;", scroll_target)
        except Exception:
            break
        if scroll_height <= client_height + 2:
            break
        if last_scroll_top is not None and scroll_top == last_scroll_top:
            break
        last_scroll_top = scroll_top
        step = max(int(client_height * 0.6), 50)
        try:
            driver.execute_script(
                "arguments[0].scrollTop = arguments[0].scrollTop + arguments[1];",
                scroll_target,
                step,
            )
        except Exception:
            break
        time.sleep(0.2)
    return False


def scroll_listbox_and_click(driver, county_name):
    selectors = [
        "div.slicerBody[role='listbox']",
        "[role='listbox'][aria-label*='County']",
        "[role='listbox']",
    ]
    for sel in selectors:
        try:
            listboxes = driver.find_elements(By.CSS_SELECTOR, sel)
        except Exception:
            listboxes = []
        if not listboxes:
            continue
        log_elements(f"listbox selector {sel}", listboxes)
        for listbox in listboxes:
            if not listbox.is_displayed():
                continue
            if click_matching_item_in_listbox(driver, listbox, county_name):
                return True
    return False


def click_search_icon_if_present(driver):
    """Click search icon/button to reveal search input if Power BI hides it."""
    selectors = [
        "button[aria-label*='Search']",
        "button[title*='Search']",
        "span.searchIcon",
        "span[title*='Search']",
        "span[aria-label*='Search']",
    ]
    clicked = False
    for sel in selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, sel)
        except Exception:
            elements = []
        for el in elements:
            try:
                if el.is_displayed() and el.is_enabled():
                    if safe_click(driver, el, label="search icon"):
                        clicked = True
                        break
            except Exception:
                continue
        if clicked:
            break
    try:
        return driver.execute_script(
            """
            const inp = document.querySelector('input.searchInput, input[aria-label="Search"]');
            if (!inp) return false;
            const chain = [];
            let node = inp;
            while (node && node !== document.body) {
              chain.push(node);
              node = node.parentElement;
            }
            for (const el of chain) {
              el.style.display = 'block';
              el.style.visibility = 'visible';
              el.style.opacity = '1';
            }
            inp.removeAttribute('disabled');
            inp.setAttribute('aria-hidden', 'false');
            return true;
            """
        )
    except Exception:
        return clicked


def click_first_matching_option(driver, county_name):
    county_lower = county_name.lower()
    option_xpaths = [
        "//*[contains(@role,'option') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '%s')]"
        % county_lower,
        "//*[contains(@role,'listitem') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '%s')]"
        % county_lower,
        "//*[@data-automation-id='slicer-item' and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '%s')]"
        % county_lower,
        "//*[@aria-label and contains(translate(@aria-label, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '%s')]"
        % county_lower,
    ]
    for option_xpath in option_xpaths:
        try:
            options = driver.find_elements(By.XPATH, option_xpath)
            log_elements(f"option xpath {option_xpath}", options)
            for opt in options:
                try:
                    displayed = opt.is_displayed()
                    enabled = opt.is_enabled()
                    debug(f"option state: displayed={displayed}, enabled={enabled}")
                except Exception:
                    displayed = False
                    enabled = False
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", opt)
                except Exception:
                    pass
                if not displayed:
                    try:
                        driver.execute_script(
                            """
                            arguments[0].style.display = 'block';
                            arguments[0].style.visibility = 'visible';
                            arguments[0].style.opacity = '1';
                            arguments[0].removeAttribute('aria-hidden');
                            """,
                            opt,
                        )
                    except Exception:
                        pass
                if safe_click(driver, opt, label="option"):
                    return True
                try:
                    driver.execute_script(
                        "arguments[0].dispatchEvent(new MouseEvent('click', {bubbles:true}));",
                        opt,
                    )
                    return True
                except Exception as e:
                    debug(f"option JS click failed: {e}")
        except Exception:
            continue
    debug("no matching option clicked")
    return False

def current_selected_value(driver):
    try:
        elements = driver.find_elements(By.CSS_SELECTOR, "div.slicer-dropdown-menu[aria-label*='County']")
        for el in elements:
            if el.is_displayed():
                text = (el.text or "").strip()
                if text:
                    return text
    except Exception:
        pass
    return ""


def wait_for_selected_value(driver, county_name, previous_value="", timeout=6):
    end = time.time() + timeout
    target = county_name.lower()
    while time.time() < end:
        selected = current_selected_value(driver)
        if selected and target in selected.lower():
            return True
        if previous_value and selected and selected != previous_value:
            return False
        time.sleep(0.3)
    return False


def open_popup_for_slicer(driver, popup_id):
    if not popup_id:
        return False
    try:
        return driver.execute_script(
            """
            const popupId = arguments[0];
            const el = document.getElementById(popupId);
            if (!el) return false;
            el.style.display = 'block';
            el.style.visibility = 'visible';
            el.style.opacity = '1';
            el.removeAttribute('aria-hidden');
            const headers = el.querySelectorAll('.searchHeader.collapsed');
            headers.forEach(h => h.classList.remove('collapsed'));
            const input = el.querySelector('input.searchInput, input[aria-label="Search"]');
            if (input) {
              input.style.display = 'block';
              input.style.visibility = 'visible';
              input.style.opacity = '1';
              input.removeAttribute('disabled');
              input.setAttribute('aria-hidden', 'false');
            }
            return true;
            """,
            popup_id,
        )
    except Exception as e:
        debug(f"open popup failed: {e}")
        return False


def click_search_icon_in_popup(driver, popup_id):
    try:
        return driver.execute_script(
            """
            const popupId = arguments[0];
            const root = popupId ? document.getElementById(popupId) : document;
            if (!root) return false;
            const icon = root.querySelector('.searchIcon, .powervisuals-glyph.search');
            if (!icon) return false;
            icon.click();
            return true;
            """,
            popup_id,
        )
    except Exception as e:
        debug(f"click search icon failed: {e}")
        return False


def enable_single_select_in_popup(driver, popup_id):
    try:
        return driver.execute_script(
            """
            const popupId = arguments[0];
            const root = popupId ? document.getElementById(popupId) : document;
            if (!root) return false;
            const selectors = [
              "[aria-label*='Single select']",
              "[title*='Single select']",
              "[aria-label*='单选']",
              "[title*='单选']",
              "[data-automation-id*='singleSelect']",
              "[data-automation-id*='single-select']"
            ];
            for (const sel of selectors) {
              const el = root.querySelector(sel) || document.querySelector(sel);
              if (el) { el.click(); return true; }
            }
            return false;
            """,
            popup_id,
        )
    except Exception as e:
        debug(f"single select failed: {e}")
        return False


def deselect_select_all_in_popup(driver, popup_id):
    try:
        return driver.execute_script(
            """
            const popupId = arguments[0];
            const root = popupId ? document.getElementById(popupId) : document;
            if (!root) return false;
            const labels = ['all', 'select all'];
            const nodes = Array.from(root.querySelectorAll(
              "[role='option'], [data-automation-id='slicer-item'], .slicerItemContainer, .slicerItem"
            ));
            const match = nodes.find(node => {
              const text = (node.textContent || '').trim().toLowerCase();
              const title = (node.getAttribute('title') || '').trim().toLowerCase();
              const aria = (node.getAttribute('aria-label') || '').trim().toLowerCase();
              return labels.includes(text) || labels.includes(title) || labels.includes(aria);
            });
            if (!match) return false;
            const input = match.querySelector('input[type="checkbox"], input');
            const ariaChecked = match.getAttribute('aria-checked') || match.getAttribute('aria-selected');
            const checked = (input && input.checked) || ariaChecked === 'true';
            if (!checked) return false;
            const clickTarget = input || match;
            ['mouseover', 'mousedown', 'mouseup', 'click'].forEach(evt => {
              clickTarget.dispatchEvent(new MouseEvent(evt, {bubbles: true}));
            });
            clickTarget.click();
            return true;
            """,
            popup_id,
        )
    except Exception as e:
        debug(f"deselect all failed: {e}")
        return False


def clear_search_input(driver, popup_id):
    try:
        return driver.execute_script(
            """
            const popupId = arguments[0];
            const root = popupId ? document.getElementById(popupId) : document;
            if (!root) return false;
            const inp = root.querySelector('input.searchInput, input[aria-label="Search"]');
            if (!inp) return false;
            inp.value = '';
            inp.dispatchEvent(new Event('input', {bubbles: true}));
            inp.dispatchEvent(new Event('change', {bubbles: true}));
            return true;
            """,
            popup_id,
        )
    except Exception as e:
        debug(f"clear search failed: {e}")
        return False


def wait_for_option_in_popup(driver, county_name, popup_id, timeout=5):
    end = time.time() + timeout
    target = county_name.lower()
    while time.time() < end:
        try:
            found = driver.execute_script(
                """
                const target = arguments[0];
                const popupId = arguments[1];
                const root = popupId ? document.getElementById(popupId) : document;
                if (!root) return false;
                const nodes = Array.from(root.querySelectorAll(
                  "[role='option'], [data-automation-id='slicer-item'], .slicerItemContainer, .slicerItem"
                ));
                return nodes.some(node => {
                  const text = (node.textContent || '').trim().toLowerCase();
                  const title = (node.getAttribute('title') || '').trim().toLowerCase();
                  const aria = (node.getAttribute('aria-label') || '').trim().toLowerCase();
                  if (![text, title, aria].some(val => val && val.includes(target))) return false;
                  const rect = node.getBoundingClientRect();
                  const style = window.getComputedStyle(node);
                  return rect.width > 2 && rect.height > 2 &&
                    style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0';
                });
                """,
                target,
                popup_id,
            )
            if found:
                return True
        except Exception as e:
            debug(f"wait for option failed: {e}")
            break
        time.sleep(0.2)
    return False


def click_option_anywhere_exact(driver, label, popup_id=None):
    try:
        return driver.execute_script(
            """
            const target = arguments[0].toLowerCase().trim();
            const popupId = arguments[1];
            let root = document;
            if (popupId) {
              const el = document.getElementById(popupId);
              if (el) root = el;
            }
            const nodes = Array.from(root.querySelectorAll(
              "[role='option'], [data-automation-id='slicer-item'], .slicerItemContainer, .slicerItem"
            ));
            const match = nodes.find(node => {
              const text = (node.textContent || '').trim().toLowerCase();
              const title = (node.getAttribute('title') || '').trim().toLowerCase();
              const aria = (node.getAttribute('aria-label') || '').trim().toLowerCase();
              return [text, title, aria].some(val => val && val === target);
            });
            if (!match) return false;
            try { match.scrollIntoView({block: 'center'}); } catch (e) {}
            match.click();
            return true;
            """,
            label,
            popup_id,
        )
    except Exception as e:
        debug(f"exact click failed: {e}")
        return False


def scroll_option_into_view(driver, county_name, popup_id):
    try:
        return driver.execute_script(
            """
            const target = arguments[0].toLowerCase();
            const popupId = arguments[1];
            const root = popupId ? document.getElementById(popupId) : document;
            if (!root) return false;
            const list = root.querySelector('.slicerBody, [role="listbox"]');
            const nodes = Array.from(root.querySelectorAll(
              "[role='option'], [data-automation-id='slicer-item'], .slicerItemContainer, .slicerItem"
            ));
            const match = nodes.find(node => {
              const text = (node.textContent || '').trim().toLowerCase();
              const title = (node.getAttribute('title') || '').trim().toLowerCase();
              const aria = (node.getAttribute('aria-label') || '').trim().toLowerCase();
              return [text, title, aria].some(val => val && val.includes(target));
            });
            if (!match) return false;
            if (list && match.offsetTop) {
              list.scrollTop = Math.max(match.offsetTop - list.clientHeight / 2, 0);
            }
            try { match.scrollIntoView({block: 'center'}); } catch (e) {}
            return true;
            """,
            county_name,
            popup_id,
        )
    except Exception as e:
        debug(f"scroll option failed: {e}")
        return False


def force_click_option_exact(driver, label, popup_id=None):
    try:
        return driver.execute_script(
            """
            const target = arguments[0].toLowerCase().trim();
            const popupId = arguments[1];
            let root = document;
            if (popupId) {
              const el = document.getElementById(popupId);
              if (el) root = el;
            }
            const nodes = Array.from(root.querySelectorAll(
              "[role='option'], [data-automation-id='slicer-item'], .slicerItemContainer, .slicerItem"
            ));
            const match = nodes.find(node => {
              const text = (node.textContent || '').trim().toLowerCase();
              const title = (node.getAttribute('title') || '').trim().toLowerCase();
              const aria = (node.getAttribute('aria-label') || '').trim().toLowerCase();
              return [text, title, aria].some(val => val && val === target);
            });
            if (!match) return false;
            const input = match.querySelector('input[type="checkbox"], input');
            const clickTarget = match.querySelector(
              'input, .slicerCheckbox, .slicerItemContainer, .slicerItem'
            ) || match;
            try { clickTarget.scrollIntoView({block: 'center'}); } catch (e) {}
            ['mouseover', 'mousedown', 'mouseup', 'click'].forEach(evt => {
              clickTarget.dispatchEvent(new MouseEvent(evt, {bubbles: true}));
            });
            clickTarget.click();
            return true;
            """,
            label,
            popup_id,
        )
    except Exception as e:
        debug(f"force click exact failed: {e}")
        return False


def close_popup(driver):
    try:
        driver.execute_script("document.body.click();")
    except Exception:
        pass
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys(Keys.ESCAPE)
    except Exception:
        pass


def select_first_filtered_option(driver, popup_id):
    try:
        input_el = driver.execute_script(
            """
            const popupId = arguments[0];
            const root = popupId ? document.getElementById(popupId) : document;
            if (!root) return null;
            return root.querySelector('input.searchInput, input[aria-label="Search"]');
            """,
            popup_id,
        )
        if not input_el:
            return False
        try:
            input_el.click()
            input_el.send_keys(Keys.CONTROL + "a")
            input_el.send_keys(Keys.ARROW_DOWN)
            input_el.send_keys(Keys.ENTER)
        except Exception:
            driver.execute_script(
                """
                const inp = arguments[0];
                const fire = (key, code) => {
                  inp.dispatchEvent(new KeyboardEvent('keydown', {key, keyCode: code, which: code, bubbles: true}));
                  inp.dispatchEvent(new KeyboardEvent('keyup', {key, keyCode: code, which: code, bubbles: true}));
                };
                inp.focus();
                fire('ArrowDown', 40);
                fire('Enter', 13);
                """,
                input_el,
            )
        return True
    except Exception as e:
        debug(f"arrow select failed: {e}")
        return False


def force_click_option(driver, county_name, popup_id=None):
    try:
        return driver.execute_script(
            """
            const target = arguments[0].toLowerCase();
            const popupId = arguments[1];
            let root = document;
            if (popupId) {
              const el = document.getElementById(popupId);
              if (el) root = el;
            }
            const nodes = Array.from(root.querySelectorAll(
              "[role='option'], [data-automation-id='slicer-item'], .slicerItemContainer, .slicerItem"
            ));
            const matches = nodes.filter(node => {
              const text = (node.textContent || '').trim().toLowerCase();
              const title = (node.getAttribute('title') || '').trim().toLowerCase();
              const aria = (node.getAttribute('aria-label') || '').trim().toLowerCase();
              return [text, title, aria].some(val => val && val.includes(target));
            });
            if (!matches.length) return false;
            const list = root.querySelector('.slicerBody, [role="listbox"]');
            const visible = matches.find(node => {
              const rect = node.getBoundingClientRect();
              const style = window.getComputedStyle(node);
              return rect.width > 2 && rect.height > 2 &&
                style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0';
            });
            const match = visible || matches[0];
            if (list && match.offsetTop) {
              list.scrollTop = Math.max(match.offsetTop - list.clientHeight / 2, 0);
            }
            const clickTarget = match.querySelector(
              'input, .slicerCheckbox, .slicerItemContainer, .slicerItem'
            ) || match;
            try { clickTarget.scrollIntoView({block: 'center'}); } catch (e) {}
            ['mouseover', 'mousedown', 'mouseup', 'click'].forEach(evt => {
              clickTarget.dispatchEvent(new MouseEvent(evt, {bubbles: true}));
            });
            clickTarget.click();
            return true;
            """,
            county_name,
            popup_id,
        )
    except Exception as e:
        debug(f"force click option failed: {e}")
        return False


def default_work_directory():
    """Fallback work directory setup (if import failed)"""
    try:
        work_dir = os.getenv("SALA_WORK_DIR")
        if not work_dir:
            current_month = datetime.now().strftime("%B %Y")
            downloads_path = os.path.expanduser("~/Downloads")
            work_dir = os.path.join(downloads_path, f"SALA Report {current_month}")

        if not os.path.exists(work_dir):
            os.makedirs(work_dir)
            print(f"📁 Created work directory: {work_dir}")
        else:
            print(f"📁 Using work directory: {work_dir}")
        return work_dir
    except Exception as e:
        print(f"❌ Error setting up work directory: {e}")
        return os.getcwd()


def build_driver():
    """Configure Chrome driver with minimal, stable options"""
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1400,900")
    return webdriver.Chrome(options=chrome_options)


def switch_to_powerbi_iframe(driver):
    """Switch into the Power BI iframe; return True on success"""
    driver.switch_to.default_content()
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    debug(f"iframe count: {len(iframes)}")
    for iframe in iframes:
        src = iframe.get_attribute("src")
        debug(f"iframe src: {_shorten(src)}")
        if src and "app.powerbi.com" in src:
            driver.switch_to.frame(iframe)
            return True
    return False


def find_county_slicer(driver):
    """
    Try multiple heuristics to find the county slicer element.
    Returns a clickable element or None.
    """
    selectors = [
        "[aria-label='County']",
        "[aria-label*='County']",
        "[aria-label*='county']",
        "button[title='County']",
        "button[title*='County']",
        "[role='combobox'][aria-label*='County']",
        "[role='button'][aria-label*='County']",
        "[data-automation-id*='slicer'][aria-label*='County']",
    ]
    for sel in selectors:
        elements = driver.find_elements(By.CSS_SELECTOR, sel)
        log_elements(f"slicer selector {sel}", elements)
        for el in elements:
            try:
                if el.is_displayed() and el.is_enabled():
                    debug(f"using slicer: {summarize_element(el)}")
                    return el
            except Exception:
                continue

    # Broad fallback: any element whose visible text mentions county
    try:
        elements = driver.find_elements(
            By.XPATH,
            "//*[contains(translate(normalize-space(string(.)),'COUNTY','county'),'county')]",
        )
        log_elements("slicer text fallback", elements)
        for el in elements:
            if el.is_displayed() and el.is_enabled():
                debug(f"using slicer text fallback: {summarize_element(el)}")
                return el
    except Exception:
        pass
    debug("no slicer found")
    return None


def click_clear_if_present(driver):
    """Click slicer clear/eraser if present to avoid stale selection."""
    clear_selectors = [
        "[aria-label*='Clear']",
        "[title*='Clear']",
        "[aria-label*='清除']",
        "[aria-label*='重置']",
    ]
    for sel in clear_selectors:
        buttons = driver.find_elements(By.CSS_SELECTOR, sel)
        log_elements(f"clear selector {sel}", buttons)
        for btn in buttons:
            try:
                if btn.is_displayed() and btn.is_enabled():
                    safe_click(driver, btn, label="clear button")
                    return True
            except Exception:
                continue
    return False


def type_search_in_popup(driver, popup_id, county_name):
    """Type county name into the slicer search input inside the popup."""
    try:
        # First, clear the search input completely
        driver.execute_script(
            """
            const popupId = arguments[0];
            const root = popupId ? document.getElementById(popupId) : document;
            if (!root) return false;
            const inp = root.querySelector('input.searchInput, input[aria-label="Search"]');
            if (!inp) return false;
            inp.value = '';
            inp.dispatchEvent(new Event('input', {bubbles: true}));
            inp.dispatchEvent(new Event('change', {bubbles: true}));
            return true;
            """,
            popup_id,
        )
        time.sleep(0.3)  # Wait for search to clear

        input_el = driver.execute_script(
            """
            const popupId = arguments[0];
            const root = popupId ? document.getElementById(popupId) : document;
            if (!root) return null;
            return root.querySelector('input.searchInput, input[aria-label="Search"]');
            """,
            popup_id,
        )
        if not input_el:
            return False
        try:
            input_el.click()
            input_el.clear()  # Use Selenium's clear method
            input_el.send_keys(county_name)
            input_el.send_keys(Keys.ENTER)
            debug(f"typed search in: {summarize_element(input_el)}")
            return True
        except Exception:
            driver.execute_script(
                """
                const inp = arguments[0];
                const val = arguments[1];
                inp.value = val;
                inp.dispatchEvent(new Event('input', {bubbles: true}));
                inp.dispatchEvent(new Event('change', {bubbles: true}));
                ['keydown', 'keyup'].forEach(evt => {
                  inp.dispatchEvent(new KeyboardEvent(evt, {key: 'Enter', keyCode: 13, which: 13, bubbles: true}));
                });
                """,
                input_el,
                county_name,
            )
            debug(f"typed search via JS in: {summarize_element(input_el)}")
            return True
    except Exception as e:
        debug(f"search input failed: {e}")
        return False


def type_into_combobox(driver, combobox, county_name):
    try:
        if not safe_click(driver, combobox, label="combobox"):
            return False
        combobox.send_keys(Keys.CONTROL + "a")
        combobox.send_keys(Keys.DELETE)
        combobox.send_keys(county_name)
        combobox.send_keys(Keys.ENTER)
        debug(f"typed into combobox: {summarize_element(combobox)}")
        return True
    except Exception as e:
        debug(f"combobox type failed: {e}")
    return False


def ensure_option_clicked(driver, county_name, popup_id=None):
    """Try to click a matching option; return True if clicked."""
    if click_first_matching_option(driver, county_name):
        return True
    if popup_id and force_click_option(driver, county_name, popup_id=popup_id):
        return True
    return force_click_option(driver, county_name)


def select_county(driver, county_name, timeout=30):
    """
    Select the county in the Power BI slicer using heuristic UI automation.
    Returns True on best-effort success.
    """
    print(f"🔁 正在尝试自动选择 {county_name}...")
    if not switch_to_powerbi_iframe(driver):
        print("❌ 未找到 Power BI iframe")
        return False

    end = time.time() + timeout
    success = False

    attempt = 0
    while time.time() < end:
        if not switch_to_powerbi_iframe(driver):
            time.sleep(1)
            continue
        slicer = find_county_slicer(driver)
        if not slicer:
            time.sleep(1)
            continue

        current_value = current_selected_value(driver)
        if current_value and county_name.lower() in current_value.lower():
            success = True
            break

        try:
            debug_element_attrs(slicer, label="slicer")
            safe_click(driver, slicer, label="slicer")
            time.sleep(0.5)
            if click_clear_if_present(driver):
                print("🧹 已点击清除按钮")
        except Exception:
            time.sleep(1)
            continue

        popup_id = None
        try:
            popup_id = slicer.get_attribute("aria-controls") or slicer.get_attribute("aria-owns")
        except Exception:
            popup_id = None

        if popup_id and open_popup_for_slicer(driver, popup_id):
            debug(f"popup opened: {popup_id}")

        listbox_from_combo = find_listbox_from_combobox(driver, slicer)
        if listbox_from_combo:
            log_listbox_html(driver, listbox_from_combo)

        if popup_id and click_search_icon_in_popup(driver, popup_id):
            debug("clicked search icon")
        if popup_id and enable_single_select_in_popup(driver, popup_id):
            debug("enabled single select")
        if click_search_icon_if_present(driver):
            debug("clicked search icon (global)")
            time.sleep(0.3)

        if popup_id and current_value and current_value.strip().lower() == "all":
            if force_click_option_exact(driver, "All", popup_id=popup_id):
                time.sleep(0.2)
            elif force_click_option_exact(driver, "Select all", popup_id=popup_id):
                time.sleep(0.2)
        if popup_id:
            if deselect_select_all_in_popup(driver, popup_id):
                debug("deselected All")

        typed = type_search_in_popup(driver, popup_id, county_name)
        if typed:
            print("⌨️  已输入并回车")
        wait_for_option_in_popup(driver, county_name, popup_id, timeout=4)
        time.sleep(0.2)

        typed_combobox = False
        if not typed:
            typed_combobox = type_into_combobox(driver, slicer, county_name)
            if typed_combobox:
                print("⌨️  已在下拉框输入并回车")
                time.sleep(0.3)

        scroll_option_into_view(driver, county_name, popup_id)
        clicked = ensure_option_clicked(driver, county_name, popup_id=popup_id)
        if not clicked:
            clicked = select_first_filtered_option(driver, popup_id)
        if clicked:
            print("🖱️  已点击匹配选项")
            time.sleep(0.4)
        listbox_clicked = scroll_listbox_and_click(driver, county_name)
        if listbox_clicked:
            print("🖱️  已点击列表项")

        selected_value = current_selected_value(driver)
        if selected_value:
            debug(f"current selected value: {selected_value}")

        if clicked and selected_value and selected_value.strip().lower() == "all":
            if select_first_filtered_option(driver, popup_id):
                time.sleep(0.3)
                selected_value = current_selected_value(driver)
                if selected_value:
                    debug(f"current selected value: {selected_value}")

        if selected_value and county_name.lower() in selected_value.lower():
            success = True
            break
        if (clicked or listbox_clicked) and wait_for_selected_value(
            driver,
            county_name,
            previous_value=selected_value,
        ):
            success = True
            break
        attempt += 1
        print(f"⏳ 重试选择 {county_name} (attempt {attempt})")
        time.sleep(1)

    driver.switch_to.default_content()
    if not success:
        print(f"⚠️  未能自动选择 {county_name}，请检查页面布局。")
    else:
        print(f"✅ 已尝试选择 {county_name}")
    return success


def wait_for_links(driver, previous_total=0, timeout=20):
    """Wait until SharePoint links appear or change from previous count."""
    end = time.time() + timeout
    best = {"county": [], "city": []}
    if not switch_to_powerbi_iframe(driver):
        return best
    while time.time() < end:
        links = get_sharepoint_links_from_powerbi(driver)
        total = len(links.get("county", [])) + len(links.get("city", []))
        print(f"   🔎 链接检测: county={len(links.get('county', []))}, city={len(links.get('city', []))}, total={total}")
        if total > 0 and total != previous_total:
            return links
        best = links
        time.sleep(1)
    return best


def current_links_total(driver):
    """Get current total link count without waiting."""
    if not switch_to_powerbi_iframe(driver):
        return 0
    links = get_sharepoint_links_from_powerbi(driver)
    return len(links.get("county", [])) + len(links.get("city", []))


def main():
    if setup_work_directory:
        work_directory = setup_work_directory()
    else:
        work_directory = default_work_directory()

    driver = None
    downloaded_files = []

    print("🚀 Final Working County & City Downloader (Auto)")
    print("=" * 60)
    print("自动选择县并下载 SharePoint PNG 文件")
    print("=" * 60)
    print(f"📁 工作目录: {work_directory}")

    try:
        driver = build_driver()
        print("📍 正在打开 Power BI 页面...")
        driver.get("https://www.car.org/marketdata/interactive/buyersguide")
        time.sleep(4)
        print("✅ 页面加载完成（如需登录请在浏览器中完成）。")

        prev_count = 0
        target_counties = load_target_counties()

        for county_name, county_order in target_counties:
            print(f"\n{'='*60}")
            print(f"🏛️  COUNTY {county_order}/{len(target_counties)}: {county_name}")
            print(f"{'='*60}")

            baseline_total = current_links_total(driver)
            driver.switch_to.default_content()
            print(f"🧭 当前链接基线: {baseline_total}")

            selected = select_county(driver, county_name)
            if not selected:
                print(f"⚠️  跳过自动选择 {county_name}，尝试直接获取链接。")

            links_data = wait_for_links(driver, previous_total=baseline_total)
            total_links = len(links_data.get("county", [])) + len(links_data.get("city", []))
            print(f"📊 发现链接: county={len(links_data.get('county', []))}, city={len(links_data.get('city', []))}")
            if total_links == baseline_total:
                print("⚠️ 链接总数与基线相同，可能未切换到目标县。")

            # Download county level image (first link if exists)
            if links_data.get("county"):
                county_filename = f"{county_order}.png"
                if download_image_file(links_data["county"][0], county_filename, work_directory):
                    downloaded_files.append(county_filename)
            else:
                print("❌ 未找到县级链接")

            # Download city level images (if any)
            if links_data.get("city"):
                for city_index, city_url in enumerate(links_data["city"], 1):
                    city_filename = f"{county_order}({city_index}).png"
                    if download_image_file(city_url, city_filename, work_directory):
                        downloaded_files.append(city_filename)
                        print(f"   ✅ City {city_index} downloaded")
                    else:
                        print(f"   ❌ City {city_index} failed")
            else:
                print("⚠️  未找到市级链接")

            prev_count = total_links if total_links else prev_count

        # Summary
        print(f"\n{'='*60}")
        print("📊 FINAL SUMMARY")
        print(f"{'='*60}")
        print(f"Target counties: {len(target_counties)}")
        print(f"Successfully downloaded: {len(downloaded_files)} files")
        if downloaded_files:
            print("✅ Files:")
            for name in sorted(downloaded_files):
                path = os.path.join(work_directory, name)
                size = os.path.getsize(path) if os.path.exists(path) else 0
                print(f"  • {name} ({size:,} bytes)")
        print("\n🎉 All done (auto mode)!")

    except Exception as e:
        print(f"❌ Error: {e}")
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass


if __name__ == "__main__":
    main()
