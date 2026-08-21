import sys
import os
import subprocess
import time
import json
import urllib.request
from pathlib import Path

import websocket


DEBUG_PORT = 9222
PROFILE = Path(os.environ["TEMP"]) / "edge_debug_profile"

#FGC_COOKIES = "C:\Macros\RDA\FGC\main.txt"
#SSO_COOKIES = "C:\Macros\RDA\FGC\main.txt"

FGC_COOKIES = sys.argv[1]
SSO_COOKIES = sys.argv[2]

URL_1 = "https://fgc-gui-app.app.paas.fedex.com/#/portal"
URL_2 = "https://fgc-lac-cairo-atl.prod.cloud.fedex.com/clearance/mainMenu.jsp"
URL_3 = "https://fgc-gui-app.app.paas.fedex.com/api/session/check"

# Find Edge
EDGE_PATHS = [
    Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Microsoft/Edge/Application/msedge.exe",
    Path(os.environ.get("PROGRAMFILES", "")) / "Microsoft/Edge/Application/msedge.exe",
]

EDGE_PATH = next(
    (path for path in EDGE_PATHS if path.exists()),
    None
)

if EDGE_PATH is None:
    raise FileNotFoundError("Microsoft Edge executable not found.")



# Start Edge
subprocess.Popen([
    str(EDGE_PATH),
    f"--remote-debugging-port={DEBUG_PORT}",
    "--remote-allow-origins=http://localhost:9222",
    f"--user-data-dir={PROFILE}",
    URL_1,
])

print("COOKIE manager Running V5.9")

time.sleep(5)


# Wait for DevTools
tabs_url = f"http://localhost:{DEBUG_PORT}/json"

for _ in range(30):
    try:
        with urllib.request.urlopen(tabs_url) as response:
            tabs = json.load(response)
        break
    except Exception:
        time.sleep(0.5)
else:
    raise RuntimeError("Edge DevTools did not become available.")


# Find page tab
tab = None

for t in tabs:
    if t.get("type") == "page":
        tab = t
        break

if tab is None:
    raise RuntimeError("No page tab found.")

print(f"Using tab: {tab['url']}")


# Connect to CDP
ws = websocket.create_connection(
    tab["webSocketDebuggerUrl"]
)

message_id = 0


def cdp_command(method, params=None):
    global message_id

    message_id += 1

    message = {
        "id": message_id,
        "method": method
    }

    if params is not None:
        message["params"] = params

    ws.send(json.dumps(message))

    while True:
        response = json.loads(ws.recv())

        if response.get("id") == message_id:
            return response


cdp_command("Page.enable")

def get_current_url():
    response = cdp_command(
        "Runtime.evaluate",
        {
            "expression": "window.location.href",
            "returnByValue": True
        }
    )

    return response["result"]["result"]["value"]


def wait_for_page_load(expected_url):

    while True:
        url_response = cdp_command(
            "Runtime.evaluate",
            {
                "expression": "window.location.href",
                "returnByValue": True
            }
        )

        current_url = url_response["result"]["result"]["value"]

        state_response = cdp_command(
            "Runtime.evaluate",
            {
                "expression": "document.readyState",
                "returnByValue": True
            }
        )

        state = state_response["result"]["result"]["value"]

        print(f"Page Status: {state}")

        url_matches = current_url.startswith(expected_url)

        if url_matches and state == "complete":
            return True

        time.sleep(1)

def wait_for_url2_with_redirect_retry(
    URL_PORTAL,
    URL_MANIFEST,
    timeout=60,
    url1_stable_seconds=1.0,
):
    overall_start = time.time()
    url1_complete_since = None

    while time.time() - overall_start < timeout:

        url_response = cdp_command(
            "Runtime.evaluate",
            {
                "expression": "window.location.href",
                "returnByValue": True
            }
        )

        current_url = url_response["result"]["result"]["value"]

        state_response = cdp_command(
            "Runtime.evaluate",
            {
                "expression": "document.readyState",
                "returnByValue": True
            }
        )

        state = state_response["result"]["result"]["value"]

        print(f"CURRENT URL:")
        print(f"{current_url}")
        print(f"CURRENT STT: {state}")

        # URL2 + complete = success
        if current_url.startswith(URL_MANIFEST) and state == "complete":
            print("Manifest Menu Loaded Sucessfully")
            return True

        # URL1 + complete = start/check stability timer
        if current_url.startswith(URL_PORTAL) and state == "complete":

            if url1_complete_since is None:
                url1_complete_since = time.time()
                print(
                    f"Still in PORTAL, Waiting for ReadyState..."
                )

            elif time.time() - url1_complete_since >= url1_stable_seconds:
                print(
                )

                print(f"Redirecting to Manifest Menu...")

                cdp_command(
                    "Page.navigate",
                    {
                        "url": URL_MANIFEST
                    }
                )

                url1_complete_since = None
                time.sleep(0.2)

        else:
            # We left URL1 or it is no longer complete.
            # Reset the URL1 stability timer.
            url1_complete_since = None

        time.sleep(1)

    print(f"TIMEOUT: Could not reach Manifest Menu within {timeout} seconds.")
    return False

# First page

time.sleep(1.5)

wait_for_page_load(URL_1)
print("Portal Ready... Launching FGCRM")

# Second page
print(" ")
print("=======================================")
print("     * FGC Redirect Manager V1.3 *     ")
print("=======================================")
print(" ")
cdp_command(
    "Page.navigate",
    {
        "url": URL_2
    }
)

time.sleep(1.5)

success = wait_for_url2_with_redirect_retry(
    URL_1,
    URL_2,
    timeout=60,
    url1_stable_seconds=1.0,
)

if not success:
    raise RuntimeError("FAILURE: Cant reach Manifest Menu...")

print("SUCESS in Portal Redirect... ")

# Save Cookies

def save_cookies(file_path, allowed_domains):
    response = cdp_command("Network.getAllCookies")

    cookies = response["result"]["cookies"]

    filtered_cookies = []

    print(" ")
    print("=============================")
    print("Checking COOKIES and Saving....")
    print(" ")

    for cookie in cookies:
        cookie_domain = cookie["domain"].lstrip(".")

        for allowed_domain in allowed_domains:
            allowed_domain = allowed_domain.lstrip(".")

            if (
                cookie_domain == allowed_domain
                or cookie_domain.endswith("." + allowed_domain)
            ):
                filtered_cookies.append(cookie)
                break

    cookie_string = ";".join(
        f"{cookie['name']}={cookie['value']}"
        for cookie in filtered_cookies
    )

    with open(file_path, "w", encoding="utf-8") as file:
        file.write(cookie_string)

    print(f"Cookies found: {len(cookies)}")


COOKIE_DOMAINS = [
    "fedex.com"
]

save_cookies(
    FGC_COOKIES,
    COOKIE_DOMAINS
)


def save_url_cookies(file_path, url):
    response = cdp_command(
        "Network.getCookies",
        {
            "urls": [url]
        }
    )

    cookies = response["result"]["cookies"]

    cookie_string = ";".join(
        f"{cookie['name']}={cookie['value']}"
        for cookie in cookies
    )

    with open(file_path, "w", encoding="utf-8") as file:
        file.write(cookie_string)

    print(f"Cookies found: {len(cookies)}")

# Check Page
print(" ")
print("=============================")
print("Redirecting to SSO Check Page")

cdp_command(
    "Page.navigate",
    {
        "url": URL_3
    }
)

time.sleep(1.5)

wait_for_page_load(URL_3)

save_url_cookies(SSO_COOKIES, URL_3)

print("=============================")
print(" ")
print("Process Finished Sucessfully!")
print("You can close this window now")

