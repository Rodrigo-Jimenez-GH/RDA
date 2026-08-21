import sys
import os
import subprocess
import time
import json
import urllib.request
from pathlib import Path

import websocket


# ============================================================
# EDGE BASICS
# ============================================================

DEBUG_PORT = 9222

PROFILE = Path(os.environ["TEMP"]) / "edge_debug_profile"


# ============================================================
# COOKIE OUTPUT FILES
# ============================================================

if len(sys.argv) < 3:
    print("ERROR: Missing cookie output paths.")
    print()
    print("Usage:")
    print("python cookies.py FGC_COOKIES SSO_COOKIES")
    sys.exit(1)

FGC_COOKIES = sys.argv[1]
SSO_COOKIES = sys.argv[2]


# ============================================================
# TARGET URLS
# ============================================================

URL_1 = "https://fgc-gui-app.app.paas.fedex.com/#/portal"
URL_2 = "https://fgc-lac-cairo-atl.prod.cloud.fedex.com/clearance/mainMenu.jsp"
URL_3 = "https://fgc-gui-app.app.paas.fedex.com/api/session/check"


# ============================================================
# LOGIN URLS
# ============================================================

LOGIN_URLS = [
    "https://purpleid.okta.com/login",
    "https://purpleid.okta.com/oauth2"
]

# ============================================================
# SETTINGS
# ============================================================

URL1_STABLE_SECONDS = 4
WRONG_URL_REDIRECT_WAIT = 4
POLL_INTERVAL = 2
DEVTOOLS_START_TIMEOUT = 25

# ============================================================
# FIND MICROSOFT EDGE
# ============================================================

EDGE_PATHS = [
    Path(
        os.environ.get("PROGRAMFILES(X86)", "")
    ) / "Microsoft/Edge/Application/msedge.exe",

    Path(
        os.environ.get("PROGRAMFILES", "")
    ) / "Microsoft/Edge/Application/msedge.exe",
]


EDGE_PATH = next(
    (
        path
        for path in EDGE_PATHS
        if path.exists()
    ),
    None
)


if EDGE_PATH is None:
    raise FileNotFoundError(
        "Microsoft Edge executable not found."
    )


# ============================================================
# START EDGE
# ============================================================

print()
print("=======================================")
print("       COOKIE MANAGER V7.1 PY Version")
print("=======================================")
print()

print(f"Edge:    {EDGE_PATH}")
print(f"Profile: {PROFILE}")
print(f"Port:    {DEBUG_PORT}")
print()

PROFILE.mkdir(
    parents=True,
    exist_ok=True
)


subprocess.Popen([
    str(EDGE_PATH),

    f"--remote-debugging-port={DEBUG_PORT}",
    "--remote-allow-origins=http://localhost:9222",
    f"--user-data-dir={PROFILE}",
    "--no-first-run",
    "--no-default-browser-check",
    "--disable-session-crashed-bubble",
    URL_1,
])

print("===================================")
print("Edge launched.")
print("Waiting for DevTools...")


# ============================================================
# WAIT FOR DEVTOOLS
# ============================================================

tabs_url = (
    f"http://localhost:{DEBUG_PORT}/json"
)

tabs = None


for _ in range(DEVTOOLS_START_TIMEOUT * 2):

    try:

        with urllib.request.urlopen(
            tabs_url,
            timeout=2
        ) as response:

            tabs = json.load(response)

        if tabs:
            break

    except Exception:
        pass

    time.sleep(0.5)


if not tabs:
    raise RuntimeError(
        "Edge DevTools did not become available."
    )


# ============================================================
# FIND FIRST PAGE TAB
# ============================================================


page_tabs = [
    tab
    for tab in tabs
    if tab.get("type") == "page"
]


if not page_tabs:
    raise RuntimeError(
        "No page tab found."
    )


tab = page_tabs[0]


print()
print("---------------------------------------")
print("Working Tab")
print("---------------------------------------")
print(
    f"Other tabs ignored: "
    f"{len(page_tabs) - 1}"
)
print()


# ============================================================
# CONNECT TO CDP
# ============================================================

ws_url = tab.get(
    "webSocketDebuggerUrl"
)


if not ws_url:
    raise RuntimeError(
        "Working tab does not have a "
        "WebSocket debugger URL."
    )


ws = websocket.create_connection(
    ws_url,
    timeout=10
)


message_id = 0


# ============================================================
# CDP COMMAND
# ============================================================

def cdp_command(
    method,
    params=None
):

    global message_id

    message_id += 1

    current_id = message_id

    message = {
        "id": current_id,
        "method": method
    }

    if params is not None:
        message["params"] = params

    ws.send(
        json.dumps(message)
    )

    while True:

        raw_response = ws.recv()

        response = json.loads(
            raw_response
        )

        if response.get("id") != current_id:
            continue

        if "error" in response:

            raise RuntimeError(
                f"CDP error in {method}: "
                f"{response['error']}"
            )

        return response

cdp_command("Page.enable")
cdp_command("Runtime.enable")
cdp_command("Network.enable")


# ============================================================
# BRING WORKING TAB TO FRONT
# ============================================================

print(
    "Activating working tab..."
)


try:

    cdp_command(
        "Page.bringToFront"
    )

except Exception as error:

    print(
        f"WARNING: Could not bring tab "
        f"to front: {error}"
    )


# ============================================================
# GET CURRENT URL
# ============================================================

def get_current_url():

    response = cdp_command(
        "Runtime.evaluate",
        {
            "expression":
                "window.location.href",
            "returnByValue": True
        }
    )

    result = (
        response
        .get("result", {})
        .get("result", {})
    )

    return result.get(
        "value",
        ""
    )


# ============================================================
# GET READY STATE
# ============================================================

def get_ready_state():

    response = cdp_command(
        "Runtime.evaluate",
        {
            "expression":
                "document.readyState",
            "returnByValue": True
        }
    )

    result = (
        response
        .get("result", {})
        .get("result", {})
    )

    return result.get(
        "value",
        ""
    )


# ============================================================
# CHECK LOGIN PAGE
# ============================================================

def is_login_page(
    current_url
):

    for login_url in LOGIN_URLS:

        if current_url.startswith(
            login_url
        ):
            return True

    return False


# ============================================================
# NAVIGATE
# ============================================================

def navigate(url):

    print()
    print(">>> NAVIGATE:")
    print(url)
    print()

    try:

        cdp_command(
            "Page.navigate",
            {
                "url": url
            }
        )

    except Exception as error:

        print(
            f"WARNING: Navigation command "
            f"failed: {error}"
        )


# ============================================================
# URL MATCH
# ============================================================

def url_matches(
    current_url,
    expected_url
):

    return current_url.startswith(
        expected_url
    )


# ============================================================
# ENSURE URL 1
# ============================================================

def ensure_url1():

    print()
    print("=======================================")
    print("      WAITING FOR PORTAL URL_1")
    print("=======================================")
    print()

    url1_stable_since = None

    wrong_url_complete_since = None

    while True:

        current_url = get_current_url()

        state = get_ready_state()

        print(
            f"URL:   {current_url}"
        )

        print(
            f"STATE: {state}"
        )


        if is_login_page(
            current_url
        ):

            print(
                "LOGIN PAGE DETECTED."
            )

            print(
                "Waiting for user to complete "
                "the login..."
            )

            print(
                "Automatic navigation is disabled "
                "while on the login page."
            )

            url1_stable_since = None

            wrong_url_complete_since = None

            time.sleep(
                POLL_INTERVAL
            )

            continue

        if url_matches(
            current_url,
            URL_1
        ):


            if state != "complete":

                print(
                    "Portal detected, "
                    "still loading..."
                )

                url1_stable_since = None

                wrong_url_complete_since = None

                time.sleep(
                    POLL_INTERVAL
                )

                continue


            wrong_url_complete_since = None


            if url1_stable_since is None:

                url1_stable_since = (
                    time.time()
                )

                print(
                    "Portal is complete."
                )

                print(
                    "Starting stability check..."
                )

            else:

                stable_time = (
                    time.time()
                    - url1_stable_since
                )

                print(
                    f"Portal stable for "
                    f"{stable_time:.2f}s / "
                    f"{URL1_STABLE_SECONDS:.2f}s"
                )


                if (
                    stable_time
                    >= URL1_STABLE_SECONDS
                ):

                    print()
                    print(
                        "SUCCESS: URL_1 is stable."
                    )
                    print()

                    return True


            time.sleep(
                POLL_INTERVAL
            )

            continue



        url1_stable_since = None


        if state != "complete":

            print(
                "Working tab is not on URL_1, "
                "but the page is still loading."
            )

            print(
                "Waiting for the current navigation "
                "to finish..."
            )

            wrong_url_complete_since = None

            time.sleep(
                POLL_INTERVAL
            )

            continue


        if (
            wrong_url_complete_since
            is None
        ):

            wrong_url_complete_since = (
                time.time()
            )

            print(
                "Working tab is on another URL "
                "and the page is complete."
            )

            print(
                "Waiting "
                f"{WRONG_URL_REDIRECT_WAIT:.2f}s "
                "for automatic redirect..."
            )

        else:

            elapsed = (
                time.time()
                - wrong_url_complete_since
            )

            print(
                f"Automatic redirect wait: "
                f"{elapsed:.2f}s / "
                f"{WRONG_URL_REDIRECT_WAIT:.2f}s"
            )



            if (
                elapsed
                >= WRONG_URL_REDIRECT_WAIT
            ):

                print(
                    "No automatic redirect "
                    "to URL_1 detected."
                )

                print(
                    "Redirecting working tab "
                    "to URL_1..."
                )

                navigate(
                    URL_1
                )

                wrong_url_complete_since = None

                url1_stable_since = None


        time.sleep(
            POLL_INTERVAL
        )


# ============================================================
# ENSURE URL 2
# ============================================================

def ensure_url2():

    print()
    print("=======================================")
    print("      NAVIGATING TO MANIFEST PAGE")
    print("=======================================")
    print()

    wrong_url_complete_since = None

    while True:

        current_url = get_current_url()

        state = get_ready_state()

        print(
            f"URL:   {current_url}"
        )

        print(
            f"STATE: {state}"
        )


        if is_login_page(
            current_url
        ):

            print(
                "LOGIN PAGE DETECTED."
            )

            print(
                "Waiting for user to complete "
                "the login..."
            )

            wrong_url_complete_since = None

            time.sleep(
                POLL_INTERVAL
            )

            continue


        # ====================================================
        # URL_2 SUCCESS
        # ====================================================

        if (
            url_matches(
                current_url,
                URL_2
            )
            and state == "complete"
        ):

            print()
            print(
                "SUCCESS: Manifest Menu loaded."
            )
            print()

            return True


        # ====================================================
        # URL_2 IS LOADING
        # ====================================================

        if url_matches(
            current_url,
            URL_2
        ):

            print(
                "Manifest URL detected, "
                "still loading..."
            )

            wrong_url_complete_since = None

            time.sleep(
                POLL_INTERVAL
            )

            continue


        # ====================================================
        # WRONG URL + LOADING
        # ====================================================

        if state != "complete":

            print(
                "Working tab is not on URL_2, "
                "but the page is still loading."
            )

            print(
                "Waiting for the current navigation "
                "to finish..."
            )

            wrong_url_complete_since = None

            time.sleep(
                POLL_INTERVAL
            )

            continue


        # ====================================================
        # WRONG URL + COMPLETE
        # ====================================================

        if (
            wrong_url_complete_since
            is None
        ):

            wrong_url_complete_since = (
                time.time()
            )

            print(
                "Wrong URL + page complete."
            )

            print(
                "Waiting for automatic redirect..."
            )

        else:

            elapsed = (
                time.time()
                - wrong_url_complete_since
            )

            print(
                f"Automatic redirect wait: "
                f"{elapsed:.2f}s / "
                f"{WRONG_URL_REDIRECT_WAIT:.2f}s"
            )


            if (
                elapsed
                >= WRONG_URL_REDIRECT_WAIT
            ):

                print(
                    "No automatic redirect "
                    "to URL_2 detected."
                )

                print(
                    "Redirecting working tab "
                    "to URL_2..."
                )

                navigate(
                    URL_2
                )

                wrong_url_complete_since = None


        time.sleep(
            POLL_INTERVAL
        )


# ============================================================
# WAIT FOR URL COMPLETE
# ============================================================

def wait_for_url_complete(
    expected_url
):

    print()
    print(
        "Waiting for:"
    )
    print(
        expected_url
    )
    print()


    while True:

        current_url = get_current_url()

        state = get_ready_state()

        print(
            f"URL:   {current_url}"
        )

        print(
            f"STATE: {state}"
        )


        # ----------------------------------------------------
        # Login page
        # ----------------------------------------------------

        if is_login_page(
            current_url
        ):

            print(
                "LOGIN PAGE DETECTED."
            )

            print(
                "Waiting for user to complete "
                "the login..."
            )

            time.sleep(
                POLL_INTERVAL
            )

            continue


        # ----------------------------------------------------
        # Success
        # ----------------------------------------------------

        if (
            url_matches(
                current_url,
                expected_url
            )
            and state == "complete"
        ):

            print()
            print(
                "SUCCESS: Target URL loaded."
            )
            print()

            return True


        # ----------------------------------------------------
        # Correct URL but still loading
        # ----------------------------------------------------

        if url_matches(
            current_url,
            expected_url
        ):

            print(
                "Correct URL, still loading..."
            )

            time.sleep(
                POLL_INTERVAL
            )

            continue


        # ----------------------------------------------------
        # Wrong URL
        # ----------------------------------------------------

        print(
            "Wrong URL. Waiting for navigation..."
        )

        time.sleep(
            POLL_INTERVAL
        )


# ============================================================
# SAVE ALL FEDEX COOKIES
# ============================================================

def save_cookies(
    file_path,
    allowed_domains
):

    response = cdp_command(
        "Network.getAllCookies"
    )

    cookies = (
        response
        .get("result", {})
        .get("cookies", [])
    )

    filtered_cookies = []

    print()
    print("=============================")
    print("Checking COOKIES and Saving")
    print("=============================")
    print()


    for cookie in cookies:

        cookie_domain = (
            cookie
            .get("domain", "")
            .lstrip(".")
        )


        for allowed_domain in allowed_domains:

            allowed_domain = (
                allowed_domain
                .lstrip(".")
            )


            if (
                cookie_domain
                == allowed_domain
                or
                cookie_domain.endswith(
                    "." + allowed_domain
                )
            ):

                filtered_cookies.append(
                    cookie
                )

                break


    cookie_string = ";".join(
        f"{cookie['name']}={cookie['value']}"
        for cookie in filtered_cookies
    )


    with open(
        file_path,
        "w",
        encoding="utf-8"
    ) as file:

        file.write(
            cookie_string
        )


    print(
        f"Total browser cookies: "
        f"{len(cookies)}"
    )

    print(
        f"Filtered cookies: "
        f"{len(filtered_cookies)}"
    )

    print(
        f"Saved to: {file_path}"
    )

    print()


# ============================================================
# SAVE COOKIES FOR SPECIFIC URL
# ============================================================

def save_url_cookies(
    file_path,
    url
):

    response = cdp_command(
        "Network.getCookies",
        {
            "urls": [url]
        }
    )

    cookies = (
        response
        .get("result", {})
        .get("cookies", [])
    )


    cookie_string = ";".join(
        f"{cookie['name']}={cookie['value']}"
        for cookie in cookies
    )


    with open(
        file_path,
        "w",
        encoding="utf-8"
    ) as file:

        file.write(
            cookie_string
        )


    print(
        f"Cookies found for URL: "
        f"{len(cookies)}"
    )

    print(
        f"Saved to: {file_path}"
    )


# ============================================================
# MAIN PROCESS
# ============================================================

try:

    # --------------------------------------------------------
    # STEP 1
    # Get to URL_1.
    #
    # If login page appears, wait indefinitely.
    # --------------------------------------------------------

    print()
    print("STEP 1: Portal")
    print()

    ensure_url1()

    print(
        "Portal Ready..."
    )


    # --------------------------------------------------------
    # STEP 2
    # Navigate to URL_2.
    # --------------------------------------------------------

    print()
    print("=======================================")
    print("       FGC Redirect Manager")
    print("=======================================")
    print()

    print(
        "Starting Portal -> Manifest navigation..."
    )

    navigate(
        URL_2
    )

    ensure_url2()


    # --------------------------------------------------------
    # STEP 3
    # Save FGC cookies.
    # --------------------------------------------------------

    COOKIE_DOMAINS = [
        "fedex.com"
    ]

    save_cookies(
        FGC_COOKIES,
        COOKIE_DOMAINS
    )


    # --------------------------------------------------------
    # STEP 4
    # Navigate to SSO check page.
    # --------------------------------------------------------

    print()
    print("=============================")
    print("Redirecting to SSO Check Page")
    print("=============================")
    print()

    navigate(
        URL_3
    )

    wait_for_url_complete(
        URL_3
    )


    # --------------------------------------------------------
    # STEP 5
    # Save SSO cookies.
    # --------------------------------------------------------

    save_url_cookies(
        SSO_COOKIES,
        URL_3
    )


    # --------------------------------------------------------
    # FINISHED
    # --------------------------------------------------------

    print()
    print("=======================================")
    print("       PROCESS FINISHED SUCCESSFULLY")
    print("=======================================")
    print()

    print(
        f"FGC cookies: {FGC_COOKIES}"
    )

    print(
        f"SSO cookies: {SSO_COOKIES}"
    )

    print()
    print(
        "The working tab will remain open."
    )

    print(
        "Closing in 5 Seconds..."
    )
    
    time.sleep(5)

    print()


except KeyboardInterrupt:

    print()
    print(
        "Process interrupted by user."
    )

    sys.exit(1)


except Exception as error:

    print()
    print("=======================================")
    print("              ERROR")
    print("=======================================")
    print()

    print("Share this with the developer before closing.")
    print()

    print("---------------------------------------")
    print("ERROR MESSAGE")
    print("---------------------------------------")
    print()

    print(error)

    print()

    try:
        print("---------------------------------------")
        print("CURRENT PAGE")
        print("---------------------------------------")
        print()

        print(f"URL:   {get_current_url()}")
        print(f"STATE: {get_ready_state()}")

        print()

    except Exception:
        pass

    print("---------------------------------------")
    print("END ERROR INFORMATION")
    print("---------------------------------------")
    print()

    input(
        "Press ENTER to close this window..."
    )

    sys.exit(1)


finally:

    try:
        ws.close()
    except Exception:
        pass