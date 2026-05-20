#!/usr/bin/env python3
"""
Refresh all data files for the Customer Pulse dashboard.

Usage:
    python3 refresh_data.py          # refresh all sources
    python3 refresh_data.py --intercom --linear  # pick sources

Writes:
    /tmp/support_data.py        (Intercom conversations, last 90 days)
    /tmp/linear_project_issues.py  (Linear in-progress issues)
    /tmp/quo_calls.py           (Quo call transcripts — via MCP; key kept for future)

Then calls generate_web.py and copies to index.html (auto-deploy via GitHub Pages).
"""

import argparse, datetime, json, os, re, sys, time, urllib.request, urllib.error

TODAY = datetime.date.today().isoformat()

# ── API credentials (loaded from .env or environment) ────────────────────────
def _load_env():
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(env_path):
        with open(env_path) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, _, v = line.partition("=")
                    os.environ.setdefault(k.strip(), v.strip())

_load_env()

INTERCOM_TOKEN = os.environ.get("INTERCOM_TOKEN", "")
LINEAR_TOKEN   = os.environ.get("LINEAR_TOKEN", "")
# QUO_TOKEN    = os.environ.get("QUO_TOKEN", "")
# HEYSAM_ORG   = os.environ.get("HEYSAM_ORG", "")  # no public API key yet

REPO_DIR = os.path.dirname(os.path.abspath(__file__))
GEN_SCRIPT = os.path.join(REPO_DIR, "generate_web.py")


# ── Helpers ──────────────────────────────────────────────────────────────────
def http_get(url, headers):
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read())


def http_post(url, headers, body, _retries=2):
    data = json.dumps(body).encode()
    req = urllib.request.Request(url, data=data, headers={**headers, "Content-Type": "application/json"})
    for attempt in range(_retries + 1):
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.loads(r.read())
        except (urllib.error.URLError, OSError) as e:
            if attempt < _retries:
                time.sleep(2 ** attempt)
                continue
            raise


def strip_html(text):
    return re.sub(r"<[^>]+>", "", text or "").strip()


# ── Intercom refresh ─────────────────────────────────────────────────────────
INTERCOM_CATEGORIES = {
    "bug": "Bug", "question": "Question", "billing": "Billing",
    "integration": "Integration",
}

def _categorize_intercom(conv):
    tag_names = [t.get("name", "").lower() for t in conv.get("tags", {}).get("tags", [])]
    for t in tag_names:
        if t in INTERCOM_CATEGORIES:
            return INTERCOM_CATEGORIES[t]
    subject = (conv.get("source", {}).get("subject") or "").lower()
    for k, v in INTERCOM_CATEGORIES.items():
        if k in subject:
            return v
    return "Other"


# Maps email domain → dashboard company name (must match create_deck_v3.py exactly).
# Verified via Intercom MCP email_domain lookups on 2026-05-20.
DOMAIN_TO_COMPANY = {
    "tomlinsongroup.com":    "TOMLINSON",
    "statewidematerials.com": "STATEWIDE MATERIALS",
    "whitcon.com":           "WHITAKER TRANSPORTATION",
    "diamondmaterials.com":  "DIAMOND MATERIALS",
    "cemex.com":             "CEMEX USA",
    "gulfshoretrucking.com": "GULFSHORE TRUCKING LLC",
    "jwgolding.com":         "JW GOLDING",
    "transphos.com":         "TRANS-PHOS INC.",
    "rockontrucks.com":      "ROCK ON TRUCKS",
    "amrize.com":            "AMRIZE: SASK + WINNIPEG",
    "tapani.com":            "TAPANI INC",
    "rpmxconstruction.com":  "RPM xCONSTRUCTION",
    "pjkeating.com":         "PJ KEATING CO",
    "holcim.com":            "HOLCIM - NORTH CENTRAL (FARGO)",
    "buesingcorp.com":       "BUESING CORP",
    "uslm.com":              "UNITED STATES LIME & MINERALS",
    "volkerstevin.ca":       "VOLKER STEVIN CONTRACTING",
}


def _fetch_convs_for_domain(headers, since, domain, canonical_name, by_company):
    """Search conversations where the source author email contains the domain."""
    seen_ids = set()
    starting_after = None
    while True:
        payload = {
            "query": {
                "operator": "AND",
                "value": [
                    {"field": "created_at", "operator": ">", "value": since},
                    {"field": "source.author.email", "operator": "~", "value": f"@{domain}"},
                ],
            },
            "pagination": {"per_page": 50},
        }
        if starting_after:
            payload["pagination"]["starting_after"] = starting_after
        try:
            data = http_post("https://api.intercom.io/conversations/search", headers, payload)
        except (urllib.error.HTTPError, urllib.error.URLError, OSError):
            break
        convs = data.get("conversations", [])
        if not convs:
            break
        for conv in convs:
            conv_id = conv.get("id", "")
            if conv_id in seen_ids:
                continue
            seen_ids.add(conv_id)
            subject = strip_html(
                conv.get("source", {}).get("subject") or
                conv.get("source", {}).get("body") or ""
            )[:120]
            created = datetime.datetime.fromtimestamp(
                conv.get("created_at", 0)
            ).date().isoformat()
            by_company.setdefault(canonical_name, []).append({
                "id": conv_id,
                "date": created,
                "state": conv.get("state", ""),
                "subject": subject,
                "category": _categorize_intercom(conv),
                "url": f"https://app.intercom.com/a/apps/m48souwv/conversations/{conv_id}",
            })
        pages = data.get("pages", {})
        next_cursor = pages.get("next", {})
        starting_after = next_cursor.get("starting_after") if isinstance(next_cursor, dict) else None
        if not starting_after:
            break
        time.sleep(0.15)


def refresh_intercom(days=90):
    print("  Fetching Intercom conversations…")
    since = int((datetime.datetime.now() - datetime.timedelta(days=days)).timestamp())
    headers = {
        "Authorization": f"Bearer {INTERCOM_TOKEN}",
        "Accept": "application/json",
        "Intercom-Version": "2.11",
    }

    by_company = {}
    total = len(DOMAIN_TO_COMPANY)
    for idx, (domain, canonical_name) in enumerate(DOMAIN_TO_COMPANY.items()):
        print(f"    [{idx+1}/{total}] {canonical_name}…", end=" ", flush=True)
        _fetch_convs_for_domain(headers, since, domain, canonical_name, by_company)
        n_convs = len(by_company.get(canonical_name, []))
        print(f"{n_convs} convs")
        time.sleep(0.15)

    # Sort companies alphabetically, conversations newest-first
    out = {}
    for k in sorted(by_company):
        out[k] = sorted(by_company[k], key=lambda x: x["date"], reverse=True)
    total_convs = sum(len(v) for v in out.values())

    lines = [
        f"# Auto-generated support data — refreshed {TODAY}",
        "# Source: Intercom conversations (last 90 days, matched via email domain → contacts)",
        "",
        f"INTERCOM_90D = {json.dumps(out, indent=4, ensure_ascii=False)}",
        "",
        "LINEAR_90D = {}",
        "ALL_INTERCOM = []",
        "ALL_LINEAR = []",
    ]
    path = "/tmp/support_data.py"
    with open(path, "w") as f:
        f.write("\n".join(lines))
    print(f"    → {path}  ({len(out)} companies, {total_convs} convs)")


# ── Linear refresh ───────────────────────────────────────────────────────────
LINEAR_CATEGORIES = {
    "billing": "Billing & Invoicing", "invoice": "Billing & Invoicing",
    "settlement": "Billing & Invoicing", "payment": "Billing & Invoicing",
    "report": "Reporting", "dashboard": "Reporting", "metric": "Reporting",
    "vendor": "Vendor Management", "hauler": "Vendor Management", "driver": "Vendor Management",
    "feature": "Feature Requests", "request": "Feature Requests",
    "mobile": "App / Mobile Issues", "ios": "App / Mobile Issues", "android": "App / Mobile Issues", "app": "App / Mobile Issues",
    "ticket": "Ticket Management", "dispatch": "Ticket Management", "job": "Ticket Management",
    "rate": "Rates & Pricing Issues", "pricing": "Rates & Pricing Issues",
    "login": "Login & Account Access", "sso": "Login & Account Access", "access": "Login & Account Access", "password": "Login & Account Access",
    "driver type": "Driver Type / Role Correction", "role": "Driver Type / Role Correction",
}

CUSTOMER_PATTERNS = [
    "rock on", "jw golding", "gulfshore", "dufferin", "tomlinson", "4m trucking",
    "rpmx", "tilcon", "amrize", "trans-phos", "crh", "us lime", "cemex",
    "pineridge", "werdco", "western", "tapani", "williams trucking",
    "statewide", "holcim", "walker", "igel", "buesing",
]

def _infer_category(title):
    t = title.lower()
    for k, v in LINEAR_CATEGORIES.items():
        if k in t:
            return v
    return "Other"

def _infer_customer(title, project):
    combined = (title + " " + (project or "")).lower()
    for pat in CUSTOMER_PATTERNS:
        if pat in combined:
            words = pat.title()
            return words
    return None

LINEAR_GQL = """
query InProgress($after: String) {
  issues(
    filter: { state: { type: { eq: "started" } } }
    first: 100
    after: $after
    orderBy: updatedAt
  ) {
    pageInfo { hasNextPage endCursor }
    nodes {
      id identifier title
      state { name }
      team { name }
      project { name }
      assignee { name }
      priority
      url
    }
  }
}
"""

def refresh_linear():
    print("  Fetching Linear in-progress issues…")
    headers = {
        "Authorization": LINEAR_TOKEN,
        "Content-Type": "application/json",
    }
    issues = []
    cursor = None
    while True:
        variables = {"after": cursor} if cursor else {}
        try:
            data = http_post("https://api.linear.app/graphql", headers, {"query": LINEAR_GQL, "variables": variables})
        except urllib.error.HTTPError as e:
            print(f"    Linear HTTP {e.code}: {e.read().decode()[:200]}")
            break
        page = data.get("data", {}).get("issues", {})
        nodes = page.get("nodes", [])
        for n in nodes:
            title = n.get("title", "")
            project = (n.get("project") or {}).get("name", "")
            issues.append({
                "id": n.get("identifier", ""),
                "identifier": n.get("identifier", ""),
                "title": title,
                "state": (n.get("state") or {}).get("name", ""),
                "team": (n.get("team") or {}).get("name", ""),
                "project": project,
                "assignee": (n.get("assignee") or {}).get("name", ""),
                "priority": n.get("priority", 0),
                "url": n.get("url", ""),
                "category": _infer_category(title),
                "customer": _infer_customer(title, project),
            })
        pi = page.get("pageInfo", {})
        if not pi.get("hasNextPage"):
            break
        cursor = pi.get("endCursor")
        time.sleep(0.2)

    rows = ",\n    ".join(repr(i) for i in issues)
    path = "/tmp/linear_project_issues.py"
    with open(path, "w") as f:
        f.write(f"# Auto-generated — refreshed {TODAY}\nLINEAR_PROJECT_ISSUES = [\n    {rows},\n]\n")
    print(f"    → {path}  ({len(issues)} issues)")


# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    p = argparse.ArgumentParser(description="Refresh Customer Pulse data and redeploy")
    p.add_argument("--intercom", action="store_true")
    p.add_argument("--linear",   action="store_true")
    p.add_argument("--deploy",   action="store_true", help="regenerate HTML and git push after refresh")
    p.add_argument("--no-deploy", dest="deploy", action="store_false")
    p.set_defaults(deploy=True)
    args = p.parse_args()

    # If no specific source flags given, refresh all
    all_sources = not (args.intercom or args.linear)

    if all_sources or args.intercom:
        refresh_intercom()
    if all_sources or args.linear:
        refresh_linear()

    if args.deploy:
        import subprocess
        out = os.path.join(REPO_DIR, "index.html")
        print(f"\n  Regenerating {out}…")
        result = subprocess.run(
            [sys.executable, GEN_SCRIPT, "--out", out],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            print("  ERROR:", result.stderr[-500:])
            sys.exit(1)
        print(" ", result.stdout.strip())
        print("  Committing and pushing…")
        subprocess.run(["git", "-C", REPO_DIR, "add", "index.html"], check=True)
        subprocess.run([
            "git", "-C", REPO_DIR, "commit", "-m",
            f"Auto-refresh data {TODAY}"
        ], check=True)
        subprocess.run(["git", "-C", REPO_DIR, "push"], check=True)
        print("  Done — GitHub Pages will update in ~30s")


if __name__ == "__main__":
    main()
