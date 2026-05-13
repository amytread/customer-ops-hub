#!/usr/bin/env python3
"""Generate ~/Desktop/tread_customers.html — run: python3 /tmp/generate_web.py"""
import sys, types, json, os, datetime

# ── Stub out pptx / Google deps ───────────────────────────────────────────────
class _Fake:
    def __init__(self, *a, **kw): pass
    def __call__(self, *a, **kw): return _Fake()
    def __getattr__(self, n): return _Fake()
    def __setattr__(self, n, v): object.__setattr__(self, n, v)
    def __getitem__(self, k): return _Fake()
    def __setitem__(self, k, v): pass
    def __iter__(self): return iter([])
    def __len__(self): return 0
    def __bool__(self): return True

for _m in [
    'pptx','pptx.util','pptx.dml','pptx.dml.color',
    'pptx.enum','pptx.enum.text','pptx.oxml','pptx.oxml.ns',
    'lxml','lxml.etree',
    'google_auth_oauthlib','google_auth_oauthlib.flow',
    'googleapiclient','googleapiclient.discovery','googleapiclient.http',
    'google','google.oauth2','google.oauth2.credentials',
    'google.auth','google.auth.transport','google.auth.transport.requests',
]:
    sys.modules[_m] = types.ModuleType(_m)

sys.modules['pptx'].Presentation        = _Fake
sys.modules['pptx.dml.color'].RGBColor  = lambda r,g,b: f'#{r:02X}{g:02X}{b:02X}'
sys.modules['pptx.util'].Inches         = lambda x: x
sys.modules['pptx.util'].Pt             = lambda x: x
sys.modules['pptx.util'].Emu            = lambda x: x
sys.modules['pptx.enum.text'].PP_ALIGN  = types.SimpleNamespace(LEFT=0,CENTER=1,RIGHT=2)
sys.modules['pptx.oxml.ns'].qn          = lambda x: x
sys.modules['pptx.oxml'].OxmlElement    = lambda x: _Fake()
sys.modules['lxml.etree'].SubElement    = lambda *a,**kw: _Fake()
sys.modules['lxml'].etree               = sys.modules['lxml.etree']

with open('/tmp/create_deck_v3.py') as f:
    src = f.read()
src = src[:src.index('# ── SLIDE 1: Cover')]

ns = {
    '__builtins__': __builtins__,
    'os': os, 'io': __import__('io'), 'json': json,
    'ssl': __import__('ssl'), 'urllib': __import__('urllib'),
    'struct': __import__('struct'),
}
exec(src, ns)

mm       = ns['midmarket_companies']
ent      = ns['enterprise_companies']
USAGE    = ns.get('USAGE_STATUS', {})
TENURE   = ns.get('TENURE_APPROX', {})
RENEWAL  = ns.get('RENEWAL_DATES', {})
FEATURES = ns.get('FEATURE_OVERRIDES', {})
INTERCOM = ns.get('INTERCOM_SUPPORT', {})
MAP_COS  = ns.get('_MAP_CUSTOMERS', [])

# ── Driver registration requests (last ~30 days, from #tread-registration-request Slack) ──
# Unique registrant count per Tread customer (as named by drivers on the self-registration form)
REGISTRATION_REQUESTS = [
    {'week': 'Apr 14',
     'RONYX LOGISTICS LLC': 2, 'PJ KEATING CO': 2, 'AMRIZE: SASK + WINNIPEG': 1,
     'VOLKER STEVIN CONTRACTING': 1, 'R.W. DUNTEMAN CO.': 1, 'TOMLINSON': 1,
     'ROCK ON TRUCKS': 1, 'NATIONAL LIME AND STONE': 1, 'QUALITY TRUCKING': 1,
     'TILCON CT INC': 1, 'SILVERKING TRUCKING': 1, 'TERRY EQUIPMENT COMPANY': 1,
    },
    {'week': 'Apr 21',
     'RONYX LOGISTICS LLC': 3, 'PJ KEATING CO': 2, 'AMRIZE: SASK + WINNIPEG': 2,
     'VOLKER STEVIN CONTRACTING': 2, 'R.W. DUNTEMAN CO.': 1, 'TOMLINSON': 1,
     'ROCK ON TRUCKS': 2, 'DUFFERIN AGGREGATES (CRH)': 1, 'NATIONAL LIME AND STONE': 1,
     'STATEWIDE MATERIALS': 1, 'TILCON CT INC': 1, 'QUALITY TRUCKING': 1,
     'MMC MATERIALS INC': 1, 'PINERIDGE FARMS INC.': 1, 'R&R TRUCKING, INC.': 1,
    },
    {'week': 'Apr 28',
     'RONYX LOGISTICS LLC': 4, 'PJ KEATING CO': 3, 'AMRIZE: SASK + WINNIPEG': 3,
     'VOLKER STEVIN CONTRACTING': 2, 'R.W. DUNTEMAN CO.': 2, 'TOMLINSON': 2,
     'ROCK ON TRUCKS': 1, 'DUFFERIN AGGREGATES (CRH)': 2, 'NATIONAL LIME AND STONE': 1,
     'STATEWIDE MATERIALS': 1, 'TILCON CT INC': 1, 'SILVERKING TRUCKING': 1,
     'MARCC TRUCKING': 1, 'PINERIDGE FARMS INC.': 1, 'DIAMOND MATERIALS': 1,
     'BUESING CORP': 1, 'MANSTEEL REBAR LTD.': 1, 'TRIO AGGREGATE HAULERS': 1,
    },
    {'week': 'May 5',
     'RONYX LOGISTICS LLC': 3, 'PJ KEATING CO': 2, 'AMRIZE: SASK + WINNIPEG': 2,
     'VOLKER STEVIN CONTRACTING': 2, 'R.W. DUNTEMAN CO.': 2, 'TOMLINSON': 1,
     'ROCK ON TRUCKS': 1, 'DUFFERIN AGGREGATES (CRH)': 1, 'NATIONAL LIME AND STONE': 1,
     'STATEWIDE MATERIALS': 1, 'QUALITY TRUCKING': 1, 'SILVERKING TRUCKING': 1,
     'MMC MATERIALS INC': 1, 'MARCC TRUCKING': 1, 'IROQUOIS BAR CORPORATION': 1,
     'RPM xCONSTRUCTION': 1, 'UNITED STATES LIME & MINERALS': 1,
    },
]

# ── Intercom stub categories ──────────────────────────────────────────────────
# IDs present in INTERCOM_SUPPORT but absent from INTERCOM_90D get stub entries.
# This dict maps each such ID to (subject, category) so they're not all "Other".
INTERCOM_STUB_CATS = {
    '215474033701302': ('Material rate selection broken on completed orders', 'Rates & Pricing Issues'),
    '215474181334923': ('Drivers can upload tickets to canceled jobs', 'Feature Requests'),
    '215473991016651': ('GPS tracking stopped after 3rd load', 'App / Mobile Issues'),
    '215473944055565': ('Cannot add new equipment type on connected vendor', 'Vendor Management'),
    '215473890010511': ('Freight rates and CSV update status', 'Rates & Pricing Issues'),
    '215474197837742': ('APEX integration tickets corrupting totals', 'Vendor Management'),
    '215473856491965': ('App loading very slowly when scrolling', 'App / Mobile Issues'),
    '215473866538222': ('Reporting an issue', 'App / Mobile Issues'),
    '215474003597339': ('How to edit missed break/shift time', 'Feature Requests'),
    '215474011047073': ('Onboarding setup email for RW Dunteman', 'Add / Onboard Driver'),
    '215473931769979': ('Escalation: $65K license, unresolved bugs', 'App / Mobile Issues'),
    '215474184671998': ('MMC account consolidation and shift hours report', 'Add / Onboard Driver'),
    '215474224958023': ('Drivers cannot connect to truck in Mansteel account', 'App / Mobile Issues'),
    '215474197049557': ('AR/AP-v3 report missing tickets', 'Reporting'),
    '215474151909814': ('Vendor rate cards not working', 'Rates & Pricing Issues'),
    '215474142675333': ('Dispatch text showing wrong load quantity', 'Ticket Management'),
    '215474083972196': ('Dispatched drivers not showing job count', 'Ticket Management'),
    '215474003749617': ('Pre-loads disappearing at midnight', 'Ticket Management'),
    '215473850527530': ('Prevailing wage custom report for OT hours', 'Reporting'),
    '215473853744955': ('Timezone wrong in SMS dispatch notifications', 'Ticket Management'),
    '215473946098787': ('Downloaded ticket PDFs showing blank pages', 'Reporting'),
    '215474230667370': ('Driver setup across multiple trucks in vendor account', 'Add / Onboard Driver'),
    '215474206679744': ('Tapani CSM intro and account update', 'Add / Onboard Driver'),
    '215473859191755': ('Old Ruckit platform invite email noise', 'No Action Needed'),
    '215474010729756': ('Geofence circle stays in wrong location', 'App / Mobile Issues'),
    '215474227880490': ('Unable to create vendor - driver already in system', 'Vendor Management'),
    '215474223429030': ('Something Went Wrong loading session', 'Login & Account Access'),
    '215474027908029': ('Civic address input issue on job lines', 'Ticket Management'),
    '215474022268366': ('Civic address input bug on Edit Job', 'Ticket Management'),
    '215474196744100': ('Adding subcontractor truck to account', 'Vendor Management'),
    '215474007719515': ('Digital ticket loads stuck in planned stage', 'Ticket Management'),
}

# ── Load 90-day support data ──────────────────────────────────────────────────
sys.path.insert(0, '/tmp')
try:
    from support_data import INTERCOM_90D as _IC90_RAW, LINEAR_90D, ALL_INTERCOM, ALL_LINEAR
except ImportError:
    _IC90_RAW = LINEAR_90D = {}
    ALL_INTERCOM = ALL_LINEAR = []

# Remap Intercom company names → canonical deck names (uppercase match + explicit aliases)
_IC90_ALIAS = {
    'ADVANCED HAULING SOLUTIONS, LLC.': 'AHS',
    'AMRIZE NCR - FARGO MOORHEAD': 'AMRIZE: NCR-TWIN CITIES',
    'AMRIZE SASKATOON': 'AMRIZE: SASK + WINNIPEG',
    'IROQUOIS BAR CORP.': 'IROQUOIS BAR CORPORATION',
    'RONYX LOGISTICS': 'RONYX LOGISTICS LLC',
    'TRANS-PHOS': 'TRANS-PHOS INC.',
    'QUALITY TRUCKING - LR': 'QUALITY TRUCKING',
    'TWIN CITY HAULING LLC': 'TWIN CITY HAULING',
    'PJ KEATING CO.': 'PJ KEATING CO',
    'PJ KEATING (A CRH COMPANY)': 'PJ KEATING CO',
    'TILCON CONNECTICUT. A CRH COMPANY': 'TILCON CT INC',
    'TILCON CT': 'TILCON CT INC',
    'MMC MATERIALS': 'MMC MATERIALS INC',
    'SILVERKING TRUCKING LLC': 'SILVERKING TRUCKING',
    'DANIELA TRUCKING': 'DANIELA TRUCKING & GRADING',
    'FRANCISCO  PEREZ': 'FRANCISCO PEREZ',
    'CARTER TRUCKING': 'CHARLES H CARTER & SON',
    'VOLKERwessels CANADA LTD.': 'VOLKER STEVIN CONTRACTING',
    'VOLKERwessels CANADA': 'VOLKER STEVIN CONTRACTING',
    'NS TRUCKING': 'N.S. TRUCKING INC.',
    'THUNDERBOLT CONTRACTING LTD.': 'THUNDERBOLT',
    'THUNDERBOLT CONTRACTING': 'THUNDERBOLT',
    'WALKER AG GROUP': 'WALKER AG GROUP',
    'PRIME AGGREGATE TRANSPORTATION': 'PRIME AGGREGATE TRANSPORTATION',
}
INTERCOM_90D = {}
for _k, _v in _IC90_RAW.items():
    _ku = _k.upper()
    INTERCOM_90D[_IC90_ALIAS.get(_ku, _ku)] = _v

_NO_ACTION_SUBJECTS = frozenset([
    'i have another question', 'i have a question', 'hello', 'hi',
    'good morning', 'good afternoon', 'good evening',
    'help', 'necesito ayuda', 'buenas tardes', 'buenas trades',
    'good morning,', 'good afternoon,', 'good evening,',
    'notifications', 'notification',
    'easy trucking', 'is marco available please?', 'is marco available',
    "can't make it", "can't make it, or will be late", "will be late",
    "i'll be late", "running late",
])

_CAT_REMAP = {
    'Login/Access':   'Login & Account Access',
    'Dispatch/Jobs':  'Ticket Management',
    'Driver App':     'App / Mobile Issues',
    'GPS/Tracking':   'App / Mobile Issues',
    'Billing':        'Billing & Invoicing',
    'Feature Request':'Feature Requests',
    'Onboarding':     'Add / Onboard Driver',
    'Integration':    'Vendor Management',
    'Performance':    'App / Mobile Issues',
    'Question':       'Other',
    'Bug':            'Other',  # bugs get reclassified by functional area in _recat; fallback to Other
}

def _recat(subject, original_cat='Other'):
    # Always remap legacy category names first
    original_cat = _CAT_REMAP.get(original_cat, original_cat)
    raw = (subject or '').strip()
    s = raw.lower()
    # Blank or placeholder subjects — no support content
    if not s:
        return 'No Action Needed'
    if len(s) < 8:
        return original_cat
    if s in _NO_ACTION_SUBJECTS or s.rstrip('.,!? ') in _NO_ACTION_SUBJECTS:
        return 'No Action Needed'
    if any(s.startswith(g) for g in ('good morning', 'good afternoon', 'good evening')) and len(s) < 28:
        return 'No Action Needed'
    # Login & Account Access
    if any(w in s for w in [
        'login', 'log in', 'password', 'sign in', 'credential', 'change email',
        'email address', 'reset', 'forgot', 'locked out',
        "can't log", 'cant log', 'cannot log',
        'already in use', 'number is already', 'already taken', 'already used',
        'said his name has alread', 'name has alread', 'name is alread',
        'portal access', 'account access',
        'registration', 'verification code', 'verification link',
        "didn't receive", 'did not receive', "haven't received", 'havent received',
        'receive my code', 'receive my link', 'receive my invite',
    ]):
        return 'Login & Account Access'
    # Ticket Management — receiving, assigning, scheduling, or managing jobs/tickets/orders
    if any(w in s for w in [
        'have not received a job', 'not received a job',
        'forward work', 'no job yet', 'pressed start',
        'adding a ticket', 'add a ticket', 'adding ticket', 'help adding a ticket',
        'change the customer on', 'change customer on', 'changing customer',
        'drag a project', 'end a project', 'end the project', 'end project',
        'job says', 'job status', 'says ended', 'says completed',
        'best practice for dispatch', 'dispatching hourly', 'hourly job',
        'customers on orders', 'customer on an order', 'customer on order',
        'dispatch a job', 'assign a job', 'job not',
        'canceled job', 'cancelled job', 'cancel job', 'cancel order',
        'proceso', 'trabajo de', 'lugar de carga',
    ]):
        return 'Ticket Management'
    # Reporting
    if any(w in s for w in ['report', 'export', 'spreadsheet', 'csv', 'payroll', 'mileage', 'milage']):
        return 'Reporting'
    # Vendor Management — vendor/subhauler management and system integrations
    if any(w in s for w in [
        'apex', 'integration', 'api', 'sync', 'import', 'interlock', 'quickbooks',
        'add vendor', 'add subhauler', 'subhauler', 'sub-hauler', 'subcontractor',
        'vendor account', 'add a vendor', 'new vendor', 'create vendor',
        'vendor name', 'vendor switch', 'vendor fix', 'vendor and driver',
        'vendor ', 'vendor,',
    ]):
        return 'Vendor Management'
    # Rates & Pricing Issues — rate configuration, pricing, fuel surcharges
    if any(w in s for w in [
        'rate card', 'rate change', 'pay rate', 'pricing', 'price per',
        'fuel surcharge', 'fsc', 'double and triple', 'triple fsc',
        'freight rate', 'haul rate', 'ton rate', 'per ton', 'per load rate',
    ]):
        return 'Rates & Pricing Issues'
    # Billing & Invoicing — post-job payments, invoices, settlements
    if any(w in s for w in [
        'billing', 'payment', 'invoice', 'charge', 'settlement',
        'amount to be paid', 'total amount', 'amount paid', 'pay on here',
    ]):
        return 'Billing & Invoicing'
    # Add / Onboard Driver — adding individual drivers and new user setup
    if any(w in s for w in [
        'add a driver', 'new driver', 'adding driver', 'add driver',
        'add their driver', 'add new driver', 'add this number',
        'as a driver', 'him as a driver',
        'change driver', 'driver to truck', 'driver number',
        'nuevo driver', 'nuevo chofer', 'un nuevo', 'tango un',
        'setup', 'set up', 'onboard', 'invite', 'configure', 'add user', 'new user',
    ]):
        return 'Add / Onboard Driver'
    # Driver Type / Role Correction
    if any(w in s for w in [
        'driver type', 'driver role', 'change role', 'wrong role', 'wrong type',
        'owner operator', 'owner-operator', 'change to owner', 'as a broker',
    ]):
        return 'Driver Type / Role Correction'
    # App / Mobile Issues — GPS, mobile app, device, performance
    if any(w in s for w in [
        'gps', 'location', 'map view', 'tracking', 'track driver',
        'app', 'mobile', 'phone', 'android', 'ios',
        'slow', 'performance',
        'notification', 'getting this message', 'explain what it means',
        'what does this mean', 'what does this message',
    ]):
        return 'App / Mobile Issues'
    # Feature Requests
    if any(w in s for w in [
        'feature', 'suggestion', 'would like', 'can we', 'could you add',
        'option to', 'ability to', 'custom report', 'update report',
    ]):
        return 'Feature Requests'
    # Bug signals — route to the functional area the bug is in
    if any(w in s for w in [
        'error', 'broken', 'not working', 'not showing', 'not appear', 'unable',
        "can't see", 'cannot see', 'missing', 'disappeared', 'wrong', 'incorrect',
        "can't find", 'not visible', 'not loading', 'offline', 'off-line',
        "won't work", 'wont work', "doesn't work", 'does not work',
        "can't connect", 'cannot connect',
        'no me deja', 'no me aparece', 'no funciona', 'no puedo',
        'clicked any boxes differently', 'something changed', 'acting strange',
    ]):
        if any(w in s for w in ['login', 'log in', 'password', 'sign in', 'account', 'access', 'session', 'verification']):
            return 'Login & Account Access'
        if any(w in s for w in ['report', 'export', 'csv', 'pdf', 'spreadsheet', 'payroll']):
            return 'Reporting'
        if any(w in s for w in ['ticket', 'dispatch', 'job', 'order', 'load', 'civic', 'address', 'pre-load', 'preload', 'timezone', 'sms']):
            return 'Ticket Management'
        if any(w in s for w in ['rate', 'pricing', 'price', 'material rate', 'rate card']):
            return 'Rates & Pricing Issues'
        if any(w in s for w in ['invoice', 'billing', 'payment', 'settlement']):
            return 'Billing & Invoicing'
        if any(w in s for w in ['vendor', 'subhauler', 'integration', 'equipment', 'apex']):
            return 'Vendor Management'
        if any(w in s for w in ['gps', 'location', 'map', 'track', 'app', 'mobile', 'phone', 'connect', 'truck', 'driver', 'offline', 'slow']):
            return 'App / Mobile Issues'
        if any(w in s for w in ['driver type', 'role']):
            return 'Driver Type / Role Correction'
        return 'App / Mobile Issues'  # default for unclassifiable bugs
    # Weak question signal
    if any(w in s for w in [
        'how ', 'what is', 'help me', 'i need help', 'why ',
        'understand', 'explain', 'wondering', 'question',
    ]):
        return original_cat if original_cat not in ('Other', 'Question') else 'Other'
    return original_cat

for _company, _convs in INTERCOM_90D.items():
    for _conv in _convs:
        _conv['category'] = _recat(_conv.get('subject', ''), _conv.get('category', 'Other'))

try:
    from linear_project_issues import LINEAR_PROJECT_ISSUES, LINEAR_PROJECT_META
except ImportError:
    LINEAR_PROJECT_ISSUES = LINEAR_PROJECT_META = {}

try:
    from heysam_data import HEYSAM_CALLS
except ImportError:
    HEYSAM_CALLS = {}

# Dispatch load counts per customer — last 90 days (from Omni Horizon RDS, 2026-05-11)
DISPATCH_LOADS = {
    # Enterprise
    "AMRIZE: SASK + WINNIPEG":      46786,
    "AMRIZE: NCR-TWIN CITIES":      20821,
    "AMRIZE: GVA (BC)":             29926,
    "AMRIZE: GTA":                  54321,
    "CEMEX USA":                   121550,
    "DUFFERIN AGGREGATES (CRH)":    41413,
    "HOLCIM - NORTH CENTRAL (FARGO)": 20821,
    "NATIONAL LIME AND STONE":      17529,
    "TOMLINSON":                    48355,
    "TRANS-PHOS INC.":             119523,
    "WHITAKER TRANSPORTATION":      32670,
    "ZEMBA INC.":                    4660,
    # Mid-market
    "4M TRUCKING":                  10153,
    "AHS":                          26167,
    "ARIZONA AGGREGATE SOLUTIONS":   4398,
    "BUESING CORP":                 13297,
    "CANTON CONCRETE (DUPLICATE)":    776,
    "DANIELA TRUCKING & GRADING":    1724,
    "DIAMOND MATERIALS":            15958,
    "EPIC MATERIALS INC":            1324,
    "FLASH TRUCKING / GOLF AGRONOMICS": 1522,
    "GERNATT ASPHALT PRODUCTS":     11839,
    "GRANITE CONSTRUCTION (SOCAL)":   431,
    "IROQUOIS BAR CORPORATION":      7199,
    "JW GOLDING":                    6856,
    "MANSTEEL REBAR LTD.":            830,
    "MARCC TRUCKING":                3693,
    "MMC MATERIALS INC":            20283,
    "N.S. TRUCKING INC.":           32754,
    "PINERIDGE FARMS INC.":         11308,
    "PRINCE GEORGE AG":              9373,
    "QUALITY TRUCKING":             83314,
    "RHINO TRUCKING INC.":           1566,
    "ROCK ON TRUCKS":                6307,
    "RONYX LOGISTICS LLC":          31310,
    "RPM xCONSTRUCTION":            87495,
    "R.W. DUNTEMAN CO.":             2424,
    "SILVERKING TRUCKING":           9022,
    "STATEWIDE MATERIALS":         128984,
    "TAPANI INC":                    7305,
    "TILCON CT INC":                16638,
    "TOP TIER TRUCKING":            22008,
    "TRIO AGGREGATE HAULERS":       12650,
    "TWIN CITY HAULING":             3734,
    "UNITED STATES LIME & MINERALS": 14562,
    "WESTERN STATES CONTRACTING":   20033,
    "WILLIAMS TRUCKING CO.":        28541,
    "D CRUPI & SONS, INC.":          4101,
    "GULFSHORE TRUCKING LLC":       38801,
    "PJ KEATING CO":                 8364,
    "R&R TRUCKING, INC.":           21363,
    "UPPAL TRUCKING LTD":           13141,
    "WERDCO BC INC.":               45293,
}

# Dispatch detail — monthly volume (Jun 2025–May 2026, 12 months) + normalized material mix
# months: [Jun, Jul, Aug, Sep, Oct, Nov, Dec, Jan, Feb, Mar, Apr, May]  (May is partial ~11 days)
DISPATCH_DETAIL = {
    # ── Enterprise ────────────────────────────────────────────────────────────
    "CEMEX USA": {
        "months": [822, 837, 767, 513, 2525, 6171, 37990, 38586, 40057, 41118, 41092, 14635],
        "materials": [["Base/Aggregate", 55], ["Stone/Rock", 25], ["Sand", 12], ["Other", 8]],
    },
    "DUFFERIN AGGREGATES (CRH)": {
        "months": [24431, 34650, 30507, 32931, 31749, 26900, 13671, 5924, 7604, 11308, 16973, 9272],
        "materials": [["Base/Aggregate", 55], ["Stone/Rock", 30], ["Sand", 10], ["Other", 5]],
    },
    "NATIONAL LIME AND STONE": {
        "months": [5154, 5631, 5362, 5105, 5228, 4260, 1970, 3066, 4913, 5398, 5715, 2932],
        "materials": [["Stone/Rock", 60], ["Base/Aggregate", 30], ["Other", 10]],
    },
    "TOMLINSON": {
        "months": [23189, 22506, 22413, 24471, 21179, 15853, 12909, 12133, 19300, 14684, 15782, 5793],
        "materials": [["Other", 51], ["Base/Aggregate", 29], ["Stone/Rock", 20]],
    },
    "TRANS-PHOS INC.": {
        "months": [17360, 19877, 19701, 42991, 53103, 38177, 13868, 18292, 18079, 19129, 18710, 6955],
        "materials": [["Fill/Dirt", 39], ["Other", 35], ["Base/Aggregate", 26]],
    },
    "WHITAKER TRANSPORTATION": {
        "months": [5802, 5221, 7226, 8789, 7411, 5036, 5482, 5159, 4980, 5857, 5745, 1943],
        "materials": [["Fill/Dirt", 39], ["Base/Aggregate", 21], ["Other", 22], ["Stone/Rock", 9], ["Sand", 9]],
    },
    "ZEMBA INC.": {
        "months": [1382, 1675, 1452, 1564, 1843, 1492, 1411, 1389, 1599, 1627, 1893, 0],
        "materials": [],
    },
    # ── Mid-Market ────────────────────────────────────────────────────────────
    "4M TRUCKING": {
        "months": [0, 0, 0, 0, 544, 2507, 1787, 2145, 3017, 2952, 3429, 2244],
        "materials": [["Base/Aggregate", 57], ["Fill/Dirt", 21], ["Asphalt", 14], ["Stone/Rock", 8]],
    },
    "ARIZONA AGGREGATE SOLUTIONS": {
        "months": [0, 0, 0, 0, 0, 0, 0, 0, 63, 1151, 2565, 649],
        "materials": [["Base/Aggregate", 70], ["Stone/Rock", 20], ["Other", 10]],
    },
    "BUESING CORP": {
        "months": [0, 0, 0, 0, 0, 0, 24, 1262, 366, 4529, 6743, 2239],
        "materials": [["Fill/Dirt", 50], ["Base/Aggregate", 35], ["Other", 15]],
    },
    "CANTON CONCRETE (DUPLICATE)": {
        "months": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 376, 461],
        "materials": [],
    },
    "DANIELA TRUCKING & GRADING": {
        "months": [250, 678, 544, 572, 712, 505, 511, 413, 494, 574, 551, 230],
        "materials": [["Fill/Dirt", 65], ["Base/Aggregate", 25], ["Other", 10]],
    },
    "DIAMOND MATERIALS": {
        "months": [0, 0, 5, 6209, 12403, 9673, 3962, 3119, 2442, 3937, 7536, 3904],
        "materials": [["Stone/Rock", 40], ["Fill/Dirt", 32], ["Base/Aggregate", 19], ["Sand", 7], ["Asphalt", 2]],
    },
    "EPIC MATERIALS INC": {
        "months": [0, 0, 0, 0, 0, 0, 0, 14, 365, 500, 446, 108],
        "materials": [["Stone/Rock", 64], ["Base/Aggregate", 27], ["Asphalt", 8], ["Fill/Dirt", 1]],
    },
    "FLASH TRUCKING / GOLF AGRONOMICS": {
        "months": [639, 837, 949, 930, 807, 477, 507, 528, 593, 632, 623, 293],
        "materials": [],
    },
    "GERNATT ASPHALT PRODUCTS": {
        "months": [3815, 4428, 11402, 10748, 8776, 5314, 2785, 2007, 1858, 1872, 6351, 2671],
        "materials": [["Asphalt", 60], ["Base/Aggregate", 25], ["Fill/Dirt", 15]],
    },
    "GRANITE CONSTRUCTION (SOCAL)": {
        "months": [2031, 0, 39, 274, 832, 35, 0, 0, 28, 202, 215, 0],
        "materials": [["Base/Aggregate", 60], ["Stone/Rock", 25], ["Other", 15]],
    },
    "GULFSHORE TRUCKING LLC": {
        "months": [0, 0, 0, 0, 1139, 268, 149, 6906, 16149, 13539, 10402, 4489],
        "materials": [["Fill/Dirt", 70], ["Base/Aggregate", 13], ["Stone/Rock", 10], ["Sand", 7]],
    },
    "IROQUOIS BAR CORPORATION": {
        "months": [0, 0, 5, 43, 768, 556, 636, 5574, 2590, 1610, 2365, 2539],
        "materials": [],
    },
    "JW GOLDING": {
        "months": [0, 66, 149, 401, 197, 59, 1150, 1728, 1694, 2277, 2663, 739],
        "materials": [["Stone/Rock", 74], ["Fill/Dirt", 12], ["Other", 8], ["Asphalt", 6]],
    },
    "MANSTEEL REBAR LTD.": {
        "months": [416, 375, 295, 316, 308, 257, 209, 189, 194, 258, 292, 183],
        "materials": [],
    },
    "MARCC TRUCKING": {
        "months": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1281, 2937],
        "materials": [["Stone/Rock", 42], ["Sand", 37], ["Base/Aggregate", 11], ["Other", 10]],
    },
    "PINERIDGE FARMS INC.": {
        "months": [0, 0, 0, 0, 4, 2885, 2701, 3428, 2994, 3506, 4209, 1614],
        "materials": [],
    },
    "PJ KEATING CO": {
        "months": [0, 0, 0, 0, 0, 0, 0, 0, 61, 450, 5023, 3029],
        "materials": [["Stone/Rock", 50], ["Base/Aggregate", 35], ["Asphalt", 15]],
    },
    "PRINCE GEORGE AG": {
        "months": [0, 896, 4377, 7831, 11134, 8936, 7514, 10887, 1347, 2313, 2759, 4305],
        "materials": [],
    },
    "QUALITY TRUCKING": {
        "months": [1807, 7508, 6456, 8820, 5682, 6488, 14297, 8566, 16136, 14510, 12481, 4015],
        "materials": [["Base/Aggregate", 85], ["Other", 15]],
    },
    "RHINO TRUCKING INC.": {
        "months": [0, 9, 0, 0, 0, 0, 0, 1, 176, 465, 720, 222],
        "materials": [],
    },
    "ROCK ON TRUCKS": {
        "months": [0, 0, 0, 0, 0, 0, 0, 3, 0, 46, 3225, 3159],
        "materials": [["Base/Aggregate", 52], ["Sand", 15], ["Fill/Dirt", 11], ["Stone/Rock", 11], ["Other", 11]],
    },
    "RONYX LOGISTICS LLC": {
        "months": [0, 0, 0, 0, 0, 0, 0, 7, 5135, 8840, 11896, 6413],
        "materials": [["Sand", 40], ["Other", 30], ["Stone/Rock", 20], ["Base/Aggregate", 10]],
    },
    "RPM xCONSTRUCTION": {
        "months": [2899, 9143, 10278, 15192, 9089, 7502, 19135, 24393, 33436, 33545, 22886, 9800],
        "materials": [["Base/Aggregate", 70], ["Fill/Dirt", 20], ["Other", 10]],
    },
    "R.W. DUNTEMAN CO.": {
        "months": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 639, 1814],
        "materials": [],
    },
    "R&R TRUCKING, INC.": {
        "months": [8, 0, 0, 0, 0, 0, 0, 19, 1233, 4586, 4191, 1472],
        "materials": [["Fill/Dirt", 60], ["Base/Aggregate", 30], ["Other", 10]],
    },
    "SILVERKING TRUCKING": {
        "months": [681, 2870, 843, 436, 924, 337, 1150, 2080, 1018, 1417, 5931, 1632],
        "materials": [],
    },
    "STATEWIDE MATERIALS": {
        "months": [64162, 62297, 45843, 46126, 45919, 38530, 35195, 25924, 27254, 46419, 46573, 19207],
        "materials": [["Base/Aggregate", 69], ["Fill/Dirt", 31]],
    },
    "TAPANI INC": {
        "months": [2697, 2574, 2663, 2735, 2731, 2437, 2346, 2263, 2206, 2501, 2771, 640],
        "materials": [],
    },
    "TILCON CT INC": {
        "months": [14350, 12389, 13306, 12928, 16223, 9679, 2122, 1301, 1012, 2098, 8389, 6038],
        "materials": [["Base/Aggregate", 50], ["Stone/Rock", 30], ["Asphalt", 15], ["Other", 5]],
    },
    "TOP TIER TRUCKING": {
        "months": [2821, 8939, 8844, 8616, 9330, 6641, 6246, 4789, 7406, 7965, 6928, 2202],
        "materials": [["Sand", 42], ["Stone/Rock", 42], ["Other", 14], ["Base/Aggregate", 2]],
    },
    "TRIO AGGREGATE HAULERS": {
        "months": [0, 0, 0, 0, 0, 0, 0, 394, 2333, 4341, 4608, 1953],
        "materials": [["Base/Aggregate", 65], ["Stone/Rock", 25], ["Other", 10]],
    },
    "TWIN CITY HAULING": {
        "months": [0, 0, 0, 0, 0, 0, 146, 705, 930, 1102, 1390, 642],
        "materials": [],
    },
    "UNITED STATES LIME & MINERALS": {
        "months": [0, 0, 0, 0, 0, 0, 113, 167, 274, 4833, 6501, 3447],
        "materials": [["Stone/Rock", 55], ["Base/Aggregate", 30], ["Other", 15]],
    },
    "UPPAL TRUCKING LTD": {
        "months": [3910, 6081, 4137, 4031, 4859, 4328, 2753, 3521, 3832, 3720, 5927, 1534],
        "materials": [],
    },
    "WERDCO BC INC.": {
        "months": [14, 4678, 20002, 22102, 22224, 13471, 17296, 18581, 15710, 15906, 15060, 4591],
        "materials": [["Fill/Dirt", 83], ["Base/Aggregate", 13], ["Asphalt", 4]],
    },
    "WESTERN STATES CONTRACTING": {
        "months": [5068, 4613, 4177, 5155, 5751, 4203, 5167, 4710, 5996, 7507, 6125, 2052],
        "materials": [["Base/Aggregate", 55], ["Fill/Dirt", 30], ["Other", 15]],
    },
    "WILLIAMS TRUCKING CO.": {
        "months": [9334, 10200, 11485, 12253, 9278, 5060, 9470, 12567, 12700, 13514, 6580, 436],
        "materials": [["Fill/Dirt", 70], ["Base/Aggregate", 23], ["Stone/Rock", 4], ["Other", 3]],
    },
    "D CRUPI & SONS, INC.": {
        "months": [0, 14, 107, 0, 0, 0, 0, 97, 269, 116, 1403, 2756],
        "materials": [],
    },
    "D CRUPI & SONS": {
        "months": [0, 14, 107, 0, 0, 0, 0, 97, 269, 116, 1403, 2756],
        "materials": [],
    },
    "AHS": {
        "months": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1],
        "materials": [],
    },
    "AMRIZE: SASK + WINNIPEG": {
        "months": [9893, 13118, 30090, 59390, 105553, 27916, 4066, 9741, 14195, 10555, 3469, 1659],
        "materials": [["Sand", 60], ["Stone/Rock", 23], ["Base/Aggregate", 14], ["Other", 3]],
    },
    "MMC MATERIALS INC": {
        "months": [0, 0, 0, 0, 0, 0, 0, 0, 22, 866, 5599, 4063],
        "materials": [["Stone/Rock", 58], ["Sand", 31], ["Other", 11]],
    },
    "PETERSON COMPANIES": {
        "months": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 95],
        "materials": [["Base/Aggregate", 99], ["Sand", 1]],
    },
    "TERRY EQUIPMENT COMPANY": {
        "months": [0, 0, 0, 0, 0, 0, 0, 0, 3, 0, 0, 370],
        "materials": [],
    },
    "CHARLES H CARTER & SON": {
        "months": [5, 7, 3, 0, 25, 402, 1828, 1315, 1634, 2030, 2132, 895],
        "materials": [["Base/Aggregate", 57], ["Other", 33], ["Stone/Rock", 10]],
    },
    "WALKER AG GROUP": {
        "months": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 156],
        "materials": [["Base/Aggregate", 100]],
    },
}

def ser(o):
    if isinstance(o, (list, tuple)): return [ser(x) for x in o]
    if isinstance(o, dict): return {k: ser(v) for k, v in o.items()}
    return o

import base64 as _b64
LOGO_PATHS = ns.get('LOGO_PATHS', {})
LOGO_PATHS.update({
    'CEMEX USA':                      '/tmp/logos/cemex_usa.img',
    'NATIONAL LIME AND STONE':        '/tmp/logos/national_lime.img',
    'TILCON CT INC':                  '/tmp/logos/tilcon.img',
    'HOLCIM - NORTH CENTRAL (FARGO)': '/tmp/logos/holcim.img',
    'CERUTTI & SONS TRANSPORTATION':  '/tmp/logos/cerutti.img',
    'CHARLES H CARTER & SON':         '/tmp/logos/carter.img',
    'JW GOLDING':                     '/tmp/logos/jw_golding.img',
    'TOP TIER TRUCKING':              '/tmp/logos/top_tier.img',
    'TRIO AGGREGATE HAULERS':         '/tmp/logos/trio_aggregate.img',
    'R&R TRUCKING, INC.':             '/tmp/logos/rr_trucking2.img',
    'UPPAL TRUCKING LTD':             '/tmp/logos/uppal.img',
    'BUESING CORP':                   '/tmp/logos/buesing.img',
    'GRANITE CONSTRUCTION (SOCAL)':   '/tmp/logos/granite_construction.img',
    'STATEWIDE MATERIALS':            '/tmp/logos/statewide_materials.img',
    'TAPANI INC':                     '/tmp/logos/tapani.img',
    'GERNATT ASPHALT PRODUCTS':       '/tmp/logos/gernatt.img',
    'GEORGE J. IGEL & CO.':           '/tmp/logos/igel.img',
    'PJ KEATING CO':                  '/tmp/logos/pj_keating.img',
    'WERDCO BC INC.':                 '/tmp/logos/werdco.img',
    'DIAMOND MATERIALS':              '/tmp/logos/diamond_materials.img',
    'RPM xCONSTRUCTION':              '/tmp/logos/rpm_construction.img',
    'QUALITY TRUCKING':               '/tmp/logos/quality_trucking.img',
    'GULFSHORE TRUCKING LLC':         '/tmp/logos/gulfshore_trucking.img',
    'D CRUPI & SONS, INC.':           '/tmp/logos/d_crupi.img',
    'TOMLINSON':                      '/tmp/logos/tomlinson.img',
    'TRANS-PHOS INC.':                '/tmp/logos/trans_phos.img',
    'ZEMBA INC.':                     '/tmp/logos/zemba.img',
    'VOLKER STEVIN CONTRACTING':      '/tmp/logos/volker_stevin.img',
    'AMRIZE: SASK + WINNIPEG':        '/tmp/logos/amrize.img',
    'AMRIZE: NCR-TWIN CITIES':        '/tmp/logos/amrize.img',
    'AMRIZE: GVA (BC)':               '/tmp/logos/amrize.img',
    'AMRIZE: GTA':                    '/tmp/logos/amrize.img',
    'DUFFERIN AGGREGATES (CRH)':      '/tmp/logos/dufferin_aggregates.img',
    'WESTERN STATES CONTRACTING':     '/tmp/logos/western_states_contracting.img',
    'TWIN CITY HAULING':              '/tmp/logos/twin_city_hauling.img',
})
_logo_cache = {}
def _logo_b64(name):
    path = LOGO_PATHS.get(name)
    if not path: return None
    if path not in _logo_cache:
        try:
            with open(path, 'rb') as f:
                data = f.read()
            if data[:4] == b'\x89PNG':                                  mime = 'image/png'
            elif data[:2] == b'\xff\xd8':                               mime = 'image/jpeg'
            elif data[:4] == b'RIFF' and data[8:12] == b'WEBP':        mime = 'image/webp'
            elif b'<svg' in data[:200] or data[:5] == b'<?xml':        mime = 'image/svg+xml'
            else:                                                        mime = 'image/png'
            _logo_cache[path] = f'data:{mime};base64,' + _b64.b64encode(data).decode()
        except OSError:
            _logo_cache[path] = None
    return _logo_cache[path]

TENURE_GROUP = {
    'Pilot':    'New',
    '< 1 mo':  'New',
    '< 3 mo':  'New',
    '~1 yr':   '~1 Year',
    '~2 yrs':  '~2 Years',
    '3+ yrs':  '3+ Years',
}

def to_json(co, seg):
    n = co['name']
    raw_tenure = TENURE.get(n, co.get('tenure',''))
    return {
        'name': n, 'segment': seg,
        'health': co.get('health','gray'),
        'type': co.get('customer_type',''),
        'arr': co.get('arr','—'), 'trucks': co.get('trucks','—'),
        'location': co.get('location','—'), 'csm': co.get('csm',''),
        'owner': co.get('owner',''), 'hubspot': co.get('hubspot',''),
        'what': co.get('what',''), 'connects_with': co.get('connects_with',''),
        'main_contacts': ser(co.get('main_contacts',[])),
        'tread_features': ser(FEATURES.get(n, co.get('tread_features',[]))),
        'personality': co.get('personality',''),
        'activity': ser(co.get('activity',[])),
        'tickets': ser(co.get('tickets',[])),
        'risks': ser(co.get('risks',[])),
        'systems': ser(co.get('systems',[])),
        'usage_status': USAGE.get(n,''),
        'tenure': raw_tenure,
        'tenure_group': TENURE_GROUP.get(raw_tenure, ''),
        'renewal_date': RENEWAL.get(n, ''),
        'intercom': INTERCOM.get(n,[]),
        'intercom_90d': (lambda ids, ic90: ic90 + [
            {'id': str(i), 'url': 'https://app.intercom.com/a/apps/m48souwv/conversations/' + str(i),
             'subject': INTERCOM_STUB_CATS.get(str(i), ('Open conversation', 'Other'))[0],
             'category': INTERCOM_STUB_CATS.get(str(i), ('Open conversation', 'Other'))[1],
             'state': 'open', 'date': '', '_stub': True}
            for i in ids if str(i) not in {str(x.get('id','')) for x in ic90}
        ])(INTERCOM.get(n,[]), INTERCOM_90D.get(n, [])),
        'linear_90d':   LINEAR_90D.get(n, []),
        'linear_project':      LINEAR_PROJECT_ISSUES.get(n, []),
        'linear_project_meta': LINEAR_PROJECT_META.get(n, {}),
        'heysam': HEYSAM_CALLS.get(n, None),
        'dispatch_loads': DISPATCH_LOADS.get(n, 0),
        'dispatch_detail': DISPATCH_DETAIL.get(n, None),
        'logo': _logo_b64(n),
    }

companies = [to_json(c,'Enterprise') for c in ent] + [to_json(c,'Mid-Market') for c in mm]
map_pts   = [{'name':n,'lat':lat,'lon':lon,'health':h,'seg':s} for n,lat,lon,h,s in MAP_COS]

# ── HTML ─────────────────────────────────────────────────────────────────────
HTML = """\
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Customer Pulse</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<style>
:root {
  --bg:      #0A1820;
  --surface: #132732;
  --surf2:   #1A3041;
  --surf3:   #223344;
  --border:  #253B4C;
  --border2: #2E4A5C;
  --yellow:  #FFE500;
  --yellow2: #FFF176;
  --green:   #22C55E;
  --amber:   #F59E0B;
  --red:     #EF4444;
  --gray:    #64748B;
  --text:    #F0F6FA;
  --text2:   #B8D0DF;
  --muted:   #6D93A8;
  --link:    #60ABDE;
  --hauler:       #3B82F6;
  --producer:     #22C55E;
  --construction: #F97316;
  --agriculture:  #A3E635;
  --mixed:        #A78BFA;
  --cat-bug:     #EF4444;
  --cat-question:#3B82F6;
  --cat-billing: #A78BFA;
  --cat-feature: #22C55E;
  --cat-onboard: #F97316;
  --cat-perf:    #F59E0B;
  --cat-integ:   #06B6D4;
  --cat-other:   #64748B;
}
body.light {
  --bg:      #F1F5F8;
  --surface: #FFFFFF;
  --surf2:   #F1F5F8;
  --surf3:   #E4ECF1;
  --border:  #D1DCE5;
  --border2: #B8CCd8;
  --yellow:  #D4A800;
  --yellow2: #B8900A;
  --text:    #0D1F2D;
  --text2:   #334E63;
  --muted:   #6B8A9E;
  --link:    #1A6FA8;
}
/* ── Light mode element overrides ── */
body.light .tab-btn:hover { background:rgba(0,0,0,.05); }
/* Remove dark overlays on card sections */
body.light .card-meta   { background:transparent; }
body.light .card-footer { background:rgba(0,0,0,.04); }
/* Tread logo: invert white SVG to dark */
body.light .header-logo img { filter:brightness(0) opacity(.75); }
/* Logo badge: add border so it's visible on white card */
body.light .badge.has-logo { background:#fff; border:1px solid var(--border2); }
/* Health chips */
body.light .chip.green  { color:#166534; border-color:rgba(34,197,94,.5); }
body.light .chip.yellow { color:#7C4A00; border-color:rgba(245,158,11,.5); }
body.light .chip.red    { color:#991B1B; border-color:rgba(239,68,68,.5); }
body.light .chip.gray   { color:#374151; border-color:rgba(100,116,139,.5); }
/* Tags */
body.light .tag.type       { background:rgba(0,0,0,.06); }
body.light .tag.tenure     { background:rgba(0,0,0,.05); }
body.light .tag.tenure.New       { color:#9A3412; }
body.light .tag.tenure.\\~1Year  { color:#78350F; }
body.light .tag.tenure.\\~2Years { color:#14532D; }
body.light .tag.tenure.3pYears   { color:#1E3A8A; }
body.light .tag.mid-market  { color:#1A6FA8; border-color:rgba(96,171,222,.45); }
body.light .tag.enterprise  { color:#92400E; border-color:rgba(217,119,6,.35); }
body.light .tag.status      { background:rgba(0,0,0,.06); }
body.light .tag.status.Primary-system { color:#166534; }
body.light .tag.status.Onboarding     { color:#1A6FA8; }
body.light .tag.status.Sporadic       { color:#7C4A00; }
body.light .tag.status.Disengaged     { color:#991B1B; }
/* Counters */
body.light .counter.tickets { color:#991B1B; border-color:rgba(239,68,68,.3); }
body.light .counter.support  { color:#1A6FA8; border-color:rgba(96,171,222,.4); }
body.light .counter.risks    { color:#7C4A00; border-color:rgba(245,158,11,.4); }
body.light .counter.none     { background:rgba(0,0,0,.06); }
/* Risk banner: stronger bg + dark readable text */
body.light .m-foot         { background:rgba(245,158,11,.13); border-color:rgba(245,158,11,.35); }
body.light .le-foot.m-foot { background:rgba(245,158,11,.13); border-color:rgba(245,158,11,.35); }
body.light .m-risk         { color:#92400E; }
/* Modal */
body.light .m-badge,
body.light .m-badge-pill { background:rgba(0,0,0,.07); border-color:rgba(0,0,0,.13); color:var(--text2); }
body.light #modal-close  { background:rgba(0,0,0,.12); border-color:rgba(0,0,0,.2); color:var(--text2); }
body.light #modal-close:hover { background:rgba(0,0,0,.2); }
* { box-sizing:border-box; margin:0; padding:0; }
body { font-family:'Inter',system-ui,sans-serif; background:var(--bg); color:var(--text); min-height:100vh; }

/* ── Header ── */
/* ── Tab bar (header) ── */
#header {
  background:var(--surface); border-bottom:1px solid var(--border);
  padding:0 28px;
  display:grid; grid-template-columns:1fr auto 1fr; align-items:stretch;
  position:sticky; top:0; z-index:200;
  min-height:50px;
}
.header-logo { display:flex; align-items:center; gap:10px; flex-shrink:0; }
.header-logo img { height:26px; }
.header-title { font-size:17px; font-weight:800; color:var(--yellow); letter-spacing:-.2px; margin-left:10px; }
#tab-nav { display:flex; align-items:stretch; }
.tab-btn {
  padding:0 20px; background:transparent; color:var(--muted);
  border:none; border-bottom:3px solid transparent;
  font-size:13px; font-weight:600; font-family:inherit;
  cursor:pointer; white-space:nowrap;
  transition:color .15s, border-color .15s, background .15s;
  margin-bottom:-1px;
}
.tab-btn:hover { color:var(--text); background:rgba(255,255,255,.04); }
.tab-btn.active { color:var(--yellow); border-bottom-color:var(--yellow); }

/* ── Controls (chips + search + filters) ── */
#controls { background:var(--surface); border-bottom:1px solid var(--border); position:sticky; top:50px; z-index:199; }
#ctrl-meta {
  padding:10px 24px; display:flex; gap:10px; align-items:center; flex-wrap:wrap;
  border-bottom:1px solid var(--border);
}
.chips { display:flex; flex-wrap:wrap; gap:8px; }
.chip { display:inline-flex; align-items:center; gap:6px; padding:5px 12px; border-radius:20px; font-size:12px; font-weight:600; border:1px solid transparent; }
.chip-count { font-size:14px; font-weight:800; }
.chip.green  { background:rgba(34,197,94,.15);  border-color:rgba(34,197,94,.3);  color:#4ADE80; }
.chip.yellow { background:rgba(245,158,11,.15); border-color:rgba(245,158,11,.3); color:#FCD34D; }
.chip.red    { background:rgba(239,68,68,.15);  border-color:rgba(239,68,68,.3);  color:#F87171; }
.chip.gray   { background:rgba(100,116,139,.15);border-color:rgba(100,116,139,.3);color:#94A3B8; }
.sub { font-size:12px; color:var(--muted); margin-left:auto; }
#ctrl-filters {
  padding:8px 24px; display:flex; gap:8px; align-items:center; flex-wrap:wrap;
}
#search {
  flex:1; min-width:180px; padding:7px 13px;
  background:var(--surf2); color:var(--text);
  border:1px solid var(--border2); border-radius:6px;
  font-size:14px; font-family:inherit; outline:none; transition:border-color .15s;
}
#search::placeholder { color:var(--muted); }
#search:focus { border-color:var(--yellow); }
select {
  padding:6px 9px; background:var(--surf2); color:var(--text2);
  border:1px solid var(--border2); border-radius:6px;
  font-size:12px; font-family:inherit; cursor:pointer; outline:none; transition:border-color .15s;
}
select:focus { border-color:var(--yellow); }
select option { background:var(--surf2); }
#result-count { font-size:12px; color:var(--muted); white-space:nowrap; }

/* ── Map ── */
#map-wrap { border-bottom:1px solid var(--border); position:relative; }
#map { height:580px; }
#map-title-overlay {
  position:absolute; top:14px; left:50%; transform:translateX(-50%);
  z-index:1000; text-align:center; pointer-events:none; white-space:nowrap;
}
#map-title-h { font-size:17px; font-weight:800; color:#fff; text-shadow:0 1px 5px rgba(0,0,0,.9); letter-spacing:.3px; }
#map-title-sub { font-size:11px; color:rgba(255,255,255,.65); text-shadow:0 1px 3px rgba(0,0,0,.8); margin-top:3px; }
.map-legend {
  background:rgba(10,24,32,.88); border:1px solid rgba(255,255,255,.12); border-radius:8px;
  padding:10px 12px; font-size:11px; color:#B8D0DF; line-height:1; backdrop-filter:blur(4px);
}
.leg-section { display:flex; flex-direction:column; gap:6px; }
.leg-row { display:flex; align-items:center; gap:6px; }
.leg-dot  { display:inline-block; width:10px; height:10px; border-radius:50%; flex-shrink:0; border:1.5px solid rgba(0,0,0,.4); }
.leg-star { font-size:13px; line-height:1; flex-shrink:0; color:#FFE500; }
.leg-box  { display:inline-block; width:12px; height:10px; border-radius:2px; flex-shrink:0; }
.leg-divider { border-top:1px solid rgba(255,255,255,.1); margin:7px 0; }
.map-ent-label {
  background:rgba(10,24,32,.82); border:1px solid rgba(255,255,255,.2); border-radius:4px;
  padding:2px 6px; font-size:11px; font-weight:700; color:#fff; white-space:nowrap;
  pointer-events:none; transform:translate(-50%,-150%); display:block;
  text-shadow:none; box-shadow:0 1px 4px rgba(0,0,0,.6);
}
.leaflet-popup-content-wrapper { background:#132732; color:#F0F6FA; border:1px solid #253B4C; border-radius:8px; }
.leaflet-popup-tip { background:#132732; }
.leaflet-popup-content { font-family:'Inter',sans-serif; font-size:12px; margin:8px 12px; }
.map-tooltip.leaflet-tooltip { background:rgba(10,24,32,.9); border:1px solid rgba(255,255,255,.15); color:#B8D0DF; font-family:'Inter',sans-serif; font-size:11px; border-radius:5px; padding:4px 8px; white-space:nowrap; }
.map-tooltip.leaflet-tooltip::before { display:none; }
.map-tip.leaflet-tooltip { background:#0F2435; border:1px solid #2E4A5C; border-radius:7px; padding:7px 11px; box-shadow:0 4px 16px rgba(0,0,0,.5); pointer-events:none; }
.map-tip.leaflet-tooltip::before { border-top-color:#2E4A5C; }
.map-tip-name { font-family:'Inter',sans-serif; font-size:12px; font-weight:600; color:#F0F6FA; white-space:nowrap; margin-bottom:2px; }
.map-tip-loc  { font-family:'Inter',sans-serif; font-size:11px; color:#B8D0DF; white-space:nowrap; margin-bottom:2px; }
.map-tip-sub  { font-family:'Inter',sans-serif; font-size:11px; color:#6D93A8; white-space:nowrap; }

/* ── Grid ── */
#grid {
  display:grid; grid-template-columns:repeat(auto-fill,minmax(300px,1fr));
  gap:12px; padding:20px 24px;
}
#grid.list-view {
  display:block; padding:8px 24px;
}
/* ── Group-by ── */
.grp-section { grid-column:1/-1; }
.grp-hdr {
  display:flex; align-items:center; gap:10px;
  padding:10px 6px 8px; margin-bottom:4px;
  border-bottom:2px solid var(--border);
  cursor:pointer; user-select:none;
}
.grp-hdr:hover { border-bottom-color:var(--border2); }
.grp-dot { width:10px; height:10px; border-radius:50%; flex-shrink:0; }
.grp-dot.green  { background:var(--green); box-shadow:0 0 5px var(--green); }
.grp-dot.yellow { background:var(--amber); box-shadow:0 0 5px var(--amber); }
.grp-dot.red    { background:var(--red);   box-shadow:0 0 5px var(--red); }
.grp-dot.gray   { background:var(--gray); }
.grp-label { font-size:12px; font-weight:700; color:var(--text); text-transform:uppercase; letter-spacing:.07em; }
.grp-count { font-size:11px; color:var(--muted); background:var(--surf2); border:1px solid var(--border); border-radius:10px; padding:1px 9px; }
.grp-chevron { font-size:11px; color:var(--muted); margin-left:auto; transition:transform .18s; display:inline-block; }
.grp-hdr.collapsed .grp-chevron { transform:rotate(-90deg); }
.grp-body { }
.grp-body.grp-hidden { display:none; }
/* card view: nested grid inside group */
.grp-cards {
  display:grid; grid-template-columns:repeat(auto-fill,minmax(300px,1fr));
  gap:12px; padding-bottom:8px;
}
/* list view: no extra wrapper needed */
.grp-lines { }
.view-toggle {
  display:flex; border:1px solid var(--border2); border-radius:6px; overflow:hidden;
}
.view-btn {
  background:none; border:none; color:var(--muted); padding:5px 9px;
  cursor:pointer; font-size:14px; line-height:1; transition:background .15s, color .15s;
}
.view-btn.active { background:var(--surf2); color:var(--yellow); }
.line-row {
  display:flex; flex-direction:column;
  padding:8px 10px 6px; border-radius:6px; cursor:pointer;
  font-size:13px; transition:background .1s;
}
.line-row:hover { background:var(--surf2); }
.line-row + .line-row { border-top:1px solid var(--border); }
.line-main {
  display:grid;
  grid-template-columns: 6px 1fr 90px 75px 110px 130px 70px;
  gap:0 12px; align-items:center;
}
.line-sub {
  display:flex; align-items:center; gap:6px; flex-wrap:wrap;
  padding:4px 0 0 18px; min-height:0;
}
.line-health-bar {
  width:4px; height:32px; border-radius:2px; flex-shrink:0; align-self:center;
}
.line-health-bar.green  { background:var(--green); }
.line-health-bar.yellow { background:var(--amber); }
.line-health-bar.red    { background:var(--red); }
.line-health-bar.gray   { background:var(--gray); }
.line-name { font-weight:600; color:var(--text); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; display:flex; align-items:center; }
.line-seg  { font-size:11px; }
.line-muted { color:var(--muted); font-size:12px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.line-call { font-size:11px; color:var(--muted); display:flex; align-items:center; gap:4px; }
.line-call .call-dot { width:6px; height:6px; border-radius:50%; flex-shrink:0; }
.line-expand {
  display:none; margin:8px 0 4px 18px;
  border:1px solid var(--border2); border-radius:8px;
  background:var(--surf2); overflow:hidden;
}
.line-expand.open { display:block; }
.le-strip {
  display:flex; flex-wrap:wrap; gap:8px 18px; padding:10px 16px;
  border-bottom:1px solid var(--border); font-size:12px; color:var(--text2);
}
.le-body {
  display:grid; grid-template-columns:1fr 1.2fr; gap:0;
}
.le-col { padding:10px 16px; }
.le-col + .le-col { border-left:1px solid var(--border); }
.le-foot { padding:8px 16px; border-top:1px solid var(--border); font-size:12px; }
.line-hdr {
  display:grid;
  grid-template-columns: 6px 1fr 90px 75px 110px 130px 70px;
  gap:0 12px;
  padding:5px 10px 6px; font-size:11px; color:var(--muted);
  text-transform:uppercase; letter-spacing:.5px; font-weight:600;
  border-bottom:1px solid var(--border);
}

/* ── Card ── */
.card {
  background:var(--surface); border:1px solid var(--border); border-radius:10px;
  cursor:pointer; transition:border-color .15s, box-shadow .15s, transform .15s;
  overflow:hidden; border-left:3px solid var(--border);
}
.card:hover {
  border-color:var(--border2); border-left-color:var(--yellow);
  box-shadow:0 4px 20px rgba(0,0,0,.4); transform:translateY(-2px);
}
.card.health-green  { border-left-color:var(--green); }
.card.health-yellow { border-left-color:var(--amber); }
.card.health-red    { border-left-color:var(--red); }
.card.health-gray   { border-left-color:var(--gray); }
.card-header { padding:13px 14px 10px; }
.card-top { display:flex; align-items:center; gap:10px; }
.badge {
  width:36px; height:36px; border-radius:8px;
  display:flex; align-items:center; justify-content:center;
  font-size:15px; font-weight:800; color:#fff; flex-shrink:0; opacity:.9;
}
.badge.Hauler       { background:var(--hauler); }
.badge.Producer     { background:var(--producer); color:#0A1820; }
.badge.Construction { background:var(--construction); }
.badge.Agriculture  { background:var(--agriculture); color:#0A1820; }
.badge.Mixed        { background:var(--mixed); }
.badge.default      { background:var(--gray); }
.badge.has-logo { background:#fff; }
.badge img { width:100%; height:100%; object-fit:contain; border-radius:6px; padding:4px; }
.card-name { font-size:13px; font-weight:700; line-height:1.3; flex:1; color:var(--text); }
.health-dot { width:8px; height:8px; border-radius:50%; flex-shrink:0; }
.health-dot.green  { background:var(--green);  box-shadow:0 0 6px var(--green); }
.health-dot.yellow { background:var(--amber);  box-shadow:0 0 6px var(--amber); }
.health-dot.red    { background:var(--red);    box-shadow:0 0 6px var(--red); }
.health-dot.gray   { background:var(--gray); }
.card-tags { display:flex; flex-wrap:wrap; gap:5px; margin-top:8px; }
.tag { font-size:10px; font-weight:600; padding:2px 8px; border-radius:10px; }
.tag.enterprise { background:rgba(255,229,0,.15); color:var(--yellow2); border:1px solid rgba(255,229,0,.2); }
.tag.mid-market { background:rgba(96,171,222,.15); color:#93C5FD; border:1px solid rgba(96,171,222,.2); }
.tag.type       { background:rgba(255,255,255,.08); color:var(--text2); }
.tag.tenure     { background:rgba(255,255,255,.06); color:var(--muted); border:1px solid var(--border); font-size:9px; }
.tag.tenure.New      { background:rgba(249,115,22,.15); color:#FED7AA; border-color:rgba(249,115,22,.25); }
.tag.tenure.\\~1Year { background:rgba(250,204,21,.12); color:#FDE68A; border-color:rgba(250,204,21,.2); }
.tag.tenure.\\~2Years{ background:rgba(34,197,94,.12); color:#A7F3D0; border-color:rgba(34,197,94,.2); }
.tag.tenure.\\3pYears{ background:rgba(59,130,246,.15); color:#BFDBFE; border-color:rgba(59,130,246,.25); }
.tag.status.Primary-system { background:rgba(34,197,94,.15); color:#4ADE80; }
.tag.status.Onboarding { background:rgba(96,171,222,.15); color:#93C5FD; }
.tag.status.Sporadic   { background:rgba(245,158,11,.15); color:#FCD34D; }
.tag.status.Disengaged { background:rgba(239,68,68,.15); color:#FCA5A5; }
.tag.status { background:rgba(255,255,255,.08); color:var(--text2); }
.card-meta { padding:8px 14px 10px; border-top:1px solid var(--border); }
.meta-row { font-size:12px; color:var(--muted); display:flex; gap:12px; flex-wrap:wrap; margin-bottom:3px; }
.meta-row span { display:flex; align-items:center; gap:3px; color:var(--text2); }
.card-footer {
  padding:8px 14px; background:rgba(0,0,0,.2); border-top:1px solid var(--border);
  display:flex; gap:8px; flex-wrap:wrap;
}
.counter { font-size:11px; font-weight:600; padding:2px 8px; border-radius:10px; }
.counter.tickets { background:rgba(239,68,68,.15); color:#FCA5A5; border:1px solid rgba(239,68,68,.2); }
.counter.support  { background:rgba(96,171,222,.15); color:#93C5FD; border:1px solid rgba(96,171,222,.2); }
.counter.risks    { background:rgba(245,158,11,.15); color:#FCD34D; border:1px solid rgba(245,158,11,.2); }
.counter.activity { background:rgba(255,229,0,.12); color:var(--yellow2); border:1px solid rgba(255,229,0,.2); }
.counter.none     { background:rgba(255,255,255,.05); color:var(--muted); }

/* ── Modal ── */
#modal-bg {
  display:none; position:fixed; inset:0; background:rgba(0,0,0,.75);
  z-index:500; overflow-y:auto; padding:24px;
}
#modal-bg.open { display:flex; align-items:flex-start; justify-content:center; }
#modal {
  background:var(--surface); border:1px solid var(--border2); border-radius:12px;
  width:100%; max-width:980px; overflow:hidden;
  box-shadow:0 24px 80px rgba(0,0,0,.6); margin:auto;
}
#modal-close {
  position:absolute; top:14px; right:14px;
  background:rgba(0,0,0,.35); color:#fff;
  border:1px solid rgba(255,255,255,.2); border-radius:50%;
  width:28px; height:28px; font-size:14px; cursor:pointer;
  display:flex; align-items:center; justify-content:center; transition:background .15s;
}
#modal-close:hover { background:rgba(0,0,0,.6); }
.m-head  { padding:18px 20px 14px; position:relative; border-bottom:1px solid var(--border); }
.m-head.green  { background:linear-gradient(135deg,rgba(34,197,94,.25),rgba(34,197,94,.05)); }
.m-head.yellow { background:linear-gradient(135deg,rgba(245,158,11,.25),rgba(245,158,11,.05)); }
.m-head.red    { background:linear-gradient(135deg,rgba(239,68,68,.25),rgba(239,68,68,.05)); }
.m-head.gray   { background:linear-gradient(135deg,rgba(100,116,139,.2),rgba(100,116,139,.03)); }
.m-head-top { display:flex; align-items:center; gap:14px; }
.m-badge {
  width:50px; height:50px; border-radius:10px;
  background:rgba(255,255,255,.1); border:1px solid rgba(255,255,255,.15);
  display:flex; align-items:center; justify-content:center;
  font-size:22px; font-weight:900; color:#fff; flex-shrink:0;
}
.m-name  { font-size:20px; font-weight:800; color:var(--text); flex:1; }
.m-badges { display:flex; gap:6px; flex-direction:column; align-items:flex-end; }
.m-badge-pill {
  font-size:10px; font-weight:700; padding:3px 9px; border-radius:10px;
  background:rgba(255,255,255,.1); color:var(--text2); border:1px solid rgba(255,255,255,.1);
}
.m-strip {
  background:var(--surf2); padding:9px 20px;
  display:flex; flex-wrap:wrap; gap:16px;
  border-bottom:1px solid var(--border); font-size:12px;
}
.m-strip span { display:flex; align-items:center; gap:4px; color:var(--muted); }
.m-strip strong { color:var(--text); font-weight:600; }
.m-strip .status-pill { font-weight:700; padding:2px 9px; border-radius:10px; font-size:11px; }
.m-body { display:grid; grid-template-columns:1fr 1.4fr; grid-template-rows:auto auto; }
.m-col  { padding:16px 20px; }
#m-left  { grid-column:1; grid-row:1; }
#m-right { grid-column:2; grid-row:1 / 3; border-left:1px solid var(--border); }
.m-section { margin-bottom:14px; }
.m-label {
  font-size:10px; font-weight:700; color:var(--muted);
  letter-spacing:1px; text-transform:uppercase; margin-bottom:6px;
}
.m-text  { font-size:13px; color:var(--text2); line-height:1.55; }
.m-text.muted  { color:var(--muted); font-style:italic; }
.m-bullet { font-size:13px; color:var(--text2); line-height:1.6; padding-left:13px; text-indent:-13px; }
.m-bullet::before { content:"• "; color:var(--yellow); }
.m-act {
  font-size:12px; color:var(--text2); line-height:1.5; margin-bottom:5px;
  padding:6px 10px; background:var(--surf2); border-radius:6px; border-left:2px solid var(--border2);
}
.m-act .src { font-size:10px; font-weight:700; color:var(--muted); text-transform:uppercase; margin-right:4px; }
.m-ticket { display:flex; gap:6px; align-items:baseline; margin-bottom:5px; font-size:12px; }
.m-ticket a { color:var(--link); text-decoration:none; font-weight:600; white-space:nowrap; }
.m-ticket a:hover { text-decoration:underline; }
.m-ticket .desc { color:var(--text2); }
.m-intercom a {
  font-size:12px; color:var(--link); text-decoration:none; margin-right:6px;
  background:rgba(96,171,222,.1); padding:2px 7px; border-radius:4px;
}
.m-intercom a:hover { text-decoration:underline; background:rgba(96,171,222,.2); }
/* 90-day interaction links */
.iact {
  display:flex; align-items:flex-start; gap:8px; margin-bottom:6px;
  padding:6px 10px; background:var(--surf2); border-radius:6px;
  border-left:2px solid var(--border2); font-size:12px;
}
.iact:hover { border-left-color:var(--yellow); }
.iact-src {
  font-size:9px; font-weight:700; padding:2px 6px; border-radius:4px;
  white-space:nowrap; flex-shrink:0; margin-top:1px;
}
.iact-src.Intercom { background:rgba(96,171,222,.2); color:#93C5FD; }
.iact-src.Linear   { background:rgba(167,139,250,.2); color:#C4B5FD; }
.iact-body { flex:1; min-width:0; }
.iact-title { color:var(--text); font-weight:500; line-height:1.4; }
.iact-title a { color:var(--text); text-decoration:none; }
.iact-title a:hover { color:var(--yellow); }
.iact-meta { color:var(--muted); font-size:11px; margin-top:2px; }
.cat-pill {
  display:inline-block; font-size:9px; font-weight:700;
  padding:1px 6px; border-radius:8px; margin-left:4px;
}
.cat-Login-Account-Access       { background:rgba(251,191,36,.2);  color:#FDE68A; }
.cat-App-Mobile-Issues          { background:rgba(167,139,250,.2); color:#DDD6FE; }
.cat-Reporting                  { background:rgba(99,102,241,.2);  color:#C7D2FE; }
.cat-Vendor-Management          { background:rgba(6,182,212,.2);   color:#A5F3FC; }
.cat-Billing-Invoicing          { background:rgba(192,132,252,.2); color:#E9D5FF; }
.cat-Rates-Pricing-Issues       { background:rgba(249,115,22,.2);  color:#FED7AA; }
.cat-Feature-Requests           { background:rgba(34,197,94,.2);   color:#6EE7B7; }
.cat-Add-Onboard-Driver         { background:rgba(52,211,153,.2);  color:#6EE7B7; }
.cat-Driver-Type-Role-Correction{ background:rgba(14,165,233,.2);  color:#7DD3FC; }
.cat-Ticket-Management          { background:rgba(45,212,191,.2);  color:#99F6E4; }
.cat-Other                      { background:rgba(100,116,139,.2); color:#94A3B8; }
/* state spans */
.state-open    { color:#4ADE80; }
.state-done    { color:#6EE7B7; }
.state-backlog { color:#94A3B8; }
.state-other   { color:#B0BEC5; }
.state-unknown { color:#64748B; font-style:italic; }
body.light .state-open    { color:#166534; }
body.light .state-done    { color:#14532D; }
body.light .state-backlog { color:#475569; }
body.light .state-other   { color:#475569; }
body.light .state-unknown { color:#6B8A9E; }
/* light mode cat pills */
body.light .cat-Login-Account-Access       { color:#92400E; }
body.light .cat-App-Mobile-Issues          { color:#4C1D95; }
body.light .cat-Reporting                  { color:#312E81; }
body.light .cat-Vendor-Management          { color:#164E63; }
body.light .cat-Billing-Invoicing          { color:#581C87; }
body.light .cat-Rates-Pricing-Issues       { color:#7C2D12; }
body.light .cat-Feature-Requests           { color:#14532D; }
body.light .cat-Add-Onboard-Driver         { color:#064E3B; }
body.light .cat-Driver-Type-Role-Correction{ color:#0C4A6E; }
body.light .cat-Ticket-Management          { color:#134E4A; }
body.light .cat-Other                      { color:#374151; }
body.light .iact-src.Intercom { color:#1A6FA8; background:rgba(96,171,222,.18); }
body.light .iact-src.Linear   { color:#4C1D95; background:rgba(167,139,250,.18); }
body.light .iact-src.Project  { color:#4C1D95; background:rgba(167,139,250,.18); }
body.light .hcall-title  { color:var(--yellow2); }
body.light .hcall-risk   { color:#991B1B; }
/* line expand notes */
.le-notes { padding:12px 16px 14px; border-top:1px solid var(--border); background:var(--surf2); }
.le-notes-title { font-size:11px; font-weight:700; color:var(--muted); text-transform:uppercase; letter-spacing:.06em; margin-bottom:8px; }
.le-notes-list  { display:flex; flex-direction:column; gap:5px; margin-bottom:8px; }
.le-notes-row   { display:flex; gap:8px; align-items:flex-end; }
.le-notes-input { flex:1; background:var(--surface); border:1px solid var(--border); border-radius:6px; color:var(--text); font-size:12px; font-family:inherit; padding:6px 9px; resize:none; min-height:44px; max-height:100px; }
.le-notes-input:focus { outline:none; border-color:var(--yellow); }
.le-notes-save  { background:var(--yellow); color:#000; border:none; border-radius:6px; font-size:12px; font-weight:700; font-family:inherit; padding:6px 12px; cursor:pointer; white-space:nowrap; flex-shrink:0; }
.le-notes-save:hover { opacity:.85; }
/* Stacked category bars */
.stacked-bar-inner { display:flex; height:8px; border-radius:4px; overflow:hidden; width:100%; }
.stacked-seg { height:100%; min-width:2px; transition:opacity .15s; }
.stacked-seg:hover { opacity:.75; }
/* Category hover tooltip */
#cat-tip {
  display:none; position:fixed; z-index:9000;
  background:#0B1E2D; border:1px solid rgba(255,255,255,.15);
  border-radius:9px; padding:12px 14px;
  min-width:220px; max-width:300px;
  box-shadow:0 8px 28px rgba(0,0,0,.55);
  pointer-events:none;
}
.cat-tip-head { font-size:12px; font-weight:600; color:var(--text); margin-bottom:8px; display:flex; align-items:center; gap:8px; }
.cat-tip-list { display:flex; flex-direction:column; gap:1px; }
.cat-tip-item { display:grid; grid-template-columns:9px 1fr 56px 24px; align-items:center; gap:6px; padding:3px 0; }
.cat-tip-dot { width:8px; height:8px; border-radius:50%; flex-shrink:0; }
.cat-tip-name { font-size:11px; color:var(--muted); overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.cat-tip-bar-track { height:4px; background:rgba(255,255,255,.1); border-radius:2px; overflow:hidden; }
.cat-tip-bar-fill { height:100%; border-radius:2px; }
.cat-tip-count { font-size:11px; font-weight:700; color:var(--text); text-align:right; }
.dash-cat-row.cat-row-hoverable:hover .stacked-bar-inner { opacity:.85; }
/* Tenure chart */
.tenure-chart { padding:14px 16px; }
.tenure-group { display:flex; align-items:center; gap:12px; margin-bottom:10px; }
.tenure-label { font-size:12px; font-weight:600; width:80px; flex-shrink:0; color:var(--text); }
.tenure-bars { flex:1; display:flex; flex-direction:column; gap:4px; }
.tenure-bar-row { display:flex; align-items:center; gap:6px; }
.tenure-bar-src { font-size:9px; font-weight:700; width:18px; flex-shrink:0; }
.tenure-bar-track { flex:1; background:var(--border); border-radius:3px; height:7px; overflow:hidden; }
.tenure-bar-fill { height:7px; border-radius:3px; transition:width .4s; }
.tenure-bar-count { font-size:11px; color:var(--muted); width:24px; text-align:right; flex-shrink:0; }
/* Sentiment view */
#senti-wrap { padding:24px; display:none; }
.senti-title { font-size:18px; font-weight:700; color:var(--text); margin-bottom:6px; }
.senti-sub { font-size:13px; color:var(--muted); margin-bottom:20px; }
.senti-layout { display:grid; grid-template-columns:1fr 1fr; gap:16px; }
@media(max-width:900px) { .senti-layout { grid-template-columns:1fr; } }
.senti-summary { display:flex; gap:10px; padding:12px 16px; flex-wrap:wrap; border-bottom:1px solid var(--border); }
.senti-band { font-size:12px; font-weight:600; padding:4px 10px; border-radius:6px; }
.senti-band.neg  { background:rgba(239,68,68,.15);  color:#FCA5A5; }
.senti-band.neut { background:rgba(245,158,11,.15); color:#FCD34D; }
.senti-band.pos  { background:rgba(34,197,94,.15);  color:#86EFAC; }
.senti-list { overflow-y:auto; max-height:460px; padding:8px 16px 16px; }
.senti-row { display:flex; align-items:center; gap:10px; margin-bottom:9px; cursor:pointer; padding:4px 6px; border-radius:6px; }
.senti-row:hover { background:var(--surf2); }
.senti-row.selected { background:rgba(255,229,0,.06); border-left:3px solid var(--yellow); padding-left:3px; }
.senti-name { font-size:12px; font-weight:600; width:160px; flex-shrink:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; color:var(--text); }
.senti-bar-wrap { flex:1; }
.senti-bar-track { background:var(--border); border-radius:3px; height:7px; overflow:hidden; }
.senti-bar-fill { height:7px; border-radius:3px; transition:width .4s; }
.senti-score { font-size:11px; font-weight:700; width:30px; text-align:right; }
.senti-detail { padding:16px; }
.senti-score-big { font-size:52px; font-weight:800; line-height:1; margin-bottom:2px; }
.senti-score-label { font-size:14px; font-weight:600; display:block; margin-bottom:14px; }
.senti-meter-track { background:var(--border); border-radius:6px; height:10px; overflow:hidden; margin-bottom:4px; }
.senti-meter-fill { height:10px; border-radius:6px; transition:width .5s; }
.senti-meter-labels { display:flex; justify-content:space-between; font-size:10px; color:var(--muted); margin-bottom:14px; }
.senti-drivers { display:flex; gap:8px; flex-wrap:wrap; margin-bottom:14px; }
.senti-driver { font-size:12px; padding:5px 10px; border-radius:6px; background:var(--surf2); color:var(--muted); }
.senti-driver.neg  { background:rgba(239,68,68,.15);  color:#FCA5A5; }
.senti-driver.warn { background:rgba(245,158,11,.15); color:#FCD34D; }
.senti-driver.pos  { background:rgba(34,197,94,.15);  color:#86EFAC; }
.senti-no-data { color:var(--muted); font-size:13px; padding:20px 16px; }
/* Feedback taxonomy */
.tag-cloud { display:flex; flex-wrap:wrap; gap:6px; padding:14px 16px; }
.tag-word { display:inline-block; padding:3px 9px; border-radius:20px; font-weight:600; cursor:default; line-height:1.5; transition:opacity .15s; }
.tag-word:hover { opacity:.75; }
.tag-neg { color:#FCA5A5; background:rgba(239,68,68,.12); }
.tag-pos { color:#86EFAC; background:rgba(34,197,94,.12); }
.tag-neu { color:#94A3B8; background:rgba(100,116,139,.08); }
.tag-word sup { font-size:9px; opacity:.7; margin-left:2px; }
.phrase-list { padding:6px 16px 14px; }
.phrase-row { display:flex; align-items:center; gap:10px; margin-bottom:7px; }
.phrase-text { font-size:12px; font-weight:600; width:160px; flex-shrink:0; }
.phrase-bar-track { flex:1; background:var(--border); border-radius:3px; height:6px; overflow:hidden; }
.phrase-bar { height:6px; border-radius:3px; }
.phrase-count { font-size:11px; color:var(--muted); width:24px; text-align:right; flex-shrink:0; }
.senti-section-label { font-size:10px; font-weight:700; color:var(--muted); letter-spacing:.8px; text-transform:uppercase; padding:10px 16px 4px; }
.sbs-wrap { padding:10px 16px 4px; }
.sbs-bar { display:flex; height:8px; border-radius:4px; overflow:hidden; background:var(--border); }
.sbs-labels { display:flex; justify-content:space-between; font-size:11px; margin-top:5px; }
.senti-co-count { font-size:11px; color:var(--muted); }
.m-foot  { padding:12px 20px; background:rgba(245,158,11,.08); border-bottom:1px solid rgba(245,158,11,.2); }
.m-foot.no-risk { display:none; }

.m-risk { font-size:12px; color:#FCD34D; font-weight:600; margin-bottom:6px; }
.m-engage { font-size:12px; color:var(--muted); font-style:italic; }
.no-results { grid-column:1/-1; text-align:center; padding:80px 20px; color:var(--muted); font-size:15px; }

/* ── Dashboard ── */
#dash-wrap { padding:24px; display:none; }
.dash-title { font-size:18px; font-weight:700; color:var(--text); margin-bottom:6px; }
.dash-sub { font-size:13px; color:var(--muted); margin-bottom:20px; }
.dash-layout { display:grid; grid-template-columns:1fr 1fr; gap:16px; }
@media(max-width:900px) { .dash-layout { grid-template-columns:1fr; } }
.dash-panel { background:var(--surface); border:1px solid var(--border); border-radius:10px; overflow:hidden; }
.dash-panel-head {
  padding:12px 16px; background:var(--surf2); border-bottom:1px solid var(--border);
  display:flex; align-items:center; justify-content:space-between;
}
.dash-panel-title { font-size:13px; font-weight:700; color:var(--text); }
.dash-panel-sub { font-size:11px; color:var(--muted); }
.dash-table { overflow-y:auto; max-height:500px; }
.dash-row {
  display:flex; align-items:center; gap:10px;
  padding:10px 16px; border-bottom:1px solid var(--border);
  cursor:pointer; transition:background .12s;
}
.dash-row:last-child { border-bottom:none; }
.dash-row:hover { background:var(--surf2); }
.dash-row.selected { background:rgba(255,229,0,.06); border-left:3px solid var(--yellow); padding-left:13px; }
.dash-row-name { font-size:13px; font-weight:600; color:var(--text); flex:1; min-width:0; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.dash-row-seg { font-size:10px; font-weight:600; padding:1px 6px; border-radius:8px; flex-shrink:0; }
.dash-row-seg.Enterprise { background:rgba(255,229,0,.15); color:var(--yellow2); }
.dash-row-seg.Mid-Market { background:rgba(96,171,222,.15); color:#93C5FD; }
.dash-row-counts { display:flex; gap:6px; align-items:center; flex-shrink:0; }
.dash-count { font-size:11px; font-weight:700; padding:2px 7px; border-radius:8px; }
.dash-count.intercom { background:rgba(96,171,222,.2); color:#93C5FD; }
.dash-count.linear   { background:rgba(167,139,250,.2); color:#C4B5FD; }
.dash-count.total    { background:rgba(255,229,0,.15); color:var(--yellow2); min-width:28px; text-align:center; }
.dash-src-bar { display:flex; align-items:center; gap:10px; margin-bottom:12px; padding:0 4px; }
.dash-src-select { background:var(--surf2); border:1px solid var(--border); border-radius:6px; color:var(--text); font-size:12px; font-family:inherit; padding:5px 10px; cursor:pointer; }
.dash-src-select:focus { outline:none; border-color:var(--yellow); }
.dash-bar-wrap { width:60px; flex-shrink:0; }
.dash-bar-track { background:var(--border); border-radius:3px; height:6px; overflow:hidden; }
.dash-bar-fill  { background:var(--yellow); height:6px; border-radius:3px; transition:width .3s; }
/* Right panel — category breakdown */
.dash-cats-wrap { padding:16px; }
.dash-cat-row { display:flex; align-items:center; gap:10px; margin-bottom:10px; }
.dash-cat-label { font-size:12px; font-weight:600; width:120px; flex-shrink:0; }
.dash-cat-track { flex:1; background:var(--border); border-radius:4px; height:8px; overflow:hidden; }
.dash-cat-fill  { height:8px; border-radius:4px; transition:width .4s; }
.dash-cat-count { font-size:11px; color:var(--muted); width:28px; text-align:right; flex-shrink:0; }
.dash-links { padding:0 16px 16px; }
.dash-link-group { margin-bottom:12px; }
.dash-link-group-label { font-size:10px; font-weight:700; color:var(--muted); letter-spacing:.8px; text-transform:uppercase; margin-bottom:6px; }
/* ── Registration panel ── */
.reg-panel { background:var(--surface); border:1px solid var(--border); border-radius:10px; overflow:hidden; margin-top:20px; }
.reg-panel-head { padding:12px 16px; background:var(--surf2); border-bottom:1px solid var(--border); display:flex; align-items:center; justify-content:space-between; }
.reg-panel-title { font-size:13px; font-weight:700; color:var(--text); }
.reg-panel-sub { font-size:11px; color:var(--muted); }
.reg-chart-wrap { padding:16px 24px 4px; }
.reg-chart-area { display:flex; align-items:flex-end; gap:10px; height:148px; }
.reg-week-col { flex:1; display:flex; flex-direction:column; align-items:center; height:100%; }
.reg-week-bar-wrap { flex:1; width:100%; display:flex; align-items:flex-end; }
.reg-week-bar { width:100%; display:flex; flex-direction:column-reverse; border-radius:4px 4px 0 0; overflow:hidden; transition:filter .12s; cursor:default; }
.reg-week-bar:hover { filter:brightness(1.12); }
.reg-week-count { font-size:11px; font-weight:700; color:var(--text2); margin-bottom:3px; text-align:center; line-height:1; }
.reg-week-label { font-size:10px; color:var(--muted); margin-top:6px; text-align:center; white-space:nowrap; }
.reg-week-seg { width:100%; }
.reg-legend { display:flex; flex-wrap:wrap; gap:6px 14px; padding:10px 20px 14px; border-top:1px solid var(--border); margin-top:10px; }
.reg-leg-item { display:flex; align-items:center; gap:5px; font-size:10px; color:var(--muted); white-space:nowrap; }
.reg-leg-dot { width:8px; height:8px; border-radius:2px; flex-shrink:0; }
#reg-tip { display:none; position:fixed; z-index:9000; background:#0B1E2D; border:1px solid rgba(255,255,255,.15); border-radius:9px; padding:12px 14px; min-width:200px; max-width:280px; box-shadow:0 8px 28px rgba(0,0,0,.55); pointer-events:none; }
.reg-tip-head { font-size:13px; font-weight:700; color:var(--text); margin-bottom:6px; }
.reg-tip-row { display:flex; justify-content:space-between; gap:12px; font-size:11px; color:var(--text2); padding:2px 0; }
.reg-tip-co { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.reg-tip-n { font-weight:700; color:var(--text); flex-shrink:0; }

/* ── Open conversations stacked bar panel ── */
.bug-panel { background:var(--surface); border:1px solid var(--border); border-radius:10px; overflow:hidden; margin-top:20px; }
.bug-panel-head { padding:12px 16px; background:var(--surf2); border-bottom:1px solid var(--border); display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:8px; }
.bug-panel-title { font-size:13px; font-weight:700; color:var(--text); }
.bug-legend { display:flex; flex-wrap:wrap; gap:10px; }
.bug-leg-item { display:flex; align-items:center; gap:4px; font-size:11px; color:var(--muted); }
.bug-leg-dot { width:8px; height:8px; border-radius:2px; flex-shrink:0; }
.bug-rows { padding:10px 16px; display:flex; flex-direction:column; gap:5px; }
.bug-row { display:flex; align-items:center; gap:10px; }
.bug-row-name { width:155px; font-size:12px; font-weight:600; color:var(--text); text-align:right; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; flex-shrink:0; }
.bug-bar-outer { flex:1; }
.bug-bar-inner { display:flex; height:18px; border-radius:4px; overflow:hidden; }
.bug-seg { height:100%; cursor:default; transition:opacity .12s; }
.bug-seg:hover { opacity:.72; }
.bug-row-count { width:26px; font-size:12px; font-weight:700; color:var(--text); text-align:right; flex-shrink:0; }
#bug-tip { display:none; position:fixed; z-index:9000; background:#0B1E2D; border:1px solid rgba(255,255,255,.15); border-radius:9px; padding:12px 14px; min-width:220px; max-width:310px; box-shadow:0 8px 28px rgba(0,0,0,.55); pointer-events:none; }
.bug-tip-head { font-size:13px; font-weight:700; margin-bottom:4px; }
.bug-tip-sub { font-size:11px; color:var(--muted); margin-bottom:8px; }
.bug-tip-list { display:flex; flex-direction:column; gap:4px; }
.bug-tip-item { font-size:11px; color:var(--text2); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; display:block; }

/* ── Notes section in modal ── */
#m-notes { grid-column:1; grid-row:2; border-top:1px solid var(--border); border-right:1px solid var(--border); padding:14px 20px 18px; }

/* ── Insights filter bar ── */
.dash-filter-bar { display:flex; align-items:center; gap:14px; padding:10px 0 14px; flex-wrap:wrap; }
.dash-filter-label { font-size:11px; font-weight:700; color:var(--muted); text-transform:uppercase; letter-spacing:.06em; white-space:nowrap; }
.dash-toggle-group { display:flex; border:1px solid var(--border); border-radius:8px; overflow:hidden; }
.dash-toggle { background:none; border:none; color:var(--muted); font-size:12px; font-weight:600; font-family:inherit; padding:5px 13px; cursor:pointer; border-right:1px solid var(--border); transition:background .15s,color .15s; white-space:nowrap; }
.dash-toggle:last-child { border-right:none; }
.dash-toggle.active { background:var(--yellow); color:#000; }
.dash-toggle:hover:not(.active) { background:var(--surf2); color:var(--text); }
.notes-head { display:flex; align-items:center; gap:8px; margin-bottom:10px; }
.notes-title { font-size:11px; font-weight:700; color:var(--muted); text-transform:uppercase; letter-spacing:.06em; }
#notes-list { display:flex; flex-direction:column; gap:6px; margin-bottom:10px; max-height:220px; overflow-y:auto; }
.notes-empty { font-size:12px; color:var(--muted); font-style:italic; padding:2px 0; }
.note-item { background:var(--surf2); border:1px solid var(--border); border-radius:7px; padding:8px 10px; }
.note-meta { display:flex; align-items:center; justify-content:space-between; margin-bottom:3px; }
.note-ts { font-size:10px; color:var(--muted); }
.note-del { background:none; border:none; color:var(--muted); font-size:15px; line-height:1; cursor:pointer; padding:0 3px; border-radius:3px; }
.note-del:hover { color:#EF4444; background:rgba(239,68,68,.12); }
.note-text { font-size:12px; color:var(--text); white-space:pre-wrap; word-break:break-word; }
.notes-input-row { display:flex; gap:8px; align-items:flex-end; }
#notes-input { flex:1; background:var(--surf2); border:1px solid var(--border); border-radius:7px; color:var(--text); font-size:12px; font-family:inherit; padding:8px 10px; resize:none; min-height:52px; max-height:120px; }
#notes-input:focus { outline:none; border-color:var(--yellow); }
#notes-save-btn { background:var(--yellow); color:#000; border:none; border-radius:7px; font-size:12px; font-weight:700; font-family:inherit; padding:8px 14px; cursor:pointer; white-space:nowrap; flex-shrink:0; }
#notes-save-btn:hover { opacity:.85; }
#notes-save-btn:disabled { opacity:.4; cursor:default; }

/* ── Topic summary on cards ── */
.card-topics { padding:5px 14px 7px; border-top:1px solid var(--border); display:flex; gap:5px; flex-wrap:wrap; align-items:center; }
.topic-label { font-size:10px; color:var(--muted); margin-right:2px; }
.topic-chip  { font-size:10px; font-weight:600; padding:2px 7px; border-radius:8px; background:rgba(255,255,255,.06); color:var(--text2); border:1px solid var(--border); }

/* ── Project history in modal ── */
.proj-header { display:flex; align-items:center; justify-content:space-between; margin-bottom:8px; padding:6px 10px; background:rgba(167,139,250,.08); border-radius:6px; border:1px solid rgba(167,139,250,.15); }
.proj-name   { font-size:12px; font-weight:700; color:#C4B5FD; }
.proj-link   { color:#C4B5FD; text-decoration:none; }
.proj-link:hover { text-decoration:underline; }
.proj-toggle { font-size:11px; color:var(--link); cursor:pointer; border:none; background:none; font-family:inherit; padding:0; }
.proj-toggle:hover { text-decoration:underline; }
.iact-src.Project { background:rgba(167,139,250,.2); color:#C4B5FD; }

/* ── Dispatch sparkline on cards ── */
.card-vol { display:flex; align-items:center; gap:10px; padding:6px 14px 7px; border-top:1px solid var(--border); }
.vol-label { font-size:10px; color:var(--muted); white-space:nowrap; flex-shrink:0; }
.vol-sparkline { flex-shrink:0; display:flex; align-items:flex-end; }
.vol-mats { display:flex; gap:4px; flex:1; align-items:center; overflow:hidden; }
.vol-mat-chip { display:inline-flex; align-items:center; gap:3px; font-size:9px; color:var(--text2); white-space:nowrap; flex-shrink:0; }
.vol-mat-swatch { width:6px; height:6px; border-radius:1px; flex-shrink:0; }

/* ── Dispatch activity panel in modal ── */
.dispatch-panel { border:1px solid rgba(96,171,222,.2); border-radius:10px; padding:14px 16px; background:rgba(96,171,222,.04); }
.dp-header { display:flex; align-items:baseline; gap:10px; margin-bottom:12px; }
.dp-big-num { font-size:28px; font-weight:800; color:var(--link); line-height:1; }
.dp-sub { font-size:11px; color:var(--muted); }
.dp-chart-label { font-size:10px; font-weight:600; color:var(--muted); text-transform:uppercase; letter-spacing:.6px; margin-bottom:6px; }
.dp-bar-chart { display:flex; align-items:flex-end; gap:3px; height:52px; }
.dp-bar-col { display:flex; flex-direction:column; align-items:center; flex:1; gap:2px; }
.dp-bar { width:100%; border-radius:2px 2px 0 0; min-height:2px; transition:opacity .15s; }
.dp-bar.partial { opacity:.5; }
.dp-bar-mon { font-size:9px; color:var(--muted); }
.dp-mats-section { margin-top:12px; }
.dp-mat-row { display:flex; align-items:center; gap:8px; margin-bottom:5px; }
.dp-mat-name { font-size:10px; color:var(--text2); width:90px; flex-shrink:0; }
.dp-mat-track { flex:1; height:6px; background:var(--border); border-radius:3px; overflow:hidden; }
.dp-mat-fill { height:6px; border-radius:3px; }
.dp-mat-pct { font-size:10px; color:var(--muted); width:28px; text-align:right; flex-shrink:0; }

/* ── Account snapshot summary in modal ── */
.acct-snap  { background:rgba(255,255,255,.03); border:1px solid var(--border2); border-radius:8px; padding:10px 13px; }
.snap-para  { font-size:12.5px; color:var(--text2); line-height:1.65; }

/* ── HeySam call strip on cards ── */
.call-strip { display:flex; align-items:center; gap:6px; padding:5px 14px 6px; border-top:1px solid var(--border); font-size:11px; overflow:hidden; }
.call-dot   { width:7px; height:7px; border-radius:50%; flex-shrink:0; }
.call-date  { color:var(--muted); white-space:nowrap; flex-shrink:0; }
.call-title { color:var(--text2); overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }

/* ── HeySam call section in modal ── */
.heysam-call     { border:1px solid rgba(255,229,0,.15); border-radius:8px; padding:10px 12px; background:rgba(255,229,0,.04); }
.hcall-title-row { display:flex; align-items:center; gap:8px; margin-bottom:6px; }
.hcall-dot       { width:8px; height:8px; border-radius:50%; flex-shrink:0; }
.hcall-title     { font-size:12px; font-weight:600; color:var(--yellow); text-decoration:none; flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.hcall-title:hover { text-decoration:underline; }
.hcall-date      { font-size:11px; color:var(--muted); white-space:nowrap; }
.hcall-topics    { font-size:12px; color:var(--text2); line-height:1.5; margin-bottom:4px; }
.hcall-risk      { font-size:11px; color:#FCA5A5; margin-top:4px; }

/* ── Volume view ── */
</style>
</head>
<body>

<div id="header">
  <div class="header-logo">
    <img src="https://tread.ai/wp-content/uploads/2025/03/Tread-Logo.svg" alt="Tread" onerror="this.style.display='none'">
    <span class="header-title">Customer Pulse</span>
  </div>
  <nav id="tab-nav">
    <button class="tab-btn active" id="toggle-grid">Customers</button>
    <button class="tab-btn" id="toggle-map">Map</button>
    <button class="tab-btn" id="toggle-dash">Insights</button>
    <button class="tab-btn" id="toggle-senti">Sentiment</button>
  </nav>
  <div style="display:flex;align-items:center;justify-content:flex-end;gap:12px;">
    <span style="font-size:11px;color:var(--muted);">Updated BUILD_DATE_PLACEHOLDER · refreshes hourly</span>
    <button id="theme-btn" title="Toggle light/dark mode" style="background:none;border:1px solid var(--border2);border-radius:6px;color:var(--muted);padding:5px 9px;cursor:pointer;font-size:14px;line-height:1;transition:color .15s,border-color .15s;flex-shrink:0;">☀︎</button>
  </div>
</div>

<div id="controls">
  <div id="ctrl-meta">
    <div class="chips" id="hdr-chips"></div>
    <div class="sub" id="hdr-sub"></div>
  </div>
  <div id="ctrl-filters">
    <input id="search" type="search" placeholder="Search companies, location, CSM…">
    <span id="result-count"></span>
    <select id="f-health">
      <option value="">All Health</option>
      <option value="green">Healthy</option>
      <option value="yellow">Needs Attention</option>
      <option value="red">At Risk</option>
      <option value="gray">Inactive</option>
    </select>
    <select id="f-seg">
      <option value="">All Segments</option>
      <option value="Enterprise">Enterprise</option>
      <option value="Mid-Market">Mid-Market</option>
    </select>
    <select id="f-type">
      <option value="">All Types</option>
      <option value="Hauler">Hauler</option>
      <option value="Producer">Producer</option>
      <option value="Construction">Construction</option>
      <option value="Agriculture">Agriculture</option>
      <option value="Mixed">Mixed</option>
    </select>
    <select id="f-usage">
      <option value="">All Statuses</option>
      <option value="Onboarding">Onboarding</option>
      <option value="Primary system">Primary System</option>
      <option value="Sporadic">Sporadic</option>
      <option value="Disengaged">Disengaged</option>
    </select>
    <select id="f-csm">
      <option value="">All CSMs</option>
      <option value="Latefa Redjouh">Latefa Redjouh</option>
      <option value="unassigned">Unassigned</option>
    </select>
    <select id="f-tenure">
      <option value="">All Tenure</option>
      <option value="New">New (&lt;3mo)</option>
      <option value="~1 Year">~1 Year</option>
      <option value="~2 Years">~2 Years</option>
      <option value="3+ Years">3+ Years</option>
    </select>
    <select id="f-sort">
      <option value="az">Sort: A → Z</option>
      <option value="za">Sort: Z → A</option>
      <option value="vol-hl">Volume: High → Low</option>
      <option value="vol-lh">Volume: Low → High</option>
      <option value="arr-hl" selected>ARR: High → Low</option>
      <option value="arr-lh">ARR: Low → High</option>
    </select>
    <select id="f-group">
      <option value="">No Grouping</option>
      <option value="health">Group by Health</option>
      <option value="seg">Group by Segment</option>
      <option value="type">Group by Type</option>
      <option value="usage">Group by Status</option>
      <option value="csm">Group by CSM</option>
      <option value="tenure">Group by Tenure</option>
    </select>
    <div class="view-toggle">
      <button class="view-btn active" id="btn-cards" title="Card view">⊞</button>
      <button class="view-btn" id="btn-lines" title="List view">☰</button>
    </div>
  </div>
</div>

<div id="map-wrap" style="display:none">
  <div id="map-title-overlay">
    <div id="map-title-h">Customer Map — North America</div>
    <div id="map-title-sub"></div>
  </div>
  <div id="map"></div>
</div>
<div id="grid"></div>
<div id="dash-wrap"></div>
<div id="senti-wrap"></div>

<div id="modal-bg">
  <div id="modal">
    <div class="m-head" id="m-head">
      <button id="modal-close" onclick="closeModal()">✕</button>
      <div class="m-head-top">
        <div class="m-badge" id="m-badge"></div>
        <div class="m-name"  id="m-name"></div>
        <div class="m-badges" id="m-badges"></div>
      </div>
    </div>
    <div class="m-strip" id="m-strip"></div>
    <div class="m-foot" id="m-foot"></div>
    <div class="m-body">
      <div class="m-col" id="m-left"></div>
      <div class="m-col" id="m-right"></div>
      <div id="m-notes">
        <div class="notes-head">
          <span class="notes-title">Notes</span>
        </div>
        <div id="notes-list"></div>
        <div class="notes-input-row">
          <textarea id="notes-input" placeholder="Add a note… (Cmd+Enter to save)" rows="2"></textarea>
          <button id="notes-save-btn">Save</button>
        </div>
      </div>
    </div>
  </div>
</div>

<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script>
// ── Data ─────────────────────────────────────────────────────────────────────
const CUSTOMERS     = CUSTOMERS_PLACEHOLDER;
const MAP_PTS       = MAP_PTS_PLACEHOLDER;
const ALL_INTERCOM  = ALL_INTERCOM_PLACEHOLDER;
const ALL_LINEAR    = ALL_LINEAR_PLACEHOLDER;
const REGISTRATIONS = REGISTRATIONS_PLACEHOLDER;

// ── Helpers ──────────────────────────────────────────────────────────────────
const HC = {green:'#22C55E',yellow:'#F59E0B',red:'#EF4444',gray:'#64748B'};
const TC = {Hauler:'#3B82F6',Producer:'#22C55E',Construction:'#F97316',Agriculture:'#A3E635',Mixed:'#A78BFA'};
const CAT_COLORS = {
  'Login & Account Access':        '#FBBF24',
  'App / Mobile Issues':           '#A78BFA',
  'Reporting':                     '#818CF8',
  'Vendor Management':             '#06B6D4',
  'Billing & Invoicing':           '#C084FC',
  'Rates & Pricing Issues':        '#F97316',
  'Feature Requests':              '#22C55E',
  'Add / Onboard Driver':          '#34D399',
  'Driver Type / Role Correction': '#0EA5E9',
  'Ticket Management':             '#2DD4BF',
  'Other':                         '#64748B',
  'No Action Needed':              '#334155',
};
const esc = s => (s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const catClass = c => 'cat-' + (c||'Other').replace(/[^a-zA-Z0-9]+/g,'-').replace(/-+/g,'-').replace(/-$/,'');
const MAX_LOADS = Math.max(...CUSTOMERS.map(c => c.dispatch_loads || 0), 1);
const DISPATCH_MONTHS = ['Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar','Apr','May'];
const MAT_COLORS = {
  'Base/Aggregate': '#3B82F6',
  'Fill/Dirt':      '#92400E',
  'Stone/Rock':     '#6B7280',
  'Sand':           '#F59E0B',
  'Asphalt':        '#374151',
  'Other':          '#7C3AED',
};
function fmtLoads(n) {
  if (!n) return '';
  if (n >= 1000) return Math.round(n/1000) + 'K';
  return n.toString();
}
function matColor(name) { return MAT_COLORS[name] || '#64748B'; }

function cardSparkline(co) {
  const d = co.dispatch_detail;
  if (!d || !d.months) return '';
  const vals = d.months;
  const mx = Math.max(...vals, 1);
  const n = vals.length;
  const barW = 3, gap = 1, h = 22;
  const totalW = n * (barW + gap) - gap;
  const last = n - 1;
  const bars = vals.map((v, i) => {
    const bh = Math.max(Math.round(v / mx * (h - 2)), v > 0 ? 2 : 0);
    const x = i * (barW + gap);
    const y = h - bh;
    const op = i === last ? '0.45' : '1';
    return `<rect x="${x}" y="${y}" width="${barW}" height="${bh}" fill="#60ABDE" opacity="${op}" rx="1"/>`;
  }).join('');
  const matChips = (d.materials || []).slice(0, 3).map(([name]) =>
    `<span class="vol-mat-chip"><span class="vol-mat-swatch" style="background:${matColor(name)}"></span>${name}</span>`
  ).join('');
  return `<div class="card-vol">
    <span class="vol-label">${fmtLoads(co.dispatch_loads)}/90d</span>
    <div class="vol-sparkline"><svg width="${totalW}" height="${h}" viewBox="0 0 ${totalW} ${h}">${bars}</svg></div>
    <div class="vol-mats">${matChips}</div>
  </div>`;
}

function dispatchPanel(co) {
  const d = co.dispatch_detail;
  if (!d) {
    if (!co.dispatch_loads) return '';
    return `<div class="dispatch-panel"><div class="dp-header"><span class="dp-big-num">${fmtLoads(co.dispatch_loads)}</span><span class="dp-sub">loads dispatched · last 90 days</span></div></div>`;
  }
  const vals = d.months;
  const n = vals.length;
  const last = n - 1;
  const mx = Math.max(...vals, 1);
  const barH = 52;
  // Show every other month label when 12 bars to avoid crowding
  const skipLabel = n >= 10;
  const bars = vals.map((v, i) => {
    const bh = Math.max(Math.round(v / mx * barH), v > 0 ? 2 : 0);
    const partial = i === last ? ' partial' : '';
    const lbl = (!skipLabel || i % 2 === 0) ? `<div class="dp-bar-mon">${DISPATCH_MONTHS[i]}</div>` : `<div class="dp-bar-mon"> </div>`;
    return `<div class="dp-bar-col"><div class="dp-bar${partial}" style="height:${bh}px;background:linear-gradient(0deg,#2563EB,#60ABDE)"></div>${lbl}</div>`;
  }).join('');
  const matRows = (d.materials || []).map(([name, pct]) =>
    `<div class="dp-mat-row"><div class="dp-mat-name">${name}</div><div class="dp-mat-track"><div class="dp-mat-fill" style="width:${pct}%;background:${matColor(name)}"></div></div><div class="dp-mat-pct">${pct}%</div></div>`
  ).join('');
  return `<div class="dispatch-panel">
    <div class="dp-header"><span class="dp-big-num">${fmtLoads(co.dispatch_loads)}</span><span class="dp-sub">loads · last 90 days · 12-month trend below</span></div>
    <div class="dp-bar-chart" style="margin-top:8px">${bars}</div>
    ${matRows ? `<div class="dp-mats-section"><div class="dp-chart-label" style="margin-top:10px">Material Mix</div>${matRows}</div>` : ''}
  </div>`;
}

// ── Header ───────────────────────────────────────────────────────────────────
function buildHeader() {
  document.getElementById('hdr-sub').textContent =
    `${CUSTOMERS.length} accounts — ${CUSTOMERS.filter(c=>c.segment==='Enterprise').length} Enterprise · ${CUSTOMERS.filter(c=>c.segment==='Mid-Market').length} Mid-Market`;
  const counts = {green:0,yellow:0,red:0,gray:0};
  CUSTOMERS.forEach(c => counts[c.health]=(counts[c.health]||0)+1);
  const labels = {green:'Healthy',yellow:'Needs Attention',red:'At Risk',gray:'Inactive'};
  const chips = document.getElementById('hdr-chips');
  Object.entries(counts).forEach(([k,v]) => {
    chips.insertAdjacentHTML('beforeend',
      `<span class="chip ${k}"><span class="chip-count">${v}</span> ${labels[k]}</span>`);
  });
}

// ── Customers ────────────────────────────────────────────────────────────────
function badgeClass(type) {
  return ['Hauler','Producer','Construction','Agriculture','Mixed'].includes(type) ? type : 'default';
}
function renewalColor(date) {
  if (!date) return 'var(--muted)';
  const d = new Date(date), now = new Date();
  const days = (d - now) / 86400000;
  if (days < 90)  return '#FCA5A5'; // red — <90 days
  if (days < 180) return '#FCD34D'; // amber — <6mo
  return 'var(--muted)';
}
const TENURE_CSS = {'New':'New','~1 Year':'\\~1Year','~2 Years':'\\~2Years','3+ Years':'\\3pYears'};

function topicSummary(co) {
  const all = [...(co.intercom_90d||[]), ...(co.linear_90d||[]), ...(co.linear_project||[])];
  if (all.length < 3) return [];
  const cats = {};
  all.forEach(i => { const c = i.category||'Other'; cats[c]=(cats[c]||0)+1; });
  return Object.entries(cats).sort((a,b)=>b[1]-a[1]).slice(0,2);
}

function tenureTag(co) {
  const g = co.tenure_group;
  if (!g) return co.tenure ? `<span class="tag tenure">${esc(co.tenure)}</span>` : '';
  const css = TENURE_CSS[g] || '';
  return `<span class="tag tenure ${css}">${esc(g)}</span>`;
}

function renderCard(co, idx) {
  const tix  = co.tickets.length;
  const sup  = co.intercom.length;
  const risk = co.risks.length;
  const act  = (co.intercom_90d||[]).filter(i=>!i._stub).length + (co.linear_90d||[]).length;
  const statusTag = co.usage_status ?
    `<span class="tag status ${esc(co.usage_status.replace(/ /g,'-'))}">${esc(co.usage_status)}</span>` : '';
  return `
  <div class="card health-${esc(co.health)}" onclick="openModal(${idx})">
    <div class="card-header">
      <div class="card-top">
        <div class="badge ${co.logo?'has-logo':esc(badgeClass(co.type))}">${co.logo?`<img src="${co.logo}" alt="">`:esc(co.name[0])}</div>
        <div class="card-name">${esc(co.name)}</div>
        <div class="health-dot ${esc(co.health)}"></div>
      </div>
      <div class="card-tags">
        <span class="tag ${co.segment==='Enterprise'?'enterprise':'mid-market'}">${esc(co.segment)}</span>
        <span class="tag type">${esc(co.type)}</span>
        ${tenureTag(co)}
        ${statusTag}
      </div>
    </div>
    <div class="card-meta">
      <div class="meta-row">
        ${co.location!=='—'?`<span>📍 ${esc(co.location)}</span>`:''}
        ${co.arr!=='—'?`<span>💰 ${esc(co.arr)}</span>`:''}
        ${co.trucks!=='—'?`<span>🚚 ${co.trucks} trucks</span>`:''}
      </div>
      <div class="meta-row">
        ${co.csm?`<span>👤 ${esc(co.csm)}</span>`:''}
        ${co.tenure?`<span>🗓 ${esc(co.tenure)}</span>`:''}
        ${co.renewal_date?`<span style="color:${renewalColor(co.renewal_date)}">🔄 ${esc(co.renewal_date)}</span>`:''}
      </div>
    </div>
    ${(co.dispatch_loads || co.dispatch_detail) ? cardSparkline(co) : ''}
    ${(()=>{ const ts=topicSummary(co); return ts.length ? `<div class="card-topics"><span class="topic-label">Mainly:</span>${ts.map(([c,n])=>`<span class="topic-chip">${esc(c)} <em style="opacity:.65">${n}</em></span>`).join('')}</div>` : ''; })()}
    ${co.heysam ? (()=>{ const c=co.heysam; const oc=c.overall==='positive'?'#4ADE80':c.overall==='neutral'?'#FCD34D':'#FCA5A5'; return `<div class="call-strip"><span class="call-dot" style="background:${oc}"></span><span class="call-date">${esc(c.date)}</span><span class="call-title">${esc(c.title)}</span></div>`; })() : ''}
    <div class="card-footer">
      ${tix>0?`<span class="counter tickets">${tix} ticket${tix>1?'s':''}</span>`:'<span class="counter none">No tickets</span>'}
      ${sup>0?`<span class="counter support">${sup} support</span>`:''}
      ${risk>0?`<span class="counter risks">${risk} risk${risk>1?'s':''}</span>`:''}
      ${act>0?`<span class="counter activity">⚡ ${act} last 90d</span>`:''}
    </div>
  </div>`;
}

function parseArr(s) {
  if (!s || s === '—') return 0;
  const m = s.replace(/[$,\s]/g,'').match(/^([\d.]+)([KkMm]?)$/);
  if (!m) return 0;
  const v = parseFloat(m[1]), u = m[2].toUpperCase();
  return u==='M' ? v*1000000 : u==='K' ? v*1000 : v;
}

function renderLine(co, idx) {
  const tix  = co.tickets.length;
  const sup  = co.intercom.length;
  const risk = co.risks.length;
  const ic   = (co.intercom_90d||[]).filter(i=>!i._stub).length;
  const li   = (co.linear_90d||[]).length;
  const act  = ic + li;
  const thumb = co.logo
    ? `<img src="${co.logo}" style="width:20px;height:20px;object-fit:contain;border-radius:3px;margin-right:7px;flex-shrink:0;" alt="">`
    : `<span style="display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;border-radius:3px;font-size:10px;font-weight:800;color:#fff;background:var(--surf2);flex-shrink:0;margin-right:7px;">${esc(co.name[0])}</span>`;
  // sub-row: tags + renewal + heysam + topics + counters
  const subParts = [];
  subParts.push(`<span class="tag ${co.segment==='Enterprise'?'enterprise':'mid-market'}" style="font-size:10px">${esc(co.segment)}</span>`);
  subParts.push(`<span class="tag type" style="font-size:10px">${esc(co.type)}</span>`);
  const tTag = tenureTag(co); if (tTag) subParts.push(tTag);
  if (co.usage_status) subParts.push(`<span class="tag status ${esc(co.usage_status.replace(/ /g,'-'))}" style="font-size:10px">${esc(co.usage_status)}</span>`);
  if (co.renewal_date) subParts.push(`<span style="font-size:11px;color:${renewalColor(co.renewal_date)}">🔄 ${esc(co.renewal_date)}</span>`);
  const ts = topicSummary(co);
  if (ts.length) subParts.push(`<span style="font-size:11px;color:var(--muted)">Mainly: ${ts.map(([c,n])=>`${esc(c)} <em style="opacity:.6">${n}</em>`).join(', ')}</span>`);
  if (co.heysam) {
    const h=co.heysam, oc=h.overall==='positive'?'#4ADE80':h.overall==='neutral'?'#FCD34D':'#FCA5A5';
    subParts.push(`<span class="line-call"><span class="call-dot" style="background:${oc}"></span><span>${esc(h.date)}</span><span style="color:var(--text2)">${esc(h.title)}</span></span>`);
  }
  // counters
  if (tix>0) subParts.push(`<span class="counter tickets" style="font-size:10px">${tix} ticket${tix>1?'s':''}</span>`);
  if (sup>0) subParts.push(`<span class="counter support" style="font-size:10px">${sup} support</span>`);
  if (risk>0) subParts.push(`<span class="counter risks" style="font-size:10px">${risk} risk${risk>1?'s':''}</span>`);
  if (act>0) subParts.push(`<span class="counter activity" style="font-size:10px">⚡ ${act} last 90d</span>`);
  return `<div class="line-row" onclick="toggleLineExpand(${idx}, this)">
    <div class="line-main">
      <div class="line-health-bar ${esc(co.health)}"></div>
      <div class="line-name">${thumb}${esc(co.name)}</div>
      <div class="line-muted">${esc(co.arr)}</div>
      <div class="line-muted">${co.trucks!=='—'?co.trucks+' trucks':'—'}</div>
      <div class="line-muted">${esc(co.csm||'—')}</div>
      <div class="line-muted">${esc(co.location!=='—'?co.location:'')}</div>
      <div class="line-muted" style="text-align:right">${act?`<span class="counter activity" style="font-size:10px">⚡ ${act}</span>`:'—'}</div>
    </div>
    ${subParts.length ? `<div class="line-sub">${subParts.join('')}</div>` : ''}
    <div class="line-expand"></div>
  </div>`;
}

let _viewMode = 'cards';
let filtered = [...CUSTOMERS.keys()];
function renderGrid() {
  const q   = document.getElementById('search').value.toLowerCase();
  const fh  = document.getElementById('f-health').value;
  const fs  = document.getElementById('f-seg').value;
  const ft  = document.getElementById('f-type').value;
  const fc  = document.getElementById('f-csm').value;
  const ftn = document.getElementById('f-tenure').value;
  const fu  = document.getElementById('f-usage').value;
  const srt = document.getElementById('f-sort').value;
  filtered = [];
  CUSTOMERS.forEach((co,i) => {
    if (fh  && co.health !== fh) return;
    if (fs  && co.segment !== fs) return;
    if (ft  && co.type !== ft) return;
    if (fc  && co.csm !== fc) return;
    if (ftn && co.tenure_group !== ftn) return;
    if (fu  && co.usage_status !== fu) return;
    if (q   && ![co.name,co.location,co.csm,co.owner,co.what,co.connects_with,co.personality,
                 ...(co.risks||[]),...(co.activity||[]).map(a=>Array.isArray(a)?a.join(' '):a)
                ].join(' ').toLowerCase().includes(q)) return;
    filtered.push(i);
  });
  if (srt === 'az')     filtered.sort((a,b) => CUSTOMERS[a].name.localeCompare(CUSTOMERS[b].name));
  else if (srt === 'za')     filtered.sort((a,b) => CUSTOMERS[b].name.localeCompare(CUSTOMERS[a].name));
  else if (srt === 'vol-hl') filtered.sort((a,b) => (CUSTOMERS[b].dispatch_loads||0) - (CUSTOMERS[a].dispatch_loads||0));
  else if (srt === 'vol-lh') filtered.sort((a,b) => (CUSTOMERS[a].dispatch_loads||0) - (CUSTOMERS[b].dispatch_loads||0));
  else if (srt === 'arr-hl') filtered.sort((a,b) => parseArr(CUSTOMERS[b].arr) - parseArr(CUSTOMERS[a].arr));
  else if (srt === 'arr-lh') filtered.sort((a,b) => parseArr(CUSTOMERS[a].arr) - parseArr(CUSTOMERS[b].arr));
  const grid = document.getElementById('grid');
  const grp  = document.getElementById('f-group').value;
  grid.classList.toggle('list-view', _viewMode === 'lines');
  if (filtered.length === 0) {
    grid.innerHTML = '<div class="no-results">No customers match the current filters.</div>';
  } else if (grp) {
    const GRP_KEY = {
      health: co => co.health || 'gray',
      seg:    co => co.segment || '—',
      type:   co => co.type || '—',
      csm:    co => co.csm || 'unassigned',
      tenure: co => co.tenure_group || '—',
      usage:  co => co.usage_status || '—',
    };
    const GRP_ORDER = {
      health: ['red','yellow','gray','green'],
      seg:    ['Enterprise','Mid-Market'],
      type:   ['Hauler','Producer','Construction','Agriculture','Mixed'],
      csm:    ['Latefa Redjouh','unassigned'],
      tenure: ['New','~1 Year','~2 Years','3+ Years'],
      usage:  ['Onboarding','Primary system','Sporadic','Disengaged'],
    };
    const GRP_LABEL = {
      health: {red:'At Risk', yellow:'Needs Attention', gray:'Inactive', green:'Healthy'},
    };
    const getKey = GRP_KEY[grp] || (co => '—');
    const order  = GRP_ORDER[grp] || [];
    const labelMap = GRP_LABEL[grp] || {};
    // bucket
    const buckets = {};
    filtered.forEach(i => { const k=getKey(CUSTOMERS[i]); (buckets[k]=buckets[k]||[]).push(i); });
    const keys = [...order.filter(k=>buckets[k]), ...Object.keys(buckets).filter(k=>!order.includes(k)).sort()];
    const linehdr = `<div class="line-hdr"><div></div><div>Customer</div><div>ARR</div><div>Trucks</div><div>CSM</div><div>Location</div><div style="text-align:right">90d</div></div>`;
    grid.innerHTML = keys.map(k => {
      const items = buckets[k];
      const label = labelMap[k] || k || '—';
      const dotCls = grp==='health' ? `grp-dot ${k}` : 'grp-dot gray';
      const innerCls = _viewMode==='lines' ? 'grp-lines' : 'grp-cards';
      const innerHtml = _viewMode==='lines'
        ? linehdr + items.map(i=>renderLine(CUSTOMERS[i],i)).join('')
        : items.map(i=>renderCard(CUSTOMERS[i],i)).join('');
      return `<div class="grp-section">
        <div class="grp-hdr" onclick="this.classList.toggle('collapsed');this.nextElementSibling.classList.toggle('grp-hidden')">
          <span class="${dotCls}"></span>
          <span class="grp-label">${esc(label)}</span>
          <span class="grp-count">${items.length}</span>
          <span class="grp-chevron">▾</span>
        </div>
        <div class="grp-body ${innerCls}">${innerHtml}</div>
      </div>`;
    }).join('');
  } else if (_viewMode === 'lines') {
    const hdr = `<div class="line-hdr">
      <div></div><div>Customer</div><div>ARR</div><div>Trucks</div>
      <div>CSM</div><div>Location</div>
      <div style="text-align:right">90d</div></div>`;
    grid.innerHTML = hdr + filtered.map(i => renderLine(CUSTOMERS[i], i)).join('');
  } else {
    grid.innerHTML = filtered.map(i => renderCard(CUSTOMERS[i], i)).join('');
  }
  document.getElementById('result-count').textContent =
    `${filtered.length} of ${CUSTOMERS.length} customers`;
}
['search','f-health','f-seg','f-type','f-csm','f-tenure','f-usage','f-sort','f-group'].forEach(id =>
  document.getElementById(id).addEventListener('input', renderGrid));
document.getElementById('btn-cards').addEventListener('click', () => {
  _viewMode = 'cards';
  document.getElementById('btn-cards').classList.add('active');
  document.getElementById('btn-lines').classList.remove('active');
  renderGrid();
});
document.getElementById('btn-lines').addEventListener('click', () => {
  _viewMode = 'lines';
  document.getElementById('btn-lines').classList.add('active');
  document.getElementById('btn-cards').classList.remove('active');
  renderGrid();
});
(function() {
  const btn = document.getElementById('theme-btn');
  const saved = localStorage.getItem('cp-theme');
  if (saved === 'light') { document.body.classList.add('light'); btn.textContent = '☾'; }
  btn.addEventListener('click', () => {
    const isLight = document.body.classList.toggle('light');
    btn.textContent = isLight ? '☾' : '☀︎';
    localStorage.setItem('cp-theme', isLight ? 'light' : 'dark');
  });
})();

// ── Modal ────────────────────────────────────────────────────────────────────
function section(label, html) {
  return `<div class="m-section"><div class="m-label">${esc(label)}</div>${html}</div>`;
}
function iactHtml(item, source) {
  const title = source === 'Intercom' ? item.subject : item.title;
  const url   = item.url;
  const state = item.state || '';
  const cat   = item.category || 'Other';
  const date  = item.date || item.createdAt || '';
  const srcCss = source === 'Project' ? 'Project' : source;
  const stateCls = state === 'open' ? 'state-open' : state === 'Done' ? 'state-done' : state === 'Backlog' ? 'state-backlog' : 'state-other';
  const stubTag = item._stub ? `<span style="font-size:9px;margin-left:4px;font-style:italic" class="state-unknown">tracked</span>` : '';
  return `<div class="iact">
    <span class="iact-src ${srcCss}">${source}</span>
    <div class="iact-body">
      <div class="iact-title"><a href="${esc(url)}" target="_blank">${esc(title||url)}</a>
        <span class="cat-pill ${catClass(cat)}">${esc(cat)}</span>${stubTag}
      </div>
      <div class="iact-meta">${date ? esc(date) : '<span class="state-unknown">date unknown</span>'} · <span class="${stateCls}">${esc(state)}</span></div>
    </div>
  </div>`;
}

function _buildContent(idx) {
  const co = CUSTOMERS[idx];
  const stripParts = [];
  if (co.location!=='—') stripParts.push(`<span>📍 <strong>${esc(co.location)}</strong></span>`);
  if (co.arr!=='—')      stripParts.push(`<span>💰 ARR <strong>${esc(co.arr)}</strong></span>`);
  if (co.trucks!=='—')   stripParts.push(`<span>🚚 <strong>${co.trucks}</strong> trucks</span>`);
  if (co.csm)            stripParts.push(`<span>👤 CSM <strong>${esc(co.csm)}</strong></span>`);
  if (co.tenure)         stripParts.push(`<span>🗓 Tenure <strong>${esc(co.tenure)}</strong></span>`);
  if (co.usage_status) {
    const sc = {
      'Primary system':'background:rgba(34,197,94,.2);color:#4ADE80;border:1px solid rgba(34,197,94,.3)',
      'Onboarding':    'background:rgba(96,171,222,.2);color:#93C5FD;border:1px solid rgba(96,171,222,.3)',
      'Sporadic':      'background:rgba(245,158,11,.2);color:#FCD34D;border:1px solid rgba(245,158,11,.3)',
      'Disengaged':    'background:rgba(239,68,68,.2);color:#FCA5A5;border:1px solid rgba(239,68,68,.3)',
    }[co.usage_status] || 'background:rgba(255,255,255,.1);color:#B8D0DF';
    stripParts.push(`<span><span class="status-pill" style="${sc}">${esc(co.usage_status)}</span></span>`);
  }
  if (co.renewal_date) {
    const rc = renewalColor(co.renewal_date);
    stripParts.push(`<span>🔄 Renewal <strong style="color:${rc}">${esc(co.renewal_date)}</strong></span>`);
  }
  if (co.hubspot) stripParts.push(`<span><a href="${esc(co.hubspot)}" target="_blank" style="color:var(--link)">HubSpot →</a></span>`);
  let left = '';
  left += section('About', `<div class="m-text">${esc(co.what)}</div>`);
  if (co.connects_with) left += section('Connects With in Tread', `<div class="m-text muted">${esc(co.connects_with)}</div>`);
  if (co.main_contacts.length) {
    left += section('Main Contacts',
      co.main_contacts.map(([nm,ti]) =>
        `<div class="m-bullet">${esc(nm)}${ti?' — <em style="color:var(--muted)">'+esc(ti)+'</em>':''}</div>`
      ).join(''));
  }
  if (co.tread_features.length) {
    left += section('Key Tread Features',
      co.tread_features.map(f => `<div class="m-bullet">${esc(f)}</div>`).join(''));
  }
  if (co.systems.length) {
    left += section('Systems', `<div class="m-text">${esc(co.systems.join(', '))}</div>`);
  }
  let right = '';

  // ── Dispatch Activity panel ─────────────────────────────────────────────────
  {
    const dp = dispatchPanel(co);
    if (dp) right += section('Dispatch Activity', dp);
  }

  // ── Account snapshot — synthesize everything we know ──────────────────────
  {
    const ic90s = (co.intercom_90d || []).filter(i=>!i._stub);
    const ic90stubs = (co.intercom_90d || []).filter(i=>i._stub);
    const li90s = co.linear_90d  || [];
    const sentences = [];

    // 1. Dispatch + usage context
    const usageMap = {
      'Primary system': 'their primary dispatch system',
      'Onboarding':     'currently onboarding',
      'Sporadic':       'using Tread sporadically',
      'Disengaged':     'disengaged from the platform',
    };
    const usageDesc = usageMap[co.usage_status] || null;
    if (co.dispatch_loads) {
      let s = `${fmtLoads(co.dispatch_loads)} loads dispatched in the last 90 days`;
      if (usageDesc) s += `, ${usageDesc}`;
      if (co.tenure) s += ` (${co.tenure} customer)`;
      sentences.push(s + '.');
    } else if (usageDesc) {
      let s = usageDesc.charAt(0).toUpperCase() + usageDesc.slice(1);
      if (co.tenure) s += ` — ${co.tenure} customer`;
      sentences.push(s + '.');
    }

    // 2. Most recent call (HeySam)
    if (co.heysam) {
      const h = co.heysam;
      let s = `Most recent call (${h.date}) covered ${h.key_topics}`;
      if (h.challenges) s += `. Flagged concern: ${h.challenges}`;
      sentences.push(s + '.');
    }

    // 3. Support activity — Intercom + Linear combined, with category breakdown
    const all90s = [...ic90s, ...li90s];
    if (all90s.length) {
      const cats = {};
      all90s.forEach(i => { const c = i.category||'Other'; cats[c]=(cats[c]||0)+1; });
      const topCats = Object.entries(cats).sort((a,b)=>b[1]-a[1]).slice(0,2);
      const catStr = topCats.map(([c,n]) => `${c.toLowerCase()} (${n})`).join(', ');
      const icCount = ic90s.length, liCount = li90s.length;
      const parts = [];
      if (icCount) parts.push(`${icCount} Intercom`);
      if (liCount) parts.push(`${liCount} Linear`);
      const stubNote = ic90stubs.length ? ` + ${ic90stubs.length} tracked open` : '';
      sentences.push(`${all90s.length} support item${all90s.length>1?'s':''} last 90 days (${parts.join(', ')}${stubNote}), mainly ${catStr}.`);
    } else {
      sentences.push('No support activity in the last 90 days.');
    }

    right += section('Account Snapshot', `<div class="acct-snap"><div class="snap-para">${sentences.join(' ')}</div></div>`);
  }

  const ic90     = co.intercom_90d || [];
  const ic90real = ic90.filter(i=>!i._stub);
  const li90     = co.linear_90d  || [];

  // ── Recent call (HeySam) ────────────────────────────────────────────────────
  if (co.heysam) {
    const h = co.heysam;
    const oc = h.overall==='positive' ? '#4ADE80' : h.overall==='neutral' ? '#FCD34D' : '#FCA5A5';
    right += section('Recent Call',
      `<div class="heysam-call">
        <div class="hcall-title-row">
          <span class="hcall-dot" style="background:${oc}"></span>
          <a class="hcall-title" href="${esc(h.url)}" target="_blank">${esc(h.title)}</a>
          <span class="hcall-date">${esc(h.date)}</span>
        </div>
        <div class="hcall-topics">${esc(h.key_topics)}</div>
        ${h.challenges ? `<div class="hcall-risk">⚠ ${esc(h.challenges)}</div>` : ''}
      </div>`);
  }

  // ── Intercom — Last 90 Days ─────────────────────────────────────────────────
  if (ic90.length) {
    const realSorted  = [...ic90real].sort((a,b) => (b.date||'').localeCompare(a.date||''));
    const stubs       = ic90.filter(i=>i._stub);
    const stubLabel   = stubs.length ? ` + ${stubs.length} tracked open` : '';
    right += section(`Intercom — Last 90 Days (${ic90real.length}${stubLabel})`,
      [...realSorted, ...stubs].map(i => iactHtml(i, 'Intercom')).join(''));
  }

  // ── Linear: project history OR standalone tickets — never both ─────────────
  {
    const proj     = co.linear_project || [];
    const projMeta = co.linear_project_meta || {};

    if (proj.length && projMeta.project_name) {
      // Has a named project — show project history only
      const projName = projMeta.project_name;
      const projUrl  = projMeta.project_url;
      const projId   = 'proj-body-' + idx;
      const preview  = proj.slice(0, 3);
      const rest     = proj.slice(3);
      let ph = `<div class="proj-header">`;
      ph += projUrl
        ? `<a class="proj-name proj-link" href="${esc(projUrl)}" target="_blank">📋 Linear Project</a>`
        : `<span class="proj-name">📋 Linear Project</span>`;
      if (rest.length) ph += `<button class="proj-toggle" onclick="toggleProj('${projId}')">+ ${rest.length} more</button>`;
      ph += `</div>`;
      ph += preview.map(i => iactHtml(i, 'Project')).join('');
      if (rest.length) ph += `<div id="${projId}" style="display:none">${rest.map(i => iactHtml(i, 'Project')).join('')}</div>`;
      right += section(`Linear (${proj.length})`, ph);
    } else {
      // No project — show standalone tickets + 90d items
      const lsorted = [...(li90 || [])].sort((a,b) => (b.createdAt||b.date||'').localeCompare(a.createdAt||a.date||''));
      let lbody = lsorted.map(i => iactHtml(i, 'Linear')).join('');
      if (co.tickets.length) {
        lbody += co.tickets.map(([tid,desc]) => {
          const url = `https://linear.app/treadapp/issue/${tid}`;
          return `<div class="m-ticket"><a href="${url}" target="_blank">${esc(tid)}</a><span class="desc">${esc(desc)}</span></div>`;
        }).join('');
      }
      const ltotal = (li90||[]).length + co.tickets.length;
      right += section(`Linear${ltotal ? ' ('+ltotal+')' : ''}`, lbody || '<div class="m-text muted">None open</div>');
    }
  }
  let foot = '';
  if (co.risks.length) {
    foot += `<div class="m-risk">⚠ RISK: ${esc(co.risks.join('  |  '))}</div>`;
    if (co.personality) foot += `<div class="m-engage">${esc(co.personality)}</div>`;
  } else if (co.personality) {
    foot += `<div class="m-engage">💬 ${esc(co.personality)}</div>`;
  }

  return { co, stripHtml: stripParts.join(''), left, right, foot,
           hasRisk: co.risks.length > 0 };
}

function openModal(idx) {
  const { co, stripHtml, left, right, foot, hasRisk } = _buildContent(idx);
  document.getElementById('m-head').className = `m-head ${co.health}`;
  document.getElementById('m-badge').textContent = co.name[0];
  document.getElementById('m-name').textContent  = co.name;
  document.getElementById('m-badges').innerHTML  =
    `<span class="m-badge-pill">${esc(co.segment)}</span>
     <span class="m-badge-pill">${esc(co.type)}</span>
     <span class="m-badge-pill">${co.health.toUpperCase()}</span>`;
  document.getElementById('m-strip').innerHTML = stripHtml;
  document.getElementById('m-left').innerHTML  = left;
  document.getElementById('m-right').innerHTML = right;
  const footEl = document.getElementById('m-foot');
  footEl.className = hasRisk ? 'm-foot' : 'm-foot no-risk';
  footEl.innerHTML = foot;
  loadNotes(co.name);
  document.getElementById('modal-bg').classList.add('open');
  document.getElementById('modal-bg').scrollTop = 0;
}

let _expandedRow = null;
function toggleLineExpand(idx, rowEl) {
  const expandEl = rowEl.querySelector('.line-expand');
  if (!expandEl) return;
  const isOpen = expandEl.classList.contains('open');
  if (_expandedRow && _expandedRow !== rowEl) {
    _expandedRow.querySelector('.line-expand')?.classList.remove('open');
  }
  if (isOpen) {
    expandEl.classList.remove('open');
    _expandedRow = null;
    return;
  }
  if (!expandEl.dataset.built) {
    const { co, stripHtml, left, right, foot, hasRisk } = _buildContent(idx);
    expandEl.innerHTML =
      `<div class="le-strip">${stripHtml}</div>` +
      (hasRisk && foot ? `<div class="le-foot m-foot">${foot}</div>` : '') +
      `<div class="le-body"><div class="le-col">${left}</div><div class="le-col">${right}</div></div>` +
      `<div class="le-notes">
        <div class="le-notes-title">Notes</div>
        <div class="le-notes-list" id="le-nl-${idx}"></div>
        <div class="le-notes-row">
          <textarea class="le-notes-input" id="le-ni-${idx}" placeholder="Add a note… (Cmd+Enter to save)" rows="2"></textarea>
          <button class="le-notes-save" onclick="saveLeNote(${idx},'${co.name.replace(/'/g,"\\'")}',document.getElementById('le-ni-${idx}').value);document.getElementById('le-ni-${idx}').value=''">Save</button>
        </div>
      </div>`;
    expandEl.dataset.built = '1';
    _renderLeNotes(idx, co.name);
    document.getElementById('le-ni-'+idx).addEventListener('keydown', e => {
      if ((e.metaKey||e.ctrlKey) && e.key==='Enter') {
        saveLeNote(idx, co.name, e.target.value);
        e.target.value = '';
      }
    });
  }
  expandEl.classList.add('open');
  _expandedRow = rowEl;
}

function closeModal() { document.getElementById('modal-bg').classList.remove('open'); }
function toggleProj(id) {
  const el = document.getElementById(id);
  if (!el) return;
  const open = el.style.display === 'none';
  el.style.display = open ? '' : 'none';
  document.querySelectorAll(`button[onclick="toggleProj('${id}')"]`).forEach(b => {
    b.textContent = open ? 'show less' : '+ ' + (el.children.length) + ' more';
  });
}
document.getElementById('modal-bg').addEventListener('click', e => {
  if (e.target === document.getElementById('modal-bg')) closeModal();
});
document.addEventListener('keydown', e => { if (e.key==='Escape') closeModal(); });

// ── View toggles ──────────────────────────────────────────────────────────────
let mapInited = false, dashInited = false, sentiInited = false;
function setView(view) {
  document.getElementById('map-wrap').style.display    = view==='map'    ? 'block' : 'none';
  document.getElementById('grid').style.display        = view==='grid'   ? ''      : 'none';
  document.getElementById('dash-wrap').style.display   = view==='dash'   ? 'block' : 'none';
  document.getElementById('senti-wrap').style.display  = view==='senti'  ? 'block' : 'none';
  ['grid','map','dash','senti'].forEach(v =>
    document.getElementById('toggle-'+v).classList.toggle('active', v===view)
  );
  const isGrid = view === 'grid';
  document.getElementById('ctrl-filters').style.display = isGrid ? '' : 'none';
  if (view==='map'   && !mapInited)   initMap();
  if (view==='dash'  && !dashInited)  initDash();
  if (view==='senti' && !sentiInited) initSenti();
}
let currentView = 'grid';
['grid','map','dash','senti'].forEach(v => {
  document.getElementById('toggle-'+v).addEventListener('click', () => {
    currentView = v;
    setView(v);
  });
});

// ── Map ───────────────────────────────────────────────────────────────────────
function initMap() {
  mapInited = true;

  const DENSITY = {
    "Texas":9,"Florida":9,"Ontario":6,"Illinois":4,"Minnesota":4,
    "Ohio":3,"New York":2,"California":2,"Arizona":2,"Nevada":2,
    "British Columbia":2,"Colorado":1,"Delaware":1,"Indiana":1,"Oregon":1,
    "Mississippi":1,"Arkansas":1,"Alabama":1,"Connecticut":1,
    "Washington":1,"Massachusetts":1,"Hawaii":1,
    "Saskatchewan":1,"Alberta":1,"Manitoba":1
  };

  function densityColor(n) {
    if (!n || n === 0) return '#0D1B2A';
    if (n === 1)       return '#1A3A55';
    if (n <= 2)        return '#1B5276';
    if (n <= 4)        return '#1A6FA0';
    if (n <= 6)        return '#0EA5E9';
    return '#F59E0B';
  }

  function stateStyle(feature) {
    const name = feature.properties.name || feature.properties.NAME || '';
    const n = DENSITY[name] || 0;
    return {
      fillColor: densityColor(n),
      fillOpacity: 0.78,
      color: '#0A1820',
      weight: 0.8,
    };
  }

  const map = L.map('map', { zoomControl: true }).setView([40, -96], 4);
  L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_nolabels/{z}/{x}/{y}{r}.png',
    { attribution:'© CartoDB © OpenStreetMap contributors', maxZoom:19 }).addTo(map);

  // Ensure customer markers always render above the GeoJSON choropleth layers
  map.createPane('customerMarkers');
  map.getPane('customerMarkers').style.zIndex = 620;

  // Load US states choropleth
  fetch('https://raw.githubusercontent.com/PublicaMundi/MappingAPI/master/data/geojson/us-states.json')
    .then(r => r.json())
    .then(data => {
      L.geoJSON(data, {
        style: stateStyle,
        onEachFeature(feature, layer) {
          const name = feature.properties.name || feature.properties.NAME || '';
          const n = DENSITY[name] || 0;
          if (n) layer.bindTooltip(`${name}: ${n} customer${n>1?'s':''}`, { sticky:true, className:'map-tooltip' });
        }
      }).addTo(map);
    }).catch(() => {});

  // Load Canadian provinces choropleth
  fetch('https://raw.githubusercontent.com/codeforamerica/click_that_hood/master/public/data/canada.geojson')
    .then(r => r.json())
    .then(data => {
      L.geoJSON(data, {
        style(feature) {
          const name = feature.properties.name || feature.properties.NAME || '';
          const n = DENSITY[name] || 0;
          return {
            fillColor: densityColor(n),
            fillOpacity: 0.78,
            color: '#0A1820',
            weight: 0.8,
          };
        },
        onEachFeature(feature, layer) {
          const name = feature.properties.name || feature.properties.NAME || '';
          const n = DENSITY[name] || 0;
          if (n) layer.bindTooltip(`${name}: ${n} customer${n>1?'s':''}`, { sticky:true, className:'map-tooltip' });
        }
      }).addTo(map);
    }).catch(() => {});

  // Customer markers
  MAP_PTS.forEach(pt => {
    if (!pt.lat) return;
    const col = HC[pt.health] || '#64748B';
    const idx = CUSTOMERS.findIndex(c => c.name === pt.name);

    let marker;
    if (pt.seg === 'Enterprise' || pt.seg === 'ent') {
      const starSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 22 22">
        <polygon points="11,2 13.5,8.5 20.5,8.5 15,13 17,20 11,16 5,20 7,13 1.5,8.5 8.5,8.5"
          fill="${col}" stroke="#0A1820" stroke-width="1.2"/>
      </svg>`;
      const icon = L.divIcon({
        html: starSvg,
        className: '',
        iconSize: [22, 22],
        iconAnchor: [11, 11],
        popupAnchor: [0, -12],
      });
      marker = L.marker([pt.lat, pt.lon], { icon, pane: 'customerMarkers' }).addTo(map);
    } else {
      marker = L.circleMarker([pt.lat, pt.lon], {
        radius: 7, fillColor: col, color: '#0A1820', weight: 1.5, fillOpacity: 0.92,
        pane: 'customerMarkers'
      }).addTo(map);
    }

    const seg = (pt.seg==='Enterprise'||pt.seg==='ent') ? 'Enterprise' : 'Mid-Market';
    const co  = CUSTOMERS.find(c => c.name === pt.name);
    const loc = co && co.location && co.location !== '—' ? co.location : null;
    marker.bindTooltip(
      `<div class="map-tip-name">${pt.name}</div>` +
      (loc ? `<div class="map-tip-loc">📍 ${loc}</div>` : '') +
      `<div class="map-tip-sub"><span style="color:${col}">${pt.health.charAt(0).toUpperCase()+pt.health.slice(1)}</span> · ${seg}</div>`,
      { direction: 'top', offset: [0, -10], className: 'map-tip' }
    );
    if (idx >= 0) marker.on('click', () => openModal(idx));
  });

  // Legend control
  const legend = L.control({ position: 'bottomright' });
  legend.onAdd = function() {
    const div = L.DomUtil.create('div', 'map-legend');
    div.innerHTML = `
      <div style="font-size:11px;font-weight:700;color:#E2EBF0;margin-bottom:8px;letter-spacing:.4px">LEGEND</div>
      <div class="leg-section">
        <div style="font-size:10px;color:rgba(255,255,255,.45);letter-spacing:.5px;margin-bottom:4px">CUSTOMER HEALTH</div>
        <div class="leg-row"><span class="leg-dot" style="background:#22C55E"></span>Green</div>
        <div class="leg-row"><span class="leg-dot" style="background:#F59E0B"></span>Yellow</div>
        <div class="leg-row"><span class="leg-dot" style="background:#EF4444"></span>Red</div>
        <div class="leg-row"><span class="leg-dot" style="background:#64748B"></span>Gray</div>
      </div>
      <div class="leg-divider"></div>
      <div class="leg-section">
        <div style="font-size:10px;color:rgba(255,255,255,.45);letter-spacing:.5px;margin-bottom:4px">SEGMENT</div>
        <div class="leg-row"><span class="leg-star">★</span>Enterprise</div>
        <div class="leg-row"><span class="leg-dot" style="background:#64748B"></span>Mid-Market</div>
      </div>
      <div class="leg-divider"></div>
      <div class="leg-section">
        <div style="font-size:10px;color:rgba(255,255,255,.45);letter-spacing:.5px;margin-bottom:4px">STATE DENSITY</div>
        <div class="leg-row"><span class="leg-box" style="background:#F59E0B"></span>7+ customers</div>
        <div class="leg-row"><span class="leg-box" style="background:#0EA5E9"></span>5–6 customers</div>
        <div class="leg-row"><span class="leg-box" style="background:#1A6FA0"></span>3–4 customers</div>
        <div class="leg-row"><span class="leg-box" style="background:#1B5276"></span>2 customers</div>
        <div class="leg-row"><span class="leg-box" style="background:#1A3A55"></span>1 customer</div>
        <div class="leg-row"><span class="leg-box" style="background:#0D1B2A;border:1px solid rgba(255,255,255,.15)"></span>None</div>
      </div>
    `;
    return div;
  };
  legend.addTo(map);

  // Subtitle
  const entCount = MAP_PTS.filter(p => p.seg==='Enterprise' || p.seg==='ent').length;
  const mmCount  = MAP_PTS.filter(p => p.seg!=='Enterprise' && p.seg!=='ent' && p.lat).length;
  document.getElementById('map-title-sub').textContent =
    `${entCount} Enterprise  ·  ${mmCount} Mid-Market`;
}

// ── Dashboard ─────────────────────────────────────────────────────────────────
let selectedDashRow = null;
const CATS = ['Add / Onboard Driver','Login & Account Access','Rates & Pricing Issues','Reporting','Vendor Management','Ticket Management','Billing & Invoicing','App / Mobile Issues','Feature Requests','Driver Type / Role Correction','Other','No Action Needed'];
const CO_PALETTE = ['#3B82F6','#F59E0B','#8B5CF6','#EC4899','#10B981','#F97316','#06B6D4','#EF4444','#84CC16','#A78BFA','#FB7185','#34D399','#FBBF24','#60A5FA','#C084FC','#FCD34D','#6EE7B7','#93C5FD','#F9A8D4','#86EFAC'];

function buildCustomerStats() {
  return CUSTOMERS.map(co => {
    const ic = co.intercom_90d || [];
    const li = co.linear_90d  || [];
    const icItems = ic.map(i=>({...i,_src:'Intercom'}));
    const liItems = li.map(i=>({...i,_src:'Linear'}));
    // total + cats = Intercom only (actual support contacts); stubs excluded from baseline
    // Linear tickets are downstream artifacts, not support volume
    const icReal = icItems.filter(i=>!i._stub);
    const cats = {};
    icReal.forEach(i => { const c=i.category||'Other'; cats[c]=(cats[c]||0)+1; });
    const allItems = [...icItems,...liItems].sort((a,b)=>(b.date||'').localeCompare(a.date||''));
    return { name:co.name, segment:co.segment, health:co.health, icCount:icReal.length, liCount:li.length, total:icReal.length, cats, items:allItems };
  }).sort((a,b) => b.total - a.total);
}

function overallCats(stats) {
  const cats = {};
  stats.forEach(s => Object.entries(s.cats).forEach(([c,n]) => cats[c]=(cats[c]||0)+n));
  return cats;
}

function buildCatByCustomer(stats) {
  const result = {};
  stats.forEach((s, idx) => {
    Object.entries(s.cats).forEach(([cat, n]) => {
      if (n <= 0) return;
      if (!result[cat]) result[cat] = [];
      result[cat].push({ name: s.name, count: n, idx });
    });
  });
  Object.values(result).forEach(arr => arr.sort((a,b) => b.count - a.count));
  return result;
}

// ── Sentiment phrase lists (1–3 words, customer-language signals) ────────────
const _SENT_POS = [
  'love','excited','happy with','appreciate','thank','perfect','smooth',
  'easy to use','works well','going well','went well','great progress',
  'good progress','all set','pleased','glad','impressed',
  'awesome','excellent','fantastic','helpful','no issues','all good','seamless',
  'love it','love the','works great','working great','going live','go live',
  'went live','just went live','shipped','deployed','resolved','fixed',
  'really helpful','great call','good call','positive','we love',
  'they love','customers love','really like','they like','we like',
];
const _SENT_NEG = [
  'frustrated','escalation','angry','unhappy','not happy','upset',
  'disappointed','critical issue','critical bug','blocking us',
  'still broken','still not working','still not fixed','still occurring',
  'not resolved','data loss','deal breaker','considering leaving',
  'looking at','competitor','switching','really struggling',
  'major issue','serious concern','unresolved for',
];

function _scanText(text, posBonus, negPenalty) {
  const t = (text || '').toLowerCase();
  let delta = 0;
  _SENT_POS.forEach(p => { if (t.includes(p)) delta += posBonus; });
  _SENT_NEG.forEach(p => { if (t.includes(p)) delta -= negPenalty; });
  return delta;
}

function computeSentiment(s) {
  const co = CUSTOMERS.find(c => c.name === s.name);
  const hasCall = co && co.heysam;
  if (!s.total && !hasCall) return null;

  let score = 50;

  // ── HeySam call signal (strongest) ───────────────────────────────────────
  if (hasCall) {
    const h = co.heysam;
    if      (h.overall === 'positive')  score += 12;
    else if (h.overall === 'negative')  score -= 8;
    // Scan call text: topics (cap at ±12) and challenges (cap at ±8)
    score += Math.max(-12, Math.min(12, _scanText(h.key_topics, 3, 2)));
    score += Math.max(-8,  Math.min(8,  _scanText(h.challenges, 2, 3)));
  }

  // ── Activity notes ────────────────────────────────────────────────────────
  if (co) {
    (co.activity || []).forEach(entry => {
      const text = Array.isArray(entry) ? entry.slice(1).join(' ') : '';
      score += Math.max(-3, Math.min(3, _scanText(text, 1, 1)));
    });
  }

  // ── Intercom: subject text + category signal ──────────────────────────────
  (s.items || []).forEach(item => {
    if (item._src !== 'Intercom') return;
    score += Math.max(-3, Math.min(3, _scanText(item.subject, 2, 1)));
    const cat  = item.category || 'Other';
    const open = item.state === 'open';
    if      (cat === 'App / Mobile Issues')           score += open ? -4 : -1;
    else if (cat === 'Vendor Management')             score += open ? -3 :  0;
    else if (cat === 'Login & Account Access')        score += open ? -3 :  0;
    else if (cat === 'Ticket Management')             score += open ? -2 :  0;
    else if (cat === 'Reporting')                     score += open ? -2 :  1;
    else if (cat === 'Rates & Pricing Issues')        score += open ? -2 :  1;
    else if (cat === 'Billing & Invoicing')           score += open ? -2 :  1;
    else if (cat === 'Feature Requests')              score += open ?  1 :  3;
    else if (cat === 'Add / Onboard Driver')          score += open ? -1 :  2;
    else if (cat === 'Driver Type / Role Correction') score += open ? -1 :  1;
    else                                              score += open ? -1 :  1;
  });

  return Math.max(5, Math.min(95, Math.round(score)));
}

function sentimentInfo(score) {
  if (score === null) return { label:'No data',      color:'#475569', band:'none' };
  if (score >= 58)    return { label:'Low Friction',  color:'#22C55E', band:'pos' };
  if (score >= 40)    return { label:'Neutral',       color:'#F59E0B', band:'neut' };
  return                     { label:'High Friction', color:'#EF4444', band:'neg' };
}

// ── Bottom breakdown charts — single-source (src = 'ic' | 'li') ───────────────
function _bottomBar(label, n, maxBar, col, labelStyle) {
  const pct = Math.round(n / maxBar * 100);
  return `<div class="tenure-group">
    <div class="tenure-label"${labelStyle ? ` style="${labelStyle}"` : ''}>${esc(label)}</div>
    <div class="tenure-bars">
      <div class="tenure-bar-row">
        <div class="tenure-bar-track"><div class="tenure-bar-fill" style="width:${pct}%;background:${col}"></div></div>
        <span class="tenure-bar-count">${n}</span>
      </div>
    </div>
  </div>`;
}

function renderTenureChart(stats, src) {
  const GROUPS = ['New','~1 Year','~2 Years','3+ Years'];
  const by = {}; GROUPS.forEach(g => { by[g] = 0; });
  stats.forEach(s => {
    const co = CUSTOMERS.find(c => c.name === s.name);
    if (!co || !by.hasOwnProperty(co.tenure_group)) return;
    by[co.tenure_group] += src === 'ic' ? s.icCount : s.liCount;
  });
  const maxBar = Math.max(...Object.values(by), 1);
  const col = src === 'ic' ? '#60ABDE' : '#FCD34D';
  let html = '<div class="tenure-chart">';
  GROUPS.forEach(g => { if (by[g]) html += _bottomBar(g, by[g], maxBar, col, null); });
  html += '</div>';
  return html;
}

function renderArrBandChart(stats, src) {
  const BANDS = [
    { label:'< $20K',    min:0,      max:20000    },
    { label:'$20K–50K',  min:20000,  max:50000    },
    { label:'$50K–100K', min:50000,  max:100000   },
    { label:'$100K+',    min:100000, max:Infinity },
    { label:'No ARR',    min:-1,     max:-1       },
  ];
  const by = {}; BANDS.forEach(b => { by[b.label] = 0; });
  stats.forEach(s => {
    const co = CUSTOMERS.find(c => c.name === s.name);
    if (!co) return;
    const v = parseArr(co.arr);
    let band = 'No ARR';
    if (co.arr && co.arr !== '—') {
      for (const b of BANDS.slice(0,-1)) {
        if (v >= b.min && v < b.max) { band = b.label; break; }
      }
    }
    by[band] += src === 'ic' ? s.icCount : s.liCount;
  });
  const maxBar = Math.max(...Object.values(by), 1);
  const col = src === 'ic' ? '#60ABDE' : '#FCD34D';
  let html = '<div class="tenure-chart">';
  BANDS.forEach(b => { if (by[b.label]) html += _bottomBar(b.label, by[b.label], maxBar, col, null); });
  html += '</div>';
  return html;
}

function renderSegmentChart(stats, src) {
  const SEGS = ['Enterprise','Mid-Market'];
  const SEG_COLORS = { Enterprise:'#FFE500', 'Mid-Market':'#60ABDE' };
  const by = {}; SEGS.forEach(seg => { by[seg] = 0; });
  stats.forEach(s => {
    if (by.hasOwnProperty(s.segment))
      by[s.segment] += src === 'ic' ? s.icCount : s.liCount;
  });
  const maxBar = Math.max(...Object.values(by), 1);
  let html = '<div class="tenure-chart">';
  SEGS.forEach(seg => {
    if (by[seg]) html += _bottomBar(seg, by[seg], maxBar, SEG_COLORS[seg]||'#64748B', `color:${SEG_COLORS[seg]||'#64748B'}`);
  });
  html += '</div>';
  return html;
}

function renderTypeChart(stats, src) {
  const by = {};
  stats.forEach(s => {
    const co = CUSTOMERS.find(c => c.name === s.name);
    if (!co) return;
    const t = co.type || 'Other';
    by[t] = (by[t] || 0) + (src === 'ic' ? s.icCount : s.liCount);
  });
  const sorted = Object.entries(by).filter(([,n])=>n>0).sort((a,b)=>b[1]-a[1]);
  const maxBar = Math.max(...sorted.map(([,n])=>n), 1);
  let html = '<div class="tenure-chart">';
  sorted.forEach(([t, n]) => {
    const col = TC[t] || '#64748B';
    html += _bottomBar(t, n, maxBar, col, `color:${col}`);
  });
  html += '</div>';
  return html;
}

function renderBottomPanels(stats, src) {
  const label = src === 'ic' ? 'Intercom contacts' : 'Linear tickets';
  return `
    <div class="dash-panel">
      <div class="dash-panel-head">
        <span class="dash-panel-title">By Tenure</span>
        <span class="dash-panel-sub">${label} · time with Tread</span>
      </div>
      ${renderTenureChart(stats, src)}
    </div>
    <div class="dash-panel">
      <div class="dash-panel-head">
        <span class="dash-panel-title">By ARR Band</span>
        <span class="dash-panel-sub">${label} · contract size</span>
      </div>
      ${renderArrBandChart(stats, src)}
    </div>
    <div class="dash-panel">
      <div class="dash-panel-head">
        <span class="dash-panel-title">By Segment</span>
        <span class="dash-panel-sub">${label} · Enterprise vs. Mid-Market</span>
      </div>
      ${renderSegmentChart(stats, src)}
    </div>
    <div class="dash-panel">
      <div class="dash-panel-head">
        <span class="dash-panel-title">By Customer Type</span>
        <span class="dash-panel-sub">${label} · Hauler / Producer / etc.</span>
      </div>
      ${renderTypeChart(stats, src)}
    </div>`;
}

// ── Feedback taxonomy helpers ─────────────────────────────────────────────────
const STOP = new Set(['the','a','an','is','in','it','to','of','for','and','or','with','on','at','from','by','as','i','my','we','our','your','have','has','had','be','been','am','are','was','were','do','does','did','can','could','would','should','will','may','not','no','so','but','if','this','that','they','them','their','he','she','its','us','you','me','just','also','now','one','some','any','all','per','via','still','even','when','where','how','what','who','which','into','up','down','out','over','under','too','very','well','get','got','need','want','see','look','add','new','make','use','find','give','take','able','please','hello','good','morning','afternoon','hi','hey','dear','best','sure','yes','know','let','try','check','set','send','show','ask','tell','work','more','recently','using','used','like','need','another','about','after','have','been','would','something','their','would']);
const NEG_T = new Set(['error','broken','bug','issue','problem','missing','wrong','incorrect','unable','failed','fail','stuck','lost','confused','slow','crash','crashing','crashed','trouble','difficulty','issue','issues','cannot','wont','disappear','disappeared','not','broken','glitch','freeze','freezing','delay']);
const POS_T = new Set(['resolved','fixed','working','appreciate','easy','quick','perfect','improved','success','solved','better','excellent','helpful','awesome','nice','glad','happy','great','thanks','thank','love','smooth']);

function tokenize(text) {
  return (text||'').toLowerCase().replace(/[^a-z0-9\s]/g,' ').split(/\s+/).filter(w => w.length > 2 && !STOP.has(w));
}
function bigrams(tokens) {
  const out = [];
  for (let i=0; i<tokens.length-1; i++) out.push(tokens[i]+' '+tokens[i+1]);
  return out;
}
function wordTone(w) {
  if (NEG_T.has(w)) return 'neg';
  if (POS_T.has(w)) return 'pos';
  return 'neu';
}
function analyzeFeedback(items) {
  const wf = {}, bf = {};
  items.forEach(it => {
    const text = it.subject || it.title || '';
    const toks = tokenize(text);
    toks.forEach(w => { wf[w]=(wf[w]||0)+1; });
    bigrams(toks).forEach(b => { bf[b]=(bf[b]||0)+1; });
  });
  return { wf, bf };
}
function renderTaxonomy(items) {
  if (!items.length) return '<div class="senti-no-data">No conversation text available for this selection.</div>';
  const { wf, bf } = analyzeFeedback(items);
  const entries = Object.entries(wf).sort((a,b)=>b[1]-a[1]).slice(0,60);
  const maxFreq = entries[0]?.[1] || 1;
  const negWords = entries.filter(([w])=>wordTone(w)==='neg');
  const posWords = entries.filter(([w])=>wordTone(w)==='pos');
  const negTotal = negWords.reduce((s,[,n])=>s+n,0);
  const posTotal = posWords.reduce((s,[,n])=>s+n,0);
  const totalSig = negTotal + posTotal || 1;
  const negPct = Math.round(negTotal/totalSig*100);
  const posPct = Math.round(posTotal/totalSig*100);

  // Tag cloud
  const cloud = entries.map(([word, freq]) => {
    const tone = wordTone(word);
    const sz = 11 + Math.round(freq/maxFreq * 13);
    return `<span class="tag-word tag-${tone}" style="font-size:${sz}px" title="${freq} occurrence${freq>1?'s':''}">${esc(word)}<sup>${freq}</sup></span>`;
  }).join('');

  // Phrases (bigrams appearing 2+ times)
  const phrases = Object.entries(bf).filter(([,n])=>n>=2).sort((a,b)=>b[1]-a[1]).slice(0,12);
  const maxPh = phrases[0]?.[1] || 1;
  const phraseHtml = phrases.length ? phrases.map(([ph, n]) => {
    const toks = ph.split(' ');
    const tone = toks.some(w=>NEG_T.has(w)) ? 'neg' : toks.some(w=>POS_T.has(w)) ? 'pos' : 'neu';
    const col = tone==='neg'?'#FCA5A5':tone==='pos'?'#86EFAC':'#94A3B8';
    return `<div class="phrase-row">
      <span class="phrase-text" style="color:${col}">${esc(ph)}</span>
      <div class="phrase-bar-track"><div class="phrase-bar" style="width:${Math.round(n/maxPh*100)}%;background:${col}40"></div></div>
      <span class="phrase-count">${n}</span>
    </div>`;
  }).join('') : '<div style="color:var(--muted);font-size:12px;padding:0 16px 8px">No repeated phrases (need 2+ occurrences)</div>';

  return `
    <div class="sbs-wrap">
      <div class="sbs-bar">
        <div style="width:${posPct}%;background:rgba(34,197,94,.5);height:100%"></div>
        <div style="width:${negPct}%;background:rgba(239,68,68,.5);height:100%;margin-left:auto"></div>
      </div>
      <div class="sbs-labels">
        <span style="color:#86EFAC">${posTotal} positive signal${posTotal!==1?'s':''}</span>
        <span style="color:#FCA5A5">${negTotal} friction signal${negTotal!==1?'s':''}</span>
      </div>
    </div>
    <div class="senti-section-label">Common Terms</div>
    <div class="tag-cloud">${cloud}</div>
    <div class="senti-section-label">Common Phrases</div>
    <div class="phrase-list">${phraseHtml}</div>`;
}

function initSenti() {
  sentiInited = true;
  const stats = buildCustomerStats().filter(s => s.total > 0);
  const allItems = [];
  CUSTOMERS.forEach(co => {
    (co.intercom_90d||[]).forEach(it => allItems.push({...it, _co: co.name, _src:'Intercom'}));
    if (co.heysam) {
      const h = co.heysam;
      if (h.key_topics) allItems.push({subject: h.key_topics, _src:'HeySam', _co: co.name, category:'', state:'', date: h.date, url: h.url});
      if (h.challenges) allItems.push({subject: h.challenges, _src:'HeySam', _co: co.name, category:'', state:'', date: h.date, url: h.url});
    }
  });

  const wrap = document.getElementById('senti-wrap');
  wrap.innerHTML = `
    <div class="senti-title">Feedback Taxonomy</div>
    <div class="senti-sub">Language patterns across ${allItems.length} support interactions · last 90 days · <span style="color:#86EFAC">green = positive signal</span> · <span style="color:#FCA5A5">red = friction signal</span></div>
    <div class="senti-layout">
      <div class="dash-panel" style="min-width:0">
        <div class="dash-panel-head">
          <span class="dash-panel-title">Filter by Customer</span>
          <span class="dash-panel-sub">click to drill in</span>
        </div>
        <div class="senti-list" id="senti-list"></div>
      </div>
      <div class="dash-panel" id="senti-right" style="min-width:0;overflow-y:auto;max-height:680px">
        <div class="dash-panel-head">
          <span class="dash-panel-title" id="senti-right-hd">All Customers</span>
          <span class="dash-panel-sub" id="senti-right-ct">${allItems.length} interactions</span>
        </div>
        <div id="senti-right-body">${renderTaxonomy(allItems)}</div>
      </div>
    </div>`;

  const listEl = document.getElementById('senti-list');
  // All row
  const allRow = document.createElement('div');
  allRow.className = 'senti-row selected';
  allRow.innerHTML = `<span class="senti-name" style="color:var(--yellow)">All Customers</span><span class="senti-co-count">${allItems.length}</span>`;
  allRow.addEventListener('click', () => {
    document.querySelectorAll('#senti-list .senti-row').forEach(r=>r.classList.remove('selected'));
    allRow.classList.add('selected');
    document.getElementById('senti-right-hd').textContent = 'All Customers';
    document.getElementById('senti-right-ct').textContent = allItems.length+' interactions';
    document.getElementById('senti-right-body').innerHTML = renderTaxonomy(allItems);
  });
  listEl.appendChild(allRow);

  stats.forEach(s => {
    const coItems = allItems.filter(it => it._co === s.name);
    const row = document.createElement('div');
    row.className = 'senti-row';
    row.innerHTML = `<span class="senti-name">${esc(s.name)}</span><span class="senti-co-count">${coItems.length}</span>`;
    row.addEventListener('click', () => {
      document.querySelectorAll('#senti-list .senti-row').forEach(r=>r.classList.remove('selected'));
      row.classList.add('selected');
      document.getElementById('senti-right-hd').textContent = s.name;
      document.getElementById('senti-right-ct').textContent = coItems.length+' interaction'+(coItems.length!==1?'s':'');
      document.getElementById('senti-right-body').innerHTML = renderTaxonomy(coItems);
    });
    listEl.appendChild(row);
  });
}

function renderCatBreakdown(cats, title, items, catByCustomer) {
  const total = Object.values(cats).reduce((s,n)=>s+n,0);
  const maxVal = Math.max(...Object.values(cats), 1);
  let html = `<div class="dash-cats-wrap">`;
  if (title) html += `<div style="font-size:13px;font-weight:700;color:var(--text);margin-bottom:12px">${esc(title)}</div>`;
  const sortedCats = Object.entries(cats).filter(([,n])=>n>0).sort((a,b)=>b[1]-a[1]).map(([c])=>c);
  sortedCats.forEach(cat => {
    const n = cats[cat] || 0;
    const pct = Math.round((n/maxVal)*100);
    const col = CAT_COLORS[cat] || '#64748B';
    const coList = catByCustomer && catByCustomer[cat];
    if (coList && coList.length > 0) {
      // Stacked bar: each segment = one customer; hover tooltip shows client breakdown
      const segs = coList.map(co => {
        const sw = (co.count / maxVal * 100).toFixed(1);
        const cc = CO_PALETTE[co.idx % CO_PALETTE.length];
        return `<div class="stacked-seg" style="width:${sw}%;background:${cc}"></div>`;
      }).join('');
      const custJson = JSON.stringify(coList).replace(/&/g,'&amp;').replace(/"/g, '&quot;');
      html += `<div class="dash-cat-row cat-row-hoverable" data-cat="${esc(cat)}" data-total="${n}" data-customers="${custJson}">
        <span class="dash-cat-label"><span class="cat-pill cat-${cat.replace(/ /g,'-')}">${esc(cat)}</span></span>
        <div class="dash-cat-track"><div class="stacked-bar-inner">${segs}</div></div>
        <span class="dash-cat-count">${n}</span>
      </div>`;
    } else {
      html += `<div class="dash-cat-row">
        <span class="dash-cat-label"><span class="cat-pill cat-${cat.replace(/ /g,'-')}">${esc(cat)}</span></span>
        <div class="dash-cat-track"><div class="dash-cat-fill" style="width:${pct}%;background:${col}"></div></div>
        <span class="dash-cat-count">${n}</span>
      </div>`;
    }
  });
  if (total === 0) html += `<div style="color:var(--muted);font-size:13px;padding:8px 0">No interactions in the last 90 days</div>`;
  html += `</div>`;

  if (items && items.length) {
    const icItems = items.filter(i=>i._src==='Intercom');
    if (icItems.length) {
      html += `<div class="dash-links"><div class="dash-link-group">
        <div class="dash-link-group-label">Intercom (${icItems.length})</div>`;
      icItems.forEach(i => {
        html += `<div class="iact">
          <span class="iact-src Intercom">Intercom</span>
          <div class="iact-body">
            <div class="iact-title"><a href="${esc(i.url)}" target="_blank">${esc(i.subject||i.url)}</a>
              <span class="cat-pill ${catClass(i.category)}">${esc(i.category||'Other')}</span>
            </div>
            <div class="iact-meta">${esc(i.date)} · <span style="color:${i.state==='open'?'#4ADE80':'#94A3B8'}">${esc(i.state||'')}</span></div>
          </div>
        </div>`;
      });
      html += `</div></div>`;
    }
  }
  return html;
}

function buildOpenConvsPanel(stats, stateFilter) {
  const sf = stateFilter || 'all';
  const panelTitle = sf === 'closed' ? 'Closed Conversations by Customer' : sf === 'open' ? 'Open Conversations by Customer' : 'Conversations by Customer';
  const rows = [];
  stats.forEach(s => {
    const byCat = {};
    (s.items || []).filter(i => i._src === 'Intercom').forEach(i => {
      const c = i.category || 'Other';
      if (!byCat[c]) byCat[c] = [];
      byCat[c].push({subject: i.subject || 'Conversation', url: i.url || ''});
    });
    const total = Object.values(byCat).reduce((a, v) => a + v.length, 0);
    if (total === 0) return;
    rows.push({ name: s.name, byCat, total });
  });
  if (rows.length === 0) return '<div style="padding:20px;font-size:13px;color:var(--muted)">No conversations match the current filter.</div>';
  rows.sort((a, b) => b.total - a.total);
  const maxTotal = rows[0].total;
  const grandTotal = rows.reduce((a, r) => a + r.total, 0);

  // Category order: by overall volume descending
  const catTotals = {};
  rows.forEach(r => Object.entries(r.byCat).forEach(([c, items]) => {
    catTotals[c] = (catTotals[c] || 0) + items.length;
  }));
  const usedCats = Object.keys(catTotals).sort((a, b) => catTotals[b] - catTotals[a]);

  const legend = usedCats.map(c =>
    `<span class="bug-leg-item"><span class="bug-leg-dot" style="background:${CAT_COLORS[c]||'#64748B'}"></span>${esc(c)}</span>`
  ).join('');

  let html = `<div class="bug-panel">
    <div class="bug-panel-head">
      <span class="bug-panel-title">${panelTitle}</span>
      <div class="bug-legend">${legend}</div>
    </div>
    <div style="font-size:11px;color:var(--muted);padding:8px 16px 2px">${grandTotal} across ${rows.length} customers · sorted highest to lowest · hover for details</div>
    <div class="bug-rows">`;

  rows.forEach(row => {
    const barPct = Math.round(row.total / maxTotal * 100);
    html += `<div class="bug-row">
      <div class="bug-row-name" title="${esc(row.name)}">${esc(row.name)}</div>
      <div class="bug-bar-outer"><div class="bug-bar-inner" style="width:${barPct}%">`;
    usedCats.forEach(cat => {
      const items = row.byCat[cat];
      if (!items || !items.length) return;
      const segPct = Math.round(items.length / row.total * 100);
      const color = CAT_COLORS[cat] || '#64748B';
      const convJson = JSON.stringify(items)
        .replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
      html += `<div class="bug-seg" style="width:${segPct}%;background:${color}" data-cat="${esc(cat)}" data-co="${esc(row.name)}" data-count="${items.length}" data-convs="${convJson}"></div>`;
    });
    html += `</div></div>
      <div class="bug-row-count">${row.total}</div>
    </div>`;
  });

  html += `</div></div>`;
  return html;
}

// ── Insights filter state ────────────────────────────────────────────────────
let _baseStats     = null;
let _filteredStats = null;
let _dashState     = 'all';
let _dashDays      = 90;

function filterStats(base, stateFilter, days) {
  const cutoff = days ? new Date(Date.now() - days * 864e5).toISOString().slice(0,10) : null;
  return base.map(s => {
    const icFiltered = (s.items || []).filter(i => {
      if (i._src !== 'Intercom') return false;
      if (stateFilter !== 'all' && i.state !== stateFilter) return false;
      if (cutoff && !i.date) return false;
      if (cutoff && i.date && i.date < cutoff) return false;
      return true;
    });
    const cats = {};
    icFiltered.forEach(i => { const c = i.category || 'Other'; cats[c] = (cats[c] || 0) + 1; });
    const total = icFiltered.length;
    const liItems = (s.items || []).filter(i => i._src !== 'Intercom');
    return { ...s, items: [...icFiltered, ...liItems], cats, total, icCount: total };
  }).sort((a, b) => b.total - a.total);
}

function refreshDash() {
  if (!_baseStats) return;
  _filteredStats = filterStats(_baseStats, _dashState, _dashDays);
  const stats   = _filteredStats;
  const overall = overallCats(stats);
  const maxTotal = Math.max(...stats.map(s => s.total), 1);
  const totalInteractions = stats.reduce((s, r) => s + r.total, 0);
  const catByCo = buildCatByCustomer(stats);

  const daysLabel = _dashDays === 30  ? 'Last 30 Days'
                  : _dashDays === 60  ? 'Last 60 Days'
                  : _dashDays === 90  ? 'Last 90 Days'
                  : _dashDays === 183 ? 'Last 6 Months'
                  : 'All Time';
  const stateLabel = _dashState === 'open' ? 'open ' : _dashState === 'closed' ? 'closed ' : '';
  document.getElementById('dash-title-el').textContent = `Support Contacts — ${daysLabel}`;
  document.getElementById('dash-sub-el').textContent   = `${totalInteractions} ${stateLabel}Intercom conversations across ${stats.filter(s => s.total > 0).length} customers · ${ALL_LINEAR.length} Linear tickets tracked separately`;

  document.getElementById('dash-right-sub').textContent   = `${totalInteractions} Intercom conversations`;
  document.getElementById('dash-right-body').innerHTML    = renderCatBreakdown(overall, null, null, catByCo);

  const srcSelect = document.getElementById('dash-src-select');
  document.getElementById('dash-bottom-row').innerHTML = renderBottomPanels(stats, srcSelect ? srcSelect.value : 'ic');
  document.getElementById('dash-open-convs').innerHTML  = buildOpenConvsPanel(stats, _dashState);
}

function initDash() {
  dashInited = true;
  _baseStats = buildCustomerStats();

  // Build static registration panel HTML
  const regHtml = buildRegWeeklyChart();

  // Set all HTML in one shot — avoids innerHTML += destroying listeners
  const wrap = document.getElementById('dash-wrap');
  wrap.innerHTML = `
    <div class="dash-filter-bar">
      <span class="dash-filter-label">Status</span>
      <div class="dash-toggle-group" id="dash-state-group">
        <button class="dash-toggle active" data-state="all">All</button>
        <button class="dash-toggle" data-state="open">Open</button>
        <button class="dash-toggle" data-state="closed">Closed</button>
      </div>
      <span class="dash-filter-label" style="margin-left:8px">Date Range</span>
      <div class="dash-toggle-group" id="dash-days-group">
        <button class="dash-toggle" data-days="30">30d</button>
        <button class="dash-toggle" data-days="60">60d</button>
        <button class="dash-toggle active" data-days="90">90d</button>
        <button class="dash-toggle" data-days="183">6 mo</button>
        <button class="dash-toggle" data-days="0">All time</button>
      </div>
    </div>
    <div id="dash-title-el" class="dash-title"></div>
    <div id="dash-sub-el" class="dash-sub"></div>
    <div class="dash-panel" id="dash-right-panel">
      <div class="dash-panel-head">
        <span class="dash-panel-title" id="dash-right-title">All Customers — Category Breakdown</span>
        <span class="dash-panel-sub" id="dash-right-sub"></span>
      </div>
      <div id="dash-right-body"></div>
    </div>
    <div id="dash-open-convs"></div>
    <div style="margin-top:20px">
      <div class="dash-src-bar">
        <span style="font-size:12px;font-weight:600;color:var(--muted)">Data source</span>
        <select id="dash-src-select" class="dash-src-select">
          <option value="ic">Intercom — support contacts</option>
          <option value="li">Linear — internal tickets</option>
        </select>
      </div>
      <div class="dash-layout" id="dash-bottom-row" style="grid-template-columns:repeat(4,1fr)"></div>
    </div>
    ${regHtml}`;

  // Populate data panels via refreshDash (uses _baseStats + current filter state)
  refreshDash();

  // Event delegation on wrap — survives any child innerHTML changes
  wrap.addEventListener('click', e => {
    // Filter toggle: state
    const stBtn = e.target.closest('.dash-toggle[data-state]');
    if (stBtn) {
      document.querySelectorAll('#dash-state-group .dash-toggle').forEach(b => b.classList.remove('active'));
      stBtn.classList.add('active');
      _dashState = stBtn.dataset.state;
      refreshDash();
      return;
    }
    // Filter toggle: days
    const dayBtn = e.target.closest('.dash-toggle[data-days]');
    if (dayBtn) {
      document.querySelectorAll('#dash-days-group .dash-toggle').forEach(b => b.classList.remove('active'));
      dayBtn.classList.add('active');
      _dashDays = +dayBtn.dataset.days || 0;
      refreshDash();
      return;
    }
  });
  wrap.addEventListener('change', e => {
    if (e.target.id === 'dash-src-select') {
      document.getElementById('dash-bottom-row').innerHTML = renderBottomPanels(_filteredStats, e.target.value);
    }
  });
}



// ── New Driver Registrations — weekly stacked bar chart ──────────────────────
const REG_PALETTE = [
  '#3B82F6','#F59E0B','#10B981','#8B5CF6','#EF4444',
  '#06B6D4','#F97316','#EC4899','#84CC16','#6366F1',
];
const REG_OTHER_COLOR = '#475569';

function _regTitleCase(s) {
  const SMALL = new Set(['and','or','of','the','a','an','co','inc','llc','ltd','corp']);
  return s.toLowerCase().split(' ').map((w,i) => {
    if (i > 0 && SMALL.has(w)) return w;
    return w.charAt(0).toUpperCase() + w.slice(1);
  }).join(' ');
}

function buildRegWeeklyChart() {
  if (!REGISTRATIONS || !REGISTRATIONS.length) return '';

  // Per-customer totals to assign colors to top 10
  const coTotals = {};
  REGISTRATIONS.forEach(wk => {
    Object.entries(wk).forEach(([co, n]) => {
      if (co === 'week') return;
      coTotals[co] = (coTotals[co] || 0) + n;
    });
  });
  const topCos = Object.entries(coTotals).sort((a,b) => b[1]-a[1]).slice(0,10).map(([co]) => co);
  const coColor = {};
  topCos.forEach((co, i) => { coColor[co] = REG_PALETTE[i]; });

  const weekTotals = REGISTRATIONS.map(wk =>
    Object.entries(wk).filter(([k]) => k !== 'week').reduce((s,[,n]) => s+n, 0)
  );
  const maxTotal  = Math.max(...weekTotals, 1);
  const grandTotal = weekTotals.reduce((s,n) => s+n, 0);
  const maxBarPx  = 110;  // pixels for tallest bar

  let bars = '';
  REGISTRATIONS.forEach((wk, wi) => {
    const wkTotal  = weekTotals[wi];
    const barH     = Math.max(Math.round(wkTotal / maxTotal * maxBarPx), 2);
    const entries  = Object.entries(wk).filter(([k]) => k !== 'week').sort((a,b) => b[1]-a[1]);
    const bdJson   = JSON.stringify(entries).replace(/&/g,'&amp;').replace(/"/g,'&quot;');

    let segs = '';
    entries.forEach(([co, n]) => {
      const color = coColor[co] || REG_OTHER_COLOR;
      segs += `<div class="reg-week-seg" style="flex:${n};background:${color}"></div>`;
    });

    bars += `<div class="reg-week-col">
      <div class="reg-week-count">${wkTotal}</div>
      <div class="reg-week-bar-wrap">
        <div class="reg-week-bar" style="height:${barH}px"
             data-week="${esc(wk.week)}" data-total="${wkTotal}" data-bd="${bdJson}">
          ${segs}
        </div>
      </div>
      <div class="reg-week-label">${esc(wk.week)}</div>
    </div>`;
  });

  // Check if any customers outside top 10 appear in the data
  const hasOther = REGISTRATIONS.some(wk =>
    Object.keys(wk).filter(k => k !== 'week').some(co => !topCos.includes(co))
  );
  const legend = topCos.map((co, i) =>
    `<span class="reg-leg-item"><span class="reg-leg-dot" style="background:${REG_PALETTE[i]}"></span>${esc(_regTitleCase(co))}</span>`
  ).join('') + (hasOther ? `<span class="reg-leg-item"><span class="reg-leg-dot" style="background:${REG_OTHER_COLOR}"></span><em>Other customers</em></span>` : '');

  return `<div class="reg-panel">
    <div class="reg-panel-head">
      <span class="reg-panel-title">New Driver Registrations</span>
      <span class="reg-panel-sub">${grandTotal} self-registrations naming a Tread customer · last 4 weeks</span>
    </div>
    <div class="reg-chart-wrap"><div class="reg-chart-area">${bars}</div></div>
    <div class="reg-legend">${legend}</div>
  </div>`;
}

// ── Reg chart tooltip ─────────────────────────────────────────────────────────
(function(){
  let tip = document.getElementById('reg-tip');
  if (!tip) { tip = document.createElement('div'); tip.id = 'reg-tip'; document.body.appendChild(tip); }

  function showRegTip(e, week, total, entries) {
    const rows = entries.slice(0,10).map(([co, n]) =>
      `<div class="reg-tip-row"><span class="reg-tip-co">${esc(_regTitleCase(co))}</span><span class="reg-tip-n">${n}</span></div>`
    ).join('');
    const more = entries.length > 10 ? `<div style="font-size:10px;color:var(--muted);margin-top:4px">+ ${entries.length-10} more</div>` : '';
    tip.innerHTML = `<div class="reg-tip-head">Week of ${esc(week)} &mdash; ${total} total</div>${rows}${more}`;
    tip.style.display = 'block';
    moveRegTip(e);
  }
  function moveRegTip(e) {
    const tw = tip.offsetWidth||240, th = tip.offsetHeight||120;
    let x = e.clientX+16, y = e.clientY+14;
    if (x+tw > window.innerWidth-8)  x = e.clientX-tw-8;
    if (y+th > window.innerHeight-8) y = e.clientY-th-8;
    tip.style.left = x+'px'; tip.style.top = y+'px';
  }
  function hideRegTip() { tip.style.display = 'none'; }

  document.addEventListener('mousemove', e => { if (tip.style.display!=='none') moveRegTip(e); });
  document.addEventListener('mouseover', e => {
    const bar = e.target.closest('.reg-week-bar');
    if (!bar) { hideRegTip(); return; }
    try {
      const entries = JSON.parse(bar.dataset.bd);
      showRegTip(e, bar.dataset.week, +bar.dataset.total, entries);
    } catch(err) { hideRegTip(); }
  });
  document.addEventListener('mouseleave', hideRegTip);
})();

// ── Init ──────────────────────────────────────────────────────────────────────
buildHeader();
renderGrid();

// ── Category bar hover tooltip ────────────────────────────────────────────────
(function(){
  let tip = document.getElementById('cat-tip');
  if (!tip) { tip = document.createElement('div'); tip.id = 'cat-tip'; document.body.appendChild(tip); }

  function showTip(e, cat, coList, total) {
    const catColor = CAT_COLORS[cat] || '#64748B';
    const maxCount = Math.max(...coList.map(c=>c.count), 1);
    const rows = coList.map(co => {
      const cc = CO_PALETTE[co.idx % CO_PALETTE.length];
      const bw = Math.round(co.count / maxCount * 100);
      return `<div class="cat-tip-item">
        <span class="cat-tip-dot" style="background:${cc}"></span>
        <span class="cat-tip-name">${esc(co.name)}</span>
        <div class="cat-tip-bar-track"><div class="cat-tip-bar-fill" style="width:${bw}%;background:${cc}"></div></div>
        <span class="cat-tip-count">${co.count}</span>
      </div>`;
    }).join('');
    tip.innerHTML = `
      <div class="cat-tip-head">
        <span class="cat-pill cat-${cat.replace(/ /g,'-')}">${esc(cat)}</span>
        <span style="color:var(--muted);font-size:11px;font-weight:400">${total} ticket${total!==1?'s':''}</span>
      </div>
      <div class="cat-tip-list">${rows}</div>`;
    tip.style.display = 'block';
    moveTip(e);
  }

  function moveTip(e) {
    const tw = tip.offsetWidth || 260, th = tip.offsetHeight || 160;
    let x = e.clientX + 16, y = e.clientY + 14;
    if (x + tw > window.innerWidth  - 8) x = e.clientX - tw - 8;
    if (y + th > window.innerHeight - 8) y = e.clientY - th - 8;
    tip.style.left = x + 'px';
    tip.style.top  = y + 'px';
  }

  function hideTip() { tip.style.display = 'none'; }

  document.addEventListener('mousemove', e => {
    if (tip.style.display === 'none') return;
    moveTip(e);
  });

  document.addEventListener('mouseover', e => {
    const row = e.target.closest('.cat-row-hoverable');
    if (!row) { hideTip(); return; }
    try {
      const coList = JSON.parse(row.getAttribute('data-customers'));
      if (!coList || !coList.length) { hideTip(); return; }
      showTip(e, row.dataset.cat, coList, parseInt(row.dataset.total || '0'));
    } catch(err) { hideTip(); }
  });

  document.addEventListener('mouseleave', hideTip);
})();

// ── Notes (localStorage) ─────────────────────────────────────────────────────
const NOTES_KEY = 'tread_notes';

function _allNotes() {
  try { return JSON.parse(localStorage.getItem(NOTES_KEY) || '{}'); } catch { return {}; }
}

function loadNotes(name) {
  document.getElementById('m-notes').dataset.customer = name;
  document.getElementById('notes-input').value = '';
  _renderNotesList(name);
}

function _renderNotesList(name) {
  const all = _allNotes();
  const notes = (all[name] || []).slice().sort((a,b) => b.id - a.id);
  const list = document.getElementById('notes-list');
  if (!notes.length) {
    list.innerHTML = '<div class="notes-empty">No notes yet</div>';
    return;
  }
  list.innerHTML = notes.map(n => {
    const d = new Date(n.id);
    const ts = d.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'})
             + ' · '
             + d.toLocaleTimeString('en-US',{hour:'numeric',minute:'2-digit'});
    const safeId = JSON.stringify(n.id);
    const safeName = JSON.stringify(name);
    return `<div class="note-item">
      <div class="note-meta">
        <span class="note-ts">${ts}</span>
        <button class="note-del" onclick="deleteNote(${safeName},${safeId})" title="Delete">&#x2715;</button>
      </div>
      <div class="note-text">${esc(n.text)}</div>
    </div>`;
  }).join('');
}

function saveNote(name, text) {
  if (!text.trim()) return;
  const all = _allNotes();
  if (!all[name]) all[name] = [];
  all[name].push({ id: Date.now(), text: text.trim() });
  localStorage.setItem(NOTES_KEY, JSON.stringify(all));
  _renderNotesList(name);
}

function deleteNote(name, id) {
  const all = _allNotes();
  if (all[name]) all[name] = all[name].filter(n => n.id !== id);
  localStorage.setItem(NOTES_KEY, JSON.stringify(all));
  _renderNotesList(name);
}

function _renderLeNotes(idx, name) {
  const list = document.getElementById('le-nl-'+idx);
  if (!list) return;
  const all = _allNotes();
  const notes = (all[name] || []).slice().sort((a,b) => b.id - a.id);
  if (!notes.length) { list.innerHTML = '<div class="notes-empty">No notes yet</div>'; return; }
  list.innerHTML = notes.map(n => {
    const d = new Date(n.id);
    const ts = d.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'}) + ' · ' + d.toLocaleTimeString('en-US',{hour:'numeric',minute:'2-digit'});
    return `<div class="note-item">
      <div class="note-meta"><span class="note-ts">${ts}</span>
        <button class="note-del" onclick="deleteLeNote(${idx},${JSON.stringify(name)},${n.id})" title="Delete">&#x2715;</button>
      </div>
      <div class="note-text">${esc(n.text)}</div>
    </div>`;
  }).join('');
}
function saveLeNote(idx, name, text) {
  if (!text.trim()) return;
  const all = _allNotes();
  if (!all[name]) all[name] = [];
  all[name].push({ id: Date.now(), text: text.trim() });
  localStorage.setItem(NOTES_KEY, JSON.stringify(all));
  _renderLeNotes(idx, name);
}
function deleteLeNote(idx, name, id) {
  const all = _allNotes();
  if (all[name]) all[name] = all[name].filter(n => n.id !== id);
  localStorage.setItem(NOTES_KEY, JSON.stringify(all));
  _renderLeNotes(idx, name);
}

document.getElementById('notes-save-btn').addEventListener('click', () => {
  const name = document.getElementById('m-notes').dataset.customer;
  const inp  = document.getElementById('notes-input');
  saveNote(name, inp.value);
  inp.value = '';
});

document.getElementById('notes-input').addEventListener('keydown', e => {
  if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
    const name = document.getElementById('m-notes').dataset.customer;
    saveNote(name, e.target.value);
    e.target.value = '';
    e.preventDefault();
  }
});

// ── Open conversations stacked bar tooltip ────────────────────────────────────
(function(){
  let tip = document.getElementById('bug-tip');
  if (!tip) { tip = document.createElement('div'); tip.id = 'bug-tip'; document.body.appendChild(tip); }

  function showTip(e, cat, coName, count, convs) {
    const color = CAT_COLORS[cat] || '#64748B';
    const shown = convs.slice(0, 8);
    const items = shown.map(c =>
      `<span class="bug-tip-item">· ${esc(c.subject)}</span>`
    ).join('');
    const more = convs.length > 8 ? `<span class="bug-tip-item" style="color:var(--muted)">+ ${convs.length - 8} more</span>` : '';
    tip.innerHTML = `
      <div class="bug-tip-head" style="color:${color}">${esc(cat)}</div>
      <div class="bug-tip-sub">${esc(coName)} &mdash; ${count} conversation${count!==1?'s':''}</div>
      <div class="bug-tip-list">${items}${more}</div>`;
    tip.style.display = 'block';
    moveTip(e);
  }

  function moveTip(e) {
    const tw = tip.offsetWidth || 280, th = tip.offsetHeight || 140;
    let x = e.clientX + 16, y = e.clientY + 14;
    if (x + tw > window.innerWidth  - 8) x = e.clientX - tw - 8;
    if (y + th > window.innerHeight - 8) y = e.clientY - th - 8;
    tip.style.left = x + 'px';
    tip.style.top  = y + 'px';
  }

  function hideTip() { tip.style.display = 'none'; }

  document.addEventListener('mousemove', e => {
    if (tip.style.display === 'none') return;
    moveTip(e);
  });

  document.addEventListener('mouseover', e => {
    const seg = e.target.closest('.bug-seg');
    if (!seg) { hideTip(); return; }
    try {
      const convs = JSON.parse(seg.dataset.convs || '[]');
      showTip(e, seg.dataset.cat, seg.dataset.co, +seg.dataset.count, convs);
    } catch(err) { hideTip(); }
  });

  document.addEventListener('mouseleave', hideTip);
})();
</script>
<div id="cat-tip"></div>
<div id="bug-tip"></div>
<div id="reg-tip"></div>
</body>
</html>
"""

# Embed the data
data_blob = (
    HTML
    .replace('CUSTOMERS_PLACEHOLDER',    json.dumps(companies,   ensure_ascii=False))
    .replace('MAP_PTS_PLACEHOLDER',      json.dumps(map_pts,     ensure_ascii=False))
    .replace('ALL_INTERCOM_PLACEHOLDER', json.dumps(ALL_INTERCOM, ensure_ascii=False))
    .replace('ALL_LINEAR_PLACEHOLDER',   json.dumps(ALL_LINEAR,   ensure_ascii=False))
    .replace('REGISTRATIONS_PLACEHOLDER', json.dumps(REGISTRATION_REQUESTS, ensure_ascii=False))
    .replace('BUILD_DATE_PLACEHOLDER',   datetime.datetime.now().strftime('%-m/%-d/%Y'))
)

out = os.path.expanduser('~/Desktop/tread_customers.html')
with open(out, 'w') as f:
    f.write(data_blob)
print(f"Written {len(data_blob):,} bytes → {out}")
