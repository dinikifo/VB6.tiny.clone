import json
from copy import deepcopy


class VBJsonRuntime:
    SCHEMAS = {
        "Root": {
            "ledger": {
                "journals": [],
                "postings": [],
                "postingsText": ""
            },
            "meta": {
                "nextJournalId": 1,
                "nextPostingSeq": 1
            }
        },
        "Customer": {
            "name": "",
            "age": 0
        }
    }

    @staticmethod
    def json_new(type_name: str):
        tpl = VBJsonRuntime.SCHEMAS.get(type_name)
        if tpl is None:
            # fallback: empty object
            return {}
        return deepcopy(tpl)

    @staticmethod
    def json_parse(text: str):
        return json.loads(text)

    @staticmethod
    def json_stringify(value):
        return json.dumps(value)

    @staticmethod
    def _split_path(path: str):
        # Parse "a.b[0].c" -> ["a","b",0,"c"]
        tokens = []
        buf = ""
        i = 0
        while i < len(path):
            ch = path[i]
            if ch == '.':
                if buf:
                    tokens.append(buf)
                    buf = ""
                i += 1
            elif ch == '[':
                if buf:
                    tokens.append(buf)
                    buf = ""
                i += 1
                num = ""
                while i < len(path) and path[i] != ']':
                    num += path[i]
                    i += 1
                i += 1  # skip ]
                tokens.append(int(num))
            else:
                buf += ch
                i += 1
        if buf:
            tokens.append(buf)
        return tokens

    @staticmethod
    def json_get(obj, path: str):
        parts = VBJsonRuntime._split_path(path)
        cur = obj
        for p in parts:
            if isinstance(p, int):
                cur = cur[p]
            else:
                cur = cur[p]
        return cur

    @staticmethod
    def json_set(obj, path: str, value):
        parts = VBJsonRuntime._split_path(path)
        if not isinstance(obj, (dict, list)):
            raise TypeError(
                f"JsonSet: first argument must be dict/list, got {type(obj).__name__}"
            )

        cur = obj
        for p in parts[:-1]:
            if isinstance(p, int):
                if not isinstance(cur, list):
                    raise TypeError(
                        f"JsonSet: trying to index non-list with [{p}], got {type(cur).__name__}"
                    )
                cur = cur[p]
            else:
                if not isinstance(cur, dict):
                    raise TypeError(
                        f"JsonSet: trying to use key '{p}' on non-dict {type(cur).__name__}"
                    )
                if p not in cur or cur[p] is None:
                    cur[p] = {}
                cur = cur[p]

        last = parts[-1]
        if isinstance(last, int):
            if not isinstance(cur, list):
                raise TypeError(
                    f"JsonSet: trying to index non-list with [{last}], got {type(cur).__name__}"
                )
            cur[last] = value
        else:
            if not isinstance(cur, dict):
                raise TypeError(
                    f"JsonSet: trying to use key '{last}' on non-dict {type(cur).__name__}"
                )
            cur[last] = value


class VBContext:
    def __init__(self):
        self.globals = {}

    def set_var(self, name, value):
        self.globals[name] = value

    def get_var(self, name):
        return self.globals.get(name, None)

    def has_var(self, name):
        return name in self.globals

# --- Added by assistant: JSON ledger load/save helpers ---

from pathlib import Path
from datetime import datetime

# Default ledger file next to this runtime.py
LEDGER_FILE = Path(__file__).with_name("ledger.json")


def _ensure_root_defaults(root):
    """Ensure the loaded JSON has the expected ledger/meta structure."""
    if not isinstance(root, dict):
        root = {}

    ledger = root.get("ledger")
    if not isinstance(ledger, dict):
        ledger = {}
    if "journals" not in ledger or not isinstance(ledger.get("journals"), list):
        ledger["journals"] = []
    if "postings" not in ledger or not isinstance(ledger.get("postings"), list):
        ledger["postings"] = []
    if "postingsText" not in ledger or not isinstance(ledger.get("postingsText"), str):
        ledger["postingsText"] = ""
    root["ledger"] = ledger

    meta = root.get("meta")
    if not isinstance(meta, dict):
        meta = {}
    if "nextJournalId" not in meta:
        meta["nextJournalId"] = 1
    if "nextPostingSeq" not in meta:
        meta["nextPostingSeq"] = 1
    root["meta"] = meta

    return root


def load_app_data(path: str | Path | None = None):
    """Load the Root object from JSON file, or create a fresh one from schema."""
    p = Path(path) if path is not None else LEDGER_FILE
    if p.exists():
        with p.open("r", encoding="utf-8") as f:
            raw = json.load(f)
        return _ensure_root_defaults(raw)
    # No file yet: start from Root schema if available, otherwise empty dict
    try:
        root = VBJsonRuntime.json_new("Root")
    except Exception:
        root = {}
    return _ensure_root_defaults(root)


def save_app_data(app_data, path: str | Path | None = None):
    """Save the Root object to JSON, plus a timestamped backup."""
    p = Path(path) if path is not None else LEDGER_FILE

    # Write to temp then atomically replace
    tmp = p.with_suffix(p.suffix + ".tmp")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(app_data, f, indent=2)
    tmp.replace(p)

    # Timestamped backup e.g. ledger.2025.11.21.13.45.00.json
    ts = datetime.now().strftime("%Y.%m.%d.%H.%M.%S")
    backup = p.with_name(f"{p.stem}.{ts}{p.suffix}")
    with backup.open("w", encoding="utf-8") as f:
        json.dump(app_data, f, indent=2)


# === Accounting-oriented Root schema and helpers (added) ===

# Override the Root schema to include accounts, assetTypes, batches,
# journals, postings, and postingsText, plus meta counters for IDs.
VBJsonRuntime.SCHEMAS["Root"] = {
    "ledger": {
        "accounts": [],
        "assetTypes": [],
        "batches": [],
        "journals": [],
        "postings": [],
        "postingsText": ""
    },
    "meta": {
        "nextAccountId": 1,
        "nextAssetTypeId": 1,
        "nextBatchId": 1,
        "nextJournalId": 1,
        "nextPostingSeq": 1
    }
}


def _ensure_root_defaults(root):
    """Ensure the loaded JSON has the expected ledger/meta structure.

    This version is accounting-aware and matches the Root schema above.
    """
    if not isinstance(root, dict):
        root = {}

    ledger = root.get("ledger")
    if not isinstance(ledger, dict):
        ledger = {}

    # Core tables as arrays
    for key in ("accounts", "assetTypes", "batches", "journals", "postings"):
        val = ledger.get(key)
        if not isinstance(val, list):
            ledger[key] = []
    # Extra free-form text for manual postings or debugging
    if not isinstance(ledger.get("postingsText"), str):
        ledger["postingsText"] = ""

    root["ledger"] = ledger

    meta = root.get("meta")
    if not isinstance(meta, dict):
        meta = {}

    for key in ("nextAccountId", "nextAssetTypeId", "nextBatchId",
                "nextJournalId", "nextPostingSeq"):
        if key not in meta or not isinstance(meta.get(key), int):
            meta[key] = 1

    root["meta"] = meta
    return root


def _find_by_code(seq, code):
    """Utility: find object with .code == code (case-sensitive)."""
    for obj in seq:
        if isinstance(obj, dict) and obj.get("code") == code:
            return obj
    return None


def get_or_create_asset_type(root, code: str, description: str | None = None):
    """Look up an asset type by code; create it if missing."""
    root = _ensure_root_defaults(root)
    ledger = root["ledger"]
    meta = root["meta"]

    if not code:
        code = "CUR"  # default currency code

    asset_types = ledger["assetTypes"]
    existing = _find_by_code(asset_types, code)
    if existing is not None:
        return existing

    at_id = meta.get("nextAssetTypeId", 1)
    meta["nextAssetTypeId"] = at_id + 1
    obj = {
        "id": at_id,
        "code": code,
        "description": description or code,
    }
    asset_types.append(obj)
    return obj


def get_or_create_account(
    root,
    code: str,
    name: str | None = None,
    account_type: str = "generic",
    asset_type_code: str | None = None,
):
    """Look up an account by code; create it if missing.

    This keeps the JSON model loosely aligned with the relational
    ACCOUNT table from the article: it has an id, a code, a name, and
    a default asset type.
    """
    root = _ensure_root_defaults(root)
    ledger = root["ledger"]
    meta = root["meta"]

    if not code:
        code = "UNSPEC"

    accounts = ledger["accounts"]
    existing = _find_by_code(accounts, code)
    if existing is not None:
        return existing

    at = get_or_create_asset_type(root, asset_type_code or "CUR")

    acc_id = meta.get("nextAccountId", 1)
    meta["nextAccountId"] = acc_id + 1
    obj = {
        "id": acc_id,
        "code": code,
        "name": name or code,
        "type": account_type,
        "assetTypeId": at["id"],
    }
    accounts.append(obj)
    return obj


def create_journal(
    root,
    date: str,
    description: str,
    period: str | None = None,
    batch_id: int | None = None,
):
    """Create a JOURNAL row in the JSON ledger and return its id."""
    root = _ensure_root_defaults(root)
    ledger = root["ledger"]
    meta = root["meta"]

    if not period:
        # simple period: YYYY-MM from the date, or raw date if unknown
        if isinstance(date, str) and len(date) >= 7 and date[4] == "-" and date[7] == "-":
            period = date[:7]
        else:
            period = "0000-00"

    j_id = meta.get("nextJournalId", 1)
    meta["nextJournalId"] = j_id + 1

    journal_obj = {
        "id": j_id,
        "date": date,
        "description": description,
        "period": period,
    }
    if batch_id is not None:
        journal_obj["batchId"] = batch_id

    ledger["journals"].append(journal_obj)
    return j_id


def post_entry(
    root,
    account_code: str,
    asset_type_code: str,
    period: str,
    journal_id: int,
    amount: float,
    memo: str | None = None,
):
    """Append a POSTING row for the given journal.

    This corresponds to inserting a row into the POSTING table in the
    relational design, but here we just extend the JSON array.
    """
    root = _ensure_root_defaults(root)
    ledger = root["ledger"]
    meta = root["meta"]

    # Ensure account and asset type exist
    account = get_or_create_account(root, account_code, asset_type_code=asset_type_code)
    asset_type = get_or_create_asset_type(root, asset_type_code)

    if not period:
        period = "0000-00"

    seq = meta.get("nextPostingSeq", 1)
    meta["nextPostingSeq"] = seq + 1

    posting = {
        "id": seq,
        "journalId": int(journal_id),
        "accountId": account["id"],
        "assetTypeId": asset_type["id"],
        "period": period,
        "amount": float(amount),
    }
    if memo:
        posting["memo"] = memo

    ledger["postings"].append(posting)
    return seq
