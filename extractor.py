"""
extractor.py

Uses Claude (claude-opus-4-6) to extract real estate deal assumptions from
free-form deal summary text. Returns a structured dict keyed by the
assumption names defined in the template manifest.

For multi-tenant templates (those with a "rent_roll" section in the manifest),
also extracts a tenants list via a separate Claude call.
"""

from __future__ import annotations

import json
import os
from typing import Any

import anthropic
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# Type aliases
# ---------------------------------------------------------------------------
ExtractionResult = dict[str, dict[str, Any]]

# Tenant entry returned for multi-tenant models:
# [{ "tenant_name": str, "unit": str, "sq_ft": int, "base_rent_psf": float,
#    "ti_psf": float, "annual_pct_increase": float,
#    "tax_reimbursement": "y"|"n", "ins_reimbursement": "y"|"n",
#    "cam_reimbursement": "y"|"n",
#    "confidence": "high"|"low", "note": str }, ...]
TenantList = list[dict[str, Any]]


# ---------------------------------------------------------------------------
# Single-tenant / flat assumption extraction
# ---------------------------------------------------------------------------

def build_system_prompt(
    assumption_keys: list[str],
    assumptions_meta: dict | None = None,
    extraction_notes: str | None = None,
) -> str:
    if assumptions_meta:
        labelled: dict = {}
        for k in assumption_keys:
            if k.startswith("_"):
                labelled[k] = k
            else:
                meta = assumptions_meta.get(k, {})
                label = meta.get("label", k)
                cell_type = meta.get("type", "string")
                labelled[k] = f"{label} ({cell_type})"
        keys_json = json.dumps(labelled, indent=2)
    else:
        keys_json = json.dumps(assumption_keys, indent=2)

    notes_block = f"\n\n{extraction_notes}" if extraction_notes else ""

    return f"""You are a precise data-extraction assistant for a real estate investment firm.

Your task is to extract deal assumptions from an unstructured deal summary provided by the user.{notes_block}

You MUST return a valid JSON object with EXACTLY these keys:
{keys_json}

For each key, return an object with three fields:
  - "value"      : the extracted value (number, string, or null if not found)
  - "confidence" : "high" if clearly stated in the summary, "low" if inferred, estimated, or absent
  - "note"       : a short explanation ONLY when confidence is "low" (empty string otherwise)

Value formatting rules:
  - currency / dollar amounts  → plain number (no $ sign, no commas), e.g. 5200000
  - percentages               → decimal form, e.g. 0.065 for 6.5%
  - integers (months, units)  → plain integer, e.g. 24
  - dates                     → ISO format string, e.g. "2026-01-01"
  - strings                   → plain text
  - not found                 → null

Return ONLY the JSON object. No markdown, no explanation, no extra text."""


def extract_assumptions(
    deal_summary: str,
    manifest: dict[str, Any],
    client: anthropic.Anthropic | None = None,
) -> ExtractionResult:
    """
    Call Claude to extract flat assumptions defined in the manifest.

    For multi-tenant templates also calls extract_tenants() automatically
    and stores the result under the special key "_tenants" in the returned dict.

    Args:
        deal_summary: Raw text pasted by the user.
        manifest:     Parsed JSON manifest dict (contains "assumptions" key).
        client:       Optional pre-constructed Anthropic client.

    Returns:
        ExtractionResult dict keyed by assumption names.
        For multi-tenant templates, also contains "_tenants" key with TenantList.
    """
    if client is None:
        api_key = os.environ.get("ANTHROPIC_API_KEY")
        if not api_key:
            raise EnvironmentError(
                "ANTHROPIC_API_KEY is not set. Add it to your .env file."
            )
        client = anthropic.Anthropic(api_key=api_key)

    assumption_keys = list(manifest.get("assumptions", {}).keys())
    if not assumption_keys:
        raise ValueError("Manifest contains no assumptions to extract.")

    # Inject a virtual key so Claude tells us the lease term explicitly.
    # Used below to null out rent years beyond the term. Never written to Excel
    # because it starts with "_" and is not in the manifest's assumptions dict.
    has_rent_years = any(k.startswith("rent_year_") for k in assumption_keys)
    if has_rent_years:
        assumption_keys = ["_lease_term_years"] + assumption_keys

    system_prompt = build_system_prompt(
        assumption_keys,
        manifest.get("assumptions", {}),
        manifest.get("_extraction_notes"),
    )

    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=4096,
        temperature=0,
        system=system_prompt,
        messages=[{"role": "user", "content": deal_summary}],
    )

    raw_text = message.content[0].text.strip()

    if raw_text.startswith("```"):
        lines = raw_text.splitlines()
        raw_text = "\n".join(
            line for line in lines if not line.startswith("```")
        ).strip()

    try:
        result: ExtractionResult = json.loads(raw_text)
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"Claude returned non-JSON response:\n{raw_text}"
        ) from exc

    missing = [k for k in assumption_keys if k not in result]
    if missing:
        raise ValueError(
            f"Claude response is missing expected keys: {missing}\n"
            f"Raw response:\n{raw_text}"
        )

    for key in assumption_keys:
        entry = result[key]
        if not isinstance(entry, dict):
            result[key] = {"value": entry, "confidence": "low", "note": "auto-wrapped"}
        entry = result[key]
        entry.setdefault("confidence", "low")
        entry.setdefault("note", "")

    # Enforce lease term: null out rent_year_N beyond the stated term
    if has_rent_years:
        try:
            lease_term = int(result["_lease_term_years"]["value"])
            for n in range(lease_term + 1, 21):
                key = f"rent_year_{n}"
                if key in result:
                    result[key] = {"value": None, "confidence": "high", "note": "beyond lease term"}
        except (KeyError, TypeError, ValueError):
            pass  # no lease term found — leave as-is

    # If this is a multi-tenant template, also extract the rent roll
    if "rent_roll" in manifest:
        result["_tenants"] = extract_tenants(deal_summary, client)

        # Safeguard: if any tenant has a ti_psf, null out additional_cost_amount
        # to prevent Claude from double-counting TI (PSF goes to rent roll only)
        tenants = result.get("_tenants", [])
        if any(t.get("ti_psf") for t in tenants):
            if "additional_cost_amount" in result:
                result["additional_cost_amount"] = {
                    "value": None,
                    "confidence": "high",
                    "note": "TI was given as PSF — amount goes to rent roll only, not Deal Summary T20",
                }
            if "additional_cost_amount_label" in result:
                result["additional_cost_amount_label"] = {
                    "value": None,
                    "confidence": "high",
                    "note": "cleared — TI is PSF-based",
                }

    return result


# ---------------------------------------------------------------------------
# Multi-tenant rent roll extraction
# ---------------------------------------------------------------------------

_TENANT_SYSTEM_PROMPT = """You are a precise data-extraction assistant for a real estate investment firm.

Your task is to extract the TENANT RENT ROLL from an unstructured deal summary.

Return a JSON array where each element represents one tenant with these fields:
  - "tenant_name"        : string — tenant / business name
  - "unit"               : string — suite, unit label, or position (e.g. "A Endcap", "B", "C Endcap")
  - "sq_ft"              : integer — tenant square footage
  - "lease_term_years"   : integer — lease term in years (null if not stated)
  - "base_rent_psf"      : number — base rent per square foot per year
  - "ti_psf"             : number — tenant improvement allowance per square foot (0 if not stated)
  - "annual_pct_increase": number — annual rent escalation as decimal (e.g. 0.02 for 2%). 0 if flat.
  - "tax_reimbursement"  : "y" or "n" — does tenant reimburse taxes? Default "y" if NNN, "n" if gross
  - "ins_reimbursement"  : "y" or "n" — does tenant reimburse insurance?
  - "cam_reimbursement"  : "y" or "n" — does tenant reimburse CAM?
  - "confidence"         : "high" if all key fields clearly stated, "low" if any inferred
  - "note"               : short explanation if confidence is "low", else ""

If no tenant information is found, return an empty array [].
Return ONLY the JSON array. No markdown, no explanation."""


def extract_tenants(
    deal_summary: str,
    client: anthropic.Anthropic,
) -> TenantList:
    """
    Extract the tenant rent roll from a multi-tenant deal summary.

    Returns a list of tenant dicts. Empty list if no tenants found.
    """
    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=2048,
        temperature=0,
        system=_TENANT_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": deal_summary}],
    )

    raw_text = message.content[0].text.strip()

    if raw_text.startswith("```"):
        lines = raw_text.splitlines()
        raw_text = "\n".join(
            line for line in lines if not line.startswith("```")
        ).strip()

    try:
        tenants = json.loads(raw_text)
    except json.JSONDecodeError:
        return []

    if not isinstance(tenants, list):
        return []

    # Normalise each tenant entry
    for t in tenants:
        t.setdefault("tenant_name", "")
        t.setdefault("unit", "")
        t.setdefault("sq_ft", None)
        t.setdefault("lease_term_years", None)
        t.setdefault("base_rent_psf", None)
        t.setdefault("ti_psf", 0)
        t.setdefault("annual_pct_increase", 0)
        t.setdefault("tax_reimbursement", "y")
        t.setdefault("ins_reimbursement", "y")
        t.setdefault("cam_reimbursement", "y")
        t.setdefault("confidence", "low")
        t.setdefault("note", "")

    return tenants


# ---------------------------------------------------------------------------
# Display helpers
# ---------------------------------------------------------------------------

def summarise_extraction(result: ExtractionResult) -> str:
    """Return a human-readable summary (for CLI testing)."""
    lines = [f"{'Assumption':<35} {'Value':<22} {'Confidence':<12} Note"]
    lines.append("-" * 85)
    for key, data in result.items():
        if key == "_tenants":
            continue
        val = str(data.get("value", "null"))[:21]
        conf = data.get("confidence", "?")
        note = data.get("note", "")
        lines.append(f"{key:<35} {val:<22} {conf:<12} {note}")

    tenants = result.get("_tenants")
    if tenants:
        lines.append(f"\nTenants ({len(tenants)} extracted):")
        for t in tenants:
            lines.append(
                f"  {t.get('tenant_name','?'):<20} unit={t.get('unit','?'):<12} "
                f"sf={t.get('sq_ft','?'):<6} psf={t.get('base_rent_psf','?'):<6} "
                f"conf={t.get('confidence','?')}"
            )
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Smoke-test
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    _sample_manifest_single = {
        "assumptions": {
            "acquisition_cost":     {"sheet": "Deal Summary", "cell": "T12", "type": "currency"},
            "financing_rate":       {"sheet": "Deal Summary", "cell": "T30", "type": "percent"},
            "ltc":                  {"sheet": "Deal Summary", "cell": "T33", "type": "percent"},
            "lender_fee_rate":      {"sheet": "Deal Summary", "cell": "T34", "type": "percent"},
            "hold_period_months":   {"sheet": "Deal Summary", "cell": "T35", "type": "integer"},
        }
    }

    _sample_summary_single = """
    We are under contract on a ground lease NNN retail building in Phoenix, AZ.
    5th Third Bank occupies 100% of the 3,200 SF building. Purchase price $2.75M.
    We plan to use 80% LTC bridge financing at 6.5% with a 1% origination fee.
    Anticipated hold is 24 months before a 1031 exchange sale.
    Annual rent is $227,000 with 3 months of free rent at start.
    """

    print("=== Single Tenant Extraction ===")
    extraction = extract_assumptions(_sample_summary_single, _sample_manifest_single)
    print(summarise_extraction(extraction))

    _sample_manifest_multi = {
        "assumptions": {
            "acquisition_cost": {"sheet": "Deal Summary", "cell": "T11", "type": "currency"},
            "building_sf":      {"sheet": "Deal Summary", "cell": "S9",  "type": "integer"},
            "sale_cap_rate_goal": {"sheet": "Deal Summary", "cell": "T30", "type": "percent"},
        },
        "rent_roll": {"sheet": "Rent Roll", "first_tenant_row": 11}
    }

    _sample_summary_multi = """
    Knox Abbot 3 Tenant Strip Center — Columbia, SC. Build-to-suit, 5,500 SF.
    Land cost $1.2M. Three tenants: Swig (A Endcap, 1,500 SF, $68 PSF NNN),
    Sleep Number (B, 2,500 SF, $48 PSF NNN), Verizon (C Endcap, 1,500 SF, $68 PSF NNN).
    All tenants reimburse taxes, insurance, and CAM. TI allowance $50 PSF.
    Financing at 6.25%, 80% LTC, 1% fee, 8 months construction, 12 IO months.
    Target sale cap rate 6.5%.
    """

    print("\n=== Multi Tenant Extraction ===")
    extraction_multi = extract_assumptions(_sample_summary_multi, _sample_manifest_multi)
    print(summarise_extraction(extraction_multi))
