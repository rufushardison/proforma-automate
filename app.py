"""
app.py

Streamlit UI for the Proforma Automation tool.

Run:  streamlit run app.py
"""

from __future__ import annotations

import datetime
import json
import os
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from dotenv import load_dotenv

from excel_writer import fill_template
from extractor import extract_assumptions

load_dotenv()

# Inject Streamlit Cloud secrets into os.environ so extractor can find them
if "ANTHROPIC_API_KEY" not in os.environ:
    try:
        os.environ["ANTHROPIC_API_KEY"] = st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        pass

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent
MANIFESTS_DIR = BASE_DIR / "manifests"
TEMPLATES_DIR = BASE_DIR / "templates"
HISTORY_DIR = BASE_DIR / "history"

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Cason Development Proforma Generator",
    page_icon=":office_building:",
    layout="wide",
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def load_manifest(manifest_path: str) -> dict:
    with open(manifest_path) as f:
        return json.load(f)


def get_available_templates() -> dict[str, dict]:
    templates = {}
    for manifest_file in sorted(MANIFESTS_DIR.glob("*.json")):
        manifest = load_manifest(str(manifest_file))
        name = manifest.get("template_name", manifest_file.stem)
        tpl_path = TEMPLATES_DIR / manifest.get("template_file", "")
        templates[name] = {
            "manifest": manifest,
            "template_path": str(tpl_path),
        }
    return templates


def get_template_defaults(template_path: str, manifest: dict) -> dict[str, Any]:
    """Read the current value of every mapped cell from the template workbook."""
    defaults = {}
    try:
        wb = openpyxl.load_workbook(template_path, keep_vba=True, data_only=True)
    except Exception:
        return defaults
    for key, cell_meta in manifest.get("assumptions", {}).items():
        sheet = cell_meta.get("sheet", "")
        cell = cell_meta.get("cell", "")
        if sheet in wb.sheetnames and cell:
            val = wb[sheet][cell].value
            # Format for display
            if hasattr(val, "strftime"):
                val = val.strftime("%Y-%m-%d")
            defaults[key] = val
    return defaults


def build_results_dataframe(
    extraction: dict,
    manifest: dict,
    defaults: dict[str, Any],
) -> pd.DataFrame:
    """
    Build the editable results table.

    Columns:
      _key             — internal, hidden in data_editor
      Label            — human-readable assumption name
      Template Default — current value baked into the template (read-only)
      Value            — Claude's extracted value (EDITABLE)
      Confidence       — high / low (read-only)
      Note             — Claude's reasoning note (read-only)
    """
    rows = []
    assumptions_meta = manifest.get("assumptions", {})
    for key, data in extraction.items():
        if key.startswith("_"):
            continue
        meta = assumptions_meta.get(key, {})
        label = meta.get("label", key)
        default_val = defaults.get(key)

        # Format default for display
        if default_val is not None:
            if isinstance(default_val, float) and default_val == int(default_val):
                display_default = str(int(default_val))
            elif isinstance(default_val, float):
                display_default = str(default_val)
            else:
                display_default = str(default_val)
        else:
            display_default = ""

        extracted_val = data.get("value")
        # Store as string for the editable cell; empty string means null
        if extracted_val is None:
            display_val = ""
        else:
            display_val = str(extracted_val)

        rows.append({
            "_key":              key,
            "Label":             label,
            "Template Default":  display_default,
            "Value":             display_val,
            "Confidence":        data.get("confidence", "low"),
            "Note":              data.get("note", ""),
        })
    return pd.DataFrame(rows)


def df_to_extraction(edited_df: pd.DataFrame, original_extraction: dict) -> dict:
    """
    Merge the user's table edits back into the extraction dict so
    fill_template receives the corrected values.
    """
    updated = dict(original_extraction)
    for _, row in edited_df.iterrows():
        key = row["_key"]
        if key not in updated:
            continue
        raw = str(row["Value"]).strip()
        # Treat empty string as null (no value)
        if raw == "" or raw.lower() == "none":
            value = None
        else:
            # Try to coerce to number if it looks like one
            try:
                value = int(raw) if "." not in raw else float(raw)
            except ValueError:
                value = raw  # keep as string

        updated[key] = {
            "value":      value,
            "confidence": row["Confidence"],
            "note":       row["Note"],
        }
    return updated


def _row_style(row: pd.Series) -> list[str]:
    if row["Value"] == "" or pd.isna(row["Value"]):
        return ["background-color: #ffe0b2"] * len(row)   # orange — not found
    if row["Confidence"] == "low":
        return ["background-color: #fff3cd"] * len(row)   # yellow — low confidence
    return [""] * len(row)


def save_to_history(deal_summary: str, template_name: str, extraction: dict) -> None:
    try:
        HISTORY_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.datetime.now(datetime.timezone.utc)
        filename = ts.strftime("%Y%m%dT%H%M%S%f") + ".json"
        entry = {
            "timestamp": ts.isoformat(),
            "template_name": template_name,
            "deal_summary": deal_summary,
            "extraction": extraction,
        }
        (HISTORY_DIR / filename).write_text(
            json.dumps(entry, indent=2, default=str), encoding="utf-8"
        )
    except Exception:
        pass  # never crash the main workflow


def load_history_entries(limit: int = 25) -> list[dict]:
    if not HISTORY_DIR.exists():
        return []
    entries = []
    for path in sorted(HISTORY_DIR.glob("*.json"), reverse=True):
        if len(entries) >= limit:
            break
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if all(k in data for k in ("timestamp", "template_name", "deal_summary", "extraction")):
                entries.append(data)
        except Exception:
            continue
    return entries


# ---------------------------------------------------------------------------
# Dictation component (Web Speech API — works in Chrome/Edge)
# ---------------------------------------------------------------------------

_DICTATION_HTML = """
<html><head><style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: sans-serif; padding: 2px 0; }
  #btn {
    background: #ff4b4b; color: white; border: none;
    padding: 5px 16px; border-radius: 4px; cursor: pointer;
    font-size: 13px; font-weight: 600;
  }
  #btn.on { background: #1a7a4a; }
  #status { font-size: 12px; color: #666; margin-left: 10px; }
</style></head>
<body>
<button id="btn" onclick="toggle()">🎤 Dictate</button>
<span id="status"></span>
<script>
  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  let rec, active = false, accumulated = '';

  function toggle() {
    if (!SR) {
      document.getElementById('status').textContent = 'Not supported — use Chrome or Edge';
      return;
    }
    active ? stop() : start();
  }

  function start() {
    rec = new SR();
    rec.continuous = true;
    rec.interimResults = true;
    rec.lang = 'en-US';
    active = true;
    document.getElementById('btn').textContent = '⏹ Stop Dictating';
    document.getElementById('btn').classList.add('on');
    document.getElementById('status').textContent = 'Listening…';

    rec.onresult = (e) => {
      let interim = '';
      for (let i = e.resultIndex; i < e.results.length; i++) {
        if (e.results[i].isFinal) accumulated += e.results[i][0].transcript + ' ';
        else interim = e.results[i][0].transcript;
      }
      document.getElementById('status').textContent = interim || 'Listening…';
      Streamlit.setComponentValue(accumulated.trim());
    };

    rec.onerror = (e) => {
      document.getElementById('status').textContent = 'Error: ' + e.error;
      stop();
    };

    // Auto-restart keeps mic alive for long dictations
    rec.onend = () => { if (active) rec.start(); };
    rec.start();
  }

  function stop() {
    active = false;
    if (rec) rec.stop();
    document.getElementById('btn').textContent = '🎤 Dictate';
    document.getElementById('btn').classList.remove('on');
    document.getElementById('status').textContent = '';
  }

  Streamlit.setFrameHeight(36);
</script>
</body></html>
"""


# ---------------------------------------------------------------------------
# Main UI
# ---------------------------------------------------------------------------

def main():
    st.title("Proforma Automation")
    st.caption(
        "Paste a deal summary, select a template, review and edit the extracted "
        "assumptions, then download the filled Excel proforma."
    )

    available = get_available_templates()
    if not available:
        st.error("No template manifests found in `manifests/`.")
        st.stop()

    # ------------------------------------------------------------------
    # Sidebar — extraction history
    # ------------------------------------------------------------------
    with st.sidebar:
        st.header("History")
        history_entries = load_history_entries(limit=25)
        if not history_entries:
            st.caption("No extractions yet.")
        else:
            for idx, entry in enumerate(history_entries):
                try:
                    ts_dt = datetime.datetime.fromisoformat(entry["timestamp"])
                    ts_display = ts_dt.strftime("%b %d, %Y  %H:%M")
                except Exception:
                    ts_display = entry.get("timestamp", "")[:16]
                tpl = entry.get("template_name", "?")
                preview = entry.get("deal_summary", "")[:80].replace("\n", " ")
                if len(entry.get("deal_summary", "")) > 80:
                    preview += "..."
                with st.expander(f"{tpl} — {ts_display}", expanded=False):
                    st.caption(preview)
                    col_load, col_dl = st.columns(2)
                    with col_load:
                        if st.button("Load", key=f"hist_load_{idx}"):
                            st.session_state["deal_summary_text"] = entry["deal_summary"]
                            st.session_state["extraction"]        = entry["extraction"]
                            st.session_state["template_name"]     = entry["template_name"]
                            st.session_state["selected_template"] = entry["template_name"]
                            st.session_state["_prev_dictated"]    = ""
                            st.rerun()
                    with col_dl:
                        if entry["template_name"] in available:
                            hist_sel = available[entry["template_name"]]
                            try:
                                hist_buf = fill_template(
                                    hist_sel["template_path"],
                                    hist_sel["manifest"],
                                    entry["extraction"],
                                )
                                fname = (
                                    f"proforma_{entry['template_name'].lower().replace(' ', '_')}"
                                    f"_{ts_display[:6].replace(' ', '')}.xlsm"
                                )
                                st.download_button(
                                    "Download", data=hist_buf, file_name=fname,
                                    mime="application/vnd.ms-excel.sheet.macroenabled.12",
                                    key=f"hist_dl_{idx}",
                                )
                            except Exception as e:
                                st.error(str(e))
                        else:
                            st.caption("Template unavailable.")

    # ------------------------------------------------------------------
    # Inputs
    # ------------------------------------------------------------------
    col_left, col_right = st.columns([2, 1])

    with col_left:
        # Dictation button — appends spoken text to the deal summary box
        dictated = components.html(_DICTATION_HTML, height=36)
        prev_dictated = st.session_state.get("_prev_dictated", "")
        if dictated and isinstance(dictated, str) and dictated != prev_dictated:
            # Component sends cumulative text; append only the new portion
            if dictated.startswith(prev_dictated):
                new_part = dictated[len(prev_dictated):].strip()
            else:
                new_part = dictated.strip()  # component was reset
            existing = st.session_state.get("deal_summary_text", "")
            st.session_state["deal_summary_text"] = (
                existing + (" " if existing else "") + new_part
            ).strip()
            st.session_state["_prev_dictated"] = dictated

        deal_summary = st.text_area(
            "Deal Summary",
            height=300,
            key="deal_summary_text",
            placeholder=(
                "Paste the deal memo, email, or investment summary here. "
                "Include purchase price, LTC, interest rate, hold period, etc."
            ),
        )

    with col_right:
        template_name = st.selectbox(
            "Select Template",
            options=list(available.keys()),
            key="selected_template",
        )
        st.markdown("---")
        st.markdown(
            "**Tips for best results:**\n"
            "- Include specific numbers (not ranges)\n"
            "- Mention LTC %, interest rate, and hold period\n"
            "- Edit any cell in the table before downloading"
        )

    submitted = st.button("Extract Assumptions", type="primary", use_container_width=True)

    # ------------------------------------------------------------------
    # Run Claude extraction (only when submit clicked)
    # ------------------------------------------------------------------
    if submitted:
        if not deal_summary.strip():
            st.warning("Please paste a deal summary before extracting.")
            return

        selected = available[template_name]
        manifest = selected["manifest"]
        template_path = selected["template_path"]

        if not Path(template_path).exists():
            st.error(f"Template file not found: `{template_path}`")
            return

        if not os.environ.get("ANTHROPIC_API_KEY"):
            st.error("ANTHROPIC_API_KEY is not set. Add it to your `.env` file and restart.")
            return

        with st.spinner("Claude is extracting assumptions..."):
            try:
                extraction = extract_assumptions(deal_summary, manifest)
            except Exception as exc:
                st.error(f"Extraction failed: {exc}")
                return

        # Persist to session state so table edits survive reruns
        st.session_state["extraction"]    = extraction
        st.session_state["template_name"] = template_name
        save_to_history(deal_summary, template_name, extraction)

    # ------------------------------------------------------------------
    # Show results table + download (whenever extraction is in state)
    # ------------------------------------------------------------------
    if "extraction" not in st.session_state:
        return

    # If user switched template, clear stale state
    if st.session_state.get("template_name") != template_name:
        del st.session_state["extraction"]
        del st.session_state["template_name"]
        return

    extraction   = st.session_state["extraction"]
    selected     = available[template_name]
    manifest     = selected["manifest"]
    template_path = selected["template_path"]

    defaults = get_template_defaults(template_path, manifest)
    df = build_results_dataframe(extraction, manifest, defaults)

    # Counts for the warning banner
    not_found_count      = (df["Value"] == "").sum()
    low_confidence_count = (df["Confidence"] == "low").sum()

    st.markdown("### Extracted Assumptions")

    if not_found_count > 0 or low_confidence_count > 0:
        msg_parts = []
        if low_confidence_count:
            msg_parts.append(f"**{low_confidence_count}** low-confidence (yellow in Excel)")
        if not_found_count:
            msg_parts.append(f"**{not_found_count}** not found — template default kept (orange in Excel)")
        st.warning("Review before downloading: " + " · ".join(msg_parts))

    st.markdown(
        "<div style='font-size:0.85em; margin-bottom:8px'>"
        "<span style='background:#FFFF00; padding:2px 8px; border-radius:3px; margin-right:6px'>Yellow</span>"
        "Extracted but low-confidence&nbsp;&nbsp;"
        "<span style='background:#FFC000; padding:2px 8px; border-radius:3px; margin-right:6px'>Orange</span>"
        "Not found — template default kept"
        "</div>",
        unsafe_allow_html=True,
    )

    # ------------------------------------------------------------------
    # Editable table
    # ------------------------------------------------------------------
    edited_df = st.data_editor(
        df.style.apply(_row_style, axis=1),
        column_config={
            "_key": None,  # hidden
            "Label": st.column_config.TextColumn(
                "Assumption", disabled=True, width="medium"
            ),
            "Template Default": st.column_config.TextColumn(
                "Template Default", disabled=True, width="small",
                help="Current value baked into the template"
            ),
            "Value": st.column_config.TextColumn(
                "Extracted Value", width="small",
                help="Edit this cell to override before downloading"
            ),
            "Confidence": st.column_config.TextColumn(
                "Confidence", disabled=True, width="small"
            ),
            "Note": st.column_config.TextColumn(
                "Note", disabled=True, width="large"
            ),
        },
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
    )

    st.caption("Edit the **Extracted Value** column directly — changes will be reflected in the downloaded file.")

    # ------------------------------------------------------------------
    # Build and download
    # ------------------------------------------------------------------
    st.markdown("---")
    col_dl, col_stats = st.columns([2, 1])

    with col_dl:
        with st.spinner("Building Excel file..."):
            try:
                updated_extraction = df_to_extraction(edited_df, extraction)
                excel_buffer = fill_template(template_path, manifest, updated_extraction)
            except Exception as exc:
                st.error(f"Failed to build Excel file: {exc}")
                return

        output_filename = f"proforma_{template_name.lower().replace(' ', '_')}.xlsm"
        st.download_button(
            label="Download Filled Proforma (.xlsm)",
            data=excel_buffer,
            file_name=output_filename,
            mime="application/vnd.ms-excel.sheet.macroenabled.12",
            type="primary",
            use_container_width=True,
        )

    with col_stats:
        high_count = ((df["Value"] != "") & (df["Confidence"] == "high")).sum()
        st.metric("High Confidence", int(high_count))
        st.metric("Low Confidence", int(low_confidence_count))
        st.metric("Not Found", int(not_found_count))


if __name__ == "__main__":
    main()
