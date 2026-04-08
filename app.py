import os
import uuid
import copy
import json
import threading
import traceback
import re
from pathlib import Path

from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd
from pptx import Presentation
from pptx.oxml.ns import qn
import anthropic

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------

app = Flask(__name__)

BASE_DIR = Path(__file__).parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
OUTPUT_FOLDER = BASE_DIR / "outputs"
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

# In-memory job registry  {job_id: {"status": ..., "message": ..., "progress": ..., "output": ...}}
jobs: dict = {}
jobs_lock = threading.Lock()


# ---------------------------------------------------------------------------
# PPTX helpers
# ---------------------------------------------------------------------------

def clone_slide(prs: Presentation, source_index: int = 0):
    """Clone the slide at source_index and append a copy to the presentation."""
    source = prs.slides[source_index]
    new_slide = prs.slides.add_slide(source.slide_layout)

    # Clear auto-generated placeholder shapes from the new slide
    sp_tree = new_slide.shapes._spTree
    for child in list(sp_tree):
        sp_tree.remove(child)

    # Deep-copy every shape from the source slide
    for child in source.shapes._spTree:
        sp_tree.append(copy.deepcopy(child))

    # Copy any image / media relationships that live directly on the slide
    for rel in source.part.rels.values():
        if "image" in rel.reltype:
            try:
                new_slide.part.relate_to(rel.target_part, rel.reltype)
            except Exception:
                pass

    return new_slide


def replace_text_in_slide(slide, replacements: dict, ordered: dict = None):
    """
    Replace placeholder tokens in every text frame on a slide.

    replacements  – {old: new} for unique tokens
    ordered       – {old: [val1, val2, val3]} for tokens that appear
                    multiple times; replaced in document order (top→bottom)
    """
    order_counts = {k: 0 for k in (ordered or {})}

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            full_text = "".join(run.text for run in para.runs)
            if not full_text.strip():
                continue

            modified = full_text
            changed = False

            # Unique replacements
            for placeholder, value in replacements.items():
                if placeholder in modified:
                    modified = modified.replace(
                        placeholder, str(value) if value is not None else ""
                    )
                    changed = True

            # Ordered replacements (same token appears N times)
            for placeholder, values in (ordered or {}).items():
                if placeholder in modified:
                    idx = order_counts[placeholder]
                    if idx < len(values):
                        modified = modified.replace(
                            placeholder,
                            str(values[idx]) if values[idx] is not None else "",
                        )
                        order_counts[placeholder] += 1
                        changed = True

            if changed and para.runs:
                para.runs[0].text = modified
                for run in para.runs[1:]:
                    run.text = ""


# ---------------------------------------------------------------------------
# AI content generation
# ---------------------------------------------------------------------------

def generate_site_content(site: dict, client: anthropic.Anthropic) -> dict:
    """
    Call Claude to generate tagline, descriptions, and landmarks for one site.
    Returns a dict with keys: tagline, location_desc, visibility_desc,
    audience_desc, landmark_1, landmark_2, landmark_3.
    """
    site_name = site.get("Site Name", "")
    location = site.get("Location", "")
    market = site.get("Market", "")
    fmt = site.get("Format", "")
    size = site.get("Size", "")
    is_mobile = str(location).strip().lower() == "various"

    if is_mobile:
        landmark_instruction = (
            "Since this is a bus/transit route covering various locations across the city, "
            "provide 3 short lines about city-wide coverage, reach, and route highlights "
            "instead of specific nearby landmarks."
        )
        landmark_format = (
            '"landmark_1": "City-wide coverage line 1",\n'
            '  "landmark_2": "City-wide coverage line 2",\n'
            '  "landmark_3": "City-wide coverage line 3"'
        )
    else:
        landmark_instruction = (
            "Provide 3 real nearby landmarks for this address with approximate walking distances. "
            "Use your knowledge of the city to name specific, recognisable places."
        )
        landmark_format = (
            '"landmark_1": "Landmark Name — X min walk",\n'
            '  "landmark_2": "Landmark Name — X min walk",\n'
            '  "landmark_3": "Landmark Name — X min walk"'
        )

    prompt = f"""You are writing punchy, professional copy for an OOH (Out-of-Home) advertising proposal.

Site details:
- Name: {site_name}
- Location / Address: {location}
- City / Market: {market}
- Format: {fmt}
- Size: {size}

Return ONLY valid JSON (no markdown fences, no extra text) with exactly these keys:

{{
  "tagline": "<4–7 word punchy advertising tagline for this site>",
  "location_desc": "<2–3 sentences describing where the site is and what surrounds it>",
  "visibility_desc": "<2–3 sentences about viewing angles, physical size, and sightlines>",
  "audience_desc": "<2–3 sentences about who passes by and approximate daily volume>",
  {landmark_format}
}}

{landmark_instruction}"""

    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=800,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = response.content[0].text.strip()
    # Strip markdown fences if the model adds them anyway
    raw = re.sub(r"^```[a-z]*\n?", "", raw)
    raw = re.sub(r"\n?```$", "", raw)
    return json.loads(raw)


# ---------------------------------------------------------------------------
# Core processing function (runs in background thread)
# ---------------------------------------------------------------------------

def process_job(job_id: str, excel_path: Path, pptx_path: Path):
    def update(status: str, message: str, progress: int = 0):
        with jobs_lock:
            jobs[job_id]["status"] = status
            jobs[job_id]["message"] = message
            jobs[job_id]["progress"] = progress

    try:
        # ── 1. Read Excel ───────────────────────────────────────────────────
        update("processing", "Reading Excel file…", 5)
        df = pd.read_excel(excel_path, engine="openpyxl")

        # Normalise column names (strip whitespace)
        df.columns = [c.strip() for c in df.columns]

        # Forward-fill the Market column to handle merged cells
        if "Market" in df.columns:
            df["Market"] = df["Market"].ffill()

        # Drop rows where Site Name is empty
        if "Site Name" not in df.columns:
            raise ValueError("Excel file must have a 'Site Name' column.")
        df = df[df["Site Name"].notna() & (df["Site Name"].astype(str).str.strip() != "")]
        df = df.reset_index(drop=True)

        if df.empty:
            raise ValueError("No valid site rows found in the Excel file.")

        total_sites = len(df)
        update("processing", f"Found {total_sites} site(s). Loading template…", 10)

        # ── 2. Load template PPTX ──────────────────────────────────────────
        prs = Presentation(str(pptx_path))
        if not prs.slides:
            raise ValueError("The PowerPoint template has no slides.")

        # Snapshot the original template slide XML BEFORE any modifications.
        # clone_slide always clones from this frozen copy so every slide
        # starts from the blank template — not from a previously-filled slide.
        template_spTree = copy.deepcopy(prs.slides[0].shapes._spTree)
        template_layout = prs.slides[0].slide_layout

        # ── 3. Set up Anthropic client ─────────────────────────────────────
        api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if not api_key:
            raise ValueError(
                "ANTHROPIC_API_KEY environment variable is not set. "
                "Add it to your environment before running the app."
            )
        client = anthropic.Anthropic(api_key=api_key)

        # ── 4. Process each row ────────────────────────────────────────────
        for idx, row in df.iterrows():
            pct = 10 + int((idx / total_sites) * 80)
            site_name = str(row.get("Site Name", "")).strip()
            update("processing", f"Generating slide {idx + 1}/{total_sites}: {site_name}…", pct)

            # --- Build frequency string ---
            spot_dur = str(row.get("Spot Duration", "")).strip()
            sov_loop = str(row.get("SOV/Loop", "")).strip()
            if spot_dur.lower() in ("", "nan", "n/a", "na"):
                frequency = sov_loop
            else:
                frequency = f"{spot_dur} {sov_loop}".strip()

            # --- Format traffic with commas ---
            raw_impacts = row.get("Impacts", "")
            try:
                traffic = f"{int(float(str(raw_impacts).replace(',', ''))):,}"
            except (ValueError, TypeError):
                traffic = str(raw_impacts).strip()

            # --- AI content ---
            site_dict = row.to_dict()
            ai = generate_site_content(site_dict, client)

            # --- Values shorthand ---
            size     = str(row.get("Size", "")).strip()
            location = str(row.get("Location", "")).strip()
            units    = str(row.get("Units/Faces", "")).strip()
            fmt      = str(row.get("Format", "")).strip()
            market   = str(row.get("Market", "")).strip()

            # --- Build replacement map ---
            # Matches the exact placeholder strings in the "xyz format" template.
            # Also supports {TOKEN} style for custom templates.
            replacements = {
                # ── Title ──────────────────────────────────────────────────
                "Site Name":        site_name,
                "Headline":         ai.get("tagline", ""),
                # ── Additional Information boxes ───────────────────────────
                "Size: xyz":        f"Size: {size}",
                "Format: xyz":      f"Format: {fmt}",
                "Location: xyz":    f"Location: {location}",
                "Frequency: xyz":   f"Frequency: {frequency}",
                "Units: xyz":       f"Units: {units}",
                "Traffic: xyz":     f"Traffic: {traffic}",
                # ── {TOKEN} style (for custom templates) ──────────────────
                "{SITE_NAME}":      site_name,
                "{TAGLINE}":        ai.get("tagline", ""),
                "{LOCATION_DESC}":  ai.get("location_desc", ""),
                "{VISIBILITY_DESC}": ai.get("visibility_desc", ""),
                "{AUDIENCE_DESC}":  ai.get("audience_desc", ""),
                "{SIZE}":           size,
                "{LOCATION}":       location,
                "{UNITS}":          units,
                "{FORMAT}":         fmt,
                "{FREQUENCY}":      frequency,
                "{TRAFFIC}":        traffic,
                "{LANDMARK_1}":     ai.get("landmark_1", ""),
                "{LANDMARK_2}":     ai.get("landmark_2", ""),
                "{LANDMARK_3}":     ai.get("landmark_3", ""),
                "{MARKET}":         market,
            }

            # --- Ordered replacements (same token appears multiple times) ---
            # "Text"        → Location desc, Visibility desc, Audience desc (top→bottom)
            # "Xyz \u20130.5km" → Landmark 1, 2, 3  (en-dash variant)
            ordered = {
                "Text": [
                    ai.get("location_desc", ""),
                    ai.get("visibility_desc", ""),
                    ai.get("audience_desc", ""),
                ],
                # en-dash "–" (U+2013) as used in the template
                "Xyz \u20130.5km": [
                    ai.get("landmark_1", ""),
                    ai.get("landmark_2", ""),
                    ai.get("landmark_3", ""),
                ],
                # hyphen fallback just in case
                "Xyz -0.5km": [
                    ai.get("landmark_1", ""),
                    ai.get("landmark_2", ""),
                    ai.get("landmark_3", ""),
                ],
            }

            # --- Select / create the slide ---
            if idx == 0:
                slide = prs.slides[0]
            else:
                # Add a fresh slide and stamp the ORIGINAL template XML onto it
                slide = prs.slides.add_slide(template_layout)
                new_tree = slide.shapes._spTree
                for child in list(new_tree):
                    new_tree.remove(child)
                for child in template_spTree:
                    new_tree.append(copy.deepcopy(child))

            replace_text_in_slide(slide, replacements, ordered)

        # ── 5. Save output ─────────────────────────────────────────────────
        update("processing", "Saving output file…", 95)
        output_filename = f"OOH_Proposal_{job_id[:8]}.pptx"
        output_path = OUTPUT_FOLDER / output_filename
        prs.save(str(output_path))

        with jobs_lock:
            jobs[job_id]["status"] = "done"
            jobs[job_id]["message"] = f"Done! {total_sites} slide(s) generated."
            jobs[job_id]["progress"] = 100
            jobs[job_id]["output"] = output_filename

    except Exception as exc:
        with jobs_lock:
            jobs[job_id]["status"] = "error"
            jobs[job_id]["message"] = f"Error: {exc}"
            jobs[job_id]["progress"] = 0
        print(traceback.format_exc())

    finally:
        # Clean up uploaded files
        try:
            excel_path.unlink(missing_ok=True)
            pptx_path.unlink(missing_ok=True)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/generate", methods=["POST"])
def generate():
    """Accept the two uploaded files and kick off background processing."""
    if "excel" not in request.files or "template" not in request.files:
        return jsonify({"error": "Both 'excel' and 'template' files are required."}), 400

    excel_file = request.files["excel"]
    template_file = request.files["template"]

    if not excel_file.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Excel file must be .xlsx or .xls"}), 400
    if not template_file.filename.endswith(".pptx"):
        return jsonify({"error": "Template file must be .pptx"}), 400

    job_id = uuid.uuid4().hex

    excel_path = UPLOAD_FOLDER / f"{job_id}_data.xlsx"
    pptx_path = UPLOAD_FOLDER / f"{job_id}_template.pptx"
    excel_file.save(str(excel_path))
    template_file.save(str(pptx_path))

    with jobs_lock:
        jobs[job_id] = {
            "status": "queued",
            "message": "Queued…",
            "progress": 0,
            "output": None,
        }

    thread = threading.Thread(
        target=process_job, args=(job_id, excel_path, pptx_path), daemon=True
    )
    thread.start()

    return jsonify({"job_id": job_id})


@app.route("/api/status/<job_id>")
def status(job_id: str):
    with jobs_lock:
        job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify(job)


@app.route("/api/download/<job_id>")
def download(job_id: str):
    with jobs_lock:
        job = jobs.get(job_id)
    if not job or job["status"] != "done":
        return jsonify({"error": "File not ready"}), 404

    output_filename = job["output"]
    output_path = OUTPUT_FOLDER / output_filename

    if not output_path.exists():
        return jsonify({"error": "Output file missing"}), 404

    def cleanup():
        try:
            output_path.unlink(missing_ok=True)
        except Exception:
            pass
        with jobs_lock:
            jobs.pop(job_id, None)

    response = send_file(
        str(output_path),
        as_attachment=True,
        download_name="OOH_Proposal.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
    # Schedule cleanup after response is sent
    @response.call_on_close
    def _cleanup():
        cleanup()

    return response


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
