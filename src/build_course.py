#!/usr/bin/env python3
"""
build_course.py — generic slide-deck-to-narrated-MP4 builder
=============================================================
Turns any PowerPoint slide deck + optional narration doc into voice-only
narrated MP4s (no avatar), one file per course section.

Inputs
------
  --pptx        PATH   slide deck (required)
  --script      PATH   Word doc with per-slide narration; two formats accepted:
                         • "Slide N:" or "Slide N -" lines  (ACAS style)
                         • "Slide N" as a heading with paragraph body (Pentesting style)
  --lessons     PATH   Word doc with lesson prose (headings + body paragraphs);
                       used for auto-matching narration when --script is absent.
  --courseinfo  PATH   Extra Word doc to extract a course blurb for the title slide.

Pauses
------
  Insert "<pause X sec>" anywhere in the narration text to get X seconds of
  silence at that point (X may be decimal). Accepts <pause 2>, <pause 2s>,
  <pause 2 sec>, <pause 2.5 seconds>. Works with every engine, and the marker
  is removed from burned-in subtitles.

Section detection — choose one (default: whole deck = one section)
------------------------------------------------------------------
  --divider-mark TEXT   Text snippet present on every section-divider slide
                        (e.g. the footer phrase used in MMT-Pentester).
  --modules-json PATH   JSON file with explicit slide ranges:
                        [{"name":"Intro","slug":"intro","start":1,"end":10}, ...]
                        (slide numbers are 1-based)

Course identity
---------------
  --course-name  TEXT   Defaults to pptx filename stem (underscores/dashes → spaces).
  --org          TEXT   Organisation phrase added to title slide, e.g.
                        "CyberSuite Academy by Montimage".
  --title-text   TEXT   Fully override the title-slide spoken text.

Slide-text filters (suppress footer / section-label lines from narration)
-------------------------------------------------------------------------
  --footer-re   REGEX  e.g. "^(MMT-Pentester|MONTIMAGE|\\d{1,3}$)"
  --label-re    REGEX  e.g. "^(PART\\s*\\d+|Module\\s*\\d+|CAPSTONE)"

Engines (--engine)
------------------
  edge        (default) Microsoft Edge neural voices; free, online.
              pip install edge-tts
  kokoro      local, CPU, Apache-2.0.
              pip install kokoro soundfile  (+ espeak-ng; needs Python ≤3.12)
  piper       local, tiny, MIT.
              pip install piper-tts
  elevenlabs  set env vars ELEVENLABS_API_KEY + ELEVENLABS_VOICE_ID.
  voicecast   voice-cloning via local VoiceCast install.
              --voice-sample <your-voice.wav>  --voicecast-dir <repo-path>
              Long narration is split into <= --voicecast-max-chars chunks and the
              model is loaded once for the whole course, so long slides are not
              truncated and the voice does not loop/repeat. Keep voicecast_synth.py
              next to this script.

Examples
--------
  # ACAS — explicit slide ranges, script doc, Edge voice
  python build_course.py \\
      --pptx ACAS_course_slides_v2.pptx \\
      --script ACAS_course_script_v2.docx \\
      --modules-json acas_modules.json \\
      --engine edge --edge-voice en-US-AndrewNeural \\
      --org "CyberSuite Academy by Montimage"

  # Pentesting — auto-detect sections from divider slides, lessons doc
  python build_course.py \\
      --pptx MMT-Pentester-Course-Slides_v3.pptx \\
      --lessons CyberSuite_MMT-Pentester_Lessons_v3.docx \\
      --divider-mark "uptake of innovative security-as-a-service" \\
      --footer-re "^(MMT-Pentester|MONTIMAGE|Montimage|\\\\d{1,3}$)" \\
      --label-re "^(PART\\\\s*\\\\d+|Module\\\\s*\\\\d+|CAPSTONE|Uptake of Innovative)" \\
      --engine edge --edge-voice en-US-AndrewNeural

  # Any course, voice-cloned with VoiceCast
  python build_course.py \\
      --pptx MyCourse.pptx --script MyCourse_script.docx \\
      --engine voicecast --voice-sample my-voice.wav

  # Dry run (no audio — just map narration to slides for review)
  python build_course.py --pptx MyCourse.pptx --script script.docx --dry-run

Dependencies (system): LibreOffice, Poppler (pdftoppm + pdfinfo), ffmpeg.
Python: python-docx, python-pptx  (+ the engine package, e.g. edge-tts).
"""

import argparse, json, os, re, sys, shutil, subprocess, asyncio, difflib
from pathlib import Path

try:
    from pptx import Presentation
except ModuleNotFoundError:
    sys.exit("ERROR: python-pptx not found. Run: pip install python-pptx")

W, H = 1920, 1080
MIN_HOLD = 3.0
TOOLS = {"soffice": None, "pdfinfo": None, "pdftoppm": None, "ffmpeg": None, "ffprobe": None}


# ------------------------------------------------------------------ tooling
def _find(names, extra_dirs):
    for n in names:
        p = shutil.which(n)
        if p:
            return p
    for d in extra_dirs:
        if not d:
            continue
        for n in names:
            cand = Path(d) / n
            if cand.exists():
                return str(cand)
    return None


def resolve_tools(args, require_ffmpeg):
    pf   = os.environ.get("ProgramFiles",       r"C:\Program Files")
    pf86 = os.environ.get("ProgramFiles(x86)",  r"C:\Program Files (x86)")
    lo  = [args.soffice and str(Path(args.soffice).parent),
           fr"{pf}\LibreOffice\program", fr"{pf86}\LibreOffice\program"]
    pop = [args.poppler_bin, fr"{pf}\poppler\Library\bin", fr"{pf}\Calibre2\app\bin",
           r"C:\poppler\Library\bin", r"C:\ProgramData\chocolatey\bin"]
    ff  = [args.ffmpeg_bin, r"C:\ffmpeg\bin", r"C:\ProgramData\chocolatey\bin"]
    TOOLS["soffice"]  = (args.soffice if (args.soffice and Path(args.soffice).exists())
                         else _find(["soffice.com", "soffice", "soffice.exe"], lo))
    TOOLS["pdfinfo"]  = _find(["pdfinfo",  "pdfinfo.exe"],  pop)
    TOOLS["pdftoppm"] = _find(["pdftoppm", "pdftoppm.exe"], pop)
    if require_ffmpeg:
        TOOLS["ffmpeg"]  = _find(["ffmpeg",  "ffmpeg.exe"],  ff)
        TOOLS["ffprobe"] = _find(["ffprobe", "ffprobe.exe"], ff)
    hints = {
        "soffice":  "LibreOffice — winget install TheDocumentFoundation.LibreOffice  (or --soffice <path>)",
        "pdfinfo":  "Poppler — choco install poppler  (or --poppler-bin <folder>)",
        "pdftoppm": "Poppler — choco install poppler  (or --poppler-bin <folder>)",
        "ffmpeg":   "ffmpeg  — winget install Gyan.FFmpeg  (or --ffmpeg-bin <folder>)",
        "ffprobe":  "ffmpeg  — winget install Gyan.FFmpeg  (or --ffmpeg-bin <folder>)",
    }
    need = ["soffice", "pdfinfo", "pdftoppm"] + (["ffmpeg", "ffprobe"] if require_ffmpeg else [])
    missing = [t for t in need if not TOOLS[t]]
    if missing:
        sys.exit("ERROR: required tool(s) not found:\n" +
                 "\n".join(f"  - {hints[t]}" for t in dict.fromkeys(missing)) +
                 "\nAfter installing, open a NEW terminal so PATH updates, then re-run.")
    print("[tools] " + " | ".join(f"{k}={Path(v).name}" for k, v in TOOLS.items() if v))


def run(cmd, **kw):
    p = subprocess.run(cmd, capture_output=True, text=True, **kw)
    if p.returncode != 0:
        raise RuntimeError(f"Command failed: {' '.join(map(str, cmd))}\n{p.stderr[-800:]}")
    return p.stdout


# ------------------------------------------------------------------ slides
def render_slides(pptx, workdir):
    pdf = workdir / (pptx.stem + ".pdf")
    print(f"[slides] {pptx.name} -> PDF ...")
    run([TOOLS["soffice"], "--headless", "--convert-to", "pdf", "--outdir", str(workdir), str(pptx)])
    if not pdf.exists():
        raise RuntimeError("LibreOffice produced no PDF (close any open LibreOffice window and retry).")
    n = int(re.search(r"Pages:\s+(\d+)", run([TOOLS["pdfinfo"], str(pdf)])).group(1))
    print(f"[slides] {n} pages -> PNG ...")
    sdir = workdir / "slides"
    sdir.mkdir(exist_ok=True)
    run([TOOLS["pdftoppm"], "-png", "-r", "150", str(pdf), str(sdir / "slide")])
    return sorted(sdir.glob("slide-*.png"))


# ---------------------------------------------------------- slide text helpers
def slide_lines(slide):
    out = []
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            for ln in sh.text_frame.text.split("\n"):
                t = ln.strip()
                if t:
                    out.append(t)
    return out


def slide_body_text(slide, footer_re, divider_mark):
    lines = slide_lines(slide)
    keep = []
    for l in lines:
        if footer_re.search(l):
            continue
        if divider_mark and divider_mark in l.lower():
            continue
        keep.append(l)
    return " ".join(keep)


def slide_title_text(slide, footer_re, label_re, divider_mark):
    for l in slide_lines(slide):
        if footer_re.search(l):
            continue
        if label_re.search(l):
            continue
        if divider_mark and divider_mark in l.lower():
            continue
        return l
    return (slide_lines(slide) or [""])[0]


def clean_for_speech(t):
    t = (t.replace("→", " then ").replace("×", " by ").replace("◄", " ").replace("►", " ")
           .replace("·", ", ").replace("•", " ").replace("⚠", " ").replace("&", " and ")
           .replace("‘", "'").replace("’", "'")
           .replace("“", '"').replace("”", '"').replace("—", " — "))
    return re.sub(r"\s+", " ", t).strip()


def make_slug(text, maxlen=34):
    slug = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")
    return slug[:maxlen] or "section"


# ---- <pause X sec> support -------------------------------------------------
# Insert a literal marker like "<pause 2 sec>" anywhere in the narration text to
# get X seconds of silence at that point. Accepts <pause 2>, <pause 2s>,
# <pause 2 sec>, <pause 2.5 seconds> (case-insensitive).
PAUSE_RE = re.compile(r"<\s*pause\s+([0-9]*\.?[0-9]+)\s*(?:s|sec|secs|second|seconds)?\s*>", re.I)


def split_pauses(text):
    """Split narration into an ordered list of ('text', str) / ('pause', seconds)."""
    out, last = [], 0
    for m in PAUSE_RE.finditer(text or ""):
        pre = (text or "")[last:m.start()]
        if pre.strip():
            out.append(("text", pre.strip()))
        out.append(("pause", float(m.group(1))))
        last = m.end()
    tail = (text or "")[last:]
    if tail.strip() or not out:
        out.append(("text", tail.strip()))
    return out


def strip_pauses(text):
    """Remove pause markers (for on-screen subtitles / display)."""
    return re.sub(r"\s+", " ", PAUSE_RE.sub(" ", text or "")).strip()


# ---------------------------------------------------------- script / narration doc parsing
def parse_script(docx):
    """Parse a narration Word doc.

    Tries the heading style first (Pentesting: 'Slide N' as a heading paragraph,
    body text in following paragraphs).  Falls back to the colon/dash style
    (ACAS: 'Slide N: text' or 'Slide N - text' inline).
    Returns {1-based slide number: narration text}.
    """
    result = _parse_heading_style(docx)
    if result:
        return result
    return _parse_colon_style(docx)


def _parse_heading_style(docx):
    from docx import Document
    marker = re.compile(r"^\s*Slide\s+(\d+)\s*$", re.I)
    out, cur = {}, None
    for para in Document(str(docx)).paragraphs:
        t = para.text.strip()
        if not t:
            continue
        m = marker.match(t)
        if m:
            cur = int(m.group(1))
            out.setdefault(cur, [])
        elif cur is not None:
            out[cur].append(t)
    return {k: " ".join(v).strip() for k, v in out.items() if " ".join(v).strip()}


def _parse_colon_style(docx):
    from docx import Document
    marker  = re.compile(r"^\s*Slide\s+(\d+)\s*[:\-]", re.I)
    section = re.compile(r"^\s*(Module\s+\d+\s*:|Final\s+Hands-?On|Course\s+Wrap-?Up)", re.I)
    out, cur = {}, None
    for para in Document(str(docx)).paragraphs:
        txt = para.text.strip()
        if not txt:
            continue
        m = marker.match(txt)
        if m:
            cur = int(m.group(1))
            out.setdefault(cur, [])
            rest = txt[m.end():].strip()
            if rest:
                out[cur].append(rest)
            continue
        if cur is None or section.match(txt):
            continue
        out[cur].append(txt)
    return {k: " ".join(v).strip() for k, v in out.items() if " ".join(v).strip()}


# ---------------------------------------------------------- lessons index & matching
STOP = {"the", "a", "an", "is", "are", "of", "to", "for", "its", "how",
        "with", "and", "in", "on", "this"}


def norm(s):
    s = re.sub(r"[^a-z0-9 ]", " ", s.lower())
    return re.sub(r"\s+", " ", s).strip()


def norm_ns(s):
    return " ".join(w for w in norm(s).split() if w not in STOP)


def build_lessons_index(docx):
    from docx import Document
    doc = Document(str(docx))
    secs, cur = [], None
    for para in doc.paragraphs:
        t  = para.text.strip()
        st = (para.style.name or "").lower()
        if st.startswith("heading"):
            lvl = int(re.sub(r"\D", "", st) or 9)
            cur = {"heading": t, "level": lvl, "body": []}
            secs.append(cur)
        elif cur is not None and t:
            cur["body"].append(t)
    for s in secs:
        s["body"] = " ".join(s["body"]).strip()
    return secs


def match_lesson(title, secs):
    nt, nts = norm(title), norm_ns(title)
    best, bscore = None, 0.0
    for s in secs:
        h = s["heading"]
        if re.match(r"^(quiz|summary|in this lesson|meet the team)", h, re.I):
            continue
        if not s["body"] or len(s["body"].split()) < 8:
            continue
        r = max(difflib.SequenceMatcher(None, nt,  norm(h)).ratio(),
                difflib.SequenceMatcher(None, nts, norm_ns(h)).ratio())
        if r > bscore:
            best, bscore = s, r
    return best, bscore


def trim_words(text, limit=130):
    words = text.split()
    if len(words) <= limit:
        return text
    cut = " ".join(words[:limit])
    end = cut.rfind(". ")
    return cut[:end + 1] if end > 40 else cut


def split_sentences(text):
    return [s.strip() for s in re.split(r"(?<=[.!?])\s+", text) if s.strip()]


def distribute(sentences, n):
    if n <= 0:
        return []
    base, rem, pos, chunks = len(sentences) // n, len(sentences) % n, 0, []
    for j in range(n):
        take = base + (1 if j < rem else 0)
        chunks.append(sentences[pos:pos + take])
        pos += take
    return chunks


def course_blurb(courseinfo_path, lessons_secs):
    if courseinfo_path and Path(courseinfo_path).exists():
        try:
            from docx import Document
            for p in Document(str(courseinfo_path)).paragraphs:
                t = p.text.strip()
                if t.lower().startswith("objectives") and len(t.split()) > 8:
                    return t.split(":", 1)[1].strip() if ":" in t else t
        except Exception:
            pass
    w = next((s for s in lessons_secs if "welcome" in s["heading"].lower()), None)
    if w:
        return w["body"].split(". ")[0] + "."
    return ""


# ---------------------------------------------------------- section detection
def detect_sections_divider(slides, divider_mark, footer_re, label_re):
    """Auto-detect sections from slides that contain divider_mark."""

    def is_div(sl):
        return any(divider_mark in l.lower() for l in slide_lines(sl))

    def div_title(sl):
        lines = slide_lines(sl)
        for j, l in enumerate(lines):
            if re.match(r"^PART\s*\d+", l, re.I):
                for nxt in lines[j + 1:]:
                    if not (label_re.search(nxt) or footer_re.search(nxt)
                            or divider_mark in nxt.lower()):
                        return nxt
        desc = ""
        for l in lines:
            if (label_re.search(l) or footer_re.search(l)
                    or divider_mark in l.lower()
                    or re.match(r"^CAPSTONE", l, re.I)):
                continue
            if len(l) > len(desc):
                desc = l
        return desc

    def div_spoken(sl):
        lines  = slide_lines(sl)
        mod    = next((l for l in lines if re.match(r"^Module\s*\d+", l,   re.I)), "")
        part   = next((l for l in lines if re.match(r"^PART\s*\d+",   l,   re.I)), "")
        cap    = next((l for l in lines if re.match(r"^CAPSTONE",      l,   re.I)), "")
        desc   = div_title(sl)
        pieces = []
        if cap:  pieces.append("Capstone")
        if mod:  pieces.append(mod)
        if part: pieces.append(part.title())
        prefix = (". ".join(pieces) + ". ") if pieces else ""
        return (prefix + desc).strip()

    starts = [i for i, s in enumerate(slides) if i == 0 or is_div(s)]
    secs = []
    for k, st in enumerate(starts):
        end = (starts[k + 1] - 1) if k + 1 < len(starts) else len(slides) - 1
        if st == 0:
            name, slug = "Introduction", "introduction"
        else:
            name = div_spoken(slides[st])
            slug = make_slug(div_title(slides[st]))
        secs.append({"start": st, "end": end, "name": name, "slug": slug})
    return secs


def detect_sections_modules(modules_json_path):
    """Load explicit section map from JSON (1-based slide numbers -> 0-based internally)."""
    data = json.loads(Path(modules_json_path).read_text(encoding="utf-8"))
    secs = []
    for item in data:
        secs.append({
            "start": item["start"] - 1,
            "end":   item["end"]   - 1,
            "name":  item.get("name", item.get("slug", "section")),
            "slug":  item.get("slug", make_slug(item.get("name", "section"))),
        })
    return secs


def detect_sections_single(n_slides, course_name):
    return [{"start": 0, "end": n_slides - 1,
             "name": course_name, "slug": make_slug(course_name)}]


# ---------------------------------------------------------- narration synthesis
def _lesson_module_key(sec_name):
    m = re.search(r"Module\s*(\d+)", sec_name, re.I)
    if m: return m.group(1)
    if re.search(r"Capstone|Final\s*Hands",  sec_name, re.I): return "capstone"
    if re.search(r"Wrap-?Up",                sec_name, re.I): return "wrapup"
    if re.search(r"^Intro",                  sec_name, re.I): return "intro"
    return None


def build_narration(slides, sections, lessons_secs, mode,
                    course_name, org, title_text, divider_mark, footer_re, label_re):
    """Return list of {text, source} aligned to slides (0-based)."""
    blurb = course_blurb(None, lessons_secs)

    # pre-distribute lessons prose across module slides (lessons mode)
    dist = {}
    if mode == "lessons" and lessons_secs:
        groups_prose = {}
        cur_key = None
        for s in lessons_secs:
            h, lvl = s["heading"], s["level"]
            if lvl == 1:
                if   re.match(r"^Introduction",              h, re.I): cur_key = "intro"
                elif re.match(r"^Module\s*(\d+)",            h, re.I): cur_key = re.match(r"^Module\s*(\d+)", h, re.I).group(1)
                elif re.search(r"Capstone|Final\s*Hands-?On",h, re.I): cur_key = "capstone"
                elif re.search(r"Course\s*Wrap-?Up",         h, re.I): cur_key = "wrapup"
                else:                                                    cur_key = None
                if cur_key:
                    groups_prose.setdefault(cur_key, [])
            if cur_key and not re.match(r"^\s*quiz", h, re.I) and s["body"]:
                groups_prose.setdefault(cur_key, []).append(s["body"])
        groups_prose = {k: " ".join(v) for k, v in groups_prose.items()}

        from collections import defaultdict
        sec_idxs = defaultdict(list)
        for sec in sections:
            key = _lesson_module_key(sec["name"])
            if key is None:
                continue
            # skip the divider slide itself (sec["start"]) — include content slides only
            first_content = sec["start"] + (1 if divider_mark and sec["start"] > 0 else 0)
            for i in range(first_content, sec["end"] + 1):
                sec_idxs[key].append(i)

        for key, idxs in sec_idxs.items():
            chunks = distribute(split_sentences(groups_prose.get(key, "")), len(idxs))
            for idx, ch in zip(idxs, chunks):
                if ch:
                    dist[idx] = " ".join(ch)

    out = []
    for i, sl in enumerate(slides):
        # title slide
        if i == 0:
            if title_text:
                txt = title_text
            else:
                t        = slide_title_text(sl, footer_re, label_re, divider_mark)
                name_in  = t or course_name
                org_part = f" This is a {org} course." if org else ""
                txt      = f"Welcome to {name_in}.{org_part} {blurb}".strip()
            out.append({"text": clean_for_speech(txt), "source": "title"})
            continue

        # divider slide
        if divider_mark and any(divider_mark in l.lower() for l in slide_lines(sl)):
            sec    = next((s for s in sections if s["start"] == i), None)
            spoken = sec["name"] if sec else slide_body_text(sl, footer_re, divider_mark)
            out.append({"text": clean_for_speech(spoken), "source": "divider"})
            continue

        # lessons distribute
        if mode == "lessons" and i in dist:
            out.append({"text": clean_for_speech(dist[i]), "source": "lesson-dist"})
            continue

        # auto: match lesson prose to this slide's title
        if mode == "auto" and lessons_secs:
            t   = slide_title_text(sl, footer_re, label_re, divider_mark)
            sec, score = match_lesson(t, lessons_secs)
            if sec and score >= 0.62:
                out.append({"text": clean_for_speech(trim_words(sec["body"])),
                            "source": f"lesson:{sec['heading'][:40]} ({score:.2f})"})
                continue

        # fallback: slide text
        out.append({"text": clean_for_speech(slide_body_text(sl, footer_re, divider_mark)),
                    "source": "slide-text"})
    return out


# -------------------------------------------------------------------- TTS
class TTS:
    def __init__(self, a):
        self.a = a
        self._kok = None

    def ext(self):
        return ".mp3" if self.a.engine in ("edge", "elevenlabs") else ".wav"

    def synth(self, text, dest):
        method = "_voicecast" if self.a.engine == "voicecast" else "_" + self.a.engine
        return getattr(self, method)(text, dest)

    def _kokoro(self, text, dest):
        import numpy as np, soundfile as sf
        if self._kok is None:
            from kokoro import KPipeline
            self._kok = KPipeline(lang_code=self.a.lang)
        ch = [x for _, _, x in self._kok(text, voice=self.a.voice)]
        sf.write(str(dest), np.concatenate(ch) if ch else np.zeros(1), 24000)
        return dest

    def _piper(self, text, dest):
        run(["piper", "--model", self.a.piper_model, "--output_file", str(dest)], input=text)
        return dest

    def _edge(self, text, dest):
        import edge_tts
        asyncio.run(edge_tts.Communicate(text, self.a.edge_voice).save(str(dest)))
        return dest

    def _voicecast(self, text, dest):
        dest = dest.with_suffix(".wav")
        # Reuse the batch pre-pass output (chunked, full-length) if present.
        if dest.exists() and dest.stat().st_size > 0:
            return dest
        # Fallback: single call (only reached with --voicecast-no-batch; long text may truncate/repeat).
        python  = _vc_python(self.a)
        vcloner = Path(self.a.voicecast_dir) / "vcloner.py"
        if not Path(python).exists():
            raise RuntimeError(f"VoiceCast venv python not found at {python}. "
                               f"Check --voicecast-dir or pass --voicecast-python.")
        if not self.a.voice_sample or not Path(self.a.voice_sample).exists():
            raise RuntimeError(f"Voice sample not found: {self.a.voice_sample}. Provide --voice-sample.")
        p = subprocess.run(
            [str(python), str(vcloner),
             "-i", self.a.voice_sample, "-t", text, "-o", str(dest),
             "-e", self.a.voicecast_engine, "-l", self.a.voicecast_lang, "--no-play"],
            capture_output=True, text=True)
        if p.returncode != 0:
            raise RuntimeError(f"VoiceCast failed:\n{p.stderr[-800:]}")
        return dest

    def _elevenlabs(self, text, dest):
        import requests
        key  = os.environ["ELEVENLABS_API_KEY"]
        v    = os.environ["ELEVENLABS_VOICE_ID"]
        base = os.environ.get("ELEVENLABS_BASE_URL", "https://api.elevenlabs.io")
        r = requests.post(
            f"{base.rstrip('/')}/v1/text-to-speech/{v}",
            headers={"xi-api-key": key, "accept": "audio/mpeg",
                     "content-type": "application/json"},
            json={"text": text, "model_id": "eleven_multilingual_v2"}, timeout=180)
        r.raise_for_status()
        dest.write_bytes(r.content)
        return dest


def audio_duration(path):
    return float(run([TOOLS["ffprobe"], "-v", "error", "-show_entries", "format=duration",
                      "-of", "default=noprint_wrappers=1:nokey=1",
                      str(Path(path).resolve())]).strip())


# ------------------------------------------------------------------ clips
def srt_ts(t):
    h = int(t // 3600); m = int((t % 3600) // 60); s = int(t % 60); ms = int((t - int(t)) * 1000)
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


def make_silence(dest, sec):
    run([TOOLS["ffmpeg"], "-y", "-f", "lavfi", "-i",
         "anullsrc=channel_layout=stereo:sample_rate=44100",
         "-t", f"{sec:.2f}", "-c:a", "libmp3lame", "-q:a", "9", str(dest)])


def slide_clip(png, audio, dest, srt):
    vf = (f"scale={W}:{H}:force_original_aspect_ratio=decrease,"
          f"pad={W}:{H}:(ow-iw)/2:(oh-ih)/2:color=black,setsar=1")
    if srt and srt.exists():
        vf += (f",subtitles='{srt.as_posix()}':force_style="
               f"'Fontsize=20,PrimaryColour=&H00FFFFFF&,BackColour=&H80000000&,"
               f"BorderStyle=4,MarginV=30'")
    run([TOOLS["ffmpeg"], "-y", "-loop", "1", "-i", str(png), "-i", str(audio),
         "-c:v", "libx264", "-tune", "stillimage", "-c:a", "aac", "-b:a", "192k",
         "-pix_fmt", "yuv420p", "-shortest", "-vf", vf, "-r", "25", str(dest)])


def concat_clips(clips, dest, workdir):
    lst = workdir / f"concat_{dest.stem}.txt"
    lst.write_text("".join(f"file '{c.resolve().as_posix()}'\n" for c in clips), encoding="utf-8")
    run([TOOLS["ffmpeg"], "-y", "-f", "concat", "-safe", "0", "-i", str(lst),
         "-c", "copy", str(dest)])


def concat_audio(parts, dest):
    """Concatenate audio files (speech + silence) into dest, normalising sample
    rate / channels so mixed formats (e.g. wav speech + mp3 silence) join cleanly."""
    inputs = []
    for p in parts:
        inputs += ["-i", str(p)]
    n = len(parts)
    pre = "".join(f"[{k}:a]aresample=44100,aformat=sample_fmts=s16:channel_layouts=stereo[a{k}];"
                  for k in range(n))
    filt = pre + "".join(f"[a{k}]" for k in range(n)) + f"concat=n={n}:v=0:a=1[out]"
    codec = (["-c:a", "libmp3lame", "-q:a", "2"] if str(dest).lower().endswith(".mp3")
             else ["-c:a", "pcm_s16le"])
    run([TOOLS["ffmpeg"], "-y", *inputs, "-filter_complex", filt, "-map", "[out]", *codec, str(dest)])


def synth_with_pauses(tts, text, dest):
    """Synthesise a slide, honouring <pause X sec> markers by inserting silence."""
    # voicecast batch pre-pass already produced the whole-slide audio (pauses
    # handled inside voicecast_synth.py), so reuse it as-is.
    if tts.a.engine == "voicecast":
        pre = dest.with_suffix(".wav")
        if pre.exists() and pre.stat().st_size > 0:
            return pre
    parts = split_pauses(text)
    if len(parts) == 1 and parts[0][0] == "text":
        return tts.synth(parts[0][1], dest)            # no pauses → unchanged behaviour
    ext = tts.ext()
    pieces = []
    for k, (kind, val) in enumerate(parts):
        if kind == "text":
            if val.strip():
                seg = dest.with_name(f"{dest.stem}_seg{k:02d}{ext}")
                tts.synth(val, seg)
                pieces.append(seg)
        else:                                          # pause
            sil = dest.with_name(f"{dest.stem}_sil{k:02d}.mp3")
            make_silence(sil, float(val))
            pieces.append(sil)
    if not pieces:                                     # e.g. text was only a pause with 0 length
        make_silence(dest, 0.3); return dest
    if len(pieces) == 1:
        Path(pieces[0]).replace(dest); return dest
    concat_audio(pieces, dest)
    for p in pieces:
        try: Path(p).unlink()
        except OSError: pass
    return dest


# ------------------------------------------------------------------- main
def _vc_python(args):
    """Resolve the voice-cast venv python (cross-platform), honouring --voicecast-python."""
    if getattr(args, "voicecast_python", None):
        return args.voicecast_python
    d = Path(args.voicecast_dir)
    for c in (d / ".venv" / "Scripts" / "python.exe",   # Windows
              d / ".venv" / "bin" / "python"):           # Linux/macOS
        if c.exists():
            return str(c)
    return str(d / ".venv" / "Scripts" / "python.exe")   # best guess; validated by caller


def _voicecast_helper():
    """Locate voicecast_synth.py (shipped alongside this script)."""
    for c in (Path(__file__).resolve().parent / "voicecast_synth.py",
              Path.cwd() / "voicecast_synth.py"):
        if c.exists():
            return c
    sys.exit("ERROR: voicecast_synth.py not found. Keep it in the same folder as build_course.py "
             "(it does the chunked, single-load voice cloning that prevents truncation/repetition).")


def run_voicecast_batch(args, work, jobs):
    """Synthesise ALL voicecast slides in ONE model load, splitting each slide's
    narration into short chunks so the engine can't truncate long text or loop/repeat."""
    helper = _voicecast_helper()
    python = _vc_python(args)
    if not Path(python).exists():
        sys.exit(f"ERROR: VoiceCast venv python not found: {python}. "
                 f"Pass --voicecast-python or check --voicecast-dir.")
    if not args.voice_sample or not Path(args.voice_sample).exists():
        sys.exit("ERROR: --engine voicecast needs --voice-sample <your-voice.wav>.")
    manifest = work / "voicecast_jobs.json"
    manifest.write_text(json.dumps([{"out": j["out"], "text": j["text"]} for j in jobs],
                                   ensure_ascii=False), encoding="utf-8")
    cmd = [python, str(helper),
           "--vc-dir", str(Path(args.voicecast_dir).resolve()),
           "--speaker", str(Path(args.voice_sample).resolve()),
           "--engine", args.voicecast_engine, "--language", args.voicecast_lang,
           "--manifest", str(manifest),
           "--max-chars", str(args.voicecast_max_chars), "--gap", str(args.voicecast_gap)]
    print(f"[voicecast] batch-synthesising {len(jobs)} slides (model loads once; "
          f"chunk <= {args.voicecast_max_chars} chars) ...")
    subprocess.run(cmd, check=True)   # stream progress live


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)

    # --- inputs
    ap.add_argument("--pptx", required=True,
                    help="Path to the slide deck (.pptx). Required.")
    ap.add_argument("--script", default=None,
                    help="Word doc with per-slide narration (primary narration source).")
    ap.add_argument("--lessons", default=None,
                    help="Word doc with lesson prose for auto-matching narration.")
    ap.add_argument("--courseinfo", default=None,
                    help="Extra Word doc to extract a course blurb for the title slide.")

    # --- section detection (mutually exclusive)
    sec_grp = ap.add_mutually_exclusive_group()
    sec_grp.add_argument("--divider-mark", default=None,
                         help="Text snippet present on every section-divider slide.")
    sec_grp.add_argument("--modules-json", default=None,
                         help='JSON file: [{"name":"...","slug":"...","start":1,"end":10}, ...]')

    # --- course identity
    ap.add_argument("--course-name", default=None,
                    help="Course name for title narration (default: inferred from pptx filename).")
    ap.add_argument("--org", default=None,
                    help='Organisation phrase, e.g. "CyberSuite Academy by Montimage".')
    ap.add_argument("--title-text", default=None,
                    help="Fully override the title-slide spoken text.")

    # --- slide-text filters
    ap.add_argument("--footer-re", default=None,
                    help="Regex matching footer/page-number lines to suppress from narration.")
    ap.add_argument("--label-re", default=None,
                    help="Regex matching section-label lines (PART N, Module N …) to suppress.")

    # --- output
    ap.add_argument("--out", default="./build",
                    help="Output directory (default: ./build).")
    ap.add_argument("--full", action="store_true",
                    help="Also emit one combined full-course MP4.")

    # --- narration synthesis
    ap.add_argument("--narration", choices=["auto", "slides", "lessons"], default="auto",
                    help="Narration synthesis mode when --script is absent or a slide is missing. "
                         "auto = lessons matched per slide + slide-text fallback; "
                         "lessons = distribute module prose across slides; "
                         "slides = slide text only.  (default: auto)")
    ap.add_argument("--dry-run", action="store_true",
                    help="Map narration without generating audio or video.")

    # --- TTS engine
    ap.add_argument("--engine",
                    choices=["edge", "kokoro", "piper", "elevenlabs", "voicecast"],
                    default="edge")
    ap.add_argument("--lang",        default="a",                  help="Kokoro lang_code (a=EN-US, b=EN-GB, f=FR).")
    ap.add_argument("--voice",       default="af_heart",           help="Kokoro voice name.")
    ap.add_argument("--piper-model", default="en_US-lessac-high",  help="Piper model name.")
    ap.add_argument("--edge-voice",  default="en-US-AndrewNeural", help="Edge TTS voice name.")
    ap.add_argument("--burn-subs",   action="store_true",          help="Burn subtitles into video.")

    # --- VoiceCast engine options
    ap.add_argument("--voice-sample", default=None,
                    help="voicecast: path to your WAV voice recording (5–30 seconds).")
    ap.add_argument("--voicecast-engine", default="chatterbox-turbo",
                    choices=["coqui", "chatterbox", "chatterbox-turbo"],
                    help="voicecast: TTS sub-engine (default: chatterbox-turbo).")
    ap.add_argument("--voicecast-dir",
                    default=r"E:\Projects\CyberSuite\Courses\voice-cast",
                    help="voicecast: path to the voice-cast repo root.")
    ap.add_argument("--voicecast-python", default=None,
                    help="voicecast: python of the voice-cast venv "
                         "(default: <dir>/.venv/Scripts/python.exe on Windows, <dir>/.venv/bin/python elsewhere).")
    ap.add_argument("--voicecast-lang", default="en",
                    help="voicecast: language code passed to the cloner (default: en).")
    ap.add_argument("--voicecast-max-chars", type=int, default=220,
                    help="voicecast: split each slide's narration into chunks no longer than this many "
                         "characters before cloning, then join them. Prevents the engine from truncating "
                         "long text or looping/repeating. Lower it (e.g. 160) if you still hear a repeat.")
    ap.add_argument("--voicecast-gap", type=float, default=0.15,
                    help="voicecast: seconds of silence inserted between chunks of the same slide.")
    ap.add_argument("--voicecast-no-batch", action="store_true",
                    help="voicecast: disable the one-load batch pass; synthesise per slide in a single call "
                         "(long slides may truncate/repeat — not recommended).")

    # --- system tools
    ap.add_argument("--soffice",     default=None, help="Path to soffice.com / soffice.exe.")
    ap.add_argument("--poppler-bin", default=None, help="Folder containing pdftoppm + pdfinfo.")
    ap.add_argument("--ffmpeg-bin",  default=None, help="Folder containing ffmpeg + ffprobe.")

    args = ap.parse_args()

    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

    resolve_tools(args, require_ffmpeg=not args.dry_run)

    pptx = Path(args.pptx)
    if not pptx.exists():
        sys.exit(f"ERROR: --pptx not found: {pptx}")

    course_name = args.course_name or re.sub(r"[_\-]+", " ", pptx.stem).strip()
    # compile filter regexes — use a never-matching pattern when not provided
    footer_re = re.compile(args.footer_re, re.I) if args.footer_re else re.compile(r"(?!x)x")
    label_re  = re.compile(args.label_re,  re.I) if args.label_re  else re.compile(r"(?!x)x")

    out    = Path(args.out).resolve()   # absolute — keeps all paths stable regardless of cwd
    videos = out / "videos"
    work   = out / "work"
    clips  = work / "clips"
    srtd   = out / "subtitles"
    narrd  = out / "narration"
    for d in (videos, work, clips, srtd, narrd):
        d.mkdir(parents=True, exist_ok=True)

    slides = list(Presentation(str(pptx)).slides)
    pngs   = render_slides(pptx, work)
    if len(pngs) != len(slides):
        print(f"[warn] {len(slides)} slides but {len(pngs)} images rendered; aligning by order.")

    # --- detect sections
    if args.divider_mark:
        sections = detect_sections_divider(slides, args.divider_mark, footer_re, label_re)
    elif args.modules_json:
        sections = detect_sections_modules(args.modules_json)
    else:
        sections = detect_sections_single(len(slides), course_name)

    print(f"[sections] {len(sections)} section(s):")
    for k, s in enumerate(sections):
        print(f"  {k:02d}-{s['slug']:<36} slides {s['start']+1}-{s['end']+1}  ({s['name'][:50]})")

    # --- build narration
    lessons_secs = (build_lessons_index(Path(args.lessons))
                    if args.lessons and Path(args.lessons).exists() else [])
    if args.lessons and not lessons_secs:
        print(f"[warn] --lessons {args.lessons} not found or empty; skipping.")

    blurb = course_blurb(args.courseinfo, lessons_secs)

    script_path = Path(args.script) if args.script else None
    if script_path and script_path.exists():
        doc_narr = parse_script(script_path)
        narration = [
            {"text":   clean_for_speech(doc_narr.get(i + 1, "")),
             "source": script_path.name if (i + 1) in doc_narr else "MISSING"}
            for i in range(len(slides))
        ]
        miss = [i + 1 for i in range(len(slides)) if (i + 1) not in doc_narr]
        print(f"[narration] {script_path.name}: {len(doc_narr)}/{len(slides)} slides"
              + (f"; missing slides (silent hold): {miss}" if miss else "."))
        # fill gaps with lessons/slide-text synthesis when available
        if miss and (lessons_secs or args.narration == "slides"):
            fallback = build_narration(slides, sections, lessons_secs, args.narration,
                                       course_name, args.org, args.title_text,
                                       args.divider_mark, footer_re, label_re)
            for i, n in enumerate(narration):
                if not n["text"] and i < len(fallback) and fallback[i]["text"]:
                    narration[i] = {**fallback[i],
                                    "source": f"fallback:{fallback[i]['source']}"}
    else:
        if script_path:
            print(f"[narration] {script_path} not found — using synthesis ({args.narration} mode).")
        narration = build_narration(slides, sections, lessons_secs, args.narration,
                                    course_name, args.org, args.title_text,
                                    args.divider_mark, footer_re, label_re)
        # override title-slide blurb from courseinfo if available
        if blurb and narration:
            txt = narration[0]["text"]
            if blurb not in txt:
                narration[0]["text"] = clean_for_speech(txt.rstrip(".") + ". " + blurb)

    from collections import Counter
    tally = Counter(n["source"].split(":")[0] for n in narration)
    print(f"[narration] {len(narration)} slides — sources: {dict(tally)}")

    # --- generate clips
    tts = None if args.dry_run else TTS(args)

    # voicecast: synthesise every slide up front in ONE model load, with sentence
    # chunking (fixes long-text truncation and repeated/looping speech).
    if args.engine == "voicecast" and not args.dry_run and not args.voicecast_no_batch:
        jobs = [{"i": i + 1, "out": str((work / f"audio-{i+1:03d}.wav").resolve()),
                 "text": narration[i]["text"]}
                for i in range(len(pngs))
                if i < len(narration) and narration[i]["text"].strip()]
        if jobs:
            run_voicecast_batch(args, work, jobs)

    clip_paths = {}
    for i, png in enumerate(pngs):
        nz   = narration[i] if i < len(narration) else {"text": "", "source": "none"}
        text = nz["text"]
        (narrd / f"slide-{i+1:03d}.txt").write_text(
            f"[{nz['source']}]\n{text}\n", encoding="utf-8")
        srt  = srtd  / f"slide-{i+1:03d}.srt"
        clip = clips / f"clip-{i+1:03d}.mp4"

        if text and not args.dry_run:
            print(f"[tts] slide {i+1:03d}: {len(text.split()):3d} words  <{nz['source']}>")
            audio = work / f"audio-{i+1:03d}{tts.ext()}"
            audio = synth_with_pauses(tts, text, audio)   # honours <pause X sec>; voicecast may change suffix
            if not audio.exists():
                raise RuntimeError(
                    f"TTS engine '{args.engine}' produced no audio file for slide {i+1}.\n"
                    f"Expected: {audio}\n"
                    f"Check that the engine is installed and working (try --dry-run first).")
            dur = audio_duration(audio)
            srt.write_text(f"1\n{srt_ts(0)} --> {srt_ts(dur)}\n{strip_pauses(text)}\n", encoding="utf-8")
            slide_clip(png, audio, clip, srt if args.burn_subs else None)
        elif text and args.dry_run:
            dur = max(MIN_HOLD, len(text.split()) / 2.6)
            srt.write_text(f"1\n{srt_ts(0)} --> {srt_ts(dur)}\n{text}\n", encoding="utf-8")
        elif not text and not args.dry_run:
            audio = work / f"audio-{i+1:03d}.mp3"
            make_silence(audio, MIN_HOLD)
            slide_clip(png, audio, clip, None)
        clip_paths[i] = clip

    # --- assemble section videos
    print("\n[assemble] per-section videos:")
    all_clips = []
    for k, s in enumerate(sections):
        present = [clip_paths[i] for i in range(s["start"], s["end"] + 1) if i in clip_paths]
        all_clips += present
        fname = f"{k:02d}-{s['slug']}.mp4"
        print(f"  {fname:<45}  ({len(present)} clips)")
        if not args.dry_run and present:
            concat_clips(present, videos / fname, work)

    if args.full and not args.dry_run and all_clips:
        full_name = make_slug(course_name) + "_full-course.mp4"
        concat_clips(all_clips, videos / full_name, work)
        print(f"  {full_name}  (combined)")

    print("\nDry run complete — review build/narration/*.txt, then re-run without --dry-run."
          if args.dry_run else "\nDone.")
    print(f"Output: {videos}")


if __name__ == "__main__":
    main()
