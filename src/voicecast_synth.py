#!/usr/bin/env python3
"""
voicecast_synth.py  —  batch voice-cloning helper for the course builders.

Why this exists
---------------
voice-cast's vcloner.py sends the WHOLE text to the TTS engine in one shot
(VoiceCloner.say -> engine.generate). XTTS/Coqui truncates long inputs and can
loop/repeat words. This helper fixes both problems and is far faster for a whole
course, by:
  * loading the cloning model ONCE (not once per slide/sentence), and
  * splitting each slide's text into short sentence-sized chunks, synthesising
    each chunk separately, and concatenating them with a small gap.

It runs INSIDE voice-cast's own venv (heavy deps live there). The caller passes
--vc-dir so `voice_cloner` can be imported regardless of the working directory.

Usage (normally invoked by build_acas_course.py, but usable standalone):
  <voice-cast venv python> voicecast_synth.py \
      --vc-dir /path/to/voice-cast --speaker my_voice.wav \
      --engine coqui --language en --max-chars 220 --gap 0.15 \
      --manifest jobs.json

jobs.json is a list of {"out": "<output wav path>", "text": "<narration>"}.
"""
import argparse, json, os, re, sys, tempfile


def split_text(text, max_chars):
    """Split into sentence-sized chunks <= max_chars. Over-long sentences are
    hard-split on word boundaries so nothing is ever dropped."""
    text = " ".join((text or "").split())
    if not text:
        return []
    sents = re.split(r"(?<=[.!?])\s+", text)
    chunks, cur = [], ""
    for s in sents:
        # hard-split a single sentence that exceeds the limit
        while len(s) > max_chars:
            cut = s.rfind(" ", 0, max_chars)
            if cut <= 0:
                cut = max_chars
            if cur:
                chunks.append(cur); cur = ""
            chunks.append(s[:cut].strip())
            s = s[cut:].strip()
        if not cur:
            cur = s
        elif len(cur) + 1 + len(s) <= max_chars:
            cur += " " + s
        else:
            chunks.append(cur); cur = s
    if cur:
        chunks.append(cur)
    return [c for c in chunks if c.strip()]


PAUSE_RE = re.compile(r"<\s*pause\s+([0-9]*\.?[0-9]+)\s*(?:s|sec|secs|second|seconds)?\s*>", re.I)


def split_pauses(text):
    """Ordered list of ('text', str) / ('pause', seconds) from <pause X sec> markers."""
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


def build_pieces(text, max_chars):
    """Flatten into ordered ('say', chunk) / ('pause', seconds) items."""
    pieces = []
    for kind, val in split_pauses(text):
        if kind == "pause":
            pieces.append(("pause", val))
        else:
            for ch in split_text(val, max_chars):
                pieces.append(("say", ch))
    return pieces


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--vc-dir", required=True)
    ap.add_argument("--speaker", required=True)
    ap.add_argument("--engine", default="coqui")
    ap.add_argument("--language", default="en")
    ap.add_argument("--manifest", required=True)
    ap.add_argument("--max-chars", type=int, default=220)
    ap.add_argument("--gap", type=float, default=0.15)
    ap.add_argument("--temperature", type=float, default=0.7)
    ap.add_argument("--cfg-weight", type=float, default=0.5)
    ap.add_argument("--exaggeration", type=float, default=0.5)
    a = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8"); sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

    sys.path.insert(0, a.vc_dir)
    import numpy as np, soundfile as sf
    from voice_cloner import VoiceCloner

    ekw = {"language": a.language}
    if a.engine == "coqui":
        ekw["temperature"] = a.temperature
    elif "chatterbox" in a.engine:
        ekw["cfg_weight"] = a.cfg_weight; ekw["exaggeration"] = a.exaggeration

    print(f"[voicecast] loading model once (engine={a.engine}) ...", flush=True)
    # Construct with only speaker + engine. (Passing model-download flags here is
    # unsafe: some voice-cast versions forward them to engines that reject them.)
    cloner = VoiceCloner(speaker_wav=a.speaker, engine=a.engine)

    items = json.load(open(a.manifest, encoding="utf-8"))
    tmp = tempfile.mkdtemp(prefix="vc_")
    for n, item in enumerate(items, 1):
        out, text = item["out"], item["text"]
        pieces = build_pieces(text, a.max_chars)
        n_say = sum(1 for k, _ in pieces if k == "say")
        n_pause = sum(1 for k, _ in pieces if k == "pause")
        print(f"[voicecast] {n}/{len(items)} -> {os.path.basename(out)} "
              f"({n_say} chunk(s), {n_pause} pause(s), {len(text)} chars)", flush=True)
        rendered, sr = [], None                        # ('audio', ndarray) | ('pause', seconds)
        for j, (kind, val) in enumerate(pieces):
            if kind == "pause":
                rendered.append(("pause", float(val))); continue
            part = os.path.join(tmp, f"s{n:04d}_{j:02d}.wav")
            for attempt in (1, 2):                      # one retry on transient failure
                try:
                    cloner.say(val, play_audio=False, save_audio=True, output_file=part, **ekw)
                    break
                except Exception as e:
                    if attempt == 2:
                        print(f"[voicecast]   WARNING: chunk {j} failed ({e}); skipping.", flush=True)
                        part = None
                    else:
                        print(f"[voicecast]   retry chunk {j} ...", flush=True)
            if part and os.path.exists(part):
                data, sr = sf.read(part)
                rendered.append(("audio", data))
        os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
        if not any(k == "audio" for k, _ in rendered):
            total = sum(v for k, v in rendered if k == "pause") or 0.3
            sf.write(out, np.zeros(int(total * (sr or 24000))), sr or 24000); continue
        sr = sr or 24000
        gap = np.zeros(int(a.gap * sr))
        segs, prev_audio = [], False
        for kind, val in rendered:
            if kind == "audio":
                if prev_audio and a.gap > 0:
                    segs.append(gap)
                segs.append(val); prev_audio = True
            else:                                        # explicit pause
                segs.append(np.zeros(int(val * sr))); prev_audio = False
        sf.write(out, np.concatenate(segs) if segs else np.zeros(1), sr)
    print("[voicecast] done.", flush=True)


if __name__ == "__main__":
    main()
