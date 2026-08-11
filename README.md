# Video\_course\_generation: converts slide deck + text to narrated MP4

Turn a PowerPoint deck and a Word doc narration text into voice‑only narrated MP4 videos — one file per course section. Narration can come from a script/lessons Word document, and speech is produced by a selectable text‑to‑speech engine, including **voice cloning** via the external *voice‑cast* toolkit.

This README covers everything needed to **run** "build\_course.py".





## 1\. Python version

|Text-to-speech engines|Python version|
|-|-|
|`edge`, `piper`, `elevenlabs`, `voicecast` engines|**Python 3.10 – 3.13** (3.13 is fine)|
|`kokoro` engine|**Python 3.10 – 3.12 only** (Kokoro/its deps have no 3.13 wheels)|
|The **voice‑cast** tool (used by `--engine voicecast`)|Its **own** venv on **Python 3.10 – 3.12**|

"build\_course.py" itself runs fine on **Python 3.13**. Only the optional `kokoro` engine, and the separate voice‑cast install, require 3.10–3.12.





## 2\. System tools (required)

"build\_course.py" shells out to three external programs. Install them and either put them on your `PATH` or pass the folder/exe explicitly.

|Tool|Used for|Install (Windows)|Override flag|
|-|-|-|-|
|**LibreOffice**|render the .pptx to PDF|download from libreoffice.org<br />or<br />winget install TheDocumentFoundation.LibreOffice|`--soffice "C:\Program Files\LibreOffice\program\soffice.com"`|
|**Poppler** (`pdftoppm`, `pdfinfo`)|PDF → slide images|`choco install poppler` (also found in a Calibre install)|`--poppler-bin "C:\Program Files\Calibre2\app\bin"`|
|**ffmpeg** (`ffmpeg`, `ffprobe`)|build/concatenate video + audio|download from ffmpeg.org (e.g. `C:\ffmpeg\bin`)<br />or<br />winget install Gyan.FFmpeg|`--ffmpeg-bin "C:\ffmpeg\bin"`|

The script auto‑discovers these on `PATH` and in common locations; use the override flags if it can’t find them.





## 3\. Python packages

Install into the environment you’ll run "build\_course.py" with.

**Always required (for any engine):**



pip install python-pptx python-docx





**Then install the package(s) for the engine you choose:**

|`--engine`|Install|Notes|
|-|-|-|
|`edge` *(default)*|pip install edge-tts|Free, natural, **online** (no key).|
|`kokoro`|pip install kokoro soundfile|Local, CPU. Also needs **espeak‑ng** installed, and **Python ≤ 3.12**.|
|`piper`|pip install piper-tts|Local, tiny, offline.|
|`elevenlabs`|pip install requests|Also set env vars `ELEVENLABS\_API\_KEY` and `ELEVENLABS\_VOICE\_ID`.|
|`voicecast`|*(nothing here)*|Uses the external voice‑cast install — see **§4**.|

Quick start for the default engine:



pip install python-pptx python-docx edge-tts





## 4\. Voice cloning with `--engine voicecast`

The `voicecast` engine narrates in a **cloned voice** using the external
[voice‑cast](https://github.com/luongnv89/voice-cast) toolkit (MIT‑licensed).

**Important points**

* voice‑cast is **not part of this project** and is **not modified** — it’s used as‑is, as a library. Install it separately.
* It has heavy dependencies (PyTorch, XTTS/Chatterbox) and must live in **its own Python 3.10–3.12 virtual environment**, independent of the environment running "build\_course.py".
* "build\_course.py" does **not** import voice‑cast directly. It calls the helper **"voicecast\_synth.py"** (shipped with this project) using voice‑cast’s own venv Python. Keep "voicecast\_synth.py" in the **same folder as "build\_course.py"**.





### 4.1 Install voice‑cast (one time)



git clone https://github.com/luongnv89/voice-cast.git
cd voice-cast
py -3.11 -m venv .venv           :: any Python 3.10–3.12
.venv\Scripts\activate
pip install -e ".\[chatterbox]"   :: \[chatterbox] is needed for the chatterbox engines



Notes:

* The default `--voicecast-engine` is **`chatterbox-turbo`**, so install the **`\chatterbox]`** extra as shown. (Use plain `pip install -e .` only if you will use the `coqui` engine.)
* If the helper later reports that `soundfile` is missing, install it **into voice‑cast’s venv**: `pip install soundfile`.
* On first run the model (\~2 GB) downloads from Hugging Face and is then cached locally; later runs are offline.





### 4.2 Prepare a voice sample

Record a **clean 5–30 sec. WAV** of the voice to clone (your own, or a consented sample). Mono, 16‑bit is ideal, e.g.:



ffmpeg -i recording.m4a -ac 1 -ar 22050 -c:a pcm\_s16le my-voice.wav





### 4.3 Run with voice cloning



python build\_course.py ^
--pptx my\_deck.pptx ^
--script my\_narration.docx ^
--engine voicecast ^
--voice-sample "C:\path\to\my-voice.wav" ^
--voicecast-dir "C:\path\to\voice-cast" ^
--soffice "C:\Program Files\LibreOffice\program\soffice.com" ^
--poppler-bin "C:\Program Files\Calibre2\app\bin" ^
--ffmpeg-bin "C:\ffmpeg\bin"



The script auto‑detects voice‑cast’s venv at `<voicecast-dir>\.venv`. If it’s elsewhere, add:



\--voicecast-python "C:\path\to\voice-cast\.venv\Scripts\python.exe"



**Voice‑cast options**

|Flag|Default|Meaning|
|-|-|-|
|`--voice-sample`|*(required)*|WAV/MP3 of the voice to clone|
|`--voicecast-dir`|`E:\…\voice-cast`|path to the cloned voice‑cast repo|
|`--voicecast-engine`|`chatterbox-turbo`|`coqui` \| `chatterbox` \| `chatterbox-turbo`|
|`--voicecast-lang`|`en`|language code|
|`--voicecast-python`|*(auto)*|voice‑cast venv Python, if not at `<dir>\.venv`|
|`--voicecast-max-chars`|`220`|max characters per synthesis chunk (see below)|
|`--voicecast-gap`|`0.15`|seconds of silence between chunks|
|`--voicecast-no-batch`|off|disable the one‑load batch pass (not recommended)|

**Why the chunking exists:** voice‑cast sends the whole slide to the engine at once, which makes long text truncate and the voice loop/repeat. "voicecast\_synth.py" loads the model **once** and splits each slide into chunks of at most `--voicecast-max-chars`, then joins them. If you ever still hear a repeat on a tricky sentence, lower it (e.g. `--voicecast-max-chars 160`).





## 5\. Running (all engines)

### Minimum



python build\_course.py --pptx my\_deck.pptx



With no narration document, each slide is narrated from its on‑slide text; with no section flags, the whole deck becomes one video.

### 

### Provide narration text (choose one)

* `--script narration.docx` — a Word doc with either `Slide N:` lines or `Slide N` headings followed by the narration paragraph.
* `--lessons lessons.docx` — lesson prose (headings + body); the script auto‑matches it to slides when `--script` is absent.

### 

### Split into sections (choose one)

* `--modules-json ranges.json` — explicit slide ranges: `\[{"name":"Intro","slug":"intro","start":1,"end":10}, ...]`
* `--divider-mark "text on every divider slide"` — auto‑detect sections from divider slides.

### 

### Pauses in the narration



Insert <pause X sec> anywhere in the narration text to get X seconds of silence at that point (X may be a decimal). Accepted forms: <pause 2>, <pause 2s>, <pause 2 sec>, <pause 2.5 seconds> (case‑insensitive). It works with every engine (including voicecast), and the marker is stripped from burned‑in subtitles. Example narration line:



Welcome to the course. <pause 1.5 sec> Let's begin with the first topic.



### Typical examples



:: Edge voice, explicit module ranges, script narration
python build\_course.py ^
--pptx file.pptx ^
--script text.docx ^
--modules-json modules.json ^
--engine edge --edge-voice en-US-AndrewNeural ^
--org "Name of organisation"

:: Preview only — no audio/video, just report the plan (dry-run)
python build\_course.py --pptx my\_deck.pptx --script my\_narration.docx --dry-run



:: Using voicecast with the provided example files (note tat the execution take some time: several hours)

py -3.13 .\src\build\_course.py --pptx .\Example\ACAS\_course\_slides\_v4.pptx --script .\Example\ACAS\_course\_script\_v4.docx --modules-json .\\Example\acas\_modules.json --org "CyberSuite Academy by Montimage" --engine voicecast  --voicecast-engine chatterbox-turbo --voice-sample ".\\Example\my-voice.wav" --soffice "C:\Program Files\LibreOffice\program\soffice.com" --poppler-bin "C:\Program Files\Calibre2\app\bin"



:: Results will be found in .\build\videos



### 

### Useful general options

|Flag|Meaning|
|-|-|
|`--out DIR`|output folder for the MP4s|
|`--engine`|`edge` (default) \| `kokoro` \| `piper` \| `elevenlabs` \| `voicecast`|
|`--edge-voice`|Edge voice name (default `en-US-AndrewNeural`)|
|`--lang` / `--voice`|Kokoro language code / voice|
|`--piper-model`|Piper model name|
|`--org` / `--course-name` / `--title-text`|title‑slide identity|
|`--footer-re` / `--label-re`|regexes to drop footer / section‑label lines from narration|
|`--burn-subs`|burn subtitles into the video|
|`--dry-run`|build nothing; print the section/narration plan|

Run "python build\_course.py -h" for the full list.





## 6\. Troubleshooting

* **`Warning: You are sending unauthenticated requests to the HF Hub … set a HF\_TOKEN`** — harmless. It appears while a model downloads (Whisper, or voice‑cast). Downloads work without a token; it stops once the model is cached. Only set `HF\_TOKEN` if you actually hit an HTTP 429 rate‑limit.
* **`soffice` / `pdftoppm` / `ffmpeg` not found** — install the tool (§2) and open a **new** terminal, or pass `--soffice` / `--poppler-bin` / `--ffmpeg-bin`.
* **voicecast: `ChatterboxEngine.\_\_init\_\_() got an unexpected keyword argument 'auto\_download'`** — you’re running an old "voicecast\_synth.py". Use the current one (it never passes `auto\_download`); confirm by searching the file — the correct version contains the text `auto\_download` **zero** times.
* **kokoro install/run fails on Python 3.13** — Kokoro needs Python ≤ 3.12; use `edge` (works on 3.13) or make a 3.12 environment for kokoro.
* **voice cloning truncates or repeats** — make sure the batch pass is on (don’t use `--voicecast-no-batch`) and lower `--voicecast-max-chars` (e.g. 160).





## 7\. Files

* "build\_course.py" — the builder.
* "voicecast\_synth.py" — voice‑cast batch/chunking helper; **keep it beside "build\_course.py"** (only needed for `--engine voicecast`).

