# copy-all

Watch the Windows clipboard and append everything you copy to a single
document until you stop it. Choose when it starts: a **Word** file (text and
pictures) or a **Markdown** / **plain text** file (words only).

You never touch the output file while you work. You just browse: on each page
press **Ctrl+A** then **Ctrl+C**, and a moment later that page is another
section of the document. Copy a screenshot or an image from a page and that
lands in there too.

## Setup (once)

1. Install Python 3.9+ from <https://python.org> and tick **"Add python.exe to
   PATH"** during setup.
2. Put this folder anywhere you like (Desktop is fine).
3. Double-click **`run-capture.bat`**. The first run installs the three
   packages it needs (`pywin32`, `Pillow`, `python-docx`) and then starts.

Prefer the command line?

```
pip install -r requirements.txt
python capture_clipboard.py
```

## Using it

```
> run-capture.bat
What should the captures be written to?

  [1] Word document    .docx   - text and pictures
  [2] Markdown file    .md     - text only
  [3] Plain text file  .txt    - text only

Choose 1-3 [1]: 1

Clipboard capture running.
  Writing to: C:\Users\you\Documents\clipboard-capture_2026-10-24_1412.docx
  Copy anything (Ctrl+A then Ctrl+C on a page) and it gets appended.
  Press Ctrl+C here to stop.

[1] 4821 chars, 6 image(s)  <- https://en.wikipedia.org/wiki/Clipboard
[2] 1190 chars, 1 image(s)  <- https://example.com/pricing
[3] 0 chars, 1 image(s)  <- image
```

Leave that window open (minimised is fine) and go browse. Press **Ctrl+C in
that window** to stop; the path is printed again on the way out.

The file is saved after *every* capture, so killing the window or losing power
never costs you more than the capture in flight.

## Options

```
python capture_clipboard.py --format md
python capture_clipboard.py --out "C:\Users\you\Desktop\research.md"
```

Give `--out` a `.docx`, `.md` or `.txt` name and the format is taken from it,
so you are not asked.

| Option | What it does |
| --- | --- |
| `--format`, `-f` | `docx`, `md` or `txt` — skips the startup question |
| `--out PATH`, `-o` | Where to write it. Default: `Documents\clipboard-capture_<date>.<ext>` |
| `--no-images` | Words only, even for Word format |
| `--min-image-size N` | Skip images under N pixels a side (default 32, keeps out spacers and tracking pixels) |
| `--image-timeout S` | How long to wait per image download (default 10s) |
| `--min-chars N` | Ignore copies shorter than N characters |
| `--allow-duplicates` | Record a capture even if it's identical to the one before |
| `--interval S` | How often to check the clipboard (default 0.4s) |

Pass them through the launcher too: `run-capture.bat --no-images`

## What ends up in the document

Each capture is headed with a number, a timestamp and the page's URL — a new
page in Word, a `---` section in Markdown, a `====` banner in text. Underneath,
in the order they appeared:

- headings, paragraphs and bullet lists
- tables, flattened to one line per row with `|` between cells
- preformatted / code blocks (fenced in Markdown, indented in text)
- images, in the Word format only

Pictures only apply to the Word format, where they are embedded in the file
itself, scaled to the page width, with the `alt` text as a caption — one
self-contained file to send. **Markdown and plain text are words only**:
pictures are skipped, and because of that those two formats never touch the
network, so each capture is instant.

## Worth knowing

- **It watches the clipboard, so it records *everything* you copy** while it
  runs — including a password or a card number pasted between apps. Stop it
  before you do anything private, and keep in mind the output file is plain,
  unencrypted `.docx`.
- **Images are fetched over the network** (Word format only) from the page you copied, using that
  page's URL as the referer. An image behind a login or a hotlink block can't
  be fetched; you get an `[image not saved: alt text]` line instead so you know
  something was there. Images pasted as data URIs, and screenshots copied
  straight to the clipboard, need no network at all.
- **Formatting is simplified on purpose** — structure and pictures survive,
  bold/italic/colour/fonts do not. The goal is a readable, searchable record,
  not a pixel copy of the page.
- **Copying the same thing twice in a row is ignored** (that's what
  `--allow-duplicates` turns off).
- **If the output file is open in Word**, Windows locks it and the save fails.
  The script says so and keeps everything in memory, retrying on the next
  capture — so don't open the document in Word until you've stopped the script,
  or open a copy of it.
- Some sites block Ctrl+A/Ctrl+C entirely. Nothing reaches the clipboard then,
  so there's nothing for this to capture.
- Windows only — it uses the Windows clipboard API directly.
