# copy-all

Watch the Windows clipboard and append everything you copy to a single Word
document — text *and* pictures — until you stop it.

You never touch the output file while you work. You just browse: on each page
press **Ctrl+A** then **Ctrl+C**, and a moment later that page is another
section of the `.docx`. Copy a screenshot or an image from a page and that
lands in the document too.

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
python capture_clipboard.py --out "C:\Users\you\Desktop\research.docx"
```

| Option | What it does |
| --- | --- |
| `--out PATH`, `-o` | Where to write the `.docx`. Default: `Documents\clipboard-capture_<date>.docx` |
| `--no-images` | Text only — nothing is downloaded |
| `--min-image-size N` | Skip images under N pixels a side (default 32, keeps out spacers and tracking pixels) |
| `--image-timeout S` | How long to wait per image download (default 10s) |
| `--min-chars N` | Ignore copies shorter than N characters |
| `--allow-duplicates` | Record a capture even if it's identical to the one before |
| `--interval S` | How often to check the clipboard (default 0.4s) |

Pass them through the launcher too: `run-capture.bat --no-images`

## What ends up in the document

Each capture becomes its own page, headed with a number, a timestamp, and the
page's URL. Underneath, in the order they appeared:

- headings, paragraphs and bullet lists
- tables, flattened to one line per row with `|` between cells
- preformatted / code blocks
- images, scaled to the page width, with the `alt` text as a caption

## Worth knowing

- **It watches the clipboard, so it records *everything* you copy** while it
  runs — including a password or a card number pasted between apps. Stop it
  before you do anything private, and keep in mind the output file is plain,
  unencrypted `.docx`.
- **Images are fetched over the network** from the page you copied, using that
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
