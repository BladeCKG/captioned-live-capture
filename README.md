# Captioned Live Capture

Captioned Live Capture is a small Windows UI that reads transcript text directly from Caption.Ed through Windows UI Automation and displays it in a simpler text-focused window.

The important part is that the app does not use OCR anymore. It does not read pixels, run image preprocessing, or try to guess words from screenshots. Instead, it reads the transcript from Caption.Ed's accessibility tree, which is much more accurate and much less noisy for this specific app.

## What Problem This App Solves

Caption.Ed renders its main content inside a Chromium surface with class `Chrome_RenderWidgetHostHWND`.

That surface is only a render target. Normal Win32 text APIs do not expose the transcript content there. If you call standard APIs such as `GetWindowText`, you usually get only the outer window title or a generic label, not the actual transcript text.

The key discovery for this project was that, although the Chromium render widget does not expose transcript text through classic Win32 APIs, Caption.Ed does expose the transcript through Windows accessibility APIs.

That means the transcript can be read as structured text instead of guessed from pixels.

## How Text Extraction Works

The extraction flow is:

1. Find the target Chromium widget by:
   - Process: `Caption.Ed.exe`
   - Class: `Chrome_RenderWidgetHostHWND`
2. Resolve the widget's root top-level window.
3. Open the root window through Windows UI Automation.
4. Find the `Transcript` control in the accessibility tree.
5. Read the transcript text from that control.
6. Normalize the text into paragraph blocks for display.

The main implementation lives in [capture_backend.py](C:/Users/aaa/Documents/dev/captioned-live-capture/capture_backend.py).

### Why `Chrome_RenderWidgetHostHWND` Is Still Used

The app still uses `Chrome_RenderWidgetHostHWND` for discovery because it is the most reliable way to identify the active Caption.Ed transcript surface.

However, the transcript is not read from that child window directly through Win32. Instead:

- the child widget is used to identify the correct app window
- then the app switches to UI Automation on the root window
- then it finds the `Transcript` control in the accessibility tree

So the widget is the entry point for finding the right app instance, not the actual text source.

### The Actual Text Source

The actual text source is the `Transcript` UI Automation control exposed by Caption.Ed.

The code reads that control in this order:

1. `ValuePattern.Value`
2. `TextPattern.DocumentRange`
3. child paragraph controls as a fallback

That order matters:

- `ValuePattern` is the most useful live text source for this app
- `TextPattern` is also usable, but often flatter and less convenient to parse
- child paragraph traversal is kept only as a last-resort fallback

The reason this matters is performance and freshness. The paragraph tree can lag behind and sometimes reflects only finalized chunks. The `ValuePattern` on the `Transcript` control updates more continuously.

## Why OCR Was Removed

OCR was the original approach, but it had several unavoidable problems:

- it only saw what was currently rendered
- it could misread UI elements as transcript text
- it produced garbage lines that never existed in the real transcript
- it depended on Tesseract and image preprocessing
- it was slower and less stable than reading actual text

Examples of OCR-specific failure modes included:

- toolbar text leaking into transcript output
- fake symbols or broken words
- duplicate blocks caused by small OCR differences between captures
- text only being available when visible in the rendered region

For Caption.Ed specifically, UI Automation is the correct solution because the transcript is already exposed as text by the app.

## How the Transcript Is Parsed

The raw `Transcript` control value contains speaker labels, timestamps, blank lines, and transcript content.

Typical raw structure looks like:

```text
Speaker 1
0 seconds

Developed a fear of planning for any responsibility...

Speaker 1
20 seconds

She upbraided me for giving in...
```

The backend normalizes this by:

- removing lines like `Speaker 1`
- removing timestamp lines like `20 seconds` or `7 minutes 9 seconds`
- joining the remaining spoken content into paragraphs
- collapsing extra whitespace

This parsing logic is in `parse_transcript_value(...)` and `normalize_transcript_paragraph(...)` in [capture_backend.py](C:/Users/aaa/Documents/dev/captioned-live-capture/capture_backend.py).

## How Updates Work

The UI polls the transcript on an interval. The default interval is `0.3` seconds.

That means the app is not event-driven right now. It checks repeatedly for the latest transcript text.

The update flow is:

1. background worker requests the latest transcript text
2. backend reads the `Transcript` UI Automation control
3. backend converts raw text into normalized paragraphs
4. UI compares the new paragraph list to the currently displayed paragraph list
5. only the changed tail is replaced in the text widget

That last step is important.

The app used to treat updates more like OCR output and perform heavier full-text merge/rewrite behavior. After moving to direct transcript extraction, that approach was no longer appropriate. Since the backend now returns an ordered transcript, the UI uses a simpler and more correct strategy:

- keep the unchanged prefix
- replace only the changed suffix
- append genuinely new transcript content at the end

This is implemented in `_apply_transcript_update(...)` in [ui_app.py](C:/Users/aaa/Documents/dev/captioned-live-capture/ui_app.py).

## How Pointer Selection Works

The UI supports a pointer-based selection model.

When you click in the text area:

- the click position becomes the pointer
- the pointer is stored as a character offset, not as fragile widget-relative state
- after each update, the app reselects from that stored offset to the end of the current text

This was done because ordinary Tk selection can be disturbed by widget updates. Storing the pointer as an absolute offset makes the behavior much more stable across text refreshes.

Relevant logic is in [ui_app.py](C:/Users/aaa/Documents/dev/captioned-live-capture/ui_app.py):

- `_set_pointer_from_click(...)`
- `_restore_pointer_selection(...)`
- `_select_pointer_to_end(...)`

## Performance Notes

The app is much faster now than the OCR version, but there are still practical limits.

Current performance characteristics:

- no screenshot capture
- no image preprocessing
- no OCR engine
- no fuzzy merge over noisy OCR blocks in the main display path
- only paragraph-based tail replacement in the UI

The remaining cost mostly comes from:

- polling on a timer
- querying UI Automation
- reading the transcript control value
- rewriting the changed tail in the Tk text widget

The app also keeps a per-thread UI Automation session so it does not rebuild COM and control state from scratch on every interval.

## Project Structure

The project is split into a few focused modules:

- [capture_text_app.py](C:/Users/aaa/Documents/dev/captioned-live-capture/capture_text_app.py)
  - thin entrypoint
- [ui_app.py](C:/Users/aaa/Documents/dev/captioned-live-capture/ui_app.py)
  - Tkinter UI
  - interval loop
  - pointer selection
  - auto-scroll
  - auto-copy
  - incremental text updates
- [capture_backend.py](C:/Users/aaa/Documents/dev/captioned-live-capture/capture_backend.py)
  - target window discovery
  - UI Automation transcript access
  - transcript parsing and normalization
- [text_processing.py](C:/Users/aaa/Documents/dev/captioned-live-capture/text_processing.py)
  - retained text-cleaning and matching utilities from earlier iterations
  - no longer the primary extraction mechanism

## Setup

Install Python packages:

```powershell
py -m pip install -r requirements.txt
```

## Run

Run with default Caption.Ed discovery:

```powershell
py capture_text_app.py
```

Override process/class/interval if needed:

```powershell
py capture_text_app.py --process-name Caption.Ed.exe --class-name Chrome_RenderWidgetHostHWND --interval 0.3
```

Force a specific window handle if necessary:

```powershell
py capture_text_app.py --hwnd 394966
```

## Portable Build

Create the portable Windows build with:

```powershell
.\build_release.ps1
```

The build output is:

```text
release\CaptionedLiveCapture-portable.zip
```

To use it on another PC:

1. Extract the zip.
2. Run `CaptionedLiveCapture.exe`.

The portable build includes the packaged Python runtime and the Python dependencies used by the app.

## Limitations

This approach is much better than OCR for Caption.Ed, but it still depends on Caption.Ed continuing to expose transcript text through UI Automation.

If a future app update changes:

- the accessibility tree
- the `Transcript` control name
- the transcript value format
- the Chromium accessibility behavior

then the extraction logic may need adjustment.

It also means this approach is app-specific. It works because Caption.Ed exposes accessible text. It is not a general solution for every Chromium-based app.

## Summary

The main breakthrough was realizing that the transcript did not need to be read from pixels at all.

Instead of:

- capture rendered image
- OCR the image
- clean noisy text

the app now does:

- find Caption.Ed's Chromium widget
- resolve the root app window
- read the `Transcript` accessibility control
- normalize the returned text
- update only the changed tail in the UI

That is why the current version is more accurate, simpler, and easier to maintain than the original OCR-based design.
