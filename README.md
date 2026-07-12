# DCSB Config Path Fixer

A Windows desktop utility designed to organize a local audio library and repair broken audio paths used by **Winamp** and **DCSB (Deathcounter and Soundboard)**. The project is built around practical library maintenance: fixing broken entries in `config.xml`, normalizing filenames, tracking rename history, and supporting related audio-management workflows.[1]

## Overview

This application was created for setups where a soundboard or media tool depends on local audio files whose paths may change over time. After renaming files, moving folders, or cleaning up a music library, stored paths inside DCSB can become invalid; this tool scans the library, compares it with the XML configuration, and repairs references wherever a reliable match can be found.[1]

The application also creates a backup of the target configuration file before saving changes, which makes the workflow safer for real-world use with larger libraries. In addition to path repair, the project includes supporting features such as rename history, checkpoints, revert operations, playlist-related workflow, and audio-library normalization tools.[1]

## Built for

This project was primarily created for the following programs:

- **Winamp** — for managing and using a local music library and playlists.
- **DCSB (Deathcounter and Soundboard)** — for maintaining valid audio file references inside `config.xml` and keeping the soundboard in sync with the actual library.[1]

## Features

- Repairs broken audio paths stored in DCSB `config.xml` files.[1]
- Creates a backup copy of the XML configuration before writing changes.[1]
- Scans a local audio library recursively and builds an internal file index.[1]
- Normalizes audio filenames using spaces, hyphens, underscores, or no rename mode.[1]
- Optionally converts filenames to lowercase.[1]
- Tracks rename history and stores operation checkpoints.[1]
- Supports reverting file rename operations and related config path changes.[1]
- Provides a Tkinter-based desktop GUI with logs, history, and missing-file reporting.[1]
- Recognizes multiple audio formats including MP3, WAV, FLAC, OGG, M4A, AAC, WMA, OPUS, AIFF, APE, and more.[1]

## How it works

The repair workflow follows a straightforward sequence. The application first validates the input paths, optionally renames files in the selected library, builds an index of available audio files, creates a backup of the XML config, and then scans XML text nodes and attributes for supported audio file paths that can be repaired.[1]

Matching is not limited to a single exact lookup. The resolver can generate filename variants, try repaired path candidates, apply rename mappings from the current session, and fall back to indexed filename matching when direct file existence checks do not succeed.[1]

If the tool finds exactly one reliable match, it updates the path automatically. If multiple possible matches exist, the path is marked as ambiguous instead of being changed blindly, which reduces the risk of corrupting a working soundboard setup.[1]

## Project structure

```text
main.py
core/
    __init__.py
    constants.py
    models.py
    settings_manager.py
    history_manager.py
    filename_normalizer.py
    path_resolver.py
    audio_indexer.py
    config_repair_service.py
    revert_service.py
    service.py
gui/
    __init__.py
    app.py
    dialogs.py
    theme.py
    constants.py
errors/
    __init__.py
    base.py
    validation_errors.py
    config_errors.py
    settings_errors.py
    history_errors.py
```

The project is split into clear layers. GUI code lives in `gui/`, the core application logic lives in `core/`, and exception types are grouped under `errors/`, which keeps the codebase easier to maintain and extend.[1]

## Core modules

| Module | Responsibility |
|---|---|
| `core/service.py` | Coordinates the main repair and normalization workflow.[1] |
| `core/config_repair_service.py` | Validates input, creates backups, reads and writes XML, and updates broken paths.[1] |
| `core/audio_indexer.py` | Recursively scans the audio library and builds the filename index.[1] |
| `core/path_resolver.py` | Generates path candidates and resolves possible file matches.[1] |
| `core/filename_normalizer.py` | Normalizes audio filenames and generates name variants.[1] |
| `core/history_manager.py` | Stores rename history, checkpoints, and operation metadata.[1] |
| `core/revert_service.py` | Reverts file rename operations and restores related XML paths.[1] |
| `gui/app.py` | Provides the main desktop user interface.[1] |

## Supported audio formats

The application treats the following extensions as audio files:[1]

- `.mp3`
- `.wav`
- `.flac`
- `.ogg`
- `.oga`
- `.m4a`
- `.aac`
- `.wma`
- `.opus`
- `.aiff`
- `.aif`
- `.ape`
- `.alac`
- `.mp2`
- `.mpga`
- `.m4b`

## Main workflow

1. Select the local audio library folder.
2. Select the DCSB `config.xml` file.
3. Choose a rename mode, if needed.
4. Optionally enable lowercase conversion.
5. Run the repair process.
6. Review the log output, missing files, and operation history.
7. Create checkpoints or revert to a previous state when necessary.[1]

## History and revert system

The application stores operation history in a JSON file inside the user's application data directory. The history system supports rename entries, grouped operations, checkpoints, and log records, and it is exposed in the GUI so changes can be reviewed before a revert is performed.[1]

When a revert is triggered, the tool attempts to restore old filenames and also update the corresponding XML paths for any affected configuration files. This makes the project useful not only as a repair tool, but also as a safer batch-renaming companion for a live soundboard setup.[1]

## Filename normalization

The built-in normalization logic can clean up separators, simplify repeated punctuation, normalize selected patterns such as `feat.` and `vs.`, remove specific unwanted fragments, and apply a consistent naming style across the library. That is especially useful when the same files are expected to work cleanly in both Winamp playlists and DCSB soundboard references.[1]

## Playlist Builder

Within this project, the Playlist Builder should export playlists as **`.m3u8`** files so playlist data is stored in UTF-8 and remains safer for filenames that contain non-ASCII characters or special symbols. That format is the better fit for modern local-library use in Winamp.[1]

## FLAC -> MP3 tool

The project also includes a FLAC -> MP3 workflow as a helper feature for preparing audio files for broader playback compatibility. The exact conversion behavior depends on the implementation of that module, but its purpose is to make source audio easier to reuse across the rest of the toolchain.[1]

## Requirements

- Windows
- Python 3.11 or newer
- A local audio library
- A valid DCSB `config.xml` file
- Optional: `ffmpeg` or another converter backend if the FLAC -> MP3 module depends on an external encoder

## Installation

```bash
git clone <repo-url>
cd .DCSB_Config_Path_Fixer
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

If the repository does not include a `requirements.txt`, the main desktop application still relies heavily on Python's standard library, especially `tkinter`. Extra dependencies may apply only to optional modules such as audio conversion.[1]

## Running the app

```bash
python main.py
```

After launch, the GUI allows selecting the library folder, choosing the XML config, setting rename options, and reviewing progress through logs, status indicators, missing-file output, and the built-in operation history panel.[1]

## Safety

Before writing any repaired XML back to disk, the application creates a backup copy of the config file. Combined with checkpoints and revert support, this gives the project a safer workflow for batch maintenance on live libraries used by Winamp and DCSB.[1]

## Roadmap

- Refine the Playlist Builder with full `.m3u8` export behavior
- Finalize and document the FLAC -> MP3 module in more detail
- Improve collision handling during rename operations
- Expand examples and end-user documentation

## Author note

This tool was created for practical use with **Winamp** and **DCSB — Deathcounter and Soundboard**.[1]
