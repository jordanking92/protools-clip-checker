#!/usr/bin/env python3
"""
Pro Tools Animation Session - Clip Name Checker
================================================
Called from Keyboard Maestro via Execute Shell Script.
All inputs come from KBM environment variables (KMVAR_*).

Required environment variables:
  KMVAR_EditType              "Full Edit" or "Select Edit"
  KMVAR_PerformanceLetters    "Yes" or "No"
                              (irrelevant for Select Edit - KBM sets to "Yes")
  KMVAR_ScriptCheck           "Yes" or "No"
  KMVAR_EpisodePathList       Multiline CSV, one episode per line:
                                102,/path/to/script.pdf,IN LINE
                                103,/path/to/script.pdf,NEXT LINE
                                106,/path/to/script.pdf,DISNEY ADR/PU
                                115,/path/to/sheet.csv,SPREADSHEET,EpCol,CharCol,LineCol
                              Episode numbers WITHOUT the EP prefix.
                              Formats: IN LINE, NEXT LINE, DISNEY ADR/PU, SPREADSHEET
  KMVAR_CharacterLookupList   Multiline CSV built by KBM, one character per line:
                                RegionName,ScriptName1[,ScriptName2,...]
                              The type (ALIAS/ACTOR) is already stripped by KBM.
                              Examples:
                                MOUSE,MOUSE
                                WHALE_inhome,WHALE
                                HULK_BRUCE,HULK,BRUCE
                                CITYGOERS_GREY,CITYGOERS
                                STEVE_CAPTAIN AMERICA,STEVE,CAPTAIN AMERICA

Session text files (always on Desktop):
  ~/Desktop/EPISODE_REGIONS.txt
  ~/Desktop/CHARACTER_REGIONS.txt
  ~/Desktop/CLIP_REGIONS.txt

Output:
  Human-readable report to stdout.
  KBM captures this into a variable for display.

PDF handling:
  .pdf script paths are auto-converted via pdftotext.
  Temp file written to /tmp, read, then deleted.
"""

import re
import os
import sys
import subprocess
from pathlib import Path
from dataclasses import dataclass
from typing import Optional


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

TRACK_ALL_TAKES = 'ALL TAKES'
TRACK_SELECTS   = 'SELECTS'
TRACK_UNKNOWN   = 'UNKNOWN'

PDFTOTEXT       = '/usr/local/bin/pdftotext'


# ---------------------------------------------------------------------------
# Line Number / Range Helpers
# ---------------------------------------------------------------------------

def normalize_range_separator(s: str) -> str:
    """Normalize ~ to - for range detection. They are interchangeable."""
    return s.replace('~', '-')


def is_range(line_number: str) -> bool:
    """
    True if line_number represents a range like 20-22 or 15A-17.
    Only applies to the line number portion (before _PREF/_ALT).
    """
    ln = normalize_range_separator(line_number)
    return bool(re.match(r'^\d+[A-Z]*-\d+[A-Z]*$', ln))


def expand_range(line_number: str) -> list:
    """
    Expand a range into individual line numbers.
    20-22  -> ['20', '21', '22']
    15A-17 -> ['15A', '16', '17']
    20~22  -> ['20', '21', '22']
    """
    ln = normalize_range_separator(line_number)
    m  = re.match(r'^(\d+)([A-Z]*)-(\d+)([A-Z]*)$', ln)
    if not m:
        return [line_number]

    start_num    = int(m.group(1))
    start_letter = m.group(2)
    end_num      = int(m.group(3))
    end_letter   = m.group(4)

    result = []
    for n in range(start_num, end_num + 1):
        if n == start_num and start_letter:
            result.append(f"{n}{start_letter}")
        elif n == end_num and end_letter:
            result.append(f"{n}{end_letter}")
        else:
            result.append(str(n))
    return result


def expand_line_number(line_number: str) -> list:
    """Return all individual line numbers a line_number string covers."""
    if is_range(line_number):
        return expand_range(line_number)
    return [line_number]


def normalize_line_number(ln: str) -> str:
    """Normalize for comparison: strip whitespace, uppercase."""
    return ln.strip().upper()


# ---------------------------------------------------------------------------
# Clip
# ---------------------------------------------------------------------------

@dataclass
class Clip:
    name:  str
    start: int
    end:   int
    track: str = TRACK_UNKNOWN

    @property
    def is_pref(self) -> bool:
        """
        True if this clip is a selected take (_PREF).
        Everything after _PREF is ignored (build notes, perf letters).
        _PREF followed by _ALT counts as ALT not PREF.
        50A_ALT_PREF -> True  (line number is 50A_ALT, this is its PREF)
        """
        if re.search(r'_PREF.*_ALT', self.name, re.IGNORECASE):
            return False
        return bool(re.search(r'_PREF', self.name, re.IGNORECASE))

    @property
    def is_alt(self) -> bool:
        """
        True if this clip is an alternate take (_ALT).
        _ALT alone does NOT count as a select. Only _PREF does.
        50A_ALT_PREF -> False (it's a PREF of line 50A_ALT)
        """
        if re.search(r'_PREF', self.name, re.IGNORECASE):
            return False
        return bool(re.search(r'_ALT', self.name, re.IGNORECASE))

    @property
    def is_select(self) -> bool:
        """Only _PREF counts as a select. _ALT alone does not."""
        return self.is_pref

    @property
    def line_number(self) -> Optional[str]:
        """
        Extract line number from clip name.

        Rules:
        - Everything before the first _PREF is the line number
        - Everything before the first _ALT is the line number (no _PREF)
        - After _PREF or _ALT, everything ignored (build notes, perf letters)
        - Raw ALL TAKES clips: strip trailing perf letter if present
            15a   -> 15       15A_b  -> 15A
            20~21 -> 20~21    83A    -> 83A

        Examples:
          15a              -> 15
          15A_b            -> 15A
          20~21            -> 20~21
          83A              -> 83A
          52AA_PREF        -> 52AA
          15_PREF          -> 15
          15A_PREF         -> 15A
          50A_ALT_PREF     -> 50A_ALT
          50A_ALT          -> 50A_ALT
          15_YELL_PREF     -> 15_YELL
          17_PREF_14f-13c  -> 17
          77~78_PREF       -> 77~78
          77~78_PREF_18b   -> 77~78
        """
        name = self.name

        # Strip _PREF and everything after
        m = re.match(r'^(.+?)_PREF.*$', name, re.IGNORECASE)
        if m:
            return m.group(1)

        # Strip _ALT and everything after
        m = re.match(r'^(.+?)_ALT.*$', name, re.IGNORECASE)
        if m:
            return m.group(1)

        # Raw ALL TAKES: strip underscore-separated perf letter: 15A_b -> 15A
        m = re.match(r'^(.+)_([a-z])$', name)
        if m:
            return m.group(1)

        # Raw ALL TAKES: strip directly attached lowercase letter: 15a -> 15
        m = re.match(r'^(\d+[A-Z]*)([a-z])$', name)
        if m:
            return m.group(1)

        # No stripping needed
        if re.match(r'^\d', name):
            return name

        return None

    @property
    def expanded_line_numbers(self) -> list:
        """All individual line numbers this clip covers, normalized."""
        ln = self.line_number
        if ln is None:
            return []
        return [normalize_line_number(x) for x in expand_line_number(ln)]


# ---------------------------------------------------------------------------
# Region
# ---------------------------------------------------------------------------

@dataclass
class Region:
    episode:   str
    character: str
    windows:   list   # list of (start, end) tuples

    def contains(self, clip: Clip) -> bool:
        for (start, end) in self.windows:
            if clip.start >= start and clip.end <= end:
                return True
        return False


# ---------------------------------------------------------------------------
# Session Text Parsing
# ---------------------------------------------------------------------------

def parse_session_text(filepath: Path) -> list:
    """
    Parse a Pro Tools session text export into Clip objects.
    Handles multiple tracks separated by TRACK NAME: headers.
    Tags each clip with its track (ALL TAKES, SELECTS, or UNKNOWN).
    """
    clips           = []
    current_track   = TRACK_UNKNOWN
    in_data_section = False

    if not filepath.exists():
        print(f"WARNING: File not found: {filepath}")
        return clips

    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        for line in f:
            line = line.rstrip('\n')

            m = re.match(r'^TRACK NAME:\s*(.+)$', line, re.IGNORECASE)
            if m:
                track_name = m.group(1).strip().upper()
                if 'ALL TAKES' in track_name:
                    current_track = TRACK_ALL_TAKES
                elif 'SELECT' in track_name:
                    current_track = TRACK_SELECTS
                else:
                    current_track = TRACK_UNKNOWN
                in_data_section = False
                continue

            if re.match(r'^CHANNEL\s+EVENT', line, re.IGNORECASE):
                in_data_section = True
                continue

            if not in_data_section:
                continue

            parts = re.split(r'\t', line)
            parts = [p.strip() for p in parts if p.strip()]

            if len(parts) < 5:
                continue
            if parts[0] != '1':
                continue
            if not re.match(r'^\d+$', parts[1]):
                continue

            clip_name = parts[2]
            try:
                start = int(parts[3])
                end   = int(parts[4])
            except ValueError:
                continue

            clips.append(Clip(
                name=clip_name, start=start, end=end, track=current_track
            ))

    return clips


def normalize_episode_name(name: str) -> str:
    """
    Normalize episode name for matching.
    Strips EP prefix but preserves everything after including S, letters etc.
    EP102   -> 102
    EPS102  -> S102
    EP104A  -> 104A
    EP102 PU -> 102 PU
    """
    name = name.strip().upper()
    name = re.sub(r'^EP(?=\S)', '', name)
    return name


def build_regions(episode_clips: list, character_clips: list) -> list:
    """
    Match episode clips to character clips by time containment.
    Multiple windows for the same EP+Character pair are combined.
    """
    region_map = {}

    for char_clip in character_clips:
        matching_ep = None
        for ep_clip in episode_clips:
            if (char_clip.start >= ep_clip.start and
                    char_clip.end <= ep_clip.end):
                matching_ep = ep_clip
                break

        if matching_ep is None:
            continue

        ep_name   = normalize_episode_name(matching_ep.name)
        char_name = char_clip.name
        key       = (ep_name, char_name)
        window    = (char_clip.start, char_clip.end)

        if key in region_map:
            region_map[key].windows.append(window)
        else:
            region_map[key] = Region(
                episode=ep_name,
                character=char_name,
                windows=[window]
            )

    return list(region_map.values())


# ---------------------------------------------------------------------------
# PDF Conversion
# ---------------------------------------------------------------------------

def pdf_to_text(pdf_path: str) -> Optional[str]:
    """
    Convert a PDF to plain text using pdftotext.
    Writes to /tmp, reads it, deletes it.
    Returns text content or None if conversion fails.
    """
    tmp_path = f"/tmp/clip_check_script_{os.getpid()}.txt"
    try:
        result = subprocess.run(
            [PDFTOTEXT, '-layout', '-nodiag', '-enc', 'UTF-8',
             pdf_path, tmp_path],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            print(f"  WARNING: pdftotext failed for {pdf_path}: {result.stderr}")
            return None
        with open(tmp_path, 'r', encoding='utf-8', errors='replace') as f:
            return f.read()
    except Exception as e:
        print(f"  WARNING: Could not convert PDF {pdf_path}: {e}")
        return None
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


def load_script_text(script_path: str) -> Optional[str]:
    """
    Load script text from file. Handles PDFs via pdftotext.
    """
    path = Path(script_path)
    if not path.exists():
        print(f"  WARNING: Script file not found: {script_path}")
        return None

    if path.suffix.lower() == '.pdf':
        return pdf_to_text(script_path)
    else:
        try:
            with open(path, 'r', encoding='utf-8', errors='replace') as f:
                return f.read()
        except Exception as e:
            print(f"  WARNING: Could not read script {script_path}: {e}")
            return None


# ---------------------------------------------------------------------------
# Check 1: Misplaced Clips
# ---------------------------------------------------------------------------

def check_misplaced_clips(regions: list, all_clips: list) -> dict:
    """
    Flag clips on the wrong track:
      - _PREF on ALL TAKES (should be on SELECTS)
      - Non-select, non-ALT on SELECTS (should be on ALL TAKES)
    """
    results = {}

    for region in regions:
        region_clips = [c for c in all_clips if region.contains(c)]

        pref_on_all_takes = []
        takes_on_selects  = []

        for clip in region_clips:
            if clip.track == TRACK_ALL_TAKES and clip.is_pref:
                pref_on_all_takes.append(clip.name)
            elif clip.track == TRACK_SELECTS and not clip.is_pref and not clip.is_alt:
                takes_on_selects.append(clip.name)

        if pref_on_all_takes or takes_on_selects:
            key = f"{region.episode} — {region.character}"
            results[key] = {
                'pref_on_all_takes': sorted(pref_on_all_takes),
                'takes_on_selects':  sorted(takes_on_selects),
            }

    return results


# ---------------------------------------------------------------------------
# Check 2: Duplicate Clip Names
# ---------------------------------------------------------------------------

def check_duplicates(regions: list, all_clips: list,
                     performance_letters: bool) -> dict:
    """
    Find clips sharing the same name within a region.
    SELECTS: always checked.
    ALL TAKES: only checked if performance_letters=True.
    """
    results = {}

    for region in regions:
        region_clips = [c for c in all_clips if region.contains(c)]
        duplicates   = set()

        # Always check SELECTS
        name_counts = {}
        for clip in region_clips:
            if clip.track == TRACK_SELECTS:
                name_counts[clip.name] = name_counts.get(clip.name, 0) + 1
        for name, count in name_counts.items():
            if count > 1:
                duplicates.add(name)

        # Only check ALL TAKES if performance letters are used
        if performance_letters:
            name_counts = {}
            for clip in region_clips:
                if clip.track == TRACK_ALL_TAKES:
                    name_counts[clip.name] = name_counts.get(clip.name, 0) + 1
            for name, count in name_counts.items():
                if count > 1:
                    duplicates.add(name)

        if duplicates:
            key = f"{region.episode} — {region.character}"
            results[key] = sorted(duplicates)

    return results


# ---------------------------------------------------------------------------
# Check 3: Missing Selects
# ---------------------------------------------------------------------------

def check_missing_selects(regions: list, all_clips: list) -> dict:
    """
    Find line numbers on ALL TAKES with no _PREF on SELECTS.
    Full Edit mode only. _ALT alone does not satisfy the requirement.
    """
    results = {}

    for region in regions:
        region_clips = [c for c in all_clips if region.contains(c)]

        selected_lines = set()
        for clip in region_clips:
            if clip.track == TRACK_SELECTS and clip.is_pref:
                for ln in clip.expanded_line_numbers:
                    selected_lines.add(normalize_line_number(ln))

        recorded_lines = set()
        for clip in region_clips:
            # ALL TAKES clips (excluding misplaced _PREF/_ALT)
            if clip.track == TRACK_ALL_TAKES:
                if clip.is_pref or clip.is_alt:
                    continue
                for ln in clip.expanded_line_numbers:
                    recorded_lines.add(normalize_line_number(ln))
            # _ALT clips on SELECTS track also count as recorded
            # A line with only _ALT and no _PREF still needs a select
            elif clip.track == TRACK_SELECTS and clip.is_alt:
                for ln in clip.expanded_line_numbers:
                    recorded_lines.add(normalize_line_number(ln))

        missing = sorted(
            recorded_lines - selected_lines,
            key=lambda x: (
                int(re.match(r'\d+', x).group()) if re.match(r'\d+', x) else 0,
                x
            )
        )

        if missing:
            key = f"{region.episode} — {region.character}"
            results[key] = missing

    return results


# ---------------------------------------------------------------------------
# Script Parsing Helpers
# ---------------------------------------------------------------------------

def strip_revision_markers(text: str) -> str:
    """Strip revision markers like ****, *3****, ****3 etc."""
    text = re.sub(r'\*+\d*\*+', '', text)
    text = re.sub(r'\*+',       '', text)
    return text


def clean_script_text(text: str) -> str:
    """Normalize script text for parsing."""
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    text = strip_revision_markers(text)
    text = text.replace('#', '').replace("'", '')
    return text


def normalize_script_format(fmt: str) -> str:
    """
    Normalize script format string from KBM to internal format key.
    Accepts: IN LINE, NEXT LINE, DISNEY ADR/PU, SPREADSHEET
    Returns: inline, nextline, disneyadr, spreadsheet
    """
    fmt = fmt.strip().upper()
    if fmt == 'IN LINE':
        return 'inline'
    elif fmt == 'NEXT LINE':
        return 'nextline'
    elif fmt in ('DISNEY ADR/PU', 'DISNEY ADR', 'DISNEY PU'):
        return 'disneyadr'
    elif fmt == 'SPREADSHEET':
        return 'spreadsheet'
    else:
        return fmt.lower().replace(' ', '')


# ---------------------------------------------------------------------------
# Script Parsers
# ---------------------------------------------------------------------------

def parse_script_inline(script_text: str, character_name: str) -> list:
    """
    IN LINE format:
        2                        PIGGIE                                  2
    Line number left AND/OR right, character name in between.
    """
    script_text  = clean_script_text(script_text)
    char_pattern = re.escape(character_name)
    seen         = set()
    line_numbers = []

    pattern_left = re.compile(
        r'^\s*(\d+[A-Za-z]*(?:_[A-Za-z]+)*)\s+.*?' + char_pattern,
        re.IGNORECASE | re.MULTILINE
    )
    for m in pattern_left.finditer(script_text):
        ln = normalize_line_number(m.group(1))
        if ln not in seen:
            seen.add(ln)
            line_numbers.append(ln)

    pattern_right = re.compile(
        r'^\s*.*?' + char_pattern + r'.*?\s+(\d+[A-Za-z]*(?:_[A-Za-z]+)*)\s*$',
        re.IGNORECASE | re.MULTILINE
    )
    for m in pattern_right.finditer(script_text):
        ln = normalize_line_number(m.group(1))
        if ln not in seen:
            seen.add(ln)
            line_numbers.append(ln)

    return line_numbers


def parse_script_nextline(script_text: str, character_name: str) -> list:
    """
    NEXT LINE format:
                      ELEPHANT GERALD
        6         I was thinking...                          6
    Character name on its own line, line number on next line.
    """
    script_text  = clean_script_text(script_text)
    char_pattern = re.escape(character_name)
    seen         = set()
    line_numbers = []
    lines        = script_text.split('\n')

    i = 0
    while i < len(lines):
        line         = lines[i]
        char_match   = re.search(char_pattern, line, re.IGNORECASE)
        has_line_num = bool(re.search(r'^\s*\d', line))

        if char_match and not has_line_num:
            for j in range(i + 1, min(i + 4, len(lines))):
                next_line = lines[j]
                if not next_line.strip():
                    continue

                m = re.match(r'^\s*(\d+[A-Za-z]*(?:_[A-Za-z]+)*)\s+\S', next_line)
                if m:
                    ln = normalize_line_number(m.group(1))
                    if ln not in seen:
                        seen.add(ln)
                        line_numbers.append(ln)
                    break

                m = re.search(r'\s+(\d+[A-Za-z]*(?:_[A-Za-z]+)*)\s*$', next_line)
                if m:
                    ln = normalize_line_number(m.group(1))
                    if ln not in seen:
                        seen.add(ln)
                        line_numbers.append(ln)
                    break

                if re.match(r'^[A-Z\s&]+$', next_line.strip()):
                    break

        i += 1

    return line_numbers


def parse_script_disney_adr(script_text: str, character_name: str) -> list:
    """
    DISNEY ADR/PU format:
        LINE #6
        GRAMMA (P)
        (timecode)
            Dialogue
    """
    script_text  = clean_script_text(script_text)
    char_pattern = re.escape(character_name)
    seen         = set()
    line_numbers = []
    lines        = script_text.split('\n')

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m    = re.match(r'^LINE\s+(\d+[A-Z]*)', line, re.IGNORECASE)
        if m:
            ln_str = m.group(1)
            for j in range(i + 1, min(i + 4, len(lines))):
                next_line  = lines[j].strip()
                if not next_line:
                    continue
                next_clean = re.sub(r'\s*\([^)]*\)', '', next_line).strip()
                if re.search(char_pattern, next_clean, re.IGNORECASE):
                    ln = normalize_line_number(ln_str)
                    if ln not in seen:
                        seen.add(ln)
                        line_numbers.append(ln)
                break
        i += 1

    return line_numbers


def parse_script_spreadsheet(script_text: str, character_name: str,
                              episode_number: Optional[str] = None,
                              episode_col: int = 0,
                              character_col: int = 1,
                              line_number_col: int = 2) -> list:
    """
    SPREADSHEET (CSV) format. Column indices are 0-based.
    episode_col=0 means N/A (single episode spreadsheet).
    """
    seen         = set()
    line_numbers = []

    for row_line in script_text.split('\n'):
        row_line = row_line.strip()
        if not row_line:
            continue

        cols = re.split(r',(?=(?:[^"]*"[^"]*")*[^"]*$)', row_line)
        cols = [c.strip().strip('"') for c in cols]

        if len(cols) <= max(character_col, line_number_col):
            continue

        char_val = cols[character_col]
        ln_val   = cols[line_number_col].strip()

        if not ln_val:
            continue
        if not re.search(re.escape(character_name), char_val, re.IGNORECASE):
            continue

        if episode_col > 0 and episode_number:
            ep_val = cols[episode_col - 1] if episode_col - 1 < len(cols) else ''
            if episode_number not in ep_val:
                continue

        ln = normalize_line_number(ln_val)
        if ln not in seen:
            seen.add(ln)
            line_numbers.append(ln)

    return line_numbers


def parse_script_for_character(script_text: str, script_format: str,
                                character_name: str,
                                episode_number: Optional[str] = None,
                                spreadsheet_config: Optional[dict] = None
                                ) -> list:
    """Dispatch to the correct parser based on normalized format key."""
    if script_format == 'inline':
        return parse_script_inline(script_text, character_name)
    elif script_format == 'nextline':
        return parse_script_nextline(script_text, character_name)
    elif script_format == 'disneyadr':
        return parse_script_disney_adr(script_text, character_name)
    elif script_format == 'spreadsheet' and spreadsheet_config:
        return parse_script_spreadsheet(
            script_text, character_name, episode_number,
            spreadsheet_config.get('episode_col', 0),
            spreadsheet_config.get('character_col', 1),
            spreadsheet_config.get('line_number_col', 2)
        )
    else:
        return []


# ---------------------------------------------------------------------------
# Character Lookup
# ---------------------------------------------------------------------------

def parse_character_lookup_list(raw: str) -> dict:
    """
    Parse CharacterLookupList multiline CSV into a lookup dict.

    Each line: RegionName,ScriptName1[,ScriptName2,...]
    The ALIAS/ACTOR type has already been stripped by KBM before
    passing this variable to Python.

    Examples:
      MOUSE,MOUSE
      WHALE_inhome,WHALE
      HULK_BRUCE,HULK,BRUCE
      CITYGOERS_GREY,CITYGOERS
      STEVE_CAPTAIN AMERICA,STEVE,CAPTAIN AMERICA

    Returns:
      { region_name_upper: [script_name1, script_name2, ...] }
    """
    lookup = {}
    if not raw.strip():
        return lookup

    for line in raw.strip().split('\n'):
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in line.split(',')]
        if len(parts) < 2:
            continue
        region_name  = parts[0].upper()
        script_names = [p for p in parts[1:] if p]
        if script_names:
            lookup[region_name] = script_names

    return lookup


def resolve_script_names(character_name: str,
                         character_lookup: dict) -> list:
    """
    Given a character region name, return the list of script names to
    search for.

    Resolution order:
      1. Strip _inhome if present
      2. Check character_lookup (case-insensitive)
      3. If found -> return the script names from the lookup
      4. If not found and underscore present -> strip after last underscore
      5. If not found and no underscore -> return name as-is

    Returns a list of one or more script name strings.
    """
    # Step 1: always strip _inhome
    name = re.sub(r'_inhome$', '', character_name, flags=re.IGNORECASE).strip()

    # Step 2 & 3: check lookup
    if name.upper() in character_lookup:
        return character_lookup[name.upper()]

    # Step 4: underscore present -> strip after last underscore
    if '_' in name:
        stripped = name.rsplit('_', 1)[0].strip()
        return [stripped]

    # Step 5: no underscore -> use as-is
    return [name]


# ---------------------------------------------------------------------------
# Check 4: Script Check
# ---------------------------------------------------------------------------

def check_script(regions: list, all_clips: list,
                 episode_scripts: dict,
                 character_lookup: dict = None) -> dict:
    """
    Compare script line numbers against _PREF clips in session.

    episode_scripts:   { episode_number: (script_text, format, spreadsheet_config) }
    character_lookup:  { region_name_upper: [script_name1, script_name2, ...] }
                       Built from KMVAR_CharacterLookupList by KBM.
                       If None or empty, falls back to default name resolution.

    For alias characters (e.g. HULK_BRUCE -> HULK, BRUCE), the script is
    searched for ALL names and line numbers are combined.

    PU sessions fall back to base episode number if no exact match.
    """
    if character_lookup is None:
        character_lookup = {}

    results = {}

    for region in regions:
        ep_key = region.episode

        script_entry = episode_scripts.get(ep_key)
        if script_entry is None:
            ep_base      = re.sub(r'\s*(PU|ADR)$', '', ep_key,
                                  flags=re.IGNORECASE).strip()
            script_entry = episode_scripts.get(ep_base)
        if script_entry is None:
            continue

        script_text, script_format, spreadsheet_config = script_entry
        region_clips = [c for c in all_clips if region.contains(c)]

        selected_lines = set()
        for clip in region_clips:
            if clip.track == TRACK_SELECTS and clip.is_pref:
                for ln in clip.expanded_line_numbers:
                    selected_lines.add(normalize_line_number(ln))

        # Resolve which script names to search for this character region
        script_names = resolve_script_names(region.character, character_lookup)

        # Search script for ALL resolved names, combine results
        all_script_lines  = []
        seen_script_lines = set()
        for script_name in script_names:
            lines = parse_script_for_character(
                script_text=script_text,
                script_format=script_format,
                character_name=script_name,
                episode_number=ep_key,
                spreadsheet_config=spreadsheet_config
            )
            for ln in lines:
                norm = normalize_line_number(ln)
                if norm not in seen_script_lines:
                    seen_script_lines.add(norm)
                    all_script_lines.append(norm)

        if not all_script_lines:
            continue

        seen_missing = set()
        missing      = []
        for ln in all_script_lines:
            norm = normalize_line_number(ln)
            if norm not in selected_lines and norm not in seen_missing:
                seen_missing.add(norm)
                missing.append(norm)

        if missing:
            key = f"{region.episode} — {region.character}"
            results[key] = missing

    return results


# ---------------------------------------------------------------------------
# Parse KBM Inputs
# ---------------------------------------------------------------------------

def parse_episode_path_list(raw: str) -> dict:
    """
    Parse EpisodePathList multiline CSV into episode_scripts dict.

    Each line format:
      Non-spreadsheet: 102,/path/to/script.pdf,IN LINE
      Spreadsheet:     102,/path/to/sheet.csv,SPREADSHEET,EpCol,CharCol,LineCol

    Returns:
      { episode_number: (script_text, format_key, spreadsheet_config_or_None) }
    """
    episode_scripts = {}

    for line in raw.strip().split('\n'):
        line = line.strip()
        if not line:
            continue

        parts = line.split(',')
        if len(parts) < 3:
            print(f"  WARNING: Skipping malformed EpisodePathList entry: {line}")
            continue

        ep_num      = parts[0].strip().upper()
        script_path = parts[1].strip()
        fmt_raw     = parts[2].strip()
        fmt         = normalize_script_format(fmt_raw)

        spreadsheet_config = None
        if fmt == 'spreadsheet':
            if len(parts) >= 6:
                try:
                    spreadsheet_config = {
                        'episode_col':     int(parts[3].strip()),
                        'character_col':   int(parts[4].strip()) - 1,
                        'line_number_col': int(parts[5].strip()) - 1,
                    }
                except ValueError:
                    print(f"  WARNING: Bad spreadsheet columns for {ep_num}, skipping.")
                    continue
            else:
                print(f"  WARNING: SPREADSHEET entry missing column info for {ep_num}")
                continue

        script_text = load_script_text(script_path)
        if script_text is None:
            print(f"  WARNING: Could not load script for episode {ep_num} - file not found or unreadable.")
            continue

        episode_scripts[ep_num] = (script_text, fmt, spreadsheet_config)

    return episode_scripts


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

def format_report(mode: str,
                  performance_letters: bool,
                  misplaced_results: dict,
                  duplicate_results: dict,
                  missing_select_results: dict,
                  script_check_results: Optional[dict]) -> str:
    lines = []

    sep  = "=" * 56
    dash = "—" * 56

    lines.append(sep)
    lines.append("  PRO TOOLS CLIP NAME CHECK REPORT")
    lines.append(f"  Mode: {'Full Edit' if mode == 'full' else 'Select Edit'}")
    if mode == 'full':
        lines.append(
            f"  Performance letters on ALL TAKES: "
            f"{'Yes' if performance_letters else 'No'}"
        )
    lines.append(sep)
    lines.append("")

    # --- Misplaced Clips ---
    lines.append("MISPLACED CLIPS")
    lines.append(dash)
    if not misplaced_results:
        lines.append("SUCCESS: No misplaced clips found.")
    else:
        lines.append("ERROR: Misplaced clips found.")
        lines.append("")
        for region_key, data in sorted(misplaced_results.items()):
            lines.append(f"  {region_key}")
            if data['pref_on_all_takes']:
                lines.append("    _PREF clips on ALL TAKES track (should be on SELECTS):")
                for name in data['pref_on_all_takes']:
                    lines.append(f"      • {name}")
            if data['takes_on_selects']:
                lines.append("    Raw takes on SELECTS track (should be on ALL TAKES):")
                for name in data['takes_on_selects']:
                    lines.append(f"      • {name}")
            lines.append("")
    lines.append("")

    # --- Duplicates ---
    lines.append("DUPLICATE CLIP NAMES")
    lines.append(dash)
    if not duplicate_results:
        if mode == 'full' and not performance_letters:
            lines.append("SUCCESS: No duplicate clip names found on SELECTS track.")
            lines.append("(ALL TAKES duplicates not checked — no performance letters)")
        else:
            lines.append("SUCCESS: No duplicate clip names found.")
    else:
        lines.append("ERROR: Duplicate clip names found.")
        lines.append("")
        for region_key, dupes in sorted(duplicate_results.items()):
            lines.append(f"  {region_key}")
            for d in dupes:
                lines.append(f"    • {d}")
            lines.append("")
    lines.append("")

    # --- Missing Selects ---
    if mode == 'full':
        lines.append("LINES WITHOUT SELECTS")
        lines.append(dash)
        if not missing_select_results:
            lines.append("SUCCESS: All recorded lines have a _PREF select.")
        else:
            lines.append("ERROR: Lines found without a _PREF select.")
            lines.append("")
            for region_key, missing in sorted(missing_select_results.items()):
                lines.append(f"  {region_key}")
                for ln in missing:
                    lines.append(f"    • Line {ln}")
                lines.append("")
        lines.append("")

    # --- Script Check ---
    if script_check_results is not None:
        lines.append("SCRIPT CHECK")
        lines.append(dash)
        if not script_check_results:
            lines.append("SUCCESS: All script lines have a _PREF in the session.")
        else:
            lines.append("ERROR: Script lines found with no _PREF in session.")
            lines.append("")
            for region_key, missing in sorted(script_check_results.items()):
                lines.append(f"  {region_key}")
                for ln in missing:
                    lines.append(f"    • Line {ln}")
                lines.append("")
        lines.append("")

    lines.append(sep)
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    desktop = Path.home() / "Desktop"

    # --- Read KBM environment variables ---
    edit_type_raw    = os.environ.get('KMVAR_EditType', 'Full Edit').strip()
    perf_letters_raw = os.environ.get('KMVAR_PerformanceLetters', 'No').strip()
    script_check_raw = os.environ.get('KMVAR_ScriptCheck', 'No').strip()
    episode_path_raw = os.environ.get('KMVAR_EpisodePathList', '').strip()
    char_lookup_raw  = os.environ.get('KMVAR_CharacterLookupList', '').strip()

    mode                = 'full' if 'Full' in edit_type_raw else 'select'
    performance_letters = perf_letters_raw.upper() in ('YES', '1', 'TRUE')
    run_script_check    = script_check_raw.upper() in ('YES', '1', 'TRUE')

    # In Select Edit mode, performance_letters is irrelevant
    if mode == 'select':
        performance_letters = True

    # --- Load session text files ---
    episode_clips   = parse_session_text(desktop / "EPISODE_REGIONS.txt")
    character_clips = parse_session_text(desktop / "CHARACTER_REGIONS.txt")
    all_clips       = parse_session_text(desktop / "CLIP_REGIONS.txt")

    if not episode_clips or not character_clips or not all_clips:
        print("ERROR: One or more session text files could not be loaded.")
        print("Make sure EPISODE_REGIONS.txt, CHARACTER_REGIONS.txt, and")
        print("CLIP_REGIONS.txt are present on the Desktop.")
        sys.exit(1)

    # --- Build regions ---
    regions = build_regions(episode_clips, character_clips)

    if not regions:
        print("ERROR: Could not match any episode regions to character regions.")
        print("Check that EPISODE and CHARACTER clip time windows overlap.")
        sys.exit(1)

    # --- Run checks ---
    misplaced_results      = check_misplaced_clips(regions, all_clips)
    duplicate_results      = check_duplicates(regions, all_clips, performance_letters)
    missing_select_results = {}

    if mode == 'full':
        missing_select_results = check_missing_selects(regions, all_clips)

    # --- Parse character lookup ---
    character_lookup = parse_character_lookup_list(char_lookup_raw)

    # --- Script check ---
    script_check_results = None
    if run_script_check and episode_path_raw:
        episode_scripts      = parse_episode_path_list(episode_path_raw)
        script_check_results = check_script(
            regions, all_clips, episode_scripts, character_lookup
        )
    elif run_script_check and not episode_path_raw:
        print("WARNING: Script check requested but EpisodePathList is empty.")

    # --- Output report ---
    report = format_report(
        mode=mode,
        performance_letters=performance_letters,
        misplaced_results=misplaced_results,
        duplicate_results=duplicate_results,
        missing_select_results=missing_select_results,
        script_check_results=script_check_results
    )

    print(report)


if __name__ == '__main__':
    main()
