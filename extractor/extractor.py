"""
TPL Extractor — Adapter
=======================
Thin bridge between app.py (Streamlit UI) and extractor/engine/core.py.

This file contains NO extraction logic.
To upgrade extraction: modify engine/core.py only.
To upgrade UI:         modify app.py only.
This file never needs to change unless the contract between them changes.

Contract
--------
app.py calls:
    result = run_extraction(zip_paths, progress_cb)

Returns ExtractionResult with:
    rows          : list[dict]  — one dict per submittal
    total         : int
    with_comments : int
    no_comments   : int
    errors        : list[str]
    debug_rows    : list[dict]  — all raw annotation data (for Debug sheet)

Each row dict:
    sr_no             : int
    submittal         : str
    document          : str
    customer_comments : str
    xylem_remarks     : str
"""

import os
import re
from dataclasses import dataclass, field
# engine imported lazily inside run_extraction() so the package loads
# even before PyMuPDF is installed (gives a clean error message instead).


@dataclass
class ExtractionResult:
    rows:          list = field(default_factory=list)
    total:         int  = 0
    with_comments: int  = 0
    no_comments:   int  = 0
    errors:        list = field(default_factory=list)
    debug_rows:    list = field(default_factory=list)   # passed to build_excel debug sheet


def run_extraction(zip_paths: list, progress_cb=None) -> ExtractionResult:
    """
    Called by app.py. Iterates zip_paths, calls process_zip(), returns results.
    Deduplicates by SUB number (e.g. SUB1448 appearing in two ZIPs → keep first).
    """
    try:
        from .engine.core import process_zip, extract_sub_number
    except ImportError as e:
        raise ImportError(
            f"Extraction engine dependency missing: {e}\n"
            "Run:  pip install PyMuPDF"
        ) from e

    result = ExtractionResult()

    # ── Deduplicate by SUB number (matches original run() logic) ─────────────
    seen_subs  = set()
    unique_paths = []
    for p in zip_paths:
        sub = extract_sub_number(os.path.basename(p))
        if sub not in seen_subs:
            seen_subs.add(sub)
            unique_paths.append(p)

    total = len(unique_paths)
    sr    = 0

    for idx, zip_path in enumerate(unique_paths):
        zip_name = os.path.basename(zip_path)
        sub      = extract_sub_number(zip_name)

        if progress_cb:
            progress_cb(idx, total, f"Processing {sub} …")

        try:
            raw = process_zip(zip_path)        # ← only call into the engine
            sr += 1

            # Stamp sub onto each debug row (same as original run())
            for d in raw['debug']:
                d['sub'] = sub
            result.debug_rows.extend(raw['debug'])

            has_comments = bool(raw['comments'])

            if has_comments:
                # RENUMBER comments sequentially (1, 2, 3, etc.) - ignore original numbering
                renumbered = []
                for comment in raw['comments']:
                    # Strip original numbering prefix (e.g., "10. " or "9. " from different pages)
                    clean = re.sub(r'^\s*\d+[\.\)]\s+', '', comment).strip()
                    if clean:
                        renumbered.append(clean)
                
                # Now renumber sequentially
                final_comments = [f"{i+1}. {c}" for i, c in enumerate(renumbered)]
                comments_text = "\n".join(final_comments)
            else:
                comments_text = ''

            result.rows.append({
                'sr_no':             sr,
                'submittal':         sub,
                'document':          raw['doc_name'] or '',
                'customer_comments': comments_text,
                'xylem_remarks':     '' if has_comments else 'Comment not Received',
            })

            if has_comments:
                result.with_comments += 1
            else:
                result.no_comments += 1

        except Exception as exc:
            result.errors.append(f"{zip_name}: {exc}")
            result.rows.append({
                'sr_no':             (sr + 1),
                'submittal':         sub,
                'document':          '',
                'customer_comments': '',
                'xylem_remarks':     f'Extraction error: {exc}',
            })

    result.total = total

    if progress_cb:
        progress_cb(total, total, "Complete")

    return result
