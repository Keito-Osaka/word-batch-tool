"""Font-preserving replacement improvements loaded automatically by Python.

This module patches the Ver.1.3.1 replacement routine without using Word COM.
"""
import re
import warnings

try:
    import core as _core
    from docx.oxml.ns import qn
except Exception:
    _core = None

_warned_fields = set()


def _direct_font(run, preferred_script=None):
    """Return a directly specified font from a run, if available."""
    rpr = run._element.rPr
    rfonts = rpr.rFonts if rpr is not None else None
    if rfonts is None:
        return None
    if preferred_script == "eastAsia":
        keys = ("eastAsia", "ascii", "hAnsi", "cs")
    elif preferred_script in ("ascii", "hAnsi"):
        keys = ("ascii", "hAnsi", "eastAsia", "cs")
    else:
        keys = ("eastAsia", "ascii", "hAnsi", "cs")
    for key in keys:
        value = rfonts.get(qn("w:" + key))
        if value:
            return value
    # run.font.name can expose a directly assigned w:rFonts value in some
    # python-docx versions, but must not be treated as a hard-coded fallback.
    return run.font.name or None


def _script_for_char(char):
    if not char:
        return None
    code = ord(char)
    if (0x3000 <= code <= 0x9FFF or 0x3040 <= code <= 0x30FF
            or 0xAC00 <= code <= 0xD7AF):
        return "eastAsia"
    if code < 0x10000:
        return "ascii"
    return None


def _placeholder_font(paragraph, start, end, mapping, original_text, field):
    """Use first placeholder character, delimiters, then nearby paragraph runs."""
    # The first ordinary placeholder character is authoritative. In particular,
    # do not use a later character (e.g. 額) when 金 has no direct formatting.
    first_run = None
    first_char = ""
    delimiter_runs = []
    for pos in range(start, end):
        index = mapping[pos]
        char = original_text[pos]
        if char in "{}<>":
            if index not in delimiter_runs:
                delimiter_runs.append(index)
        elif first_run is None:
            first_run = index
            first_char = char

    if first_run is not None:
        font = _direct_font(paragraph.runs[first_run], _script_for_char(first_char))
        if font:
            return font

    # If the placeholder text has no font, explicitly formatted {{ / }} (or
    # << / >>) is the requested second-level fallback.
    for index in delimiter_runs:
        font = _direct_font(paragraph.runs[index])
        if font:
            return font

    # Finally search the nearest runs in the same paragraph. Prefer the run
    # immediately before and after the placeholder, then expand outward.
    run_count = len(paragraph.runs)
    candidates = []
    before = mapping[start - 1] if start > 0 and mapping else None
    after = mapping[end] if end < len(mapping) else None
    for index in (before, after):
        if index is not None and index not in candidates:
            candidates.append(index)
    if first_run is not None:
        for distance in range(1, run_count + 1):
            for index in (first_run - distance, first_run + distance):
                if 0 <= index < run_count and index not in candidates:
                    candidates.append(index)
    for index in candidates:
        font = _direct_font(paragraph.runs[index])
        if font:
            return font

    if field not in _warned_fields:
        _warned_fields.add(field)
        warnings.warn(
            "一部フォント情報を取得できなかったため、ＭＳ明朝を使用しました。"
            f" 対象: {field}", RuntimeWarning, stacklevel=3)
    return "ＭＳ明朝"


def _replace_text_in_paragraph(paragraph, replacements):
    if not paragraph.runs:
        return
    original_text = "".join(run.text for run in paragraph.runs)
    if not original_text:
        return
    mapping = []
    for run_index, run in enumerate(paragraph.runs):
        mapping.extend([run_index] * len(run.text))

    groups = [("{{", "}}", replacements, "row")]
    common = getattr(paragraph, "_word_batch_common_replacements", None)
    if common:
        groups.append(("<<", ">>", common, "common"))
    matches = []
    for opener, closer, values, kind in groups:
        for key, value in values.items():
            token = f"{opener}{key}{closer}"
            cursor = 0
            while True:
                pos = original_text.find(token, cursor)
                if pos < 0:
                    break
                matches.append({"start": pos, "end": pos + len(token),
                                "value": "" if value is None else str(value),
                                "field": token})
                cursor = pos + len(token)
    if not matches:
        return
    matches.sort(key=lambda item: item["start"])
    selected, last_end = [], -1
    for match in matches:
        if match["start"] >= last_end:
            selected.append(match)
            last_end = match["end"]

    segments = []
    cursor = 0
    for match in selected:
        if cursor < match["start"]:
            pos = cursor
            while pos < match["start"]:
                index = mapping[pos]
                end = pos + 1
                while end < match["start"] and mapping[end] == index:
                    end += 1
                segments.append({"text": original_text[pos:end],
                                 "source": index, "replacement": False})
                pos = end
        source = _placeholder_source(paragraph, match["start"], match["end"], mapping, original_text)
        font = _placeholder_font(paragraph, match["start"], match["end"], mapping,
                                 original_text, match["field"])
        segments.append({"text": match["value"], "source": source,
                         "replacement": True, "font": font})
        cursor = match["end"]
    if cursor < len(original_text):
        pos = cursor
        while pos < len(original_text):
            index = mapping[pos]
            end = pos + 1
            while end < len(original_text) and mapping[end] == index:
                end += 1
            segments.append({"text": original_text[pos:end], "source": index,
                             "replacement": False})
            pos = end

    existing = list(paragraph.runs)
    for run in existing:
        run.text = ""
    for index, segment in enumerate(segments):
        target = existing[index] if index < len(existing) else paragraph.add_run()
        source = existing[segment["source"]]
        _core.copy_run_format(source, target)
        if segment["replacement"]:
            _core.set_run_font_all(target, segment["font"])
        target.text = segment["text"]


def _placeholder_source(paragraph, start, end, mapping, text):
    for pos in range(start, end):
        if text[pos] not in "{}<>":
            return mapping[pos]
    return mapping[start] if start < len(mapping) else 0


if _core is not None:
    _core.replace_text_in_paragraph = _replace_text_in_paragraph
