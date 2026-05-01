"""Translate Chinese DOCX documents to English while preserving layout as much as possible.

Workflow:
1. Unpack DOCX using bundled office helpers.
2. Translate visible text in selected XML parts.
3. Repack and validate output.

This first version intentionally skips charts, formulas, field codes, and image text.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from urllib import error as urllib_error
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"w": W_NS}
ET.register_namespace("w", W_NS)
ET.register_namespace("mc", MC_NS)

TRANSLATE_API_BASE = "https://api.longcat.chat/openai"
TRANSLATE_API_KEY = "dummy"
TRANSLATE_MODEL = "LongCat-Flash-Thinking-2601"
DEFAULT_OUTPUT_SUFFIX = "en"
TRANSLATION_BATCH_SIZE = 8

ROOT_NAMESPACE_HINTS = {
    "word/document.xml": {
        "xmlns:ns2": "http://schemas.microsoft.com/office/word/2010/wordml",
        "xmlns:wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        "xmlns:cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
        "xmlns:cx1": "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
        "xmlns:cx2": "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
        "xmlns:cx3": "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
        "xmlns:cx4": "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
        "xmlns:cx5": "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
        "xmlns:cx6": "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
        "xmlns:cx7": "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
        "xmlns:cx8": "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
        "xmlns:aink": "http://schemas.microsoft.com/office/drawing/2016/ink",
        "xmlns:am3d": "http://schemas.microsoft.com/office/drawing/2017/model3d",
        "xmlns:o": "urn:schemas-microsoft-com:office:office",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        "xmlns:v": "urn:schemas-microsoft-com:vml",
        "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        "xmlns:w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
        "xmlns:w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
        "xmlns:w16": "http://schemas.microsoft.com/office/word/2018/wordml",
        "xmlns:w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
        "xmlns:w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
        "xmlns:wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        "xmlns:wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "xmlns:w10": "urn:schemas-microsoft-com:office:word",
        "xmlns:wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        "xmlns:wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
        "xmlns:wne": "http://schemas.microsoft.com/office/word/2006/wordml",
        "xmlns:wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    },
    "word/fontTable.xml": {
        "xmlns:ns2": "http://schemas.microsoft.com/office/word/2010/wordml",
        "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        "xmlns:w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
        "xmlns:w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
        "xmlns:w16": "http://schemas.microsoft.com/office/word/2018/wordml",
        "xmlns:w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
        "xmlns:w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    },
    "word/styles.xml": {
        "xmlns:ns2": "http://schemas.microsoft.com/office/word/2010/wordml",
        "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        "xmlns:w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
        "xmlns:w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
        "xmlns:w16": "http://schemas.microsoft.com/office/word/2018/wordml",
        "xmlns:w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
        "xmlns:w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    },
}

SKIP_TAGS = {
    f"{{{W_NS}}}instrText",
    f"{{{W_NS}}}delText",
}

TEXT_CONTAINER_TAGS = {
    f"{{{W_NS}}}p",
}

PART_PATTERNS = [
    "word/document.xml",
    "word/styles.xml",
    "word/fontTable.xml",
    "word/comments.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
    "word/header*.xml",
    "word/footer*.xml",
]

CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")
WS_RE = re.compile(r"\s+")


@dataclass
class TextNode:
    element: ET.Element
    text: str


@dataclass(frozen=True)
class HelperScripts:
    unpack: Path
    pack: Path
    validate: Path


def has_cjk(text: str) -> bool:
    return bool(text and CJK_RE.search(text))


def normalize_ws(text: str) -> str:
    return WS_RE.sub(" ", text).strip()


def decode_subprocess_output(payload: bytes | None) -> str:
    return (payload or b"").decode("utf-8", errors="replace").strip()


def infer_output_path(input_path: Path, output_path: Path | None) -> Path:
    if output_path:
        return output_path
    return input_path.with_name(f"{input_path.stem}_{DEFAULT_OUTPUT_SUFFIX}{input_path.suffix}")


def run_python(script: Path, *args: str) -> None:
    cmd = ["python", str(script), *map(str, args)]
    env = os.environ.copy()
    env.setdefault("PYTHONUTF8", "1")
    result = subprocess.run(cmd, capture_output=True, text=False, check=False, env=env)
    if result.returncode != 0:
        stderr = decode_subprocess_output(result.stderr)
        stdout = decode_subprocess_output(result.stdout)
        raise RuntimeError(stderr or stdout or f"Command failed: {cmd}")


def collect_parts(unpacked_dir: Path) -> list[Path]:
    return [
        part
        for pattern in PART_PATTERNS
        for part in sorted(unpacked_dir.glob(pattern))
    ]


def is_inside_skipped_ancestor(element: ET.Element) -> bool:
    current = element
    while current is not None:
        if current.tag in SKIP_TAGS:
            return True
        current = PARENT_MAP.get(id(current))
    return False


def build_parent_map(root: ET.Element) -> dict[int, ET.Element | None]:
    parent_map: dict[int, ET.Element | None] = {id(root): None}
    for parent in root.iter():
        for child in list(parent):
            parent_map[id(child)] = parent
    return parent_map


def iter_paragraphs(root: ET.Element):
    yield from root.iterfind(f".//{{{W_NS}}}p")


def collect_text_nodes(paragraph: ET.Element) -> list[TextNode]:
    nodes: list[TextNode] = []
    for elem in paragraph.iter():
        if elem.tag != f"{{{W_NS}}}t":
            continue
        if is_inside_skipped_ancestor(elem):
            continue
        text = elem.text or ""
        if text:
            nodes.append(TextNode(elem, text))
    return nodes


def build_translation_prompt(texts: list[str], target_lang: str = "English") -> str:
    payload = json.dumps(texts, ensure_ascii=False)
    return (
        f"Translate each Chinese string in the JSON array into natural professional {target_lang}. "
        "Preserve meaning, numbers, names, list markers, acronyms, and inline punctuation where possible. "
        "Return only valid JSON with the exact shape {\"translations\": [\"...\"]} and keep the same number of items. "
        "Do not wrap the JSON in markdown fences. Do not add explanations or notes.\n\n"
        f"Input JSON:\n{payload}"
    )


def extract_json_object(text: str) -> dict:
    text = text.strip()
    if not text:
        raise RuntimeError("Translation response was empty")

    try:
        body = json.loads(text)
        if isinstance(body, dict):
            return body
    except json.JSONDecodeError:
        pass

    fence_match = re.search(r"```(?:json)?\s*(\{[\s\S]*?\})\s*```", text)
    if fence_match:
        candidate = fence_match.group(1)
        try:
            body = json.loads(candidate)
            if isinstance(body, dict):
                return body
        except json.JSONDecodeError:
            pass

    decoder = json.JSONDecoder()
    for match in re.finditer(r"\{", text):
        try:
            body, _ = decoder.raw_decode(text[match.start():])
            if isinstance(body, dict):
                return body
        except json.JSONDecodeError:
            continue

    preview = text[:500].replace("\r", "\\r").replace("\n", "\\n")
    raise RuntimeError(f"Could not extract JSON object from translation response: {preview}")


def extract_claude_result_text(stdout: str) -> str:
    text = stdout.strip()
    if not text:
        return text
    try:
        body = json.loads(text)
        if isinstance(body, dict) and "result" in body and isinstance(body["result"], str):
            return body["result"]
    except json.JSONDecodeError:
        pass
    return text


def validate_translation_payload(body: dict, expected_count: int, source: str) -> list[str]:
    translations = body.get("translations")
    if not isinstance(translations, list) or len(translations) != expected_count:
        preview = json.dumps(body, ensure_ascii=False)[:500]
        raise RuntimeError(f"{source} returned malformed translation payload: {preview}")
    if not all(isinstance(item, str) and item.strip() for item in translations):
        raise RuntimeError(f"{source} returned empty translation text")
    return translations


def translate_with_claude_cli(texts: list[str], target_lang: str = "English") -> list[str]:
    if not shutil.which("claude"):
        raise RuntimeError("Claude CLI is not available on PATH.")

    cmd = [
        "claude",
        "-p",
        build_translation_prompt(texts, target_lang),
        "--output-format",
        "json",
        "--tools",
        "",
    ]
    env = {key: value for key, value in os.environ.items() if key != "CLAUDECODE"}
    result = subprocess.run(cmd, capture_output=True, text=False, check=False, env=env)
    stdout = (result.stdout or b"").decode("utf-8", errors="replace").strip()
    stderr = (result.stderr or b"").decode("utf-8", errors="replace").strip()
    if result.returncode != 0:
        raise RuntimeError(stderr or stdout or "Claude CLI translation failed")

    stdout = extract_claude_result_text(stdout)
    body = extract_json_object(stdout)
    return validate_translation_payload(body, len(texts), "Claude CLI")


def translate_with_api(texts: list[str], target_lang: str = "English") -> list[str]:
    api_key = os.getenv("LONGCAT_API_KEY", TRANSLATE_API_KEY)
    api_base = os.getenv("LONGCAT_API_BASE", TRANSLATE_API_BASE).rstrip("/")
    model = os.getenv("LONGCAT_MODEL", TRANSLATE_MODEL)

    import urllib.request

    payload = {
        "model": model,
        "max_tokens": 4000,
        "temperature": 0,
        "messages": [
            {
                "role": "user",
                "content": build_translation_prompt(texts, target_lang),
            }
        ],
    }

    req = urllib.request.Request(
        f"{api_base}/v1/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "content-type": "application/json",
            "authorization": f"Bearer {api_key}",
        },
        method="POST",
    )
    with urllib.request.urlopen(req) as resp:
        body = json.loads(resp.read().decode("utf-8"))
    choices = body.get("choices") or []
    message = choices[0].get("message", {}) if choices else {}
    response_text = (message.get("content") or "").strip()
    if not response_text:
        raise RuntimeError("Configured translation API returned empty text")

    body = extract_json_object(response_text)
    return validate_translation_payload(body, len(texts), "Configured translation API")


def translate_text(text: str, target_lang: str = "English") -> str:
    translations = translate_texts([text], target_lang)
    return translations[0]


def translate_texts(texts: list[str], target_lang: str = "English") -> list[str]:
    if not texts:
        return []
    try:
        return translate_with_api(texts, target_lang)
    except Exception:
        if shutil.which("claude"):
            return translate_with_claude_cli(texts, target_lang)
        raise



def redistribute_translation(nodes: list[TextNode], translated: str) -> None:
    if not nodes:
        return

    source_lengths = [max(len(node.text), 1) for node in nodes]
    total = sum(source_lengths)
    if len(nodes) == 1 or total == 0:
        write_text(nodes[0].element, translated)
        for node in nodes[1:]:
            write_text(node.element, "")
        return

    words = translated.split(" ")
    if len(words) < len(nodes):
        write_text(nodes[0].element, translated)
        for node in nodes[1:]:
            write_text(node.element, "")
        return

    allocations: list[str] = []
    cursor = 0
    for index, length in enumerate(source_lengths):
        if index == len(nodes) - 1:
            chunk_words = words[cursor:]
        else:
            target_count = max(1, round(len(words) * length / total))
            remaining_slots = len(nodes) - index - 1
            max_take = len(words) - cursor - remaining_slots
            take = min(max(target_count, 1), max_take)
            chunk_words = words[cursor: cursor + take]
            cursor += take
        allocations.append(" ".join(chunk_words).strip())

    if not any(allocations):
        allocations = [translated] + [""] * (len(nodes) - 1)

    for node, chunk in zip(nodes, allocations):
        write_text(node.element, chunk)


def write_text(element: ET.Element, text: str) -> None:
    element.text = text
    if text[:1].isspace() or text[-1:].isspace():
        element.set(f"{{{XML_NS}}}space", "preserve")
    else:
        element.attrib.pop(f"{{{XML_NS}}}space", None)


def resolve_helper_scripts() -> HelperScripts:
    search_roots = [
        Path(__file__).resolve().parent,
        Path.home() / ".claude" / "plugins" / "cache" / "anthropic-agent-skills",
    ]

    for root in search_roots:
        if not root.exists():
            continue
        unpack_candidates = list(root.glob("**/skills/docx/scripts/office/unpack.py"))
        for unpack_script in unpack_candidates:
            office_dir = unpack_script.parent
            pack_script = office_dir / "pack.py"
            validate_script = office_dir / "validate.py"
            if pack_script.exists() and validate_script.exists():
                return HelperScripts(unpack=unpack_script, pack=pack_script, validate=validate_script)

    raise FileNotFoundError(
        "Could not locate docx office helper scripts (unpack.py / pack.py / validate.py). "
        "Expected them under the local skill tree or ~/.claude/plugins/cache/anthropic-agent-skills/."
    )


def ensure_runtime_requirements(input_file: Path) -> None:
    if not input_file.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")
    if input_file.suffix.lower() != ".docx":
        raise ValueError(f"Input file must be .docx: {input_file}")
    if os.getenv("LONGCAT_API_KEY") or TRANSLATE_API_KEY:
        return
    if shutil.which("claude"):
        return
    raise RuntimeError("Neither configured translation API access nor Claude CLI is available for translation.")


def restore_root_namespace_hints(xml_path: Path) -> None:
    relative = xml_path.as_posix().split("/unpacked/")[-1]
    hints = ROOT_NAMESPACE_HINTS.get(relative)
    if not hints:
        return

    text = xml_path.read_text(encoding="utf-8")
    match = re.search(r"<(?!\?)([\w:.-]+)([^>]*)>", text, re.DOTALL)
    if not match:
        return

    start, end = match.span()
    tag_name = match.group(1)
    attrs = match.group(2)
    root_open = f"<{tag_name}{attrs}>"

    namespace_uri_to_prefix = {
        W_NS: "w",
        MC_NS: "mc",
        "http://schemas.microsoft.com/office/word/2010/wordml": "w14",
        "http://schemas.microsoft.com/office/word/2012/wordml": "w15",
        "http://schemas.microsoft.com/office/word/2018/wordml/cex": "w16cex",
        "http://schemas.microsoft.com/office/word/2016/wordml/cid": "w16cid",
        "http://schemas.microsoft.com/office/word/2018/wordml": "w16",
        "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash": "w16sdtdh",
        "http://schemas.microsoft.com/office/word/2015/wordml/symex": "w16se",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing": "wp14",
    }

    for uri, prefix in namespace_uri_to_prefix.items():
        attr = f'xmlns:{prefix}'
        if f'{attr}="' not in root_open and f'"{uri}"' in root_open:
            root_open = re.sub(rf'\sxmlns:(?!{prefix}\b)[^=]+="{re.escape(uri)}"', "", root_open, count=1)
            root_open = root_open[:-1] + f' {attr}="{uri}">'

    ignorable_tokens: list[str] = []
    for attr, value in hints.items():
        if f'{attr}="' not in root_open:
            root_open = root_open[:-1] + f' {attr}="{value}">'
        if attr.startswith("xmlns:"):
            ignorable_tokens.append(attr.split(":", 1)[1])

    existing_ignorable = re.search(r'mc:Ignorable="([^"]*)"', root_open)
    merged_ignorable: list[str] = []
    if existing_ignorable:
        merged_ignorable.extend(existing_ignorable.group(1).split())
    merged_ignorable.extend(token for token in ignorable_tokens if token not in merged_ignorable)
    ignorable_value = " ".join(merged_ignorable)

    if existing_ignorable:
        root_open = re.sub(r'mc:Ignorable="([^"]*)"', f'mc:Ignorable="{ignorable_value}"', root_open)
    else:
        root_open = root_open[:-1] + f' mc:Ignorable="{ignorable_value}">'

    xml_path.write_text(text[:start] + root_open + text[end:], encoding="utf-8")


def normalize_namespace_prefixes(xml_path: Path) -> None:
    relative = xml_path.as_posix().split("/unpacked/")[-1]
    prefix_map = {
        "word/document.xml": {
            "ns0": "w",
            "ns1": "mc",
            "ns2": "w14",
            "ns3": "m",
            "ns4": "v",
            "ns5": "o",
            "ns6": "r",
            "ns7": "wp",
            "ns8": "wp14",
            "ns9": "a",
            "ns10": "pic",
            "ns11": "a14",
        },
        "word/fontTable.xml": {
            "ns0": "w",
            "ns1": "mc",
            "ns2": "r",
            "ns3": "w14",
        },
        "word/styles.xml": {
            "ns0": "w",
            "ns1": "mc",
            "ns2": "r",
            "ns3": "w14",
        },
    }.get(relative)
    if not prefix_map:
        return

    text = xml_path.read_text(encoding="utf-8")
    for old, new in sorted(prefix_map.items(), key=lambda item: -len(item[0])):
        text = re.sub(rf"(?<![\w:-]){old}:", f"{new}:", text)
        text = text.replace(f'xmlns:{old}="', f'xmlns:{new}="')
    xml_path.write_text(text, encoding="utf-8")


def process_xml_part(xml_path: Path) -> tuple[int, int]:
    tree = ET.parse(xml_path)
    root = tree.getroot()

    global PARENT_MAP
    PARENT_MAP = build_parent_map(root)

    paragraph_count = 0
    text_count = 0

    pending_nodes: list[list[TextNode]] = []
    pending_texts: list[str] = []

    def flush_pending() -> None:
        nonlocal paragraph_count, text_count, pending_nodes, pending_texts
        if not pending_texts:
            return
        translations = translate_texts(pending_texts)
        for nodes, translated in zip(pending_nodes, translations):
            redistribute_translation(nodes, translated)
            paragraph_count += 1
            text_count += len(nodes)
        pending_nodes = []
        pending_texts = []

    for paragraph in iter_paragraphs(root):
        nodes = collect_text_nodes(paragraph)
        if not nodes:
            continue
        source_text = "".join(node.text for node in nodes)
        if not has_cjk(source_text):
            continue
        visible_text = normalize_ws(source_text)
        if not visible_text:
            continue
        pending_nodes.append(nodes)
        pending_texts.append(visible_text)
        if len(pending_texts) >= TRANSLATION_BATCH_SIZE:
            flush_pending()

    flush_pending()

    tree.write(xml_path, encoding="utf-8", xml_declaration=True)
    normalize_namespace_prefixes(xml_path)
    restore_root_namespace_hints(xml_path)
    return paragraph_count, text_count


def translate_docx(input_file: Path, output_file: Path | None = None) -> Path:
    ensure_runtime_requirements(input_file)
    output = infer_output_path(input_file, output_file)
    helpers = resolve_helper_scripts()

    with tempfile.TemporaryDirectory(prefix="docx_translate_") as temp_dir:
        temp_path = Path(temp_dir)
        unpacked_dir = temp_path / "unpacked"
        run_python(helpers.unpack, str(input_file), str(unpacked_dir))

        total_paragraphs = 0
        total_nodes = 0
        for part in collect_parts(unpacked_dir):
            paragraphs, nodes = process_xml_part(part)
            total_paragraphs += paragraphs
            total_nodes += nodes

        run_python(
            helpers.pack,
            str(unpacked_dir),
            str(output),
            "--original",
            str(input_file),
            "--validate",
            "false",
        )
        run_python(helpers.validate, str(unpacked_dir), "--original", str(input_file))

        print(f"Translated {total_paragraphs} paragraphs across {total_nodes} text nodes")

    return output


def main() -> None:
    parser = argparse.ArgumentParser(description="Translate Chinese DOCX to English")
    parser.add_argument("input_file", help="Input .docx file")
    parser.add_argument("output_file", nargs="?", help="Optional output .docx path")
    args = parser.parse_args()

    input_path = Path(args.input_file)
    output_path = Path(args.output_file) if args.output_file else None
    result = translate_docx(input_path, output_path)
    print(f"Output written to: {result}")


PARENT_MAP: dict[int, ET.Element | None] = {}


if __name__ == "__main__":
    main()
