"""Deterministic artifact checks against the five papers supplied by preprocess.

Verify identities, source numerical anchors, requested structure and actual word
counts. These checks cannot judge arbitrary scientific interpretations or prove
that the agent read the papers; keyword/fact presence is not a semantic review.
"""
import json
import re
import sys
import unicodedata
from argparse import ArgumentParser
from datetime import datetime
from pathlib import Path

from docx import Document
from pptx import Presentation

CRITICAL_CHECKS = {
    "Word document readable", "All 6 heading groups found",
    "Summary table has exactly 5 data rows and 4 columns",
    "Summary rows match the five source papers",
    "Word discussion includes source benchmark results",
    "All required keywords present in Word doc", "Presentation readable",
    "PPTX has exactly 7 slides", "Required slide content present",
    "Word counts match extracted artifact text",
}
FAILED_NAMES = []


def _check(name, condition, fail_msg=""):
    if condition:
        print(f"  PASS: {name}")
        return 1, 1
    FAILED_NAMES.append(name)
    print(f"  FAIL: {name}{(': ' + fail_msg) if fail_msg else ''}")
    return 0, 1


def normalized(text):
    text = unicodedata.normalize("NFKC", str(text)).casefold().replace("’", "'")
    return " ".join(re.findall(r"[\w]+", text, flags=re.UNICODE))


def mentions(text, phrase):
    # Technical terms may use spaces, hyphens or typographic dashes.
    return f" {normalized(phrase)} " in f" {normalized(text)} "


def has_result(text, fact):
    if not mentions(text, fact["benchmark"]):
        return False
    numbers = re.findall(r"(?<![\w.])\d+(?:[.,]\d+)?(?![\w.])", str(text))
    value = float(fact["value"])
    # Accept percentages and equivalent fractional accuracies, decimal comma too.
    return any(abs(float(n.replace(",", ".")) - target) < 1e-6
               for n in numbers for target in (value, value / 100))


def word_text(doc):
    texts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            # Merged cells should not count the same XML text several times.
            seen = set()
            for cell in row.cells:
                if cell._tc not in seen:
                    texts.append(cell.text)
                    seen.add(cell._tc)
    return "\n".join(texts)


def slide_text(slide):
    texts = []
    def visit(shapes):
        for shape in shapes:
            if shape.has_text_frame:
                texts.extend(p.text for p in shape.text_frame.paragraphs)
            if shape.has_table:
                texts.extend(cell.text for row in shape.table.rows for cell in row.cells)
            if hasattr(shape, "shapes"):
                visit(shape.shapes)
    visit(slide.shapes)
    return "\n".join(texts)


def check_word_doc(agent_workspace, gt_data):
    checks = []
    path = Path(agent_workspace) / gt_data["review_doc"]["filename"]
    try:
        doc = Document(path)
    except Exception as exc:
        return _check("Word document readable", False, str(exc))
    checks.append(_check("Word document readable", True))
    headings = [p.text for p in doc.paragraphs if p.style and "Heading" in p.style.name]
    groups = gt_data["review_doc"]["required_headings"]
    heading_ok = all(any(normalized(alt) in normalized(h) for h in headings for alt in
                         (group if isinstance(group, list) else [group])) for group in groups)
    checks.append(_check("All 6 heading groups found", heading_ok))
    required_columns = gt_data["review_doc"]["required_table_columns"]
    # Other tables are allowed; identify the requested summary by its columns.
    candidates = [table for table in doc.tables if table.rows and
                  {normalized(c.text) for c in table.rows[0].cells} ==
                  {normalized(c) for c in required_columns}]
    table = candidates[0] if candidates else None
    shape_ok = table is not None and len(table.rows) == 6 and len(table.columns) == 4
    checks.append(_check("Summary table has exactly 5 data rows and 4 columns", shape_ok))
    papers_ok = shape_ok
    if shape_ok:
        headers = [normalized(c.text) for c in table.rows[0].cells]
        rows = [[c.text for c in row.cells] for row in table.rows[1:]]
        cols = {name: headers.index(normalized(name)) for name in required_columns}
        matched = set()
        for row in rows:
            paper = next((p for p in gt_data["papers_used"]
                          if normalized(row[cols["Paper Title"]]) == normalized(p["title"])), None)
            if paper is None or paper["arxiv_id"] in matched:
                papers_ok = False
                continue
            matched.add(paper["arxiv_id"])
            year_ok = str(row[cols["Year"]]).strip() in {str(y) for y in paper["years"]}
            method_ok = any(mentions(row[cols["Key Method"]], alias) for alias in paper["method_aliases"])
            result_ok = any(has_result(row[cols["Main Result"]], fact) for fact in paper["results"])
            papers_ok = papers_ok and year_ok and method_ok and result_ok
        papers_ok = papers_ok and len(matched) == 5
    checks.append(_check("Summary rows match the five source papers", papers_ok))
    prose = "\n".join(p.text for p in doc.paragraphs)
    # Numerical discussion must appear in the review, not only in its table.
    facts_ok = all(has_result(prose, fact) for fact in gt_data["review_doc"]["required_result_facts"])
    checks.append(_check("Word discussion includes source benchmark results", facts_ok))
    keyword_ok = all(mentions(word_text(doc), term) for term in gt_data["review_doc"]["required_mentions"])
    checks.append(_check("All required keywords present in Word doc", keyword_ok))
    return tuple(map(sum, zip(*checks)))


def check_pptx(agent_workspace, gt_data):
    path = Path(agent_workspace) / gt_data["slides"]["filename"]
    try:
        presentation = Presentation(path)
    except Exception as exc:
        return _check("Presentation readable", False, str(exc))
    checks = [_check("Presentation readable", True)]
    texts = [slide_text(slide) for slide in presentation.slides]
    checks.append(_check("PPTX has exactly 7 slides", len(texts) == 7))
    content_ok = len(texts) == 7
    if content_ok:
        title_ok = any(mentions(texts[0], t) for t in gt_data["slides"]["required_title_alternatives"])
        title_ok = title_ok and "2022" in texts[0] and "2023" in texts[0]
        overview_ok = all(mentions(texts[1], p["title"]) for p in gt_data["papers_used"])
        cot_ok = all(mentions(texts[2], t) for t in ["chain-of-thought", "self-consistency"])
        cot_ok = cot_ok and all(has_result(texts[2], f) for f in gt_data["review_doc"]["required_result_facts"][:2])
        tree_ok = mentions(texts[3], "tree of thoughts") and mentions(texts[3], "RAP")
        tree_ok = tree_ok and all(has_result(texts[3], f) for f in gt_data["review_doc"]["required_result_facts"][2:4])
        process_ok = any(mentions(texts[4], t) for t in ["process supervision", "process-supervised", "PRM"])
        process_ok = process_ok and all(has_result(texts[4], f) for f in gt_data["review_doc"]["required_result_facts"][4:])
        table_texts = ["\n".join(cell.text for row in shape.table.rows for cell in row.cells)
                       for shape in presentation.slides[5].shapes if shape.has_table]
        comparison_ok = any(sum(has_result(text, fact) for fact in gt_data["review_doc"]["required_result_facts"]) >= 2
                            for text in table_texts)
        conclusion_ok = any(mentions(texts[6], t) for t in ["заключение", "conclusion", "выводы"])
        content_ok = all([title_ok, overview_ok, cot_ok, tree_ok, process_ok, comparison_ok, conclusion_ok])
    checks.append(_check("Required slide content present", content_ok))
    return tuple(map(sum, zip(*checks)))


def count_variants(text):
    # Two common terminal word-count conventions: whitespace tokens and lexical
    # words (with technical hyphens/apostrophes kept inside a word).
    return {len(text.split()), len(re.findall(r"\w+(?:[-'’]\w+)*", text, flags=re.UNICODE))}


def check_word_count(agent_workspace):
    root = Path(agent_workspace)
    try:
        lines = [line.strip() for line in (root / "word_count.txt").read_text(encoding="utf-8").splitlines() if line.strip()]
        counts = {}
        for line in lines:
            match = re.fullmatch(r"(LLM_Reasoning_Review\.docx|LLM_Reasoning_Slides\.pptx):\s*(\d+)\s+words", line)
            if match is None or match[1] in counts:
                raise ValueError("Expected one actual count per filename")
            counts[match[1]] = int(match[2])
        doc = Document(root / "LLM_Reasoning_Review.docx")
        presentation = Presentation(root / "LLM_Reasoning_Slides.pptx")
        expected = {
            "LLM_Reasoning_Review.docx": count_variants(word_text(doc)),
            "LLM_Reasoning_Slides.pptx": count_variants("\n".join(slide_text(s) for s in presentation.slides)),
        }
        ok = set(counts) == set(expected) and all(counts[name] in values for name, values in expected.items())
        return _check("Word counts match extracted artifact text", ok, f"reported={counts}, actual={expected}")
    except Exception as exc:
        return _check("Word counts match extracted artifact text", False, str(exc))


def main(args):
    FAILED_NAMES.clear()
    # Preserve the shared required-artifact preflight before content checks.
    sys.path.insert(0, str(Path(__file__).resolve().parents[4]))
    from utils.evaluation.artifact_guards import require_task_artifacts
    require_task_artifacts(__file__, args.agent_workspace)
    with open(Path(args.groundtruth_workspace) / "expected_results.json", encoding="utf-8") as file:
        gt_data = json.load(file)
    results = [check_word_doc(args.agent_workspace, gt_data), check_pptx(args.agent_workspace, gt_data),
               check_word_count(args.agent_workspace)]
    passed, total = map(sum, zip(*results))
    critical = [name for name in FAILED_NAMES if name in CRITICAL_CHECKS]
    result = {"total_passed": passed, "total_checks": total, "accuracy": 100 * passed / total if total else 0,
              "critical_failed": critical, "timestamp": datetime.now().isoformat()}
    if args.res_log_file:
        Path(args.res_log_file).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    success = not critical and result["accuracy"] >= 80
    print("PASS" if success else "FAIL")
    raise SystemExit(0 if success else 1)


if __name__ == "__main__":
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", default=".")
    parser.add_argument("--groundtruth_workspace", default=".")
    parser.add_argument("--launch_time")
    parser.add_argument("--res_log_file")
    main(parser.parse_args())
