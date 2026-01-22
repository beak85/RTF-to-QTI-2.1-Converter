#!/usr/bin/env python3
"""
rtf_dir_to_qti21.py

Convert ALL RTF files in a directory into QTI 2.1 packages
compatible with Canvas New Quizzes.

Usage:
  python rtf_dir_to_qti21.py input_dir --outdir output_dir
"""

from __future__ import annotations
import argparse, html, os, re, zipfile
from pathlib import Path
from dataclasses import dataclass
from typing import Dict, List

# QTI constants
QTI_NS = "http://www.imsglobal.org/xsd/imsqti_v2p1"
IMSCP_NS = "http://www.imsglobal.org/xsd/imscp_v1p1"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
RP_MATCH_CORRECT = "http://www.imsglobal.org/question/qti_v2p1/rptemplates/match_correct"


@dataclass
class MCQ:
    num: int
    stem: str
    options: Dict[str, str]
    answer: str


def esc(s: str) -> str:
    return html.escape(s, quote=False)


# ---------- RTF PARSING ----------

def rtf_to_text(rtf: bytes) -> str:
    s = rtf.decode("utf-8", errors="ignore")
    s = re.sub(r"\\'([0-9a-fA-F]{2})",
               lambda m: bytes.fromhex(m.group(1)).decode("latin-1", errors="ignore"),
               s)
    s = s.replace("\\par", "\n").replace("\\line", "\n")
    s = s.replace("\\{", "{").replace("\\}", "}").replace("\\\\", "\\")
    s = re.sub(r"{\\\*[^{}]*}", "", s)
    s = re.sub(r"\\[a-zA-Z]+-?\d* ?", "", s)
    s = s.replace("{", "").replace("}", "")
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def clean_text(t: str) -> str:
    lines = [ln.strip() for ln in t.splitlines()]
    lines = [ln for ln in lines if ln and ln != "d"]
    return "\n".join(lines)


def parse_questions(text: str) -> List[MCQ]:
    if "MULTIPLE CHOICE" in text:
        text = text.split("MULTIPLE CHOICE", 1)[1]

    blocks = re.split(r"\n(?=\d+\.)", "\n" + text)
    questions: List[MCQ] = []

    for block in blocks:
        block = block.strip()
        if not block:
            continue

        m = re.match(r"(\d+)\.\n(.*)", block, re.S)
        if not m:
            continue

        num = int(m.group(1))
        body = m.group(2)

        ans = re.search(r"ANS:\s*([A-D])", body)
        if not ans:
            raise ValueError(f"Missing ANS for question {num}")
        answer = ans.group(1).lower()

        body = body.split("ANS:", 1)[0]
        opts = list(re.finditer(r"\n([a-d])\.\s+", "\n" + body))
        stem = ("\n" + body)[:opts[0].start()].strip()
        stem = re.sub(r"\n+", " ", stem)

        options = {}
        for i, o in enumerate(opts):
            letter = o.group(1)
            start = o.end()
            end = opts[i + 1].start() if i + 1 < len(opts) else len("\n" + body)
            options[letter] = re.sub(r"\n+", " ", ("\n" + body)[start:end]).strip()

        questions.append(MCQ(num, stem, options, answer))

    return questions


# ---------- QTI WRITER ----------

def write_qti21(questions: List[MCQ], title: str, out_zip: Path):
    work = out_zip.with_suffix("")
    (work / "items").mkdir(parents=True, exist_ok=True)

    item_files = []

    for q in questions:
        item_id = f"Q{q.num:03d}"
        correct = f"CHOICE_{q.answer.upper()}"

        choices = [
            f'      <simpleChoice identifier="CHOICE_{k.upper()}">{esc(v)}</simpleChoice>'
            for k, v in sorted(q.options.items())
        ]

        item_xml = [
            '<?xml version="1.0" encoding="UTF-8"?>',
            f'<assessmentItem xmlns="{QTI_NS}" identifier="{item_id}" title="Question {q.num}" adaptive="false" timeDependent="false">',
            '  <responseDeclaration identifier="RESPONSE" cardinality="single" baseType="identifier">',
            f'    <correctResponse><value>{correct}</value></correctResponse>',
            '  </responseDeclaration>',
            '  <itemBody>',
            '    <choiceInteraction responseIdentifier="RESPONSE" maxChoices="1">',
            f'      <prompt>{esc(q.stem)}</prompt>',
            *choices,
            '    </choiceInteraction>',
            '  </itemBody>',
            f'  <responseProcessing template="{RP_MATCH_CORRECT}"/>',
            '</assessmentItem>',
        ]

        rel = f"items/{item_id}.xml"
        (work / rel).write_text("\n".join(item_xml), encoding="utf-8")
        item_files.append(rel)

    test_xml = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        f'<assessmentTest xmlns="{QTI_NS}" identifier="TEST1" title="{esc(title)}">',
        '  <testPart identifier="part1" navigationMode="linear" submissionMode="individual">',
        '    <assessmentSection identifier="section1" visible="true">',
    ]
    for f in item_files:
        test_xml.append(f'      <assessmentItemRef href="{f}"/>')
    test_xml += ['    </assessmentSection>', '  </testPart>', '</assessmentTest>']
    (work / "assessmentTest.xml").write_text("\n".join(test_xml), encoding="utf-8")

    manifest = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        f'<manifest xmlns="{IMSCP_NS}" xmlns:imsqti="{QTI_NS}">',
        '  <resources>',
        '    <resource type="imsqti_test_xmlv2p1" href="assessmentTest.xml">',
        '      <file href="assessmentTest.xml"/>',
    ]
    for f in item_files:
        manifest.append(f'      <file href="{f}"/>')
    manifest += ['    </resource>', '  </resources>', '</manifest>']
    (work / "imsmanifest.xml").write_text("\n".join(manifest), encoding="utf-8")

    with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(work / "imsmanifest.xml", "imsmanifest.xml")
        z.write(work / "assessmentTest.xml", "assessmentTest.xml")
        for f in item_files:
            z.write(work / f, f)


# ---------- MAIN ----------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("indir", type=Path, help="Directory containing RTF files")
    ap.add_argument("--outdir", type=Path, default=Path("qti_output"))
    args = ap.parse_args()

    args.outdir.mkdir(exist_ok=True)

    for rtf in sorted(args.indir.glob("*.rtf")):
        print(f"Converting {rtf.name}...")
        text = clean_text(rtf_to_text(rtf.read_bytes()))
        questions = parse_questions(text)
        out_zip = args.outdir / f"{rtf.stem}_QTI21.zip"
        write_qti21(questions, rtf.stem, out_zip)
        print(f"  → {out_zip.name} ({len(questions)} questions)")

    print("Done.")


if __name__ == "__main__":
    main()
