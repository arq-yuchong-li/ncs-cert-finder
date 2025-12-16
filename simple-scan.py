#!/usr/bin/env python3
from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import List, Dict, Set
import pdfplumber
from pathlib import Path

import pandas as pd
from collections import Counter

ROOT = Path("/Users/yuchong.li/Downloads/CVs/pdfs")  # change as needed
OUTPUT_XLSX = Path("ai_cert_employee.xlsx")

# Put EXACT certification names here (as they appear in CVs)
CERT_METADATA = {
    "Azure AI Fundamentals": "AI-900",
    "Azure AI Engineer Associate": "AI-102",
    "Azure Data Scientist Associate": "DP-100",
    "Designing Microsoft Azure Infrastructure Solutions": "AZ-305",
    "Power Platform Fundamentals": "PL-900",
    "Power Platform Developer Associate": "PL-400",
    "Implementing Analytics Solutions Using Microsoft Fabric": "DP-600",
    "Power Platform Solution Architect Expert": "PL-600",
    "Designing and Implementing Microsoft DevOps Solutions": "AZ-400",
    "Designing and Implementing Enterprise-Scale Analytics Solutions": "DP-500",
    "Microsoft Identity and Access Administrator": "SC-300",
    "Microsoft Cybersecurity Architect": "SC-100",
    "Agentic AI Business Solutions Architect": "AB-100",
    "AI Transformation Leader": "AB-731",
    "Responsible AI": "",
}
CERT_NAMES = list(CERT_METADATA.keys())

def normalize_text(s: str) -> str:
    # remove common bullet / NBSP weirdness
    s = s.replace("\u00a0", " ")
    s = s.replace("\u2022", " ")
    s = s.replace("\n", " ")
    s = s.replace("–", "-").replace("—", "-")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def extract_pdf_text(path: Path) -> str:
    text = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text.append(t)
    return "\n".join(text)



def employee_name_from_pdf(pdf_path: Path) -> str:
    """
    PDFs are output into one folder, so infer employee name from filename.
    Example: 'Ian Qin NCS Consultant Profile - Long Form.pdf' -> 'Ian Qin'
    """
    name = pdf_path.stem  # filename without .pdf
    name = re.sub(r"\s*NCS Consultant Profile.*$", "", name, flags=re.IGNORECASE).strip()
    name = re.sub(r"\s+", " ", name).strip()
    return name


def main() -> None:
    if not ROOT.exists():
        raise SystemExit(f"ROOT does not exist: {ROOT}")
    
    total_pdfs = 0
    cert_counter = Counter()
    employee_counter = Counter()

    cert_lower = {c: normalize_text(c).lower() for c in CERT_NAMES}
    matches: Set[tuple[str, str, str, str]] = set()

    for pdf_path in ROOT.rglob("*.pdf"):
        total_pdfs += 1

        '''Ian Qin's CV used for debug and testing'''
        try:
            text = normalize_text(extract_pdf_text(pdf_path)).lower()
            text = text.replace("\n", " ")
        except Exception:
            if "Ian Qin" in pdf_path.as_posix():
                print("[DEBUG] Failed reading Ian Qin PDF:", e)
            continue
        
        if "Ian Qin" in pdf_path.as_posix():
            print("[DEBUG] Ian Qin text length:", len(text))
            idx = text.find("cert")
            print("[DEBUG] Ian Qin snippet:", text[idx:idx+500] if idx != -1 else text[:500])

        employee = employee_name_from_pdf(pdf_path)

        for cert_name, cert_norm in cert_lower.items():
            if cert_norm in text:
                cert_id = CERT_METADATA.get(cert_name, "")
                vendor = "Microsoft"

                key = (employee, cert_name, cert_id, vendor)
                if key not in matches:
                    matches.add(key)
                    cert_counter[cert_name] += 1
                    employee_counter[employee] += 1

    results = [
        {
            "employeeName": e,
            "certName": c,
            "certId": cid,
            "vendor": v,
        }
        for (e, c, cid, v) in sorted(matches)
    ]

    df = pd.DataFrame(
        results,
        columns=["employeeName", "certName", "certId", "vendor"]
    )


    if not df.empty:
        df = df.sort_values(["certName", "employeeName"], kind="stable")

    df.to_excel(OUTPUT_XLSX, index=False)
    print("\n===== CV SCAN REPORT =====")
    print("Total CVs (PDFs) scanned:", total_pdfs)
    print("Total matched rows:", len(df))
    print("Unique employees matched:", len(employee_counter))
    print("\nMatches per certification:")

    for cert, count in cert_counter.most_common():
        print(f"  {cert}: {count}")

    print("========================")


if __name__ == "__main__":
    main()
