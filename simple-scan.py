#!/usr/bin/env python3
from __future__ import annotations

import re
from pathlib import Path
from typing import Set
import pdfplumber
import pandas as pd
from collections import Counter

ROOT = Path("/Users/Your_path/pdfs")  # change as needed
OUTPUT_XLSX = Path("ai_cert_employee.xlsx")

# Put EXACT certification names here (as they appear in CVs)
CERT_CATALOG = {
    # Azure
    "Azure AI Fundamentals": {"id": "AI-900", "vendor": "Microsoft"},
    "Azure AI Engineer Associate": {"id": "AI-102", "vendor": "Microsoft"},
    "Azure Data Scientist Associate": {"id": "DP-100", "vendor": "Microsoft"},
    "Designing Microsoft Azure Infrastructure Solutions": {"id": "AZ-305", "vendor": "Microsoft"},
    "Power Platform Fundamentals": {"id": "PL-900", "vendor": "Microsoft"},
    "Power Platform Developer Associate": {"id": "PL-400", "vendor": "Microsoft"},
    "Implementing Analytics Solutions Using Microsoft Fabric": {"id": "DP-600", "vendor": "Microsoft"},
    "Power Platform Solution Architect Expert": {"id": "PL-600", "vendor": "Microsoft"},
    "Designing and Implementing Microsoft DevOps Solutions": {"id": "AZ-400", "vendor": "Microsoft"},
    "Designing and Implementing Enterprise-Scale Analytics Solutions": {"id": "DP-500", "vendor": "Microsoft"},
    "Microsoft Identity and Access Administrator": {"id": "SC-300", "vendor": "Microsoft"},
    "Microsoft Cybersecurity Architect": {"id": "SC-100", "vendor": "Microsoft"},
    "Agentic AI Business Solutions Architect (beta)": {"id": "AB-100", "vendor": "Microsoft"},
    "AI Transformation Leader": {"id": "B-731", "vendor": "Microsoft"},
    "Responsible AI": {"id": "", "vendor": "mentioned in CV, not a cert"},

    # AWS
    "AWS Certified Cloud Practitioner": {"id": "CLF-C02", "vendor": "AWS"},
    "AWS Certified Developer - Associate": {"id": "DVA-C02", "vendor": "AWS"},
    "AWS Certified Machine Learning - Specialty": {"id": "MLS-C01", "vendor": "AWS"},
    "AWS Certified Data Engineer - Associate": {"id": "DEA-C01", "vendor": "AWS"},
    "AWS Certified Solutions Architect - Associate": {"id": "SAA-C03", "vendor": "AWS"},
    "AWS Certified Solutions Architect - Professional": {"id": "SAP-C02", "vendor": "AWS"},
    "AWS Certified DevOps Engineer - Professional": {"id": "DOP-C02", "vendor": "AWS"},
    "AWS Certified Database - Specialty": {"id": "DBS-C01", "vendor": "AWS"},
    "AWS Certified Security - Specialty": {"id": "SCS-C03", "vendor": "AWS"},
    "AWS Certified Advanced Networking - Specialty": {"id": "ANS-C01", "vendor": "AWS"},

    #GCP
    "Cloud Digital Leader": {"id": "CDL", "vendor": "Google Cloud"},
    "Associate Cloud Engineer": {"id": "ACE", "vendor": "Google Cloud"},
    "Professional Cloud Architect": {"id": "PCA", "vendor": "Google Cloud"},
    "Professional Data Engineer": {"id": "PDE", "vendor": "Google Cloud"},
    "Professional Cloud DevOps Engineer": {"id": "PCDOE", "vendor": "Google Cloud"},
    "Professional Cloud Developer": {"id": "PCD", "vendor": "Google Cloud"},
    "Professional Cloud Network Engineer": {"id": "PCNE", "vendor": "Google Cloud"},
    "Professional Cloud Security Engineer": {"id": "PCSE", "vendor": "Google Cloud"},
    "Professional Cloud Database Engineer": {"id": "PCDBE", "vendor": "Google Cloud"},
    "Professional Machine Learning Engineer": {"id": "PMLE", "vendor": "Google Cloud"},
    "Google Cloud Certified Fellow": {"id": "GCCF", "vendor": "Google Cloud"},

    #Databricks
    "Generative AI Fundamentals": {"id": "-", "vendor": "Databricks"},
    "Generative AI Engineering Associate": {"id": "-", "vendor": "Databricks"},
    "Machine Learning Associate": {"id": "-", "vendor": "Databricks"},
    "Machine Learning Professional": {"id": "-", "vendor": "Databricks"},
}

CERT_NAMES = list(CERT_CATALOG.keys())

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
            '''if "Ian Qin" in pdf_path.as_posix():
                print("[DEBUG] Failed reading Ian Qin PDF:", e)'''
            continue
        
        '''if "Ian Qin" in pdf_path.as_posix():
            print("[DEBUG] Ian Qin text length:", len(text))
            idx = text.find("cert")
            print("[DEBUG] Ian Qin snippet:", text[idx:idx+500] if idx != -1 else text[:500])'''

        employee = employee_name_from_pdf(pdf_path)

        for cert_name, cert_norm in cert_lower.items():
            if cert_norm in text:
                meta = CERT_CATALOG[cert_name]
                cert_id = meta["id"]
                vendor = meta["vendor"]

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
