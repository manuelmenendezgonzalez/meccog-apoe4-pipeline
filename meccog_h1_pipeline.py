"""MecCog Hypothesis 1 Pilot Pipeline (CoResearcher + EdiTXT Integration).

Generates an audited, validator-compliant submission spreadsheet for:
  H1: "In some types of non-Alzheimer's, non-aged human brain cells, for example
       astrocytes, APOE4 causes reduced endosome pH compared with APOE3."

Verification invariants enforced:
1. Every PMID, DOI, and retraction status comes from real PubMed XML (cache_h1).
2. Every verbatim `Finding quote` is checked as an exact substring of the
   downloaded abstract or Europe PMC OA full-text XML.
3. Missing statistics are honestly tagged with official provenance codes
   (`full_text_unavailable` or `not_found_in_source`).
4. Output spreadsheet is validated via official `validate_submission.py --strict`.
"""

from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
import openpyxl

BASE_DIR = Path(r"c:\Users\usuario\Neurodiagnoses Ecosystem\proposals\MecCog_APOE4_Challenge_2026")
RES_DIR = BASE_DIR / "resources"
CACHE_DIR = BASE_DIR / "cache_h1"
TEMPLATE_PATH = RES_DIR / "submission_template.xlsx"
VALIDATOR_PATH = RES_DIR / "validate_submission.py"
OUT_XLSX = BASE_DIR / "submission_H1_pilot.xlsx"
MANIFEST_PATH = BASE_DIR / "pilot_audit_manifest.json"

HYPOTHESIS_H1 = (
    "In some types of non-Alzheimer's, non-aged human brain cells, for example astrocytes, "
    "APOE4 causes reduced endosome pH compared with APOE3."
)


def load_paper_texts() -> dict[str, dict]:
    """Load downloaded PubMed metadata and normalized OA full text per PMID."""
    with open(CACHE_DIR / "h1_papers_metadata.json", "r", encoding="utf-8") as f:
        meta = json.load(f)

    by_pmid: dict[str, dict] = {}
    for p in meta["papers"]:
        pmid = p["pmid"]
        abs_norm = " ".join((p.get("abstract") or "").split())
        ft_norm = ""
        xml_path = p.get("fulltext_xml") or ""
        if not xml_path and p.get("pmcid"):
            cand = CACHE_DIR / f"{p['pmcid']}.xml"
            if cand.exists():
                xml_path = str(cand)
        if xml_path and os.path.exists(xml_path):
            tree = ET.parse(xml_path)
            ft_norm = " ".join("".join(tree.getroot().itertext()).split())
        by_pmid[pmid] = {
            **p,
            "abs_norm": abs_norm,
            "ft_norm": ft_norm,
            "has_oa_fulltext": bool(ft_norm),
        }
    return by_pmid


def verify_quote_in_source(paper: dict, quote: str) -> str:
    """EdiTXT integrity gate: assert quote exists verbatim in abstract or OA full text."""
    q_norm = " ".join(quote.split())
    if q_norm in paper["abs_norm"]:
        return "abstract"
    if paper["ft_norm"] and q_norm in paper["ft_norm"]:
        return "full_text_xml"
    raise AssertionError(
        f"QUOTE VERIFICATION FAILED for PMID {paper['pmid']}: {q_norm!r} not found verbatim!"
    )


def build_curated_entries() -> list[dict]:
    """Structured findings curated from verified PubMed abstracts & OA full-text XMLs."""
    return [
        {
            "pmid": "29946028",
            "doi": "10.1073/pnas.1801612115",
            "source_type": "PubMed published",
            "id": "P1",
            "findings": [
                {
                    "id": "P1.F1",
                    "desc": (
                        "In ApoE4 astrocytes, aberrant endosomal acidification (lower endosome pH compared "
                        "with ApoE3) is driven by selective transcriptional downregulation of NHE6 via "
                        "nuclear HDAC4 translocation, trapping LRP1 intracellularly and impairing Aβ clearance."
                    ),
                    "quote": (
                        "Here, we show that aberrant endosomal acidification in ApoE4 astrocytes traps the "
                        "low-density lipoprotein receptor-related protein (LRP1) within intracellular compartments, "
                        "leading to loss of surface expression and Aβ clearance."
                    ),
                    "summary": "(APOE3 -> APOE4) -> (HDAC4 nuclear translocation -> NHE6 downregulation -> lower endosomal pH in astrocytes)",
                    "relevance": 0.95,
                    "exp_system": (
                        "ApoE4 vs. ApoE3 astrocytes (murine astrocytes expressing human ApoE isoforms; "
                        "in vivo validation in NHE6KO mice). NOTE: Abstract-only; full text not available in OA XML."
                    ),
                    "data_location": "Abstract (full text paywalled)",
                    "effect_size": "N/A",
                    "effect_prov": "full_text_unavailable",
                    "p_value": "N/A",
                    "p_prov": "full_text_unavailable",
                    "sample_size": "N/A",
                    "sample_prov": "full_text_unavailable",
                    "directness": "1",
                    "strength": "2 strong",
                },
                {
                    "id": "P1.F2",
                    "desc": (
                        "Pharmacological restoration of NHE6 expression using HDAC inhibitors normalized "
                        "ApoE4-specific defects in endosomal pH, LRP1 trafficking, and Aβ clearance."
                    ),
                    "quote": (
                        "HDAC inhibitors that restored NHE6 expression normalized ApoE4-specific defects in "
                        "endosomal pH, LRP1 trafficking, and amyloid clearance."
                    ),
                    "summary": "(HDACi -> NHE6 rescue in APOE4 astrocytes) -> (normalized endosomal pH & LRP1/Aβ clearance)",
                    "relevance": 0.85,
                    "exp_system": (
                        "ApoE4 vs. ApoE3 astrocytes treated with HDAC inhibitors. "
                        "NOTE: Rescue perturbation (indirect input); abstract-only."
                    ),
                    "data_location": "Abstract (full text paywalled)",
                    "effect_size": "N/A",
                    "effect_prov": "full_text_unavailable",
                    "p_value": "N/A",
                    "p_prov": "full_text_unavailable",
                    "sample_size": "N/A",
                    "sample_prov": "full_text_unavailable",
                    "directness": "2",
                    "strength": "3 moderate",
                },
            ],
        },
        {
            "pmid": "29567836",
            "doi": "10.1074/jbc.RA118.002025",
            "source_type": "PubMed published",
            "id": "P2",
            "findings": [
                {
                    "id": "P2.F1",
                    "desc": (
                        "Pharmacological targeting of the CREB/HDAC pathway to increase NHE6 expression "
                        "regulated endosomal pH and corrected defective Aβ clearance in an apoE4 astrocyte model."
                    ),
                    "quote": (
                        "Of note, pharmacological targeting of the CREB pathway to increase NHE6 expression "
                        "helped regulate endosomal pH and correct defective clearance of amyloid Aβ in an apoE4 "
                        "astrocyte model of Alzheimer's disease."
                    ),
                    "summary": "(CREB activation -> NHE6 upregulation in APOE4 astrocytes) -> (endosomal pH restoration)",
                    "relevance": 0.75,
                    "exp_system": (
                        "apoE4 astrocyte cell culture model of AD (alongside yeast/Drosophila/mouse bioinformatics). "
                        "Caveat: Abstract-only; full text not available in OA XML."
                    ),
                    "data_location": "Abstract (full text paywalled)",
                    "effect_size": "N/A",
                    "effect_prov": "full_text_unavailable",
                    "p_value": "N/A",
                    "p_prov": "full_text_unavailable",
                    "sample_size": "N/A",
                    "sample_prov": "full_text_unavailable",
                    "directness": "2",
                    "strength": "3 moderate",
                }
            ],
        },
        {
            "pmid": "30375977",
            "doi": "10.7554/eLife.40048",
            "source_type": "PubMed published",
            "id": "P3",
            "findings": [
                {
                    "id": "P3.F1",
                    "desc": (
                        "NUANCED MECHANISM: In primary rodent neurons, pharmacological NHE inhibition (EMD87580) "
                        "and shRNA knockdown of NHE6 (which further acidify early endosomes) completely reversed "
                        "ApoE4-induced Apoer2 and AMPA/NMDA receptor recycling blocks—indicating ApoE4's isoelectric "
                        "point matches physiological EE pH rather than requiring hyperacidification for toxicity."
                    ),
                    "quote": (
                        "We found that pharmacological and genetic inhibition of NHE6, the primary proton leak "
                        "channel in the early endosome, in rodents completely reverses the ApoE4-induced recycling "
                        "block of the ApoE receptor Apoer2/Lrp8 and the AMPA- and NMDA-type glutamate receptors"
                    ),
                    "summary": "(APOE4 + NHE6 inhibition -> lower EE pH) -> (restored Apoer2/glutamate receptor recycling in neurons)",
                    "relevance": 0.80,
                    "exp_system": (
                        "Primary rodent cortical neurons and ApoE4-KI vs ApoE3-KI hippocampal slices; "
                        "EMD87580 (3 uM) and 3 NHE6 shRNAs (>90% knockdown). Caveat: rodent neurons, not human astrocytes; "
                        "tests endosomal acidification as a rescue rather than APOE4-driven baseline pH reduction."
                    ),
                    "data_location": "Abstract & Fig 4-6 (PMC6261251)",
                    "effect_size": "N/A",
                    "effect_prov": "not_found_in_source",
                    "p_value": "N/A",
                    "p_prov": "not_found_in_source",
                    "sample_size": "N/A",
                    "sample_prov": "not_found_in_source",
                    "directness": "3",
                    "strength": "3 moderate",
                }
            ],
        },
        {
            "pmid": "34617884",
            "doi": "10.7554/eLife.72034",
            "source_type": "PubMed published",
            "id": "P4",
            "findings": [
                {
                    "id": "P4.F1",
                    "desc": (
                        "NUANCED MECHANISM: Genetic disruption of NHE6 (Slc9a6-) restores Apoer2 and AMPA/NMDA "
                        "receptor surface trafficking in ApoE4-treated neurons and reduces amyloid plaque deposition "
                        "in AppNL-F;APOE4 mice by accelerating early endosome acidification."
                    ),
                    "quote": (
                        "We have shown that enhancing endosomal acidification by inhibiting the EE-specific "
                        "sodium-hydrogen exchanger 6 (NHE6) restores vesicular trafficking and normalizes synaptic homeostasis."
                    ),
                    "summary": "(Slc9a6/NHE6 loss -> accelerated EE acidification) -> (rescues APOE4 trafficking & lowers plaques)",
                    "relevance": 0.75,
                    "exp_system": (
                        "Slc9a6- embryonic mouse neurons and Slc9a6fl;CAG-CreERT2;APOE4;AppNL-F mice (14-16 mo). "
                        "Caveat: mouse model; shows EE acidification rescues APOE4 defects rather than causing them."
                    ),
                    "data_location": "Abstract, Fig 4C, Fig 7C-D (PMC8547963)",
                    "effect_size": "N/A",
                    "effect_prov": "not_found_in_source",
                    "p_value": "N/A",
                    "p_prov": "not_found_in_source",
                    "sample_size": "N/A",
                    "sample_prov": "not_found_in_source",
                    "directness": "3",
                    "strength": "3 moderate",
                }
            ],
        },
        {
            "pmid": "41103078",
            "doi": "10.1080/15548627.2025.2576613",
            "source_type": "PubMed published",
            "id": "P5",
            "findings": [
                {
                    "id": "P5.F1",
                    "desc": (
                        "CONTRADICTING (endo-lysosomal compartment): Human APOE4 expression in Neuro-2a neuronal cells "
                        "causes dramatic lysosomal alkalinization (higher pH = 5.86 vs. pH = 4.30 in HsAPOE3) via "
                        "TMED5 accumulation and LGALS3BP depletion, impairing cathepsin B proteolysis."
                    ),
                    "quote": (
                        "We found that HsAPOE4-expressing cells exhibit dramatic lysosomal alkalinization (pH = 5.86) "
                        "compared to HsAPOE3 (pH = 4.30) cells (Figure 1C–E)"
                    ),
                    "summary": "(APOE3 -> APOE4) -> (lysosomal alkalinization: pH 4.30 -> pH 5.86 in neuronal cells) [CONTRADICTS lower-pH direction]",
                    "relevance": 0.85,
                    "exp_system": (
                        "Neuro-2a neuronal cells expressing human APOE4 vs. HsAPOE3 genomic locus; OregonGreen-70000 kDa "
                        "dextran ratiometric pH clamping (3 independent biological replicates). Caveat: measures mature "
                        "lysosomal pH rather than early endosome pH, in mouse neuroblastoma cells."
                    ),
                    "data_location": "Fig 1C-E (PMC12758178)",
                    "effect_size": "qualitative",
                    "effect_prov": "printed_in_text",
                    "p_value": "N/A",
                    "p_prov": "not_found_in_source",
                    "sample_size": "3",
                    "sample_prov": "printed_in_text",
                    "directness": "3",
                    "strength": "3 moderate",
                }
            ],
        },
        {
            "pmid": "42182449",
            "doi": "10.64898/2026.05.12.724612",
            "source_type": "PubMed preprint",
            "id": "P6",
            "findings": [
                {
                    "id": "P6.F1",
                    "desc": (
                        "CONTRADICTING (human iPSC microglia): Isogenic human APOE44 microglia exhibit impaired "
                        "lysosomal acidification and shift toward alkalinization upon challenge, accompanied by "
                        "downregulation of 53 of 84 lysosomal hydrolases and V-ATPases."
                    ),
                    "quote": (
                        "Upon challenge with organoid debris, UCI5 APOE44 iMG lysosomes failed to mount an "
                        "adaptive acidification response and instead shifted toward alkalinization (Figure 7M)."
                    ),
                    "summary": "(APOE33 -> APOE44 in human iPSC microglia) -> (impaired lysosomal acidification / alkalinization) [CONTRADICTS lower-pH direction]",
                    "relevance": 0.90,
                    "exp_system": (
                        "Human isogenic APOE44 vs. APOE33 iPSC-derived microglia (iMG; Jackson Lab & UCI5 lines; "
                        "n=3-4 donor lines for functional assays). Caveat: human microglia (not astrocytes) and "
                        "lysosomal compartment rather than early endosome."
                    ),
                    "data_location": "Fig 7M, Fig S3D (PMC13192593)",
                    "effect_size": "qualitative",
                    "effect_prov": "printed_in_text",
                    "p_value": "N/A",
                    "p_prov": "not_found_in_source",
                    "sample_size": "3-4",
                    "sample_prov": "printed_in_text",
                    "directness": "3",
                    "strength": "2 strong",
                }
            ],
        },
    ]


def populate_workbook() -> dict:
    """Verify every quote/PMID and write the official validated Excel workbook."""
    papers_by_pmid = load_paper_texts()
    curated = build_curated_entries()

    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb["Submission"]

    # Clear any sample/blank rows below row 2
    if ws.max_row > 2:
        ws.delete_rows(3, ws.max_row - 2)

    ws.cell(2, 1).value = HYPOTHESIS_H1

    row_idx = 3
    audit_log = []

    for entry in curated:
        pmid = entry["pmid"]
        assert pmid in papers_by_pmid, f"PMID {pmid} missing from cache_h1!"
        p_meta = papers_by_pmid[pmid]
        assert not p_meta["retracted"], f"PMID {pmid} is RETRACTED!"
        assert p_meta["doi"] == entry["doi"], f"DOI mismatch for {pmid}: {p_meta['doi']} vs {entry['doi']}"

        # Write Paper row
        ws.cell(row_idx, 2).value = entry["doi"]
        ws.cell(row_idx, 3).value = entry["source_type"]
        ws.cell(row_idx, 4).value = int(pmid)
        ws.cell(row_idx, 5).value = entry["id"]
        row_idx += 1

        for f in entry["findings"]:
            matched_in = verify_quote_in_source(p_meta, f["quote"])
            ws.cell(row_idx, 5).value = f["id"]
            ws.cell(row_idx, 6).value = f["desc"]
            ws.cell(row_idx, 7).value = f["quote"]
            ws.cell(row_idx, 8).value = f["summary"]
            ws.cell(row_idx, 9).value = f["relevance"]
            ws.cell(row_idx, 10).value = f["exp_system"]
            ws.cell(row_idx, 11).value = f["data_location"]
            ws.cell(row_idx, 12).value = f["effect_size"]
            ws.cell(row_idx, 13).value = f["effect_prov"]
            ws.cell(row_idx, 14).value = f["p_value"]
            ws.cell(row_idx, 15).value = f["p_prov"]
            ws.cell(row_idx, 16).value = f["sample_size"]
            ws.cell(row_idx, 17).value = f["sample_prov"]
            ws.cell(row_idx, 18).value = int(f["directness"])
            ws.cell(row_idx, 19).value = f["strength"]
            row_idx += 1

            audit_log.append({
                "finding_id": f["id"],
                "pmid": pmid,
                "doi": entry["doi"],
                "retracted": p_meta["retracted"],
                "quote_verified_in": matched_in,
                "has_oa_fulltext": p_meta["has_oa_fulltext"],
            })

    # Populate Method sheet
    ws_m = wb["Method"]
    method_values = {
        "Team": "FdNIA-CoResearcher",
        "Team leader": "Manuel Menendez-Gonzalez",
        "Team leader email": "administracion@fneurociencias.org",
        "Method URL": "https://github.com/Fundacion-de-Neurociencias/meccog-apoe4-pilot",
        "Commit SHA": "0000000",
        "Who did the work": "both",
        "Model and version": "CoResearcher + EdiTXT bibliographic verifier + Gemini 4 Hydrogen (pilot curation)",
    }
    for r in range(1, ws_m.max_row + 1):
        lbl = ws_m.cell(r, 1).value
        if lbl in method_values:
            ws_m.cell(r, 2).value = method_values[lbl]

    wb.save(OUT_XLSX)
    sha256 = hashlib.sha256(OUT_XLSX.read_bytes()).hexdigest()

    manifest = {
        "hypothesis": HYPOTHESIS_H1,
        "output_xlsx": str(OUT_XLSX),
        "xlsx_sha256": sha256,
        "papers_count": len(curated),
        "findings_count": len(audit_log),
        "audit_records": audit_log,
    }
    MANIFEST_PATH.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
    return manifest


if __name__ == "__main__":
    manifest = populate_workbook()
    print(f"Generated {OUT_XLSX} | SHA256={manifest['xlsx_sha256'][:16]}...")
    print(f"Papers: {manifest['papers_count']} | Findings: {manifest['findings_count']}")
    res = subprocess.run(
        [sys.executable, str(VALIDATOR_PATH), str(OUT_XLSX), "--strict"],
        capture_output=True,
        text=True,
    )
    print(res.stdout)
    if res.stderr:
        print("STDERR:", res.stderr)
    sys.exit(res.returncode)
