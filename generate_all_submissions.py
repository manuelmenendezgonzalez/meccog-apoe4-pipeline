"""Generate and strictly validate MecCog Challenge submissions for ALL 5 hypotheses (H1-H5).

Enforces:
1. Real PMID/DOI/retraction verification from local PubMed XML caches (cache_h1 & cache_all).
2. Exact verbatim substring verification for every Finding quote against abstract or OA full-text XML.
3. Exact official `validate_submission.py --strict` pass for all 5 workbooks.
4. Pins real public GitHub Method URL and 40-char Commit SHA in the Method sheet.
"""

from __future__ import annotations

import argparse
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
CACHE_H1 = BASE_DIR / "cache_h1"
CACHE_ALL = BASE_DIR / "cache_all"
TEMPLATE_PATH = RES_DIR / "submission_template.xlsx"
VALIDATOR_PATH = RES_DIR / "validate_submission.py"

HYPOTHESIS_STATEMENTS = {
    "H1": "In some types of non-Alzheimer's, non-aged human brain cells, for example astrocytes, APOE4 causes reduced endosome pH compared with APOE3.",
    "H2": "In some types of non-Alzheimer's, non-aged human brain cells, for example astrocytes, APOE4 causes impaired ARF6-related recycling of receptors between the cell surface and the endosome, compared with APOE3.",
    "H3": "In some types of non-Alzheimer's, non-aged human brain cells, APOE4 undergoes increased aggregation in the endo-lysosome compared with APOE3.",
    "H4": "In some types of non-Alzheimer's, non-aged human brain cells, APOE4 adheres more tightly to some recycling receptors than APOE3, impeding recycling.",
    "H5": "In some types of non-Alzheimer's, non-aged human brain cells, APOE4 disruption of lipid metabolism, compared with APOE3, results in impaired late endosome/lysosome function.",
}


def load_all_papers() -> dict[str, dict]:
    """Load all cached PubMed metadata and OA full-text XMLs across cache_h1 and cache_all."""
    by_pmid: dict[str, dict] = {}

    def ingest_list(papers: list[dict], cache_dir: Path) -> None:
        for p in papers:
            pmid = p["pmid"]
            abs_norm = " ".join((p.get("abstract") or "").split())
            ft_norm = ""
            xml_path = p.get("fulltext_xml") or ""
            if not xml_path and p.get("pmcid"):
                cand = cache_dir / f"{p['pmcid']}.xml"
                if cand.exists():
                    xml_path = str(cand)
            if xml_path and os.path.exists(xml_path):
                tree = ET.parse(xml_path)
                ft_norm = " ".join("".join(tree.getroot().itertext()).split())
            existing = by_pmid.get(pmid)
            if not existing or (ft_norm and not existing["ft_norm"]):
                by_pmid[pmid] = {
                    **p,
                    "abs_norm": abs_norm,
                    "ft_norm": ft_norm,
                    "has_oa_fulltext": bool(ft_norm),
                }

    with open(CACHE_H1 / "h1_papers_metadata.json", "r", encoding="utf-8") as f:
        ingest_list(json.load(f)["papers"], CACHE_H1)
    with open(CACHE_ALL / "h2_h5_metadata.json", "r", encoding="utf-8") as f:
        h25 = json.load(f)
        for blk in h25.values():
            ingest_list(blk["papers"], CACHE_ALL)
    return by_pmid


def verify_quote(paper: dict, quote: str) -> str:
    q_norm = " ".join(quote.split())
    if q_norm in paper["abs_norm"]:
        return "abstract"
    if paper["ft_norm"] and q_norm in paper["ft_norm"]:
        return "full_text_xml"
    raise AssertionError(f"Quote not found verbatim in PMID {paper['pmid']}: {q_norm!r}")


def get_curated_data() -> dict[str, list[dict]]:
    from meccog_h1_pipeline import build_curated_entries as build_h1

    h2 = [
        {
            "pmid": "31641056",
            "doi": "10.1523/JNEUROSCI.1400-19.2019",
            "source_type": "PubMed published",
            "id": "P1",
            "findings": [
                {
                    "id": "P1.F1",
                    "desc": "ApoE4 promotes greater ARF6 expression than ApoE3, trapping ABCA1 in late-endosomes and impairing ABCA1 recycling to the cell membrane, lowering cholesterol efflux.",
                    "quote": "We found that ApoE4 promoted greater expression of ARF6 compared with ApoE3, trapping ABCA1 in late-endosomes and impairing its recycling to the cell membrane.",
                    "summary": "(APOE3 -> APOE4) -> (increased ARF6 -> ABCA1 trapped in late endosomes -> impaired membrane recycling)",
                    "relevance": 0.98,
                    "exp_system": "Mouse/cellular astrocytes + human CSF (APOE e4/e4 vs e3/e3) + male ApoE4-TR mice. NOTE: Abstract-only (OA XML unavailable).",
                    "data_location": "Abstract (full text paywalled)",
                    "effect_size": "N/A", "effect_prov": "full_text_unavailable",
                    "p_value": "N/A", "p_prov": "full_text_unavailable",
                    "sample_size": "N/A", "sample_prov": "full_text_unavailable",
                    "directness": "1", "strength": "2 strong",
                }
            ],
        },
        {
            "pmid": "39901180",
            "doi": "10.1186/s13024-025-00802-7",
            "source_type": "PubMed published",
            "id": "P2",
            "findings": [
                {
                    "id": "P2.F1",
                    "desc": "Caveolin-1 promotes ABCA1 endolysosomal trapping in APOE4-TR mice and human AD brains; reducing cholesterol via cyclodextrin lowers ABCA1 lysosomal trapping.",
                    "quote": "Discovery proteomics identified caveolin-1, a sensor of cellular cholesterol accumulation, as a key promoter of ABCA1 endolysosomal trafficking. Greater caveolin-1 expression was observed in APOE4-TR mouse models and AD human brains.",
                    "summary": "(APOE4 / oxysterols) -> (caveolin-1 upregulation -> ABCA1 endolysosomal trapping)",
                    "relevance": 0.80,
                    "exp_system": "ROSMAP DLPFC (n=632 bulk, n=427 snRNA-seq), APOE4-TR mice, and human iPSC-derived astrocytes. Caveat: caveolin-1 pathway rather than ARF6 directly.",
                    "data_location": "Abstract & Fig 1 (PMC11792374)",
                    "effect_size": "N/A", "effect_prov": "not_found_in_source",
                    "p_value": "N/A", "p_prov": "not_found_in_source",
                    "sample_size": "632", "sample_prov": "printed_in_text",
                    "directness": "3", "strength": "3 moderate",
                }
            ],
        },
        {
            "pmid": "38217595",
            "doi": "10.3233/JAD-230514",
            "source_type": "PubMed published",
            "id": "P3",
            "findings": [
                {
                    "id": "P3.F1",
                    "desc": "APOE4 primary neurons exhibit markedly reduced surface-to-total membrane levels of APOER2 (53% reduction) and LRP1 (48% reduction) compared with APOE3 neurons.",
                    "quote": "The results obtained show that the ratio of the externally oriented receptors to the total, was significantly lower in APOE4 neurons of APOER2 (Fig. 4A, 53% reduction, p < 0.05) and LRP1",
                    "summary": "(APOE3 -> APOE4) -> (53% lower surface APOER2 & 48% lower surface LRP1)",
                    "relevance": 0.82,
                    "exp_system": "Primary mouse cortical neurons from APOE4 vs APOE3 targeted-replacement mice; surface biotinylation ELISA (N=3 preparations). Caveat: mouse neurons, ARF6 not directly tested.",
                    "data_location": "Fig 4A-B (PMC10894586)",
                    "effect_size": "53%", "effect_prov": "printed_in_text",
                    "p_value": "<0.05", "p_prov": "printed_in_text",
                    "sample_size": "3", "sample_prov": "printed_in_text",
                    "directness": "3", "strength": "3 moderate",
                }
            ],
        },
    ]

    h3 = [
        {
            "pmid": "39419029",
            "doi": "10.1016/j.immuni.2024.09.014",
            "source_type": "PubMed published",
            "id": "P1",
            "findings": [
                {
                    "id": "P1.F1",
                    "desc": "Fibrillary APOE aggregates form within the endo-lysosomal system of microglia in amyloidosis mice and human AD brains, seeding Aβ amyloidosis.",
                    "quote": "These APOE aggregates that stained positive for β sheet-binding dyes triggered Aβ amyloidosis within the endo-lysosomal system of microglia, in a process influenced by microglial lipid metabolism and the JAK/STAT signaling pathway.",
                    "summary": "(APOE4 / microglial uptake) -> (fibrillary APOE aggregation in endo-lysosomes -> Aβ seeding)",
                    "relevance": 0.95,
                    "exp_system": "HaloTag-APOE knock-in mice and human AD brain autopsies. NOTE: Abstract-only (OA XML unavailable).",
                    "data_location": "Abstract (full text paywalled)",
                    "effect_size": "N/A", "effect_prov": "full_text_unavailable",
                    "p_value": "N/A", "p_prov": "full_text_unavailable",
                    "sample_size": "N/A", "sample_prov": "full_text_unavailable",
                    "directness": "1", "strength": "2 strong",
                }
            ],
        },
        {
            "pmid": "39532095",
            "doi": "10.1016/j.cell.2024.10.027",
            "source_type": "PubMed published",
            "id": "P2",
            "findings": [
                {
                    "id": "P2.F1",
                    "desc": "In human neurons, PUFA-cholesteryl ester-loaded ApoE displays an allelic series (ApoE4 > ApoE3 > ApoE2) driving lysosomal lipofuscinosis and lysosomal tau fibril accumulation.",
                    "quote": "In human neurons, the addition of ApoE carrying polyunsaturated fatty acids (PUFAs)-CE revealed an allelic series (ApoE4 > ApoE3 > ApoE2) associated with lipofuscinosis, an age-related lysosomal pathology resulting from lipid peroxidation.",
                    "summary": "(APOE3 -> APOE4) -> (increased lysosomal lipofuscinosis & tau fibril accumulation)",
                    "relevance": 0.85,
                    "exp_system": "Human neurons + APOE4 mice. Caveat: Abstract-only; measures lysosomal lipofuscin/tau accumulation downstream of APOE4.",
                    "data_location": "Abstract (full text paywalled)",
                    "effect_size": "N/A", "effect_prov": "full_text_unavailable",
                    "p_value": "N/A", "p_prov": "full_text_unavailable",
                    "sample_size": "N/A", "sample_prov": "full_text_unavailable",
                    "directness": "3", "strength": "3 moderate",
                }
            ],
        },
        {
            "pmid": "41103078",
            "doi": "10.1080/15548627.2025.2576613",
            "source_type": "PubMed published",
            "id": "P3",
            "findings": [
                {
                    "id": "P3.F1",
                    "desc": "Quantitative LysoIP proteomics shows HsAPOE4-expressing lysosomes accumulate APOE itself alongside AD-associated secretases (PSEN2, CD63, TSPAN3) relative to HsAPOE3 and HsAPOE-/- controls.",
                    "quote": "Through this approach, we found that HsAPOE4-expressing lysosomes accumulated APOE itself, AD-associated secretase proteins (PSEN2, CD63, TSPAN3)",
                    "summary": "(HsAPOE3 -> HsAPOE4) -> (increased APOE accumulation inside neuronal lysosomes)",
                    "relevance": 0.82,
                    "exp_system": "Neuro-2a cells expressing human APOE4 vs HsAPOE3 genomic locus; TMEM192-3xHA LysoIP proteomics. Caveat: measures lysosomal APOE abundance, not fibrillar aggregation.",
                    "data_location": "Fig 3C (PMC12758178)",
                    "effect_size": "qualitative", "effect_prov": "printed_in_text",
                    "p_value": "N/A", "p_prov": "not_found_in_source",
                    "sample_size": "N/A", "sample_prov": "not_found_in_source",
                    "directness": "3", "strength": "3 moderate",
                }
            ],
        },
    ]

    h4 = [
        {
            "pmid": "30375977",
            "doi": "10.7554/eLife.40048",
            "source_type": "PubMed published",
            "id": "P1",
            "findings": [
                {
                    "id": "P1.F1",
                    "desc": "Cell-derived ApoE4 shows prolonged retention compared with ApoE3 in Apoer2-containing intracellular compartments and impairs Apoer2 surface recycling.",
                    "quote": "The delayed recycling of Apoer2 was also apparent by the prolonged retention of cell-derived ApoE4 compared to cell-derived ApoE3 in Apoer2-containing intracellular compartments as shown by co-immunoprecipitation of neuronal lysates (Figure 2D,E).",
                    "summary": "(APOE3 -> APOE4) -> (prolonged ApoE4-Apoer2 intracellular retention -> recycling block)",
                    "relevance": 0.96,
                    "exp_system": "Primary rat cortical neurons treated with naturally secreted human ApoE3 vs ApoE4; co-IP & surface biotinylation. Caveat: rodent neurons, not human cells.",
                    "data_location": "Fig 2D-E (PMC6261251)",
                    "effect_size": "qualitative", "effect_prov": "printed_in_text",
                    "p_value": "N/A", "p_prov": "not_found_in_source",
                    "sample_size": "N/A", "sample_prov": "not_found_in_source",
                    "directness": "1", "strength": "2 strong",
                }
            ],
        },
        {
            "pmid": "33930462",
            "doi": "10.1016/j.jbc.2021.100715",
            "source_type": "PubMed published",
            "id": "P2",
            "findings": [
                {
                    "id": "P2.F1",
                    "desc": "ApoE4 is the most potent isoform inhibitor of LRP1-mediated tau uptake, consistent with higher functional affinity/competition for LRP1.",
                    "quote": "LRP1-mediated uptake of tau is inhibited by apoE, with the apoE4 isoform being the most potent inhibitor, likely because of its higher affinity for LRP1.",
                    "summary": "(APOE3 -> APOE4) -> (stronger LRP1 receptor binding / inhibition of LRP1-mediated tau uptake)",
                    "relevance": 0.78,
                    "exp_system": "LRP1-expressing vs LRP1-deficient cells + SPR binding assays. NOTE: Abstract-only (OA XML unavailable).",
                    "data_location": "Abstract (full text paywalled)",
                    "effect_size": "N/A", "effect_prov": "full_text_unavailable",
                    "p_value": "N/A", "p_prov": "full_text_unavailable",
                    "sample_size": "N/A", "sample_prov": "full_text_unavailable",
                    "directness": "3", "strength": "3 moderate",
                }
            ],
        },
        {
            "pmid": "31796678",
            "doi": "10.3233/JAD-190914",
            "source_type": "PubMed published",
            "id": "P3",
            "findings": [
                {
                    "id": "P3.F1",
                    "desc": "ApoE4 attenuates binding of sorting nexin 17 (SNX17) to ApoER2 and inhibits ApoER2 recycling to the cell surface; carnosic acid restores SNX17-ApoER2 binding and surface ApoER2.",
                    "quote": "In this study, we demonstrated that ApoE4 attenuates the binding of sorting nexin 17 (SNX17) to ApoER2 and inhibits the recycling of ApoER2, resulting in decreased cell surface level of ApoER2.",
                    "summary": "(APOE4) -> (reduced SNX17-ApoER2 binding -> impaired ApoER2 recycling)",
                    "relevance": 0.85,
                    "exp_system": "Neuronal cell culture model. NOTE: Abstract-only (OA XML unavailable).",
                    "data_location": "Abstract (full text paywalled)",
                    "effect_size": "N/A", "effect_prov": "full_text_unavailable",
                    "p_value": "N/A", "p_prov": "full_text_unavailable",
                    "sample_size": "N/A", "sample_prov": "full_text_unavailable",
                    "directness": "2", "strength": "3 moderate",
                }
            ],
        },
        {
            "pmid": "38217595",
            "doi": "10.3233/JAD-230514",
            "source_type": "PubMed published",
            "id": "P4",
            "findings": [
                {
                    "id": "P4.F1",
                    "desc": "APOE4 neurons show significantly lower surface-to-total ratios of APOER2 (53% reduction, p < 0.05) and LRP1 (48% reduction) compared with APOE3 neurons.",
                    "quote": "The results obtained show that the ratio of the externally oriented receptors to the total, was significantly lower in APOE4 neurons of APOER2 (Fig. 4A, 53% reduction, p < 0.05) and LRP1",
                    "summary": "(APOE3 -> APOE4) -> (53% reduced surface APOER2 / 48% reduced surface LRP1)",
                    "relevance": 0.80,
                    "exp_system": "Primary APOE4 vs APOE3 mouse cortical neurons (N=3 biotinylation preparations). Caveat: measures surface receptor depletion rather than direct receptor-ApoE4 adhesion affinity.",
                    "data_location": "Fig 4A-B (PMC10894586)",
                    "effect_size": "53%", "effect_prov": "printed_in_text",
                    "p_value": "<0.05", "p_prov": "printed_in_text",
                    "sample_size": "3", "sample_prov": "printed_in_text",
                    "directness": "3", "strength": "3 moderate",
                }
            ],
        },
    ]

    h5 = [
        {
            "pmid": "37777962",
            "doi": "10.1016/j.celrep.2023.113183",
            "source_type": "PubMed published",
            "id": "P1",
            "findings": [
                {
                    "id": "P1.F1",
                    "desc": "In isogenic human iPSC-derived astrocytes, ApoE4-induced lysosomal cholesterol accumulation impairs lysosome-dependent mitophagy and oxidative phosphorylation; cholesterol depletion rescues autophagy.",
                    "quote": "Using isogenic ApoE3 and ApoE4 astrocytes derived from human induced pluripotent stem cells, we find that ApoE4 increases glycolytic activity but impairs mitochondrial respiration in astrocytes. Ultrastructural and autophagy flux analyses show that ApoE4-induced cholesterol accumulation impairs lysosome-dependent removal of damaged mitochondria.",
                    "summary": "(APOE3 -> APOE4 in human iPSC astrocytes) -> (lysosomal cholesterol accumulation -> impaired lysosomal mitophagy)",
                    "relevance": 0.98,
                    "exp_system": "Isogenic human iPSC-derived ApoE3 vs ApoE4 astrocytes. NOTE: Abstract-only (OA XML unavailable).",
                    "data_location": "Abstract (full text paywalled)",
                    "effect_size": "N/A", "effect_prov": "full_text_unavailable",
                    "p_value": "N/A", "p_prov": "full_text_unavailable",
                    "sample_size": "N/A", "sample_prov": "full_text_unavailable",
                    "directness": "1", "strength": "1 very strong",
                }
            ],
        },
        {
            "pmid": "42182449",
            "doi": "10.64898/2026.05.12.724612",
            "source_type": "PubMed preprint",
            "id": "P2",
            "findings": [
                {
                    "id": "P2.F1",
                    "desc": "Isogenic human APOE44 iPSC-derived microglia accumulate cholesterol esters in lysosomes (BODIPY/LysoTracker colocalization), downregulate 53 of 84 lysosomal hydrolases/V-ATPases, and show impaired lysosomal acidification.",
                    "quote": "Of the 84 identified lysosomal hydrolases and V-ATPases, 53 were significantly downregulated in APOE44 iMG, supporting lysosomal impairment (Figure S3D).",
                    "summary": "(APOE33 -> APOE44 in human iPSC microglia) -> (lysosomal cholesterol ester accumulation -> lysosomal hydrolase/V-ATPase downregulation)",
                    "relevance": 0.92,
                    "exp_system": "Human isogenic APOE44 vs APOE33 iPSC-derived microglia (iMG; Jackson Lab & UCI5 lines). Direct human brain-cell model.",
                    "data_location": "Fig 3E, Fig S3D, Fig 7M (PMC13192593)",
                    "effect_size": "qualitative", "effect_prov": "printed_in_text",
                    "p_value": "N/A", "p_prov": "not_found_in_source",
                    "sample_size": "3-4", "sample_prov": "printed_in_text",
                    "directness": "1", "strength": "2 strong",
                }
            ],
        },
        {
            "pmid": "42636814",
            "doi": "10.1016/j.stem.2026.08.001",
            "source_type": "PubMed published",
            "id": "P3",
            "findings": [
                {
                    "id": "P3.F1",
                    "desc": "In 3D human iPSC miBrains, cholesterol accumulation in APOE4/4 astrocytes causes endolysosomal dysfunction that impairs soluble α-synuclein degradation and seeds neuronal inclusions.",
                    "quote": "Combinatorial experiments revealed that endolysosomal dysfunction caused by cholesterol accumulation in APOE4/4 astrocytes impairs the degradation of soluble α-synuclein leading to a pathogenic transformation that seeds α-synuclein inclusions in neurons.",
                    "summary": "(APOE4/4 in human astrocytes) -> (cholesterol accumulation -> endolysosomal dysfunction -> impaired α-synuclein degradation)",
                    "relevance": 0.90,
                    "exp_system": "Human iPSC-derived 3D multicellular brain-like tissue (miBrain). NOTE: Abstract-only (OA XML unavailable).",
                    "data_location": "Abstract (full text paywalled)",
                    "effect_size": "N/A", "effect_prov": "full_text_unavailable",
                    "p_value": "N/A", "p_prov": "full_text_unavailable",
                    "sample_size": "N/A", "sample_prov": "full_text_unavailable",
                    "directness": "1", "strength": "2 strong",
                }
            ],
        },
        {
            "pmid": "40675218",
            "doi": "10.1016/j.jbc.2025.110479",
            "source_type": "PubMed published",
            "id": "P4",
            "findings": [
                {
                    "id": "P4.F1",
                    "desc": "CONTRADICTING / CONDITIONAL: Without external aging or cholesterol triggers, mature primary mouse neurons from ApoE-KO, ApoE3, and ApoE4 mice show no baseline differences in endosomal appearance or function; reduced lysosomal degradation emerges only after prolonged time in culture.",
                    "quote": "We show that without external triggers, mature primary neurons from ApoE knockout (KO), ApoE3, and ApoE4 mice show no major differences in endosomal appearance and function and adapt similarly to increased synaptic activity.",
                    "summary": "(APOE3 -> APOE4 in non-aged primary neurons) -> (no baseline endolysosomal difference without prolonged culture/cholesterol) [CONSTRAINS H5]",
                    "relevance": 0.85,
                    "exp_system": "Primary hippocampal/cortical cultures from human ApoE4-TR, ApoE3-TR, and ApoE-KO mice (18 DIV vs prolonged culture). Caveat: mouse neurons.",
                    "data_location": "Abstract & Fig 1-2 (PMC12357307)",
                    "effect_size": "qualitative", "effect_prov": "printed_in_text",
                    "p_value": "N/A", "p_prov": "not_found_in_source",
                    "sample_size": "N/A", "sample_prov": "not_found_in_source",
                    "directness": "1", "strength": "2 strong",
                }
            ],
        },
    ]

    return {"H1": build_h1(), "H2": h2, "H3": h3, "H4": h4, "H5": h5}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--method-url", required=True)
    ap.add_argument("--commit-sha", required=True)
    args = ap.parse_args()

    papers_by_pmid = load_all_papers()
    all_data = get_curated_data()
    summary = {}

    for hid, curated in all_data.items():
        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        ws = wb["Submission"]
        if ws.max_row > 2:
            ws.delete_rows(3, ws.max_row - 2)
        ws.cell(2, 1).value = HYPOTHESIS_STATEMENTS[hid]

        row_idx = 3
        f_count = 0
        for entry in curated:
            pmid = entry["pmid"]
            assert pmid in papers_by_pmid, f"{hid}: PMID {pmid} missing!"
            pm = papers_by_pmid[pmid]
            assert not pm["retracted"], f"{hid}: PMID {pmid} is RETRACTED!"
            assert pm["doi"] == entry["doi"], f"{hid}: DOI mismatch for {pmid}: {pm['doi']} vs {entry['doi']}"

            ws.cell(row_idx, 2).value = entry["doi"]
            ws.cell(row_idx, 3).value = entry["source_type"]
            ws.cell(row_idx, 4).value = int(pmid)
            ws.cell(row_idx, 5).value = entry["id"]
            row_idx += 1

            for f in entry["findings"]:
                verify_quote(pm, f["quote"])
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
                f_count += 1

        ws_m = wb["Method"]
        method_values = {
            "Team": "FdNIA-CoResearcher",
            "Team leader": "Manuel Menendez-Gonzalez",
            "Team leader email": "administracion@fneurociencias.org",
            "Method URL": args.method_url,
            "Commit SHA": args.commit_sha,
            "Who did the work": "both",
            "Model and version": "CoResearcher + EdiTXT bibliographic verifier + Gemini 4 Hydrogen",
        }
        for r in range(1, ws_m.max_row + 1):
            lbl = ws_m.cell(r, 1).value
            if lbl in method_values:
                ws_m.cell(r, 2).value = method_values[lbl]

        out_xlsx = BASE_DIR / f"submission_{hid}.xlsx"
        wb.save(out_xlsx)

        res = subprocess.run(
            [sys.executable, str(VALIDATOR_PATH), str(out_xlsx), "--strict"],
            capture_output=True,
            text=True,
        )
        assert res.returncode == 0, f"Validator failed for {hid}:\n{res.stdout}\n{res.stderr}"
        sha256 = hashlib.sha256(out_xlsx.read_bytes()).hexdigest()
        summary[hid] = {
            "file": str(out_xlsx),
            "papers": len(curated),
            "findings": f_count,
            "sha256": sha256,
            "valid_strict": True,
        }
        print(f"[{hid}] VALID_STRICT | papers={len(curated)} findings={f_count} | {out_xlsx.name}")

    (BASE_DIR / "all_submissions_manifest.json").write_text(
        json.dumps({"method_url": args.method_url, "commit_sha": args.commit_sha, "submissions": summary}, indent=2),
        encoding="utf-8",
    )
    print("ALL 5 SUBMISSIONS VALIDATED STRICTLY!")


if __name__ == "__main__":
    main()
