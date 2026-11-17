# Method — FdNIA-CoResearcher (MecCog APOE4 Challenge, Step 1)

- **Team:** FdNIA-CoResearcher (Fundación de Neurociencias, Oviedo, Spain)
- **Team Leader:** Dr. Manuel Menéndez-González (`administracion@fneurociencias.org`)
- **Hypotheses covered:** H1, H2, H3, H4, H5 (all five APOE4 endo-lysosomal mechanism hypotheses)
- **Submission date:** 2026-11-17

---

## 1. Who did the work

A human-and-agent collaborative workflow (`both`) combining deterministic Python literature retrieval/verification scripts (`discover_h1_papers.py`, `discover_h1_broad.py`, `discover_h2_h5.py`, `generate_all_submissions.py`), bibliographic/retraction checks modeled on **EdiTXT**, formal hypothesis-polarity and directness scoring modeled on **CoResearcher**, and AI-assisted curation. Automated scripts executed all PubMed/Europe PMC queries, JATS XML parsing, verbatim quote substring verification, and strict schema validation; Dr. Manuel Menéndez-González and the NeuroIA team (Fundación de Neurociencias) reviewed the pipeline design and extracted outputs.

## 2. Model and version

- **Orchestration & Curation Agent:** Antigravity IDE agent running **Gemini 4 Hydrogen (Medium)** (Windows 11, Python 3.13).
- **Verification & Pipeline Layer:** Deterministic Python scripts (`discover_h1_papers.py`, `discover_h1_broad.py`, `discover_h2_h5.py`, `generate_all_submissions.py`) validated with the official MecCog `validate_submission.py --strict` (v2, revised 2026-10-26).

## 3. How papers were found

We queried NCBI PubMed E-utilities (`esearch.fcgi` + `efetch.fcgi`) on 2026-11-17 using targeted Boolean queries for each hypothesis (H1–H5), always seeding with the official starting reference (`PMID: 29946028` for H1, `31641056` for H2, `39419029` for H3, `30375977` for H4, `37777962` for H5). Review articles were filtered out during experimental curation so that only primary experimental (and direct mechanistic) papers were retained.

## 4. How full text was obtained

1. For every discovered paper with a PMCID, we queried the **Europe PMC REST API** (`https://www.ebi.ac.uk/europepmc/webservices/rest/{PMCID}/fullTextXML`).
2. When Europe PMC lacked `<body/>` XML, we fell back to **NCBI PMC E-utilities** (`efetch.fcgi?db=pmc&id=...&retmode=xml`).
3. Across all hypotheses, open-access JATS XML full texts were downloaded whenever available (`PMC6261251`, `PMC8547963`, `PMC12758178`, `PMC13192593`, `PMC11792374`, `PMC10894586`, `PMC12357307`, `PMC13003211`, etc.). For paywalled papers lacking OA full-text XML via API (such as `PMID: 29946028`, `31641056`, `39419029`, `37777962`), extraction relied strictly on the PubMed abstract and missing statistics were tagged `full_text_unavailable`.

## 5. What a human checked

Every `PMID`, `DOI`, retraction flag (`retracted == False`), and verbatim `Finding quote` was verified programmatically via exact substring matching against downloaded abstracts or OA XMLs, and every `.xlsx` passed `validate_submission.py --strict` with zero errors and zero warnings. Human review focused on verifying hypothesis polarity (including contradicting/conditional evidence) and directness classifications against `Directness_of_evidence_v02.docx`.

## 6. Anything else

- **Contradicting & Conditional Evidence:** We explicitly included contradicting/opposing evidence where APOE4 induces *lysosomal alkalinization* (`PMID: 41103078`, `42182449`) in H1, rescue-direction acidification (`PMID: 30375977`, `34617884`), and conditional baseline null findings (`PMID: 40675218` in H5 showing no baseline endolysosomal difference in mature primary neurons without prolonged culture/cholesterol).
- **Preprint Deduplication:** Preprints superseded by peer-reviewed journal versions (`PMID: 37873080` -> `41103078`) were deduplicated to avoid double counting.

---

## Code and prompts

- `discover_h1_papers.py` & `discover_h1_broad.py` — PubMed & Europe PMC/NCBI PMC discovery for H1.
- `discover_h2_h5.py` — PubMed & Europe PMC discovery for H2–H5.
- `meccog_h1_pipeline.py` — H1 curation and strict validation module.
- `generate_all_submissions.py` — Unified builder, verbatim-quote & retraction verifier, and `--strict` validator for `submission_H1.xlsx` through `submission_H5.xlsx`.
