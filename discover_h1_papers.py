"""Real PubMed + Europe PMC literature discovery & full-text retrieval for MecCog H1."""

import json
import os
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET

OUT_DIR = r"c:\Users\usuario\Neurodiagnoses Ecosystem\proposals\MecCog_APOE4_Challenge_2026\cache_h1"
os.makedirs(OUT_DIR, exist_ok=True)

QUERIES = [
    '(APOE4[Title/Abstract] OR "ApoE4"[Title/Abstract] OR "apolipoprotein E4"[Title/Abstract]) AND ("endosome pH"[Title/Abstract] OR "endosomal pH"[Title/Abstract] OR "endosomal acidification"[Title/Abstract] OR "lysosomal pH"[Title/Abstract] OR "NHE6"[Title/Abstract] OR "V-ATPase"[Title/Abstract])',
    '(APOE4[Title/Abstract] OR "ApoE4"[Title/Abstract]) AND (endosom*[Title/Abstract] AND pH[Title/Abstract])',
    '("APOE4"[Title/Abstract] OR "ApoE4"[Title/Abstract]) AND ("endolysosomal"[Title/Abstract] OR "endo-lysosomal"[Title/Abstract]) AND ("pH"[Title/Abstract] OR "acidification"[Title/Abstract])',
]

all_pmids = ["29946028"]  # Always include starting reference
query_logs = []

for q in QUERIES:
    url = (
        "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?db=pubmed&term="
        + urllib.parse.quote(q)
        + "&retmax=50&retmode=json"
    )
    data = json.loads(
        urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=15)
        .read()
        .decode("utf-8")
    )
    ids = data["esearchresult"]["idlist"]
    query_logs.append({"query": q, "total_count": int(data["esearchresult"]["count"]), "returned": len(ids)})
    for i in ids:
        if i not in all_pmids:
            all_pmids.append(i)

print(f"Total unique PMIDs discovered: {len(all_pmids)}")

# Fetch full PubMed XML (abstracts, pubtypes, DOIs, PMCIDs)
efetch_url = (
    "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&id="
    + ",".join(all_pmids)
    + "&retmode=xml"
)
xml_bytes = urllib.request.urlopen(
    urllib.request.Request(efetch_url, headers={"User-Agent": "Mozilla/5.0"}), timeout=25
).read()
root = ET.fromstring(xml_bytes)

papers = []
for art in root.findall(".//PubmedArticle"):
    pmid = art.findtext("./MedlineCitation/PMID")
    title = "".join(art.find("./MedlineCitation/Article/ArticleTitle").itertext()) if art.find("./MedlineCitation/Article/ArticleTitle") is not None else ""
    journal = art.findtext("./MedlineCitation/Article/Journal/ISOAbbreviation") or ""
    year = art.findtext("./MedlineCitation/Article/Journal/JournalIssue/PubDate/Year") or art.findtext("./MedlineCitation/Article/Journal/JournalIssue/PubDate/MedlineDate") or ""
    pub_types = [pt.text for pt in art.findall("./MedlineCitation/Article/PublicationTypeList/PublicationType") if pt.text]
    abs_parts = []
    for ab in art.findall("./MedlineCitation/Article/Abstract/AbstractText"):
        lbl = ab.attrib.get("Label", "")
        txt = "".join(ab.itertext()).strip()
        abs_parts.append(f"{lbl}: {txt}" if lbl else txt)
    abstract = "\n".join(abs_parts)
    aids = {a.attrib.get("IdType"): a.text for a in art.findall("./PubmedData/ArticleIdList/ArticleId")}
    doi = aids.get("doi", "")
    pmc = aids.get("pmc", "")
    retracted = any("retracted" in pt.lower() for pt in pub_types)

    # Try Europe PMC full-text XML if PMCID exists
    fulltext_path = ""
    if pmc:
        epmc_url = f"https://www.ebi.ac.uk/europepmc/webservices/rest/{pmc}/fullTextXML"
        try:
            ft_bytes = urllib.request.urlopen(
                urllib.request.Request(epmc_url, headers={"User-Agent": "Mozilla/5.0"}), timeout=20
            ).read()
            if b"<article" in ft_bytes or b"<body" in ft_bytes:
                ft_file = os.path.join(OUT_DIR, f"{pmc}.xml")
                with open(ft_file, "wb") as fp:
                    fp.write(ft_bytes)
                fulltext_path = ft_file
        except Exception:
            pass

    papers.append({
        "pmid": pmid,
        "pmcid": pmc,
        "doi": doi,
        "year": year,
        "journal": journal,
        "title": title,
        "pub_types": pub_types,
        "retracted": retracted,
        "abstract": abstract,
        "fulltext_xml": fulltext_path,
    })
    print(f"PMID {pmid} | {year} | PMC={pmc or 'NONE'} | OA_XML={bool(fulltext_path)} | Retracted={retracted} | {title[:90]}")

with open(os.path.join(OUT_DIR, "h1_papers_metadata.json"), "w", encoding="utf-8") as fp:
    json.dump({"queries": query_logs, "papers": papers}, fp, indent=2, ensure_ascii=False)
print("Saved metadata to cache_h1/h1_papers_metadata.json")
