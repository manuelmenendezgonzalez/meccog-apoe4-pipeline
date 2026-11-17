"""Real PubMed + Europe PMC literature discovery & OA XML fetcher for MecCog H2, H3, H4, H5."""

import json
import os
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET

OUT_DIR = r"c:\Users\usuario\Neurodiagnoses Ecosystem\proposals\MecCog_APOE4_Challenge_2026\cache_all"
os.makedirs(OUT_DIR, exist_ok=True)

HYPOTHESES_SEARCH = {
    "H2": {
        "starting_pmid": "31641056",
        "queries": [
            '("APOE4"[Title/Abstract] OR "ApoE4"[Title/Abstract]) AND ("ARF6"[Title/Abstract] OR "ABCA1"[Title/Abstract]) AND (recycl*[Title/Abstract] OR trafficking[Title/Abstract])',
            '("APOE4"[Title/Abstract] OR "ApoE4"[Title/Abstract]) AND (endosom*[Title/Abstract] AND recycl*[Title/Abstract]) AND (astrocyte*[Title/Abstract] OR receptor*[Title/Abstract])',
        ],
    },
    "H3": {
        "starting_pmid": "39419029",
        "queries": [
            '("APOE4"[Title/Abstract] OR "ApoE4"[Title/Abstract]) AND (aggregat*[Title/Abstract] OR fibril*[Title/Abstract] OR oligomer*[Title/Abstract]) AND (lysosom*[Title/Abstract] OR endolysosom*[Title/Abstract] OR endosom*[Title/Abstract])',
            '("APOE4"[Title/Abstract] OR "ApoE4"[Title/Abstract]) AND ("microglia"[Title/Abstract] OR "brain"[Title/Abstract]) AND (aggregat*[Title/Abstract] AND "APOE3"[Title/Abstract])',
        ],
    },
    "H4": {
        "starting_pmid": "30375977",
        "queries": [
            '("APOE4"[Title/Abstract] OR "ApoE4"[Title/Abstract]) AND ("Apoer2"[Title/Abstract] OR "LRP8"[Title/Abstract] OR "LDLR"[Title/Abstract] OR "LRP1"[Title/Abstract]) AND (dissociat*[Title/Abstract] OR recycl*[Title/Abstract] OR affinity[Title/Abstract] OR trap*[Title/Abstract])',
        ],
    },
    "H5": {
        "starting_pmid": "37777962",
        "queries": [
            '("APOE4"[Title/Abstract] OR "ApoE4"[Title/Abstract]) AND ("cholesterol"[Title/Abstract] OR "lipid"[Title/Abstract]) AND ("lysosom*"[Title/Abstract] OR "late endosome"[Title/Abstract]) AND (astrocyte*[Title/Abstract] OR microglia*[Title/Abstract] OR neuron*[Title/Abstract])',
        ],
    },
}

results = {}

for hid, cfg in HYPOTHESES_SEARCH.items():
    pmids = [cfg["starting_pmid"]]
    qlogs = []
    for q in cfg["queries"]:
        url = (
            "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?db=pubmed&term="
            + urllib.parse.quote(q)
            + "&retmax=25&retmode=json"
        )
        data = json.loads(
            urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=15)
            .read()
            .decode("utf-8")
        )
        ids = data["esearchresult"]["idlist"]
        qlogs.append({"query": q, "count": int(data["esearchresult"]["count"]), "returned": len(ids)})
        for i in ids:
            if i not in pmids:
                pmids.append(i)

    # Cap to top 12 per hypothesis for focused precision
    pmids = pmids[:12]
    efetch_url = (
        "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&id="
        + ",".join(pmids)
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
        year = art.findtext("./MedlineCitation/Article/Journal/JournalIssue/PubDate/Year") or ""
        pub_types = [pt.text for pt in art.findall("./MedlineCitation/Article/PublicationTypeList/PublicationType") if pt.text]
        abs_parts = ["".join(ab.itertext()).strip() for ab in art.findall("./MedlineCitation/Article/Abstract/AbstractText")]
        abstract = "\n".join(abs_parts)
        aids = {a.attrib.get("IdType"): a.text for a in art.findall("./PubmedData/ArticleIdList/ArticleId")}
        doi = aids.get("doi", "")
        pmc = aids.get("pmc", "")
        retracted = any("retracted" in pt.lower() for pt in pub_types)
        ft_path = ""
        if pmc:
            epmc_url = f"https://www.ebi.ac.uk/europepmc/webservices/rest/{pmc}/fullTextXML"
            try:
                ft_bytes = urllib.request.urlopen(
                    urllib.request.Request(epmc_url, headers={"User-Agent": "Mozilla/5.0"}), timeout=15
                ).read()
                if b"<body" in ft_bytes:
                    ft_path = os.path.join(OUT_DIR, f"{pmc}.xml")
                    with open(ft_path, "wb") as fp:
                        fp.write(ft_bytes)
            except Exception:
                pass
        papers.append({
            "pmid": pmid, "pmcid": pmc, "doi": doi, "year": year,
            "journal": journal, "title": title, "pub_types": pub_types,
            "retracted": retracted, "abstract": abstract, "fulltext_xml": ft_path,
        })
        print(f"[{hid}] PMID {pmid} | {year} | PMC={pmc or 'NONE'} | OA={bool(ft_path)} | {title[:75]}")
    results[hid] = {"queries": qlogs, "papers": papers}

with open(os.path.join(OUT_DIR, "h2_h5_metadata.json"), "w", encoding="utf-8") as fp:
    json.dump(results, fp, indent=2, ensure_ascii=False)
print("Saved h2_h5_metadata.json")
