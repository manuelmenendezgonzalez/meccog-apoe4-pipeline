import json, os, urllib.parse, urllib.request, xml.etree.ElementTree as ET

cache_dir = r"c:\Users\usuario\Neurodiagnoses Ecosystem\proposals\MecCog_APOE4_Challenge_2026\cache_h1"

# 1. Try NCBI PMC efetch for PMC6048470, PMC7614123, PMC5936830
for pmcid in ["PMC6048470", "PMC7614123", "PMC5936830"]:
    num = pmcid.replace("PMC", "")
    url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pmc&id={num}&retmode=xml"
    try:
        data = urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=20).read()
        has_body = b"<body" in data
        print(pmcid, "len=", len(data), "has_body=", has_body)
        if has_body:
            with open(os.path.join(cache_dir, f"{pmcid}.xml"), "wb") as f:
                f.write(data)
    except Exception as e:
        print(pmcid, "ERR:", e)

# 2. Broader PubMed search
q = '("APOE4"[Title/Abstract] OR "ApoE4"[Title/Abstract] OR "apolipoprotein E4"[Title/Abstract]) AND (endosom*[Title/Abstract] OR lysosom*[Title/Abstract]) AND (pH[Title/Abstract] OR acidif*[Title/Abstract] OR alkalin*[Title/Abstract] OR NHE6[Title/Abstract] OR "V-ATPase"[Title/Abstract])'
url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?db=pubmed&term=" + urllib.parse.quote(q) + "&retmax=60&retmode=json"
data = json.loads(urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=15).read().decode("utf-8"))
ids = data["esearchresult"]["idlist"]
print("Broad query count:", data["esearchresult"]["count"], "returned:", len(ids))

# Fetch summaries for new ones
with open(os.path.join(cache_dir, "h1_papers_metadata.json"), "r", encoding="utf-8") as f:
    meta = json.load(f)
existing = {p["pmid"] for p in meta["papers"]}
new_ids = [i for i in ids if i not in existing]
print("New PMIDs:", new_ids)

if new_ids:
    efetch_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&id=" + ",".join(new_ids) + "&retmode=xml"
    xml_bytes = urllib.request.urlopen(urllib.request.Request(efetch_url, headers={"User-Agent": "Mozilla/5.0"}), timeout=25).read()
    root = ET.fromstring(xml_bytes)
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
            try:
                epmc_url = f"https://www.ebi.ac.uk/europepmc/webservices/rest/{pmc}/fullTextXML"
                ft_bytes = urllib.request.urlopen(urllib.request.Request(epmc_url, headers={"User-Agent": "Mozilla/5.0"}), timeout=20).read()
                if b"<body" in ft_bytes:
                    ft_path = os.path.join(cache_dir, f"{pmc}.xml")
                    with open(ft_path, "wb") as fp:
                        fp.write(ft_bytes)
            except Exception:
                pass
        meta["papers"].append({
            "pmid": pmid, "pmcid": pmc, "doi": doi, "year": year,
            "journal": journal, "title": title, "pub_types": pub_types,
            "retracted": retracted, "abstract": abstract, "fulltext_xml": ft_path
        })
        print(f"NEW PMID {pmid} | {year} | PMC={pmc or 'NONE'} | OA_XML={bool(ft_path)} | {title[:85]}")
    meta["queries"].append({"query": q, "total_count": int(data["esearchresult"]["count"]), "returned": len(ids)})
    with open(os.path.join(cache_dir, "h1_papers_metadata.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2, ensure_ascii=False)
