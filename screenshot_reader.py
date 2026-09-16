import re
from datetime import date

RANGE = re.compile(r"(20\d{2})[/.-](\d{1,2})[/.-](\d{1,2})[~–—-](20\d{2})[/.-](\d{1,2})[/.-](\d{1,2})")
def norm(value):
    return "".join(x for x in re.split(r"[^A-Z0-9]+", value.upper()) if x and x != "RG")

def analyze(words, size):
    tokens = [dict(text=r["text"].strip(), x=int(r["left"])/2, y=int(r["top"])/2,
                   w=int(r["width"])/2,h=int(r["height"])/2,conf=float(r.get("conf",0)))
              for r in words if r.get("text","").strip()]
    lines=[]
    for t in sorted(tokens,key=lambda t:(t["y"],t["x"])):
        line=next((l for l in lines if abs(l[0]["y"]-t["y"])<=max(5,t["h"]*.5)),None)
        if line is None: lines.append([t])
        else: line.append(t)
    titles=[]
    for line in lines:
        line.sort(key=lambda t:t["x"])
        text=" ".join(t["text"] for t in line)
        match=RANGE.search(re.sub(r"\s+","",text))
        if match and re.search(r"speed\s+summary",text,re.I): titles.append((line,match))
    if len(titles)!=1:
        raise ValueError("Tanggal dan Speed Summary tidak terbaca tunggal. Sertakan satu grafik lengkap.")
    title,match=titles[0]
    vals=[int(v) for v in match.groups()]
    start,end=date(*vals[:3]),date(*vals[3:])
    if end<start: raise ValueError("Rentang tanggal OCR tidak valid.")
    top=min(t["y"] for t in title)
    # Header suggestions require a unique sheet match and user confirmation.
    # Correct device headers can have low OCR confidence (WAGUBSU: 43.66).
    candidates={t["text"].strip("•●:") for t in tokens if t["y"]<top and t["conf"]>=0
                and re.fullmatch(r"\d{1,3}[-–][A-Za-z0-9]+(?:[-–][A-Za-z0-9]+)+",t["text"].strip("•●:"))}
    links=set()
    for desc in tokens:
        if desc["text"].rstrip(":").casefold()!="description" or desc["conf"]<75 or desc["y"]>=top: continue
        for t in tokens:
            if t["x"]>desc["x"]+desc["w"] and abs(t["y"]-desc["y"])<=max(6,desc["h"]) and t["conf"]>=75:
                if t["text"].casefold() in {"metro","broadband"}: links.add(t["text"].upper())
    empty=[l for l in lines if min(t["y"] for t in l)>top and re.search(r"\bno\s+data\b"," ".join(t["text"] for t in l),re.I)]
    legends=[t for t in tokens if t["y"]>top and t["text"].casefold() in {"uplink","downlink"}]
    if empty: bottom=max(t["y"]+t["h"] for l in empty for t in l)+48
    elif {t["text"].casefold() for t in legends}=={"uplink","downlink"}:
        bottom=max(t["y"]+t["h"] for t in legends)+8
    else: raise ValueError("Batas grafik tidak terbaca. Sertakan legenda Uplink/Downlink atau tulisan No Data.")
    box=(max(0,int(min(t["x"] for t in title))-8),max(0,int(top)-12),size[0],min(size[1],int(bottom)))
    return {"date":end,"box":box,"no_data":bool(empty),"projects":candidates,"links":links}

# Explicit alias verified against the supplied device and workbook screenshots.
# Do not match by project number alone: different devices may share a prefix.
PROJECT_ALIASES = {
    norm("03-RG-WAGUBSU-SERVER"): norm("03-RUDIN WAGUBSU"),
}


def suggest_selection(info, sheets):
    """Return a unique exact/verified-alias suggestion; never choose by number only."""
    project = ""
    if len(info["projects"]) == 1:
        candidate = norm(next(iter(info["projects"])))
        matches = [name for name in sheets if norm(name) == candidate]
        if not matches and candidate in PROJECT_ALIASES:
            matches = [name for name in sheets
                       if norm(name) == PROJECT_ALIASES[candidate]]
        if len(matches) == 1:
            project = matches[0]
    link = next(iter(info["links"])) if len(info["links"]) == 1 else ""
    return project, link


def resolve_selection(info, sheets):
    project, link = suggest_selection(info, sheets)
    return (project, link) if project and link else None
