"""
My code retrives all the papers from 2022.
Sources:
  OpenAlex (https://api.openalex.org): main source, author profiles with affiliations
  Crossref (https://api.crossref.org): the DOI registry, catches DOI papers OpenAlex misses
  ORCID (https://pub.orcid.org): whatever members list on their own ORCID profile

What I took into account:
  Cecilia Zanni-Merk: only papers with at least one Azerbaijani co-author from UFAZ or ASOIU(another lab member, an Azerbaijan-affiliated author, or a typical Azerbaijani surname)
  Excluded ones are listed in the sheet "Excluded (Cecilia)" so they can be checked
  All other members: every paper since 2022, whatever the institution or co-authors

Output:
  publications.xlsx: sheet "Publications", newest first
  publications.csv: same data, plain text
  author_matches.csv: which author profile was used for each member
"""

import html
import re
import time 
import unicodedata
from pathlib import Path

import pandas as pd
import requests


FROM_YEAR = 2022
HERE = Path(__file__).resolve().parent
WEBSITE_LIST_FILE = HERE / "website_publications.txt"

# OpenAlex requires a free API key
OPENALEX_API_KEY = "PASTE_YOUR_KEY_HERE"
# Words that identify the lab's institutions in OpenAlex affiliation names.
# Used to pick the right author profile when several people share a name.
AFFILIATION_HINTS = ["ufaz", "french-azerbaijani", "azerbaijan", "baku","insa rouen", "litis", "strasbourg", "york",]

# Members: permanent staff + PhD candidates.
MEMBERS = [
    # Permanent staff
    {"name": "Cecilia Zanni-Merk",     "role": "staff", "variants": ["Cecilia Zanni Merk", "C. Zanni-Merk"], "openalex_ids": []},
    {"name": "Ulviyya Abdulkarimova",  "role": "staff", "variants": ["Ulviya Abdulkarimova", "U. Abdulkarimova"], "openalex_ids": []},
    {"name": "Nahla El Kadhi",         "role": "staff", "variants": ["Nahla Elkadhi", "N. El Kadhi"], "openalex_ids": []},
    {"name": "Yusif Ibrahimov",        "role": "staff", "variants": ["Y. Ibrahimov"], "openalex_ids": []},
    {"name": "Rza Mustafayev",         "role": "staff", "variants": ["R. Ch. Mustafayev", "Rza Ch. Mustafayev"], "openalex_ids": []},
    # PhD candidates
    {"name": "Kamran Ismayilzada",     "role": "phd",   "variants": ["Kamran Ismayilzade"], "openalex_ids": []},
    {"name": "Rauf Fatali",            "role": "phd",   "variants": ["R. Fatali"], "openalex_ids": []},
    {"name": "Ali Khudiyev",           "role": "phd",   "variants": ["A. Khudiyev"], "openalex_ids": []},
    {"name": "Pavel Kosov",            "role": "phd",   "variants": ["P. Kosov"], "openalex_ids": []},
    {"name": "Seymur Mirzabayov",      "role": "phd",   "variants": ["Seymour Mirzabayov", "S. E. Mirzabayov"], "openalex_ids": []},
]

MEMBERS_BY_NAME = {m["name"]: m for m in MEMBERS}

SESSION = requests.Session()
SESSION.headers["User-Agent"] = "UFAZ-CS-publications/1.0"

def get_json(url, params=None, headers=None, retries=4):
    """GET with retry/back-off. Prints the real reason if it finally fails."""
    for attempt in range(retries):
        wait = 3 * (attempt + 1)
        try:
            r = SESSION.get(url, params=params, headers=headers, timeout=30)
            if r.status_code == 429:                       # rate limit: wait and retry
                wait = int(r.headers.get("Retry-After", wait))
                raise requests.HTTPError(f"429 Too Many Requests (rate limit): {r.text[:150]}")
            if r.status_code >= 400:
                raise requests.HTTPError(f"{r.status_code}: {r.text[:150]}")
            try:
                return r.json()
            except ValueError:
                # the site answered, but not with JSON (e.g. an HTML error or "slow down" page)
                snippet = " ".join(r.text[:150].split())
                raise ValueError(f"not JSON (HTTP {r.status_code}, {r.headers.get('Content-Type')}): {snippet}")
        except (requests.RequestException, ValueError) as e:
            if attempt == retries - 1:
                print(f"   request failed: {e}")
                return None
            time.sleep(wait)
    return None


def normalize(text):
    #Lowercase, remove accents and punctuation, used for comparing names/titles.
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    text = re.sub(r"[^a-z0-9 ]", " ", text.lower())
    return re.sub(r"\s+", " ", text).strip()


def clean_doi(doi):
    if not doi:
        return ""
    return re.sub(r"^https?://(dx\.)?doi\.org/", "", doi.strip(), flags=re.I).lower()


def name_matches(candidate, member):
    cand = normalize(candidate).split()
    if len(cand) < 2:
        return False
    cand_surname = "".join(cand[1:])          # "El Kadhi" and "Elkadhi" both become "elkadhi"
    spellings = [normalize(n).split() for n in [member["name"]] + member["variants"]]
    spellings = [t for t in spellings if len(t) >= 2]
    full_first_names = {t[0] for t in spellings if len(t[0]) > 1}
    if not any(cand_surname.endswith(t[-1]) for t in spellings):   # last word of surname, e.g. "kadhi"
        return False
    first = cand[0]
    if len(first) == 1:                       # only an initial was written, e.g. "N. El Kadhi"
        return first == member["name"][0].lower()
    return first in full_first_names          # full first name must be one of the member's spellings


def clean_text(text):
    """Remove HTML tags and entities that some publishers put in titles (e.g. <sub>2</sub>, &amp;)."""
    text = html.unescape(re.sub(r"<[^>]+>", "", text or ""))
    return re.sub(r"\s+", " ", text).strip()


# OpenAlex

OPENALEX = "https://api.openalex.org"


def find_openalex_authors(member):
    """Return (OpenAlex ID, name, institutions, ORCID iD) for each profile that belongs to this member."""
    if member["openalex_ids"]:
        return [(i, "manual", "", "") for i in member["openalex_ids"]]

    found, name_only = {}, {}
    for name in [member["name"]] + member["variants"]:
        data = get_json(f"{OPENALEX}/authors", {"search": name, "per-page": 25, "api_key": OPENALEX_API_KEY})
        for a in (data or {}).get("results", []):
            if not name_matches(a.get("display_name", ""), member):
                continue
            insts = [i.get("display_name", "") for i in a.get("last_known_institutions") or []]
            insts += [aff.get("institution", {}).get("display_name", "") for aff in a.get("affiliations") or []]
            inst_label = "; ".join(sorted(set(filter(None, insts))))[:200]
            aid = a["id"].split("/")[-1]
            orcid = (a.get("orcid") or "").split("/")[-1]
            if any(h in normalize(" ".join(insts)) for h in AFFILIATION_HINTS):
                found[aid] = (a["display_name"], inst_label, orcid)
            else:
                name_only[aid] = (a["display_name"], "UNVERIFIED, check this profile: " + inst_label, orcid)
        time.sleep(0.2)

    # Fallback: no profile with a known affiliation, so use name matches, flagged for review.
    if not found:
        found = name_only

    # OpenAlex often splits one person into several profiles, so we keep all that match.
    return [(aid, dn, inst, orcid) for aid, (dn, inst, orcid) in found.items()]


def fetch_openalex_works(author_ids):
    """All works of the given author IDs since FROM_YEAR"""
    works, cursor = [], "*"
    filt = f"author.id:{'|'.join(author_ids)},from_publication_date:{FROM_YEAR}-01-01"
    while cursor:
        data = get_json(f"{OPENALEX}/works", {"filter": filt, "per-page": 200, "cursor": cursor, "api_key": OPENALEX_API_KEY})
        if not data:
            break
        for w in data.get("results", []):
            loc = w.get("primary_location") or {}
            src = loc.get("source") or {}
            # co-authors whose affiliation is in Azerbaijan (country code AZ)
            az_authors = [
                a["author"]["display_name"] for a in w.get("authorships", [])
                if a.get("author") and ("AZ" in (a.get("countries") or [])
                    or any(i.get("country_code") == "AZ" for i in a.get("institutions") or [])
                )
            ]
            works.append({
                "az_affiliated": "; ".join(az_authors),
                "title": w.get("title") or w.get("display_name") or "",
                "authors": "; ".join(a["author"]["display_name"] for a in w.get("authorships", []) if a.get("author")),
                "year": w.get("publication_year"),
                "publication_date": w.get("publication_date") or "",
                "venue": src.get("display_name") or "",
                "type": w.get("type") or "",
                "doi": clean_doi(w.get("doi")),
                "url": w.get("doi") or loc.get("landing_page_url") or w.get("id"),
                "source": "OpenAlex",
            })
        cursor = data.get("meta", {}).get("next_cursor")
        time.sleep(0.2)
    return works


# Crossref
CROSSREF = "https://api.crossref.org"
AZ_WORDS = ("azerbaijan", "baku", "ufaz", "asoiu")


def fetch_crossref_works(member, orcids):
    """
    Two kinds of Crossref queries:
      1. by the member's ORCID iDs (exact, when the publisher recorded the ORCID)
      2. by name (fuzzy), kept only if there is proof it is the right person:
        the author's ORCID matches, the author's affiliation is a known lab institution or another lab member is a co-author.
    """
    select = "DOI,title,author,issued,container-title,type,URL"
    queries = [{"filter": f"orcid:{o},from-pub-date:{FROM_YEAR}-01-01", "rows": 200, "select": select} for o in orcids]
    queries += [{"query.author": name, "filter": f"from-pub-date:{FROM_YEAR}-01-01", "rows": 200, "select": select}
                for name in [member["name"]] + member["variants"]]

    other_members = [m for m in MEMBERS if m is not member]
    works = {}
    for params in queries:
        by_orcid = "query.author" not in params
        data = get_json(f"{CROSSREF}/works", params)
        for w in ((data or {}).get("message") or {}).get("items", []):
            authors = w.get("author") or []
            names = [f"{a.get('given', '')} {a.get('family', '')}".strip() for a in authors]
            mine = [a for n, a in zip(names, authors) if name_matches(n, member)]
            if not mine:
                continue
            if not by_orcid:
                orcid_ok = any((a.get("ORCID") or "").split("/")[-1] in orcids for a in mine)
                affs = normalize(" ".join(aff.get("name", "") for a in mine for aff in a.get("affiliation") or []))
                affiliation_ok = any(h in affs for h in AFFILIATION_HINTS)
                colleague_ok = any(name_matches(n, m) for n in names for m in other_members)
                if not (orcid_ok or affiliation_ok or colleague_ok):
                    continue
            parts = ((w.get("issued") or {}).get("date-parts") or [[None]])[0]
            year = parts[0] if parts else None
            if not year or year < FROM_YEAR:
                continue
            az = [n for n, a in zip(names, authors)
                  if any(word in normalize(aff.get("name", "")) for aff in a.get("affiliation") or [] for word in AZ_WORDS)]
            doi = clean_doi(w.get("DOI"))
            works[doi] = {
                "az_affiliated": "; ".join(az),
                "title": clean_text((w.get("title") or [""])[0]),
                "authors": "; ".join(names),
                "year": year,
                "publication_date": "-".join(f"{p:02d}" if i else str(p) for i, p in enumerate(parts) if p),
                "venue": clean_text((w.get("container-title") or [""])[0]),
                "type": w.get("type") or "",
                "doi": doi,
                "url": w.get("URL") or f"https://doi.org/{doi}",
                "source": "Crossref",
            }
        time.sleep(0.5)
    return list(works.values())


# ORCID

ORCID = "https://pub.orcid.org/v3.0"
ORCID_HEADERS = {"Accept": "application/json"}


def find_orcid_ids(member, from_openalex):
    """ORCID iDs from the OpenAlex profiles; if there are none, search ORCID itself."""
    ids = {o: "from OpenAlex profile" for o in from_openalex if o}
    if ids:
        return ids
    given, family = member["name"].split(" ", 1)
    query = f'family-name:"{family}" AND given-names:"{given}"'
    data = get_json(f"{ORCID}/expanded-search/", {"q": query}, headers=ORCID_HEADERS)
    for r in (data or {}).get("expanded-result") or []:
        insts = "; ".join(r.get("institution-name") or [])
        if any(h in normalize(insts) for h in AFFILIATION_HINTS):
            ids[r["orcid-id"]] = insts[:200]
    return ids


def fetch_orcid_works(orcid_id):
    data = get_json(f"{ORCID}/{orcid_id}/works", headers=ORCID_HEADERS)
    works = []
    for group in (data or {}).get("group") or []:
        w = (group.get("work-summary") or [{}])[0]          # ORCID groups duplicates, take the first
        date = w.get("publication-date") or {}
        year = int(((date.get("year") or {}).get("value")) or 0)
        if year < FROM_YEAR:
            continue
        month = (date.get("month") or {}).get("value")
        day = (date.get("day") or {}).get("value")
        ext = ((w.get("external-ids") or {}).get("external-id")) or []
        doi = next((clean_doi(e.get("external-id-value")) for e in ext if e.get("external-id-type") == "doi"), "")
        works.append({
            "az_affiliated": "",
            "title": clean_text((((w.get("title") or {}).get("title")) or {}).get("value")),
            "authors": "",                                    # ORCID summaries have no author list
            "year": year,
            "publication_date": "-".join(p for p in [str(year), month, day] if p),
            "venue": (w.get("journal-title") or {}).get("value") or "",
            "type": w.get("type") or "",
            "doi": doi,
            "url": f"https://doi.org/{doi}" if doi else ((w.get("url") or {}).get("value") or f"https://orcid.org/{orcid_id}"),
            "source": "ORCID",
        })
    return works


# Merge, compare with website, save


def merge(records):
    """
    Simple duplicate merge: same DOI, or same normalized title.
    """
    merged, by_title = {}, {}
    for r in records:
        title_key = normalize(r["title"])
        key = r["doi"] or "title:" + title_key
        # a record without DOI can still be the same paper as one with a DOI: match by title
        if key not in merged and title_key in by_title and (not r["doi"] or not merged[by_title[title_key]]["doi"]):
            key = by_title[title_key]
        if key not in merged:
            merged[key] = dict(r, members=set([r["member"]]), sources=set([r["source"]]))
            by_title.setdefault(title_key, key)
            continue
        m = merged[key]
        m["members"].add(r["member"])
        m["sources"].add(r["source"])
        if r.get("az_affiliated") and not m.get("az_affiliated"):
            m["az_affiliated"] = r["az_affiliated"]
        # fill any empty fields from the duplicate
        for f in ["publication_date", "venue", "doi", "url", "authors"]:
            longer = f in ("publication_date", "authors") and len(str(r[f])) > len(str(m[f]))
            if not m.get(f) or longer:
                m[f] = r[f]

    rows = []
    for m in merged.values():
        m["member"] = "; ".join(sorted(m.pop("members")))
        m["source"] = "; ".join(sorted(m.pop("sources")))
        rows.append(m)
    return rows


# Rule for Dr.Cecilia: keep her papers only if there is an Azerbaijani co-author which will be UFAZ or ASOIU
RESTRICTED_MEMBER = "Cecilia Zanni-Merk"

# Typical Azerbaijani surname endings, a fallback when OpenAlex has no affiliation for a co-author.
AZ_SURNAME_ENDINGS = ("ov", "ova", "ev", "eva", "yev", "yeva", "zade", "zada", "zadeh", "li", "lu", "ly")


def looks_azerbaijani(name):
    parts = normalize(name).split()
    if len(parts) < 2:
        return False
    surname = parts[-1]
    return len(surname) >= 5 and surname.endswith(AZ_SURNAME_ENDINGS)


def azerbaijani_coauthor_reason(row):
    """
    Checked in order of reliability:
      1. another lab member is a co-author
      2. a co-author is affiliated with an institution in Azerbaijan (OpenAlex)
      3. a co-author has a typical Azerbaijani surname (weakest, review these)
    """
    other_members = [m for m in row["member"].split("; ") if m != RESTRICTED_MEMBER]
    if other_members:
        return "lab member: " + ", ".join(other_members)
    if row.get("az_affiliated"):
        return "Azerbaijan affiliation: " + row["az_affiliated"]
    names = [a for a in row["authors"].split("; ") if not name_matches(a, MEMBERS_BY_NAME[RESTRICTED_MEMBER])]
    az_names = [a for a in names if looks_azerbaijani(a)]
    if az_names:
        return "surname (please check): " + ", ".join(az_names)
    return ""


def apply_cecilia_rule(df):
    """Split into (kept, excluded). Only papers where Cecilia is the ONLY lab member are checked."""
    df = df.copy()
    df["az_coauthor"] = df.apply(azerbaijani_coauthor_reason, axis=1)
    only_cecilia = df["member"] == RESTRICTED_MEMBER
    excluded = df[only_cecilia & (df["az_coauthor"] == "")]
    kept = df.drop(excluded.index)
    # the column is only meaningful for papers involving Cecilia
    kept.loc[~kept["member"].str.contains(RESTRICTED_MEMBER), "az_coauthor"] = ""
    return kept, excluded


def mark_on_website(df):
    path = Path(WEBSITE_LIST_FILE)
    if not path.exists():
        df["on_ufaz_website"] = ""
        return df
    text = path.read_text(encoding="utf-8")
    site_dois = {clean_doi(d).rstrip(".") for d in re.findall(r"10\.\d{4,9}/[^\s\]\)]+", text)}
    site_text = normalize(text)

    def check(row):
        if row["doi"] and row["doi"] in site_dois:
            return "yes"
        t = normalize(row["title"])
        return "yes" if len(t) > 20 and t in site_text else "no"

    df["on_website"] = df.apply(check, axis=1)
    return df


def main():
    if not OPENALEX_API_KEY or OPENALEX_API_KEY == "PASTE_YOUR_KEY_HERE":
        print("Please paste your OpenAlex API key in OPENALEX_API_KEY at the top of the script.")
        print("Get a free one at https://openalex.org/settings/api")
        return

    all_records, matches = [], []

    for member in MEMBERS:
        print(f"\n{member['name']} ({member['role']})")
        # OpenAlex
        authors = find_openalex_authors(member)
        for aid, dn, inst, orcid in authors:
            matches.append({"member": member["name"], "source": "OpenAlex", "profile_id": aid,
                            "profile_name": dn, "institutions": inst})
        if authors:
            works = fetch_openalex_works([a[0] for a in authors])
            print(f"   OpenAlex: {len(authors)} profile(s), {len(works)} works")
            all_records += [dict(w, member=member["name"]) for w in works]
        else:
            print("   OpenAlex: no matching profile")

        # ORCID iDs
        orcids = find_orcid_ids(member, [a[3] for a in authors])

        # Crossref
        works = fetch_crossref_works(member, list(orcids))
        print(f"   Crossref: {len(works)} works")
        all_records += [dict(w, member=member["name"]) for w in works]

        # ORCID
        if orcids:
            for oid, note in orcids.items():
                works = fetch_orcid_works(oid)
                print(f"   ORCID {oid}: {len(works)} works")
                matches.append({"member": member["name"], "source": "ORCID", "profile_id": oid, "profile_name": member["name"], "institutions": note})
                all_records += [dict(w, member=member["name"]) for w in works]
        else:
            print("ORCID: no profile found")

        time.sleep(1)

    if not all_records:
        print("\nNo publications found, check your internet connection.")
        return

    df = pd.DataFrame(merge(all_records))
    df, excluded = apply_cecilia_rule(df)
    df = mark_on_website(df)

    def sort_newest_first(d):
        d = d.copy()
        d["sort_date"] = pd.to_datetime(d["publication_date"].astype(str), errors="coerce", format="mixed")
        return d.sort_values(["sort_date", "year"], ascending=False).drop(columns="sort_date")

    cols = ["publication_date", "year", "title", "authors", "member", "venue", "type", "doi", "url", "source", "on_website", "az_coauthor"]
    df = sort_newest_first(df)[cols]
    excluded = sort_newest_first(excluded)[[c for c in cols if c != "on_website"]] if len(excluded) else excluded

    df.to_csv(HERE / "publications.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(matches).to_csv(HERE / "author_matches.csv", index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(HERE / "publications.xlsx", engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="Publications", index=False)
        # Dr.Cecilia's papers dropped because no Azerbaijani co-author was found, worth a quick look
        excluded.to_excel(xw, sheet_name="Excluded (Cecilia)", index=False)
        pd.DataFrame(matches).to_excel(xw, sheet_name="Author matches", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                width = min(max(len(str(c.value or "")) for c in col) + 2, 60)
                ws.column_dimensions[col[0].column_letter].width = width

    print(f"\nDone: {len(df)} unique publications since {FROM_YEAR}")
    print(f"{len(excluded)} of Cecilia's papers excluded (no Azerbaijani co-author found)")
    if (df["on_website"] == "no").any():
        print(f"{(df['on_website'] == 'no').sum()} are not on the website yet")
    print(f"Saved publications.xlsx, publications.csv, author_matches.csv in {HERE}")


if __name__ == "__main__":
    main()
