# UFAZ Lab Publications

A Python script that collects every publication by the lab's members (permanent staff and PhD candidates) since 2022 and saves them as one deduplicated, sortable table.

## Data sources

| Source | Used for |
| --- | --- |
| [OpenAlex](https://openalex.org) | Main source: author profiles with institutions, works with author lists and affiliations |
| [Crossref](https://www.crossref.org) | DOI registry, catches papers OpenAlex misses |
| [ORCID](https://orcid.org) | Works members list on their own ORCID profiles |

## How it works

1. Finds each member's OpenAlex profile by name and spelling variants, preferring profiles affiliated with the lab's institutions.
2. Fetches their works from OpenAlex, Crossref and ORCID.
3. Merges duplicates by DOI, or by title when there is no DOI.
4. Keeps Prof. Zanni-Merk's papers only if there is an Azerbaijani co-author (UFAZ or ASOIU). Excluded papers go to a separate sheet for review.
5. Marks which papers are already listed on the lab website.

## Output

| File | Contents |
| --- | --- |
| `publications.xlsx` | Sheets: Publications (newest first), Excluded (Cecilia), Author matches |
| `publications.csv` | Same data as the Publications sheet |
| `author_matches.csv` | Which OpenAlex and ORCID profile was used for each member |

## How to run

1. Install dependencies:
   ```
   pip install requests pandas openpyxl
   ```
2. Get a free OpenAlex API key at https://openalex.org/settings/api and paste it into `OPENALEX_API_KEY` at the top of `get_publications.py`.
3. Keep `website_publications.txt` in the same folder.
4. Run:
   ```
   python get_publications.py
   ```

To add a member, edit the `MEMBERS` list. To change the start year, edit `FROM_YEAR`.

## Limitations

Papers not indexed by OpenAlex, Crossref or ORCID (for example some local-language journals and DOI-less abstracts) need to be added manually.
