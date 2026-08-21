#!/usr/bin/env python3
"""Generate one landing page per Kanton from index.html.

index.html stays the single source of truth for design/content. Re-run this
script any time index.html changes to regenerate all Kanton pages.
"""

import pathlib

ROOT = pathlib.Path(__file__).parent
SOURCE = ROOT / "index.html"

KANTONE = [
    ("zuerich", "Zürich"),
    ("bern", "Bern"),
    ("luzern", "Luzern"),
    ("basel", "Basel"),
    ("st-gallen", "St. Gallen"),
    ("aargau", "Aargau"),
    ("thurgau", "Thurgau"),
    ("solothurn", "Solothurn"),
    ("zug", "Zug"),
    ("schwyz", "Schwyz"),
    ("graubuenden", "Graubünden"),
    ("wallis", "Wallis"),
    ("genf", "Genf"),
    ("waadt", "Waadt"),
    ("uri", "Uri"),
    ("obwalden", "Obwalden"),
    ("nidwalden", "Nidwalden"),
    ("glarus", "Glarus"),
    ("freiburg", "Freiburg"),
    ("schaffhausen", "Schaffhausen"),
    ("appenzell-ausserrhoden", "Appenzell Ausserrhoden"),
    ("appenzell-innerrhoden", "Appenzell Innerrhoden"),
    ("tessin", "Tessin"),
    ("neuenburg", "Neuenburg"),
    ("jura", "Jura"),
]

TITLE_OLD = "<title>AutoAnkauf Schweiz – Auto verkaufen schnell & fair</title>"
DESC_OLD = 'content="Wir kaufen Occasionsautos in der ganzen Schweiz. Sofortige Kaufentscheidung, Barzahlung am selben Tag. Kostenlose Offerte anfordern."'
H1_OLD = "<h1>Ihr Auto verkaufen</h1>"


def main():
    source_html = SOURCE.read_text(encoding="utf-8")

    for slug, name in KANTONE:
        html = source_html
        html = html.replace(
            TITLE_OLD,
            f"<title>Auto verkaufen in {name} – AutoAnkauf Schweiz</title>",
        )
        html = html.replace(
            DESC_OLD,
            f'content="Wir kaufen Ihr Auto in {name}. Sofortige Kaufentscheidung, Barzahlung am selben Tag. Kostenlose Offerte anfordern."',
        )
        html = html.replace(H1_OLD, f"<h1>Ihr Auto verkaufen in {name}</h1>")

        out_path = ROOT / f"{slug}.html"
        out_path.write_text(html, encoding="utf-8")
        print(f"wrote {out_path.name}")


if __name__ == "__main__":
    main()
