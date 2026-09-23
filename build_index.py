#!/usr/bin/env python3
"""Build index.html from resources.json + _index.template.html.

The manifest is the single source of truth for the resource library. Section
counts, the browse-by-section hub, and the filter chips are all derived, so
adding a resource means adding one JSON entry and re-running this script.

    python build_index.py

Output is static HTML: every card is present in the markup, so the page works
with JavaScript disabled and stays indexable. JavaScript only adds filtering.

The template is underscore-prefixed so GitHub Pages (Jekyll) keeps it out of
the published site.
"""
import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent
MANIFEST = ROOT / "resources.json"
TEMPLATE = ROOT / "_index.template.html"
OUTPUT = ROOT / "index.html"


def panel(section):
    return f"""    <a class="panel" href="#sec-{section['slug']}" data-section="{section['slug']}">
      <span class="panel-head">
        <span class="panel-label">{section['label']}</span>
        <span class="panel-count">{len(section['resources'])}</span>
      </span>
      <span class="panel-blurb">{section['blurb']}</span>
    </a>"""


def card(resource, section_slug):
    classes = "card external" if resource["external"] else "card"
    formats = list(resource["formats"])
    if resource["external"] and "external" not in formats:
        formats.append("external")
    arrow = "&nearr;" if resource["external"] else "&rarr;"
    target = ' target="_blank" rel="noopener noreferrer"' if resource["external"] else ""

    tags = []
    if resource["external"]:
        tags.append('        <span class="tag tag-external">External link</span>')
    tags += [f'        <span class="tag">{t}</span>' for t in resource["tags"]]

    source = ""
    if resource.get("source"):
        source = f'\n      <p class="card-source">{resource["source"]}</p>'

    return f"""    <a class="{classes}" data-section="{section_slug}" data-format="{' '.join(formats)}" data-topics="{' '.join(resource['topics'])}" href="{resource['href']}"{target}>
      <p class="card-title"><span class="arrow">{arrow}</span> {resource['title']}</p>
      <p class="card-desc">{resource['desc']}</p>{source}
      <div class="card-tags">
{chr(10).join(tags)}
      </div>
    </a>"""


def section_block(section):
    cards = "\n".join(card(r, section["slug"]) for r in section["resources"])
    return f"""  <h2 id="sec-{section['slug']}" data-section="{section['slug']}">{section['label']} <span class="sec-count">{len(section['resources'])}</span></h2>
  <p class="sec-blurb">{section['blurb']}</p>
  <div class="cards">
{cards}
  </div>"""


def main():
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    template = TEMPLATE.read_text(encoding="utf-8")
    sections = manifest["sections"]

    seen = {}
    for s in sections:
        for r in s["resources"]:
            if r["href"] in seen:
                sys.exit(f"Duplicate href: {r['href']} in {seen[r['href']]} and {s['slug']}")
            seen[r["href"]] = s["slug"]

    replacements = {
        "<!--BUILD:PANELS-->": "\n".join(panel(s) for s in sections),
        "<!--BUILD:TOPICS-->": "\n".join(
            f'    <button class="chip" type="button" data-filter="{t["slug"]}" aria-pressed="false">{t["label"]}</button>'
            for t in manifest["topics"]
        ),
        "<!--BUILD:FORMATS-->": "\n".join(
            f'    <button class="chip chip-sm" type="button" data-format="{f["slug"]}" aria-pressed="false">{f["label"]}</button>'
            for f in manifest["formats"]
        ),
        "<!--BUILD:SECTIONS-->": "\n\n".join(section_block(s) for s in sections),
    }

    html = template
    for marker, value in replacements.items():
        if marker not in html:
            sys.exit(f"Template is missing marker {marker}")
        html = html.replace(marker, value)

    OUTPUT.write_text(html, encoding="utf-8")

    total = sum(len(s["resources"]) for s in sections)
    print(f"Wrote {OUTPUT.name}: {total} resources across {len(sections)} sections")
    for s in sections:
        print(f"  {s['slug']:<16} {len(s['resources'])}")


if __name__ == "__main__":
    main()
