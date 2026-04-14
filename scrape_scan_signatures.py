"""
Scrape Scan Signature Identifier data from https://scmdb.net/?page=mine
Extracts mineral name, scan values, rarity/color, and category from the overlay.
Outputs JSON and CSV formats.
"""
import asyncio
import json
import csv
from pathlib import Path
from typing import List, Dict

from playwright.async_api import async_playwright

OUTPUT_DIR = Path("csv/scansig")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

TARGET_URL = "https://scmdb.net/?page=mine"

def parse_color_to_category(color: str) -> str:
    """Map RGB color to rarity/category string."""
    color_map = {
        "rgb(255, 170, 51)": "legendary",
        "rgb(204, 102, 255)": "epic",
        "rgb(51, 153, 255)": "rare",
        "rgb(51, 204, 170)": "uncommon",
        "rgb(136, 153, 170)": "common",
        "rgb(102, 221, 170)": "ROC Mineables",
        "rgb(119, 187, 221)": "FPS Mineables",
        "rgb(170, 153, 119)": "Salvage",
    }
    return color_map.get(color, color)

async def scrape_scan_signatures():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        await page.goto(TARGET_URL, wait_until="networkidle")

        # Click the Scan Signature Identifier button
        await page.click('button[title^="Scan Signature Identifier"]')
        # Wait for overlay to appear
        await page.wait_for_selector('.sigchart-overlay', timeout=10000)

        # Extract data from overlay
        data = await page.evaluate('''() => {
            const rows = Array.from(document.querySelectorAll('.sigchart-row'));
            return rows.map(row => {
                const labelDiv = row.querySelector('.sigchart-label');
                const color = labelDiv ? labelDiv.style.color : null;
                const mineral = labelDiv ? labelDiv.textContent.trim() : null;
                const pills = Array.from(row.querySelectorAll('.sigchart-pill'));
                const values = pills.map(pill => {
                    // Example: "Quantainium ×2 = 6,340"
                    const m = pill.title.match(/(.+) ×(\d+) = ([\d,]+)/);
                    return {
                        text: pill.textContent.trim(),
                        title: pill.title,
                        amount: m ? parseInt(m[2]) : null,
                        value: m ? parseInt(m[3].replace(/,/g, '')) : null,
                    };
                });
                return {
                    mineral,
                    color,
                    values,
                };
            });
        }''')

        # Add rarity/category
        for entry in data:
            entry["category"] = parse_color_to_category(entry["color"])

        await browser.close()
        return data

def save_json(data, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def save_csv(data, path):
    # Flatten for CSV: one row per mineral per value
    rows = []
    for entry in data:
        for v in entry["values"]:
            rows.append({
                "mineral": entry["mineral"],
                "category": entry["category"],
                "color": entry["color"],
                "amount": v["amount"],
                "value": v["value"],
                "pill_text": v["text"],
                "pill_title": v["title"],
            })
    fieldnames = ["mineral", "category", "color", "amount", "value", "pill_text", "pill_title"]
    with open(path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def save_summary_csv(data, path):
    # One row per mineral: base value (x1), max multiplier
    rows = []
    for entry in data:
        if not entry["values"]:
            continue
        # Find base value (amount==1) and max multiplier
        base_value = None
        max_multiplier = None
        for v in entry["values"]:
            if v["amount"] == 1:
                base_value = v["value"]
            if max_multiplier is None or (v["amount"] and v["amount"] > max_multiplier):
                max_multiplier = v["amount"]
        # Fallback: if no x1, use first value
        if base_value is None and entry["values"]:
            base_value = entry["values"][0]["value"]
        rows.append({
            "mineral": entry["mineral"],
            "category": entry["category"],
            "base_value": base_value,
            "max_multiplier": max_multiplier,
        })
    fieldnames = ["mineral", "category", "base_value", "max_multiplier"]
    with open(path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

def main():
    data = asyncio.run(scrape_scan_signatures())
    save_json(data, OUTPUT_DIR / "scan_signatures.json")
    save_csv(data, OUTPUT_DIR / "scan_signatures.csv")
    save_summary_csv(data, OUTPUT_DIR / "scan_signatures_summary.csv")
    print(f"Saved {len(data)} minerals to {OUTPUT_DIR}/scan_signatures.json, .csv, and scan_signatures_summary.csv")

if __name__ == "__main__":
    main()
