# Leftover inventory → SKU mapping

## What this is

simplehuman is transitioning the 9 oz. liquid sensor pumps (LSP) and the touch sensor mirrors to
newer models. Two inventory lists record the component stock still sitting at Lomak and at
suppliers. The raw lists say *what* we have and *what it cost* — they do not say **which finished
SKU each part belongs to**, which is what we need in order to answer:

1. How much of this excess stock is stranded, i.e. tied to a SKU we cannot finish building?
2. How many units of each SKU can we still build?

`map_inventory_skus.py` reads the raw lists and produces annotated workbooks in `outputs/` that
answer both.

## Layout

```
inputs/
  inventory_lists/
    lsp_inventory_list.xlsx            raw list, 46 parts   (do not edit)
    sensor_mirror_inventory_list.xlsx  raw list, 114 parts  (do not edit)
  lsp_skus.csv                         SKU -> label, from plytix
  mirrors_skus.csv                     SKU -> label, from plytix
  mappings/
    lsp_color_map.csv                  colour token -> SKU(s)   <- edit these
    lsp_overrides.csv                  per-part exceptions      <- edit these
    mirrors_color_map.csv
    mirrors_overrides.csv
sample.xlsx                            agreed output format (mirror family)
map_inventory_skus.py
outputs/                               generated - safe to delete and regenerate
```

## Running it

```
python map_inventory_skus.py                  # both families
python map_inventory_skus.py --family lsp     # lsp | mirrors
```

Needs Python 3 and `openpyxl`. Prints a per-family summary and any warnings.

## Output

Each output workbook has two sheets.

**Sheet 1 — the annotated inventory list.** The original sheet with three columns inserted after
the part name (`C` color, `D` SKU, `E` SKU description), a `Mapping basis / notes` column in `W`,
and a legend below the totals. Every original formula, fill, number format and the whole totals
block are preserved — the script shifts formulas with `openpyxl`'s `Translator` rather than
rebuilding the sheet.

All text is **Arial 9, black**, bold being the only variation (title, headers, legend
headings). The source lists mix typefaces and sizes, so the script sweeps every cell at the end
and also resets the workbook's default font — merged-range continuation cells and any cell you
type into later fall back to that default rather than carrying a style of their own.

Cell fills in the SKU column carry the confidence:

| Fill | Meaning |
|---|---|
| none | **Verified** — the part name states the colour (or the SKU) and it matches the plytix label |
| yellow `FFF2CC` | **Inferred** — best guess from colour codes or quantity agreement; reasoning is in the notes column |
| orange `FCE4D6` | **Needs review** — no defensible allocation; someone has to confirm it |

`ALL` in the SKU column means a common part used by every SKU. `REVIEW` means unallocated.

**Sheet 2 — `SKU buildable`.** One row per SKU:

- **Max units from its own parts** — the lowest count across parts used by that SKU and nothing
  else. This is the number that drives excess stock: those parts fit one product only.
- **Units allowed by shared / common stock** — a ceiling, not a capacity. Common stock is one
  pool, so the per-SKU figures cannot all be hit at once. Note that for the mirrors this is
  currently capped at 60 by a single line ("Pouch"), which caps every SKU at once.
- **Buildable now** — the lower of the two: what could be assembled without buying anything.
- **Stranded US$** — value of a SKU's own parts held beyond its *own* bottleneck. Measured against
  its own parts, not the common gate, because a cheap shared accessory running low can be
  re-bought and does not strand body parts.

Costs convert to US$ at the rates in the source totals block: RMB 6.8, HK$ 7.8317.

## Correcting a mapping

Mapping decisions live in CSVs, never in the Python. To change one, edit the CSV and re-run.

`<family>_color_map.csv` — the default rule for a colour token found in a part name:

```csv
color,skus,confidence,note
matte black,ST1084,verified,"colour named in part, matched to plytix label"
grey,ST1082,inferred,"Interior colour code - qty 15,098 matches the brushed exterior parts"
,ALL,inferred,
```

The blank-colour row is the fallback for parts with no colour token. Tokens are matched
**longest-first**, so `matte black` wins over `black` and `light grey` over `grey`.

`<family>_overrides.csv` — exceptions for one specific part, keyed by the **Item** number in
column A of the source list (not the part name: mirror items 29 and 33 share a name):

```csv
item,part_name,skus,confidence,qty_per_unit,note
16,Spout cover black,ST1084,verified,,"Exterior spout cover - qty 3,950 matches Cap matte black"
83,Screw WA Dia2.3 x L7mm,,,?,"Fastener - pieces per unit unknown"
```

- Blank `skus` / `confidence` / `note` fields inherit from the colour rule, so an override can
  change just the note.
- `part_name` is a safety check. If the source list is revised and item 16 becomes a different
  part, the script warns instead of silently misapplying the override.
- `qty_per_unit` defaults to 1. Set it to `?` for parts whose pieces-per-unit is unknown
  (fasteners, tape, papers) — those are then excluded from the buildable minimums rather than
  being treated as 1-per-unit and producing a false bottleneck.
- `skus` may be a single code, several joined by ` / `, `ALL`, or `REVIEW`.

## Mapping decisions worth knowing

**Mirrors.** Seeded from `sample.xlsx`, which is the agreed answer for this family. Interior
colour codes (`light grey` → ST3054, `grey` → ST3052/ST3053/ST3061) are inferred, not from
plytix. Accessories (USB cables, AC adapters), the 10x detail-mirror frames and the stickers are
flagged for review.

**LSP.** Derived from the list, since there was no agreed answer:

- Exterior finishes are named directly: matte black, white, brushed, brass, matte gold.
- `grey` interiors (15,098 / 15,110 / 15,128) match the brushed exterior parts (15,098–15,400)
  → **ST1082**. Inferred.
- `black` interiors (6,224 / 6,237 / 6,257) ≈ matte black 3,950 + brass 1,000 + matte gold 1,500
  = 6,450 → **ST1084 / ST1083 / ST1102**. Inferred.
- Item 16 "Spout cover black" is an *exterior* part despite the plain `black` token: qty 3,950
  matches item 21 "Cap matte black" exactly, so it is the matte black finish → **ST1084**. This
  is an override, not the colour rule.
- **ST1092 (polished) has no parts on this list at all.** Grey interior quantity matches brushed
  exactly, so it was not folded into ST1092. Zero polished units are buildable from this stock.

## Verification

The mirror output is checked against `sample.xlsx`. Rows 1–122, columns A–U and W:

```
python -c "
import openpyxl
a = openpyxl.load_workbook('sample.xlsx').active
b = openpyxl.load_workbook('outputs/sensor_mirror_inventory_sku_mapping.xlsx')['Touch Mirror']
d=[(r,c) for r in range(1,123) for c in list(range(1,22))+[23] if a.cell(r,c).value!=b.cell(r,c).value]
print(len(d), 'diffs;', len([x for x in d if x[1]<=21]), 'in cols A-U')
"
```

Expect **20 diffs, 4 of them in columns A–U**. All are deliberate:

- 14 notes-column diffs: added "pieces per unit unknown" markers for fasteners/tapes (these feed
  sheet 2), and two notes that cite item numbers instead of spreadsheet row numbers, which shift.
- 4 diffs in D/E on the two accessory rows (items 94, 96): same five SKUs, but joined with ` / `
  to match every other multi-SKU row instead of the sample's ad-hoc `/` and comma list.

Zero diffs in any quantity, cost or formula cell.
