# ENCON Offer Generator — Product Scope & Roadmap

The target scope is ENCON's **full product catalog** (from the encon.co.in PRODUCTS menu).
This file tracks which product lines the offer generator covers.

_Last updated: 2026-06-10_

**Legend:** ✅ Complete · 🟡 Partial / in progress · 🟦 Scaffolding only (not started in earnest) · ⬜ Pending (not started)

---

## IRON & STEEL MAKING

| Product | Status | Module |
|---|---|---|
| Vertical Ladle Preheaters | ✅ Complete | `vlph` |
| Horizontal Ladle Preheaters | ✅ Complete | `hlph` |
| Tundish Preheating System | ✅ Complete | `tundish` |
| Torpedo Preheaters (TLC) | ⬜ Pending | reference: `Combined_Offer_Test.docx` |
| Impeller Dryer | ⬜ Pending | — |
| Launder / Runner / Chutes Dryer & Preheater | ⬜ Pending | — |
| Ladle & Tundish Cooler – Refractory Cooler | ⬜ Pending | — |
| Refractory Dry-out Burner System | ⬜ Pending | — |
| RAYCON Systems – Patch-up Lining Heater (Sintering) | ⬜ Pending | — |
| Scrap Dryer / Heaters | ⬜ Pending | — |
| Sub-merged Nozzle (SEN) Heating | ⬜ Pending | — |
| Vessel Preheating Systems | ⬜ Pending | — |

## COMBUSTION EQUIPMENT

| Product | Status | Module |
|---|---|---|
| Blowers | ✅ Complete | `blower` |
| Burners | ✅ Complete | `burner` |
| Heating & Pumping Units (HPU + PU-only) | ✅ Complete | `hpu`, `pu` |
| Recuperators | ✅ Complete | `recup` |
| REGEN Burners | 🟦 Scaffolding | `regen` — not started in earnest |
| Fuel & Air Skids / Trains | 🟡 Partial | gas/air trains exist inside other offers, no standalone module |
| Self-Recuperative Burners | ⬜ Pending | — |
| Specialised Burners | ⬜ Pending | — |
| Die / Bolster Heating | ⬜ Pending | — |
| Hot Air Generator | ⬜ Pending | — |
| Refractory Dry-out Burner System | ⬜ Pending | — |

## FURNACES · OVENS · OTHER

| Product | Status | Module |
|---|---|---|
| Billet Reheating Furnace (BRF) | 🟦 Scaffolding | `brf` — not started in earnest |
| Box Type Furnace (BTF) | 🟦 Scaffolding | `btf` — not started in earnest |
| Rad Heat Gas Elements | 🟡 Partial | DB tables only (`rad_heat_master`, `rad_heat_tata_master`), no UI |
| Specialised Fabrications | ⬜ Pending | (furnaces only) |
| Ovens | ⬜ Pending | — |
| Other Furnaces | ⬜ Pending | — |
| Specialised Testing Ovens & Furnaces | ⬜ Pending | — |

---

## Cross-cutting (not product lines)

| Capability | Status |
|---|---|
| Combined Offer builder (multi-equipment, one customer/T&C) | ✅ Complete |
| Pricelist / price-master management & upload | ✅ Complete |
| Google Drive offer upload | ✅ Complete |

---

## Summary

- **Complete (8):** VLPH, HLPH, Tundish, Blowers, Burners, HPU/PU, Recuperators.
- **Scaffolding (3):** REGEN Burners, Billet Reheating Furnace, Box Type Furnace.
- **Pending:** ~17 remaining product lines.

**Module acronyms:** VLPH/HLPH = Vertical/Horizontal Ladle Preheater · HPU = Heating & Pumping Unit · PU = Pumping Unit only · BRF = Billet Reheating Furnace · BTF = Box Type Furnace · REGEN = Regenerative Burner · TLC = Torpedo Ladle Car.
