# coding: utf-8
"""
bom/regen_parser.py
-------------------
Parses the Regen Standard Costing workbook:
  - Per-KW costing BOM items (regen_costing_items)
  - Burner sizing: dimensions + weights + costs (regen_sizing)
  - Material rate assumptions used in sizing (regen_material_rates)
  - Price list summary (regen_pricelist)
"""

import openpyxl
import pandas as pd

KW_SHEETS = {
    500:  'Costing REGENBurner 500 kw',
    1000: 'Costing REGENBurner 1000 Kw',
    1500: 'Costing REGENBurner 1500 Kw',
    2000: 'Costing REGENBurner 2000 Kw',
    2500: 'Costing REGENBurner 2500 Kw',
    3000: 'Costing REGENBurner 3000 Kw',
    4500: 'Costing REGENBurner 4500 Kw',
    6000: 'Costing REGENBurner 6000 Kw',
}


def _v(x):
    return float(x) if isinstance(x, (int, float)) else None


def _s(x):
    return str(x).strip() if x is not None else ''


def parse_regen_costing(xl_path, conn):
    """
    Parse the Regen Standard Costing workbook.
    Writes four DB tables:
        regen_costing_items, regen_sizing,
        regen_material_rates, regen_pricelist
    """
    wb = openpyxl.load_workbook(xl_path, read_only=True, data_only=True)

    # ── 1. Per-KW costing sheets ─────────────────────────────────────────────
    costing_records = []

    for kw, shname in KW_SHEETS.items():
        if shname not in wb.sheetnames:
            continue
        ws = wb[shname]
        all_rows = list(ws.iter_rows(min_row=1, max_row=60, values_only=True))

        cur_section = ''
        row_order = 0

        for row in all_rows[2:]:          # skip header rows 1-2
            if all(v is None for v in (row[:10] if len(row) >= 10 else row)):
                continue

            sno         = _s(row[0]) if row[0] is not None else ''
            desc        = _s(row[1]) if len(row) > 1 and row[1] is not None else ''
            size_val    = row[2] if len(row) > 2 else None
            qty         = _v(row[3]) if len(row) > 3 else None
            cost_unit   = _v(row[4]) if len(row) > 4 else None
            total_cost  = _v(row[5]) if len(row) > 5 else None
            sp_unit     = _v(row[6]) if len(row) > 6 else None
            total_sp    = _v(row[7]) if len(row) > 7 else None

            # ── detect special rows by content (positions differ across KW) ──
            if not desc:
                # TOTAL row: both col F and col H are large non-zero numbers
                if total_cost is not None and total_sp is not None:
                    costing_records.append({
                        'kw': kw, 'row_order': row_order, 'row_type': 'total',
                        'sno': '', 'section': '', 'description': 'TOTAL',
                        'size': '', 'qty': None,
                        'cost_per_unit': None, 'total_cost': total_cost,
                        'sp_per_unit': None,  'total_sp': total_sp,
                    })
                    row_order += 1
                # FINAL row: only col H has a value
                elif total_sp is not None and total_cost is None and cost_unit is None:
                    costing_records.append({
                        'kw': kw, 'row_order': row_order, 'row_type': 'final',
                        'sno': '', 'section': '', 'description': 'FINAL SELLING PRICE',
                        'size': '', 'qty': None,
                        'cost_per_unit': None, 'total_cost': None,
                        'sp_per_unit': None,  'total_sp': total_sp,
                    })
                    row_order += 1
                # else: intermediate calc row — skip
                continue

            # ── section header: description present but no qty/cost numbers ──
            if qty is None and cost_unit is None and total_cost is None:
                cur_section = desc
                row_type = 'header'
            else:
                row_type = 'item'

            costing_records.append({
                'kw': kw, 'row_order': row_order, 'row_type': row_type,
                'sno': sno,
                'section': cur_section if row_type == 'item' else '',
                'description': desc,
                'size': _s(size_val) if size_val is not None else '',
                'qty': qty,
                'cost_per_unit': cost_unit,
                'total_cost': total_cost,
                'sp_per_unit': sp_unit,
                'total_sp': total_sp,
            })
            row_order += 1

    pd.DataFrame(costing_records).to_sql(
        'regen_costing_items', conn, if_exists='replace', index=False
    )

    # ── 2. Burner Sizing sheet ────────────────────────────────────────────────
    sizing_records = []
    if 'Burner Sizing and costing' in wb.sheetnames:
        ws    = wb['Burner Sizing and costing']
        srows = list(ws.iter_rows(min_row=1, max_row=65, values_only=True))

        # Dimensions: 0-indexed rows 19-26 (Excel rows 20-27)
        dim = {}
        for row in srows[19:27]:
            kw = _v(row[0])
            if kw is None:
                continue
            dim[int(kw)] = {
                'shell_thick': _v(row[1]), 'retainer_thick': _v(row[2]),
                'refractory_thick': _v(row[3]),
                'dim_L': _v(row[4]), 'dim_H': _v(row[5]), 'dim_W': _v(row[6]),
                'bottom_h': _v(row[7]),
                'vol_total': _v(row[8]), 'vol_effective': _v(row[9]),
                'vol_refractory': _v(row[10]), 'density_castable': _v(row[11]),
            }

        # Weights: 0-indexed rows 34-41 (Excel rows 35-42)
        wt = {}
        for row in srows[34:42]:
            kw = _v(row[3])
            if kw is None:
                continue
            wt[int(kw)] = {
                'wt_burner_ms':   _v(row[4]), 'wt_burner_refrac':  _v(row[5]),
                'wt_regen_ms':    _v(row[6]), 'wt_regen_ss':       _v(row[7]),
                'wt_regen_refrac':_v(row[8]), 'wt_ceramic_balls':  _v(row[9]),
                'wt_burner_block':_v(row[10]),'wt_total':          _v(row[11]),
            }

        # Costs: 0-indexed rows 53-61 (Excel rows 54-62 — 500 KW starts at row 54)
        cs = {}
        for row in srows[53:62]:
            kw = _v(row[3])
            if kw is None:
                continue
            cs[int(kw)] = {
                'cost_burner_ms':    _v(row[4]), 'cost_burner_refrac': _v(row[5]),
                'cost_regen_ms':     _v(row[6]), 'cost_regen_ss':      _v(row[7]),
                'cost_regen_refrac': _v(row[8]), 'cost_ceramic_balls': _v(row[9]),
                'cost_burner_block': _v(row[10]),'cost_total':         _v(row[11]),
            }

        for kw in sorted(dim.keys()):
            rec = {'kw': kw}
            rec.update(dim.get(kw, {}))
            rec.update(wt.get(kw, {}))
            rec.update(cs.get(kw, {}))
            # Pull selling price from costing items
            rec['selling_price'] = next(
                (r['total_sp'] for r in costing_records
                 if r['kw'] == kw and r['row_type'] == 'final'), None
            )
            sizing_records.append(rec)

        pd.DataFrame(sizing_records).to_sql(
            'regen_sizing', conn, if_exists='replace', index=False
        )

        # Material rates: 0-indexed rows 47-50 (Excel rows 48-51)
        mat_labels = ['MS', 'SS', 'Refractory', 'Ceramic Balls']
        mr_records = []
        for i, row in enumerate(srows[47:51]):
            label = _s(row[2]) or mat_labels[i]
            mr_records.append({
                'material':      label,
                'wastage':       _v(row[3]),
                'material_cost': _v(row[4]),
                'labor_cost':    _v(row[5]),
            })
        pd.DataFrame(mr_records).to_sql(
            'regen_material_rates', conn, if_exists='replace', index=False
        )

    # ── 3. Price List summary sheet ───────────────────────────────────────────
    if 'Price List' in wb.sheetnames:
        ws = wb['Price List']
        pl_records = []
        for row in ws.iter_rows(min_row=5, max_row=13, values_only=True):
            if not row or row[2] is None:
                continue
            model     = _s(row[2])
            kw        = _v(row[3])
            price_std = _v(row[4])
            price_wog = _v(row[5])
            per_kw    = _v(row[6])
            if model and kw:
                pl_records.append({
                    'model': model, 'kw': int(kw),
                    'price_std_complete': price_std,
                    'price_wo_gas_train': price_wog,
                    'price_per_kw': per_kw,
                })
        pd.DataFrame(pl_records).to_sql(
            'regen_pricelist', conn, if_exists='replace', index=False
        )

    wb.close()
    return {
        'costing_items': len(costing_records),
        'sizing_rows':   len(sizing_records),
    }
