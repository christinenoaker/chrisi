"""Builds the Draper cabinet quote worksheet (Draper_Cabinet_Quote.xlsx).

Source: "Draper residence presentation. 09.14.26_c.pdf" (plan designer, sent 10/8/26)
and the Draper Master Spec Sheet (Millwork & Cabinets tab).
Prices are left blank: enter them from the Pro Kitchen / KCD / Designer's Choice
pro accounts, and the totals calculate automatically.
"""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = "draper/Draper_Cabinet_Quote.xlsx"

# (room, run, sku, description, qty, line/finish, confidence, notes)
ITEMS = [
    # ---------------- KITCHEN ISLAND ----------------
    ("Kitchen", "Island (inside)", "B3D18", "Base 3-drawer 18\"W", 1, "Sun Valley Maple Brownie", "High", "Inside island elevation, left end"),
    ("Kitchen", "Island (inside)", "TRB-2 / B18 trash", "Base 18\"W single door (double trash pull-out)", 1, "Sun Valley Maple Brownie", "Medium", "Plan label TRB-2; confirm trash pull-out kit is included or add separately"),
    ("Kitchen", "Island (inside)", "BS36", "Sink base 36\"W", 1, "Sun Valley Maple Brownie", "High", "Sink TBD (Blanco 33-1/2\" undermount per spec sheet)"),
    ("Kitchen", "Island (inside)", "DW opening", "Dishwasher opening 23-3/4\" (Bosch SHP9PCM5N)", 1, "N/A", "High", "No cabinet; panel-ready? confirm if DW panel needed"),
    ("Kitchen", "Island (inside)", "B3D24", "Base 3-drawer 24\"W", 1, "Sun Valley Maple Brownie", "High", "Inside island elevation, right end"),
    ("Kitchen", "Island (ends)", "TPOST5", "Turned post / island leg", 2, "Sun Valley Maple Brownie", "Medium", "Shown on plan at both island ends; render shows turned legs"),
    ("Kitchen", "Island (outside)", "Back panel / end panels", "Finished back panel 120-3/16\" x 34-1/2\" + 2 end panels", 1, "Sun Valley Maple Brownie", "Medium", "Outside island elevation is a plain panel; quote as skins/panels"),
    ("Kitchen", "Island (end)", "18-RV18PB", "18\" end unit (label partly cut off on plan)", 1, "Sun Valley Maple Brownie", "Low", "Label cut off in PDF - confirm SKU with plan designer"),
    # ---------------- KITCHEN COOKTOP / HOOD WALL ----------------
    ("Kitchen", "Cooktop wall", "B3D24", "Base 3-drawer 24\"W", 2, "Sun Valley Alabaster", "High", "One each end of the cooktop run"),
    ("Kitchen", "Cooktop wall", "B3D30", "Base 3-drawer 30\"W", 2, "Sun Valley Alabaster", "High", "Each side of cooktop; spec: all base cabs on hood wall = drawers"),
    ("Kitchen", "Cooktop wall", "ANT4248", "36\" cooktop base / drawer unit under LG CBIS3618BE (plan label ANT4248)", 1, "Sun Valley Alabaster", "Low", "Confirm SKU - may be the hood/cooktop symbol, not a cabinet"),
    ("Kitchen", "Cooktop wall", "W2418-L / W2418-R", "Wall 24\"W x 18\"H (stack tops over tall columns), L & R", 2, "Sun Valley Alabaster", "Low", "Plan label partly obscured (W2_18-L/R)"),
    ("Kitchen", "Cooktop wall", "Tall columns flanking hood", "Tall/stacked wall columns each side of hood (see render)", 2, "Sun Valley Alabaster", "Low", "Not dimensioned in PDF - need elevation from plan designer"),
    ("Kitchen", "Cooktop wall", "Hood", "Wood/plaster hood w/ insert (vent insert TBD per spec)", 1, "Sun Valley Alabaster", "Low", "Confirm whether hood is cabinetry or separate (Hoodsly etc.)"),
    ("Kitchen", "Cooktop wall", "500VECRN10", "Crown / valance moulding", 2, "Sun Valley Alabaster", "Low", "Quote crown by LF once full elevations are in hand"),
    # ---------------- KITCHEN SINK / WINDOW WALL ----------------
    ("Kitchen", "Window wall", "B3D21", "Base 3-drawer 21\"W", 1, "Sun Valley Alabaster", "High", "Left end, next to refrigerator"),
    ("Kitchen", "Window wall", "WS3642", "36\"W unit (plan label WS3642)", 1, "Sun Valley Alabaster", "Low", "Confirm - may be wall/shelf unit over base"),
    ("Kitchen", "Window wall", "B (33\")", "Base 33\"W (label overlapped/garbled on plan)", 1, "Sun Valley Alabaster", "Low", "Confirm SKU"),
    ("Kitchen", "Window wall", "B2D36", "Base 2-drawer 36\"W", 1, "Sun Valley Alabaster", "High", "Right end"),
    ("Kitchen", "Window wall", "WS3636", "36\"W unit (plan label WS3636)", 1, "Sun Valley Alabaster", "Low", "Confirm"),
    # ---------------- KITCHEN TALL / OVEN WALL ----------------
    ("Kitchen", "Oven / fridge wall", "Oven tall cabinet", "Double oven tall cabinet for Cafe CTD70DP2NS1 (30\")", 1, "Sun Valley Alabaster", "Low", "NOT in the PDF plan - shown in render only; need elevation"),
    ("Kitchen", "Oven / fridge wall", "Tall pantry", "Tall pantry cabinet(s) left of ovens", 1, "Sun Valley Alabaster", "Low", "NOT in the PDF plan - render only"),
    ("Kitchen", "Oven / fridge wall", "Refrigerator wall cab + panels", "Wall cab over fridge + stacked tops + fridge end panels (Cafe CVE28DP2NS1)", 1, "Sun Valley Alabaster", "Low", "NOT in the PDF plan - render only"),
    ("Kitchen", "Oven / fridge wall", "Appliance garage / beverage nook", "Base drawers + open shelves at right end (render)", 1, "Sun Valley Alabaster", "Low", "Spec sheet: 'Inset Appliance Garage' - confirm"),
    ("Kitchen", "Accessories", "Fillers / toe kick / scribe / crown", "Fillers, toe kick, scribe, crown, light rail", 1, "Sun Valley Alabaster", "Low", "Add per vendor's takeoff"),
    # ---------------- PANTRY (Designer's Choice Bellview - green) ----------------
    ("Pantry", "Sink wall", "Tall pantry 30\"W x 84\"+", "Tall pantry 30\"W, 2 stacked door sets (purple on elevation)", 1, "Designer's Choice Bellview (green)", "Medium", "Elevation: 30\" tall unit, 53-7/16\" + upper section"),
    ("Pantry", "Sink wall", "W3018", "Wall 30\"W x 18\"H stacker (over tall pantry)", 1, "Designer's Choice Bellview (green)", "High", "Plan W3018"),
    ("Pantry", "Sink wall", "W3618-24", "Wall 36\"W x 18\"H x 24\"D over refrigerator", 1, "Designer's Choice Bellview (green)", "High", "Over GE GFE26JYMFS fridge"),
    ("Pantry", "Sink wall", "W3624 (stacker)", "Wall 36\"W x 24\"H stacker over W3618-24", 1, "Designer's Choice Bellview (green)", "Medium", "Elevation shows 2 tiers over fridge"),
    ("Pantry", "Sink wall", "DD30-3", "Drawer base 30\"W, 3-drawer", 3, "Designer's Choice Bellview (green)", "High", "Plan DD30-3 x3"),
    ("Pantry", "Sink wall", "SB33", "Sink base 33\"W", 1, "Designer's Choice Bellview (green)", "High", "Plan SB33 (UKD-10 = sink)"),
    ("Pantry", "Sink wall", "W3018", "Wall 30\"W x 18\"H stackers (top tier)", 3, "Designer's Choice Bellview (green)", "High", "Over the 3 drawer-base columns"),
    ("Pantry", "Sink wall", "W3042 (approx.)", "Wall 30\"W x ~42\"H (lower tier, tall uppers)", 3, "Designer's Choice Bellview (green)", "Medium", "Elevation shows tall uppers under each W3018; confirm height"),
    ("Pantry", "Sink wall", "W3330", "Wall 33\"W x 30\"H over sink", 1, "Designer's Choice Bellview (green)", "High", "Plan W3330"),
    ("Pantry", "Sink wall", "CM24 / TKM8-25 / F396", "Crown (CM24), toe kick (TKM8-25), filler 3\"x96\" (F396)", 1, "Designer's Choice Bellview (green)", "Medium", "Accessories from plan"),
    ("Pantry", "Door wall", "REPV24120", "Refrigerator-style end panel 24\"D x 120\"H (open shelving uprights)", 4, "Designer's Choice Bellview (green)", "Medium", "Plan REPV24120 x4"),
    ("Pantry", "Door wall", "REPV12120", "End panel 12\"D x 120\"H", 1, "Designer's Choice Bellview (green)", "Medium", "Plan REPV12120"),
    ("Pantry", "Door wall", "SHELF", "Open shelves between panels", 0, "Designer's Choice Bellview (green)", "Low", "Count shelves from Pantry Door Wall elevation (~4 bays)"),
    ("Pantry", "Hardware", "PAL001", "Pulls", 0, "Owner provided", "High", "Spec sheet: owner to provide pantry hardware - do not quote"),
    # ---------------- PRIMARY BATH (KCD Shaker White) ----------------
    ("Primary Bath", "Vanity 112-1/2\"", "VB12-L", "Vanity base 12\"W, hinge left", 1, "KCD Shaker White", "High", "Plan VB12 L"),
    ("Primary Bath", "Vanity 112-1/2\"", "V27", "Vanity sink base 27\"W", 2, "KCD Shaker White", "High", "UBS-04 sinks"),
    ("Primary Bath", "Vanity 112-1/2\"", "VDB15", "Vanity drawer base 15\"W", 2, "KCD Shaker White", "High", "Center pair"),
    ("Primary Bath", "Vanity 112-1/2\"", "VB12-R", "Vanity base 12\"W, hinge right", 1, "KCD Shaker White", "High", "Plan VB12 R"),
    ("Primary Bath", "Vanity 112-1/2\"", "F396", "Filler 3\"W x 96\"", 3, "KCD Shaker White", "High", "Ends + center"),
    ("Primary Bath", "Vanity 112-1/2\"", "Grooming roll-out", "Vanity grooming station roll-out 12\" w/ power", 2, "KCD Shaker White", "Medium", "Spec sheet: Vanity Roll-Outs (x2), 12\" w/power, Shaker White"),
    # ---------------- BATH 2 (KCD Shaker White) ----------------
    ("Bath 2 (Girls)", "Vanity 95-3/4\"", "VDB12", "Vanity drawer base 12\"W x 21\"D", 2, "KCD Shaker White", "High", "Each end"),
    ("Bath 2 (Girls)", "Vanity 95-3/4\"", "V24-2", "Vanity sink base 24\"W x 21\"D, 2 doors", 2, "KCD Shaker White", "High", "UBS-01 oval sinks"),
    ("Bath 2 (Girls)", "Vanity 95-3/4\"", "VDB18", "Vanity drawer base 18\"W x 21\"D", 1, "KCD Shaker White", "High", "Center"),
    ("Bath 2 (Girls)", "Vanity 95-3/4\"", "Filler 2\"", "Filler 2\" (F396 on plan)", 1, "KCD Shaker White", "High", "Left end"),
    ("Bath 2 (Girls)", "Linen", "Linen 17\"W x 96\"H", "Linen tower 17\"W (15\" door) x 96\"H: base door 40-1/2\" + open shelves above", 1, "KCD Shaker White", "Medium", "'Girls bath cabinet' elevation; confirm SKU"),
    # ---------------- BATH 3 (KCD Shaker White) ----------------
    ("Bath 3", "Vanity 48\"", "V4821 (sink base 48\")", "Vanity sink base 48\"W x 21\"D, 2 doors + false front + 2 side drawers", 1, "KCD Shaker White", "Medium", "Bath 3 elevation 48\" + 4\" filler"),
    ("Bath 3", "Vanity 48\"", "Filler 4\"", "Filler 4\"", 1, "KCD Shaker White", "High", "Right side"),
    # ---------------- BATH 4 ----------------
    ("Bath 4", "-", "None", "Vitreous china console sink - no cabinet", 0, "N/A", "High", "No cabinetry"),
]

VENDORS = ["Pro Kitchen", "KCD", "Designer's Choice"]

thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
hdr_fill = PatternFill("solid", fgColor="3F3A36")
room_fill = PatternFill("solid", fgColor="EDE6DC")
low_fill = PatternFill("solid", fgColor="FCE4D6")
med_fill = PatternFill("solid", fgColor="FFF2CC")
hdr_font = Font(bold=True, color="FFFFFF")

wb = Workbook()
ws = wb.active
ws.title = "Cabinet Quote"

ws["A1"] = "DRAPER RESIDENCE - CABINET QUOTE WORKSHEET"
ws["A1"].font = Font(bold=True, size=14)
ws["A2"] = "8794 Pine Barrens Dr, Orlando FL 32817  |  Job WORL8794PBD  |  Source: Draper residence presentation 09.14.26_c + Draper Master Spec Sheet"
ws["A3"] = "Enter UNIT price from each pro account; extended totals calculate. Orange = SKU not confirmed in PDF (verify before ordering). Yellow = inferred from elevation."
ws["A3"].font = Font(italic=True, color="7F7F7F")

headers = ["Room", "Run", "SKU / Plan Label", "Description", "Qty", "Door Style / Finish",
           "Confidence", "Notes"]
for v in VENDORS:
    headers += [f"{v} Unit $", f"{v} Ext $"]
HR = 5
for c, h in enumerate(headers, 1):
    cell = ws.cell(row=HR, column=c, value=h)
    cell.font = hdr_font
    cell.fill = hdr_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border

r = HR + 1
first = r
for room, run, sku, desc, qty, finish, conf, notes in ITEMS:
    vals = [room, run, sku, desc, qty, finish, conf, notes]
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.border = border
        cell.alignment = Alignment(vertical="top", wrap_text=True)
    for i in range(len(VENDORS)):
        ucol = 9 + i * 2
        ecol = ucol + 1
        u = ws.cell(row=r, column=ucol)
        u.number_format = '"$"#,##0.00'
        u.border = border
        e = ws.cell(row=r, column=ecol,
                    value=f'=IF({get_column_letter(ucol)}{r}="","",{get_column_letter(ucol)}{r}*E{r})')
        e.number_format = '"$"#,##0.00'
        e.border = border
    fill = low_fill if conf == "Low" else med_fill if conf == "Medium" else None
    if fill:
        ws.cell(row=r, column=7).fill = fill
        ws.cell(row=r, column=3).fill = fill
    r += 1
last = r - 1

# Subtotals by room + grand total
r += 1
ws.cell(row=r, column=4, value="SUBTOTALS BY ROOM").font = Font(bold=True)
r += 1
rooms = []
for it in ITEMS:
    if it[0] not in rooms:
        rooms.append(it[0])
for room in rooms:
    ws.cell(row=r, column=4, value=room).fill = room_fill
    for i in range(len(VENDORS)):
        ecol = get_column_letter(10 + i * 2)
        cell = ws.cell(row=r, column=10 + i * 2,
                       value=f'=SUMIF($A${first}:$A${last},"{room}",{ecol}{first}:{ecol}{last})')
        cell.number_format = '"$"#,##0.00'
        cell.fill = room_fill
    r += 1
ws.cell(row=r, column=4, value="CABINET MATERIAL TOTAL").font = Font(bold=True)
for i in range(len(VENDORS)):
    ecol = get_column_letter(10 + i * 2)
    cell = ws.cell(row=r, column=10 + i * 2, value=f"=SUM({ecol}{first}:{ecol}{last})")
    cell.number_format = '"$"#,##0.00'
    cell.font = Font(bold=True)
r += 1
for label in ["Freight / delivery", "Assembly (if RTA)", "Sales tax"]:
    ws.cell(row=r, column=4, value=label)
    for i in range(len(VENDORS)):
        ws.cell(row=r, column=10 + i * 2).number_format = '"$"#,##0.00'
    r += 1
ws.cell(row=r, column=4, value="GRAND TOTAL").font = Font(bold=True, size=12)
for i in range(len(VENDORS)):
    ecol = get_column_letter(10 + i * 2)
    cell = ws.cell(row=r, column=10 + i * 2, value=f"=SUM({ecol}{r-4}:{ecol}{r-1})")
    cell.number_format = '"$"#,##0.00'
    cell.font = Font(bold=True, size=12)

widths = [14, 18, 24, 48, 6, 26, 11, 48] + [13, 13] * len(VENDORS)
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = ws.cell(row=HR + 1, column=4)

# ---------------- Open questions sheet ----------------
q = wb.create_sheet("Before Ordering")
q["A1"] = "CHECK BEFORE SENDING / ORDERING"
q["A1"].font = Font(bold=True, size=14)
QUESTIONS = [
    "Bishop order C518332-1 was confirmed 9/4/26 as a 'draper PO' and the spec sheet lists the kitchen as Bishop (perimeter inset, island frameless). Confirm what that order covers so it isn't quoted twice.",
    "Buildertrend PO WORL8794PBD-0004 'Draper Vanities' already exists (completion 10/5/26). Confirm whether the bath vanities are already purchased.",
    "Kitchen door style conflict: the PDF shows Sun Valley Alabaster (perimeter) / Sun Valley Maple Brownie (island). The spec sheet shows Bishop Alabaster inset. Which line is being quoted?",
    "The kitchen oven/refrigerator/tall-pantry wall is in the renders but not in the PDF floor plan. Ask the plan designer (Isabela) for that elevation or the full cabinet list/report.",
    "Several kitchen plan labels are overlapped or cut off in the PDF (ANT4248, WS3642, WS3636, 18-RV18PB, W2_18). Get the cabinet report from the design software for exact SKUs.",
    "Pantry hardware is owner provided per spec sheet - excluded.",
    "Primary vanity grooming roll-outs (x2, 12\" w/ power) come from the spec sheet, not the PDF - confirm they're still wanted.",
    "Pantry open shelf count and the tall-upper heights need confirmation from the Pantry Door Wall / Sink Wall elevations.",
]
for i, t in enumerate(QUESTIONS, 3):
    q.cell(row=i, column=1, value=f"{i-2}.")
    c = q.cell(row=i, column=2, value=t)
    c.alignment = Alignment(wrap_text=True, vertical="top")
q.column_dimensions["A"].width = 4
q.column_dimensions["B"].width = 120

wb.save(OUT)
print("saved", OUT, "items:", len(ITEMS))
