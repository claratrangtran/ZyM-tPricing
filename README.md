# ZyMōt Dynamic Pricing Tool

> A Sales Operations case study — building a data-driven pricing framework for a niche medtech product across 40+ European IVF clinics, from kickoff workshop to a live Excel tool that account managers used to walk into client meetings with a defensible price.

**Author:** Clara Tran — Sales Operations & CRM Analyst **Portfolio:** [claratran.lovable.app](https://claratran.lovable.app) **LinkedIn:** [linkedin.com/in/clara-tran-thi](https://linkedin.com/in/clara-tran-thi/) **Project:** CooperSurgical EMEA, 2024 *(portfolio reconstruction with synthetic data)* **Stack:** Excel · Python (pandas, openpyxl) · python-pptx · live formulas · no macros, no VBA

---

## Table of contents

- [1\. Context](#1-context)  
- [2\. The business problem](#2-the-business-problem)  
- [3\. How we built the solution](#3-how-we-built-the-solution)  
  - [Phase 1 — Joint workshop with commercial and product teams](#phase-1--joint-workshop-with-commercial-and-product-teams)  
  - [Phase 2 — Build the model](#phase-2--build-the-model)  
  - [Phase 3 — Ship a tool account managers would actually open](#phase-3--ship-a-tool-account-managers-would-actually-open)  
- [4\. The model](#4-the-model)  
- [5\. How account managers use the tool](#5-how-account-managers-use-the-tool)  
- [6\. Results after 3 months](#6-results-after-3-months)  
- [7\. What's in this repository](#7-whats-in-this-repository)  
- [8\. How to reproduce](#8-how-to-reproduce)  
- [9\. Caveats](#9-caveats)

---

## 1\. Context

**ZyMōt** is a microfluidic sperm-separation device used in the sperm-preparation step of an IVF cycle. It's a niche medtech product with no direct microfluidic competitor at its price point — meaning its manufacturer has genuine pricing power.

CooperSurgical acquired the ZyMōt product line through the DxNow acquisition (2019–2020). The product was integrated into Cooper's broader IVF portfolio (media, cryopreservation, embryo culture, catheters, diagnostics), but **pricing had been inherited from the legacy distribution network** — set by whichever distributor handled a given country or client at the time of acquisition.

As the product moved into Cooper's own commercial structure, Sales Operations needed to answer a simple question nobody had a documented answer to:

> **What should we charge whom, and why?**

## 2\. The business problem

Same product. Same box. Same clinical outcome. **Prices ranging from €200 to €310 per unit** across the European client base — a 55% spread with no commercial logic behind it.

Three concrete symptoms:

1. Account managers had no framework for what to quote a new or existing client. Every conversation was a judgment call.  
2. Low-priced clients were being subsidised without that subsidy buying any strategic value (loyalty, volume growth, cross-sell).  
3. High-priced clients were paying a premium that reflected a legacy distributor markup, not CooperSurgical's own commercial intent — a retention risk the moment they discovered it.

>   
> \[\!NOTE\] ZyMōt has pricing power (no direct competitor), so the question was never "how do we compete on price?" — it was "how do we use price as a lever for the relationship?"

## 3\. How we built the solution

### Phase 1 — Joint workshop with commercial and product teams

The pricing analyst (me) couldn't build the model alone. Three groups had to agree before a single formula was written:

| Team | What they brought to the table |
| :---- | :---- |
| **Pricing analyst** (Sales Ops) | Data access, modelling capability, the Excel tool itself |
| **Commercial team** (EMEA Account Managers) | Ground-truth knowledge of each client, feasibility of price moves, political sensitivity |
| **ZyMōt Product Management team** | Clinical positioning, cost structure, strategic growth priorities for the product line |

In the workshop we settled four questions:

1. **What defines a strategic account?** → A combination of total ZyMōt-eligible cycle volume (clinic size) and depth of existing Cooper relationship (how many product categories they already buy).  
2. **What's the correct denominator for "penetration"?** → Not total IVF cases (overstates — ZyMōt is optional). Not ICSI alone (understates — conventional IVF also needs sperm prep). The right answer: **IVF \+ ICSI cycles combined**, as reported by the account manager into Cooper's internal system.  
3. **When do we discount, and when do we hold list price?** → Discount deepest for high-potential accounts we don't yet own (strategic acquisition). Hold list or small premium for loyal deep accounts (relationship covers price). Hold list for low-potential accounts (no investment, no discount).  
4. **What's the goal of the pricing recommendation?** → Two things, not one: (a) grow ZyMōt share within the client's eligible cycles, and (b) encourage the client to grow their total Cooper portfolio footprint over time.

### Phase 2 — Build the model

A two-axis strategic matrix: **Account Potential × ZyMōt Penetration**.

**Account Potential score (0–100)** combined:

- **70%** weight: total eligible cycles (IVF \+ ICSI) — primary signal of clinic size  
- **30%** weight: number of Cooper product categories already owned — relationship depth signal

Percentile-bucketed into **Big (top 20%) / Average (middle 30%) / Small (bottom 50%)** so the classification self-calibrates as the base grows.

**ZyMōt Penetration** \= current ZyMōt units purchased ÷ total eligible cycles. Bucketed into Low (\<20%) / Mid (20–60%) / High (\>60%).

Combining the two axes produces a 3×3 matrix with nine strategic cells:

|  | Low Pen (\<20%) | Mid Pen (20–60%) | High Pen (\>60%) |
| :---- | :---- | :---- | :---- |
| **Big Clinic** | **WIN** · −16% | **GROW** · −6% | **PROTECT** · \+2% |
| **Average Clinic** | **ACQUIRE** · −12% | **MAINTAIN** · list | **DEFEND** · \+2% |
| **Small Clinic** | **EVALUATE** · list | **MAINTAIN** · list | **HARVEST** · \+4% |

Each tier also carried a **target penetration lift** (e.g. WIN clinics should move from current penetration \+ 30 percentage points, capped at 80%) and an **acceptance confidence** (probability the client accepts the proposed pricing) so the model could produce both gross and weighted-expected revenue figures.

### Phase 3 — Ship a tool account managers would actually open

The deliverable had to be usable by a commercial team under time pressure, not an analyst stack. A single **Excel workbook with five sheets**, every classification and pricing lookup driven by live formulas (VLOOKUP, PERCENTILE, named ranges).

> \[\!TIP\] No macros, no VBA, no external dependencies. The whole model runs on native Excel formulas — any sales ops lead can inspect the logic and change the matrix in 30 seconds.

| Sheet | What it does |
| :---- | :---- |
| Cover | Business logic summary \+ list price \+ penetration cap. Two global levers. |
| Pricing Matrix | Editable 3×3 matrix with lookup table underneath. Change any cell → whole portfolio recalculates. |
| Client Database | 40 clinics × 29 columns (13 inputs \+ 16 live-computed outputs). |
| Client Lookup | Dropdown-driven single-client recommendation card. |
| Portfolio Summary | KPIs \+ by-tier \+ by-country aggregates. |

## 4\. The model

Full input → output flow per client:

```
INPUTS (from CRM):                      OUTPUTS:
  IVF cycles / year        ──┐           ┌──  Tier (WIN / GROW / ACQUIRE / etc.)
  ICSI cycles / year       ──┤           ├──  Recommended price (€)
  Current ZyMōt units       │           ├──  Price change vs current (%)
  Current price (€)        ──┼──MODEL──→ ├──  Target penetration %
  # Cooper categories       │           ├──  Target units / year
  (Cooper group spend €)    │           ├──  Additional units
  (Account manager, years) ──┘           ├──  Additional revenue (gross €)
                                         ├──  Acceptance confidence
                                         └──  Weighted expected revenue (€)
```

## 5\. How account managers use the tool

Before a quarterly client review, the AM:

1. **Opens** the Client Lookup sheet.  
2. **Picks** the client from a dropdown (e.g. `C-004 — Atlantique Reproduction`).  
3. **Reads** the recommendation card — tier, recommended price, target penetration lift, expected additional revenue.  
4. **Walks into the client meeting** with a defensible number and the business reason behind it ("you're one of our strategic Big-Clinic accounts, we want to deepen our partnership — here's the price we propose, here's the volume target we think we can hit together").  
5. **Portfolio Summary** aggregates every AM's book into a leadership-visible view of the total opportunity across EMEA.

>   
> \[\!NOTE\] The tool was never designed to replace judgment. It was designed to make the first 30 seconds of a pricing conversation defensible.

## 6\. Results after 3 months

Observed effects after the tool went into active use across the EMEA commercial team:

- **ZyMōt sales up 5–10%** across the European portfolio.  
- **Positive feedback from the account manager team** — the tool clarified price justification in client conversations and gave AMs real-time data to back their proposals, instead of relying on gut and memory.  
- **Several clients expanded their Cooper product footprint** beyond ZyMōt — the structured conversations surfaced adjacent product opportunities (media, cryo, catheters) that AMs then followed up on.  
- **Contributed to end-of-year revenue target attainment** for the product line — not the sole driver, but a measurable one.

The secondary win was cultural: AMs started asking Sales Ops for pricing support on *other* product lines, not just ZyMōt. The tool became a template.

## 7\. What's in this repository

| File | Purpose |
| :---- | :---- |
| `ZyMot_Pricing_Tool.xlsx` | The headline deliverable — 5-sheet interactive workbook |
| `ZyMot_Pricing_Case_Study.pptx` | 15-slide presentation (16:9) |
| `client_dataset.csv` | 40 synthetic European IVF clinics with full profiles |
## 8\. How to reproduce

Requires Python 3.10+ with `openpyxl` and `python-pptx` installed. LibreOffice is only needed to regenerate the PDF export.

```sh
pip install openpyxl python-pptx
python3 generate_clients.py        # writes client_dataset.csv
python3 build_pricing_tool.py      # writes ZyMot_Pricing_Tool.xlsx
python3 build_presentation.py      # writes ZyMot_Pricing_Case_Study.pptx

# Optional — PDF export of the deck
soffice --headless --impress --convert-to pdf ZyMot_Pricing_Case_Study.pptx
```

> \[\!TIP\] The dataset generator uses `random.seed(2026)` so every run produces identical data. Change the seed in `generate_clients.py` to generate a different 40-clinic sample.

## 9\. Caveats

> \[\!WARNING\] All data in this repository is synthetic. The 40 clinic names, their country/city, cycle counts, prices, and Cooper spend figures are fictional. The real CooperSurgical client base is not reproduced here.  
> 

- **Pricing anchors are illustrative.** The list price (€250) and tier modifiers (−16% to \+4%) were chosen to demonstrate the model shape, not to reflect CooperSurgical's actual pricing structure.  
- **The methodology is reconstructed from memory.** The two-axis matrix, scoring formula, and tool architecture are my best reconstruction of the real project built at CooperSurgical EMEA in 2024\. The exact weights, cell labels, and breakpoints likely differed in the live tool.  
- **The 5–10% sales uplift is a real observed result** from the deployed tool at CooperSurgical, attributed to the commercial team's adoption over roughly a quarter. The number is directional rather than a controlled measurement — commercial outcomes have multiple inputs.

---

*Built: November 2026 · Tools: Python 3.10, openpyxl 3.1, python-pptx 1.0, LibreOffice (for PDF export).*
