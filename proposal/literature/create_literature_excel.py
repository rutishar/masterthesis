"""Creates the literature screening Excel — aligned with thesis statement and SQ1–SQ3."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

papers = [
    # ── H — Hybrid AI / Symbolic / Subsymbolic / ESCO / Foundational ─────────
    {
        "cluster": "H", "bibkey": "garcez2022neurosymbolic",
        "authors": "Garcez, A. d'Avila; Lamb, L.C.",
        "year": 2023,
        "title": "Neurosymbolic AI: The 3rd Wave",
        "source": "Artificial Intelligence Review (Springer)",
        "doi": "10.1007/s10462-023-10448-w",
        "relevance": "High",
        "finding": "Defines neurosymbolic AI as the integration of symbolic reasoning with neural learning; demonstrates superior accuracy and interpretability over either paradigm in isolation — theoretical foundation for the hybrid AI architecture.",
        "category": "Research Design | SQ2",
    },
    {
        "cluster": "H", "bibkey": "gartner2024hype",
        "authors": "Gartner",
        "year": 2024,
        "title": "Gartner Hype Cycle for Artificial Intelligence",
        "source": "Gartner Research Report",
        "doi": "",
        "relevance": "High",
        "finding": "Identifies hybrid AI as an emerging high-value enterprise technology trend; relevant where domain expertise must be preserved while adapting to changing data distributions.",
        "category": "Research Design | SQ2",
    },
    {
        "cluster": "H", "bibkey": "le2014esco",
        "authors": "European Commission",
        "year": 2014,
        "title": "ESCO: European Skills, Competences, Qualifications and Occupations",
        "source": "Publications Office of the European Union",
        "doi": "",
        "relevance": "High",
        "finding": "Multilingual hierarchical taxonomy of 13,900+ skills with formal SKOS relationships — symbolic backbone for skill-demand modelling and ontological inference of skill substitutability.",
        "category": "SQ1 | SQ2",
    },
    {
        "cluster": "H", "bibkey": "peter2025ai4sme",
        "authors": "Peter, M.K.; Laurenzi, E.; Hinkelmann, K.",
        "year": 2025,
        "title": "Workshop Canvas AI-4-SME Framework",
        "source": "FHNW School of Business",
        "doi": "",
        "relevance": "High",
        "finding": "Three-phase Design–Build–Run methodology with KIA/DIA distinction and hybrid AI build path — overarching methodological structure of the thesis; maps symbolic (Neo4j, GraphDB) and subsymbolic (TensorFlow, PyTorch) components to the proposed architecture.",
        "category": "Research Design | SQ1 | SQ2 | SQ3",
    },
    {
        "cluster": "H", "bibkey": "lim2021tft",
        "authors": "Lim, B.; Arik, S.O.; Loeff, N.; Pfister, T.",
        "year": 2021,
        "title": "Temporal Fusion Transformers for Interpretable Multi-horizon Time Series Forecasting",
        "source": "International Journal of Forecasting (Elsevier)",
        "doi": "10.1016/j.ijforecast.2021.03.012",
        "relevance": "High",
        "finding": "TFT combines multi-head attention with interpretable variable selection for multi-horizon forecasting — state-of-the-art candidate for temporal skill-demand forecasting across CAPEX project lifecycle phases.",
        "category": "SQ1",
    },
    # ── A — CAPEX Portfolio Management / Strategic Alignment ──────────────────
    {
        "cluster": "A", "bibkey": "rodriguescoelho2025strategic",
        "authors": "Rodrigues Coelho, F.I.; Bizarrias, F.S.; Rabechini, R.; Martens, C.D.P.; Martens, M.L.",
        "year": 2025,
        "title": "Strategic Alignment and Value Optimization: Unveiling the Critical Role of Project Portfolio Management for a Flexible Environment",
        "source": "Global Journal of Flexible Systems Management (Springer)",
        "doi": "10.1007/s40171-024-00434-8",
        "relevance": "High",
        "finding": "SEM + NCA show strategic alignment mediates portfolio value — necessary but not sufficient condition; frames the thesis's alignment objective.",
        "category": "SQ2 | Research Design",
    },
    {
        "cluster": "A", "bibkey": "chiang2013strategic",
        "authors": "Chiang, I.R.; Nunez, M.A.",
        "year": 2013,
        "title": "Strategic alignment and value maximization for IT project portfolios",
        "source": "Information Technology and Management (Springer)",
        "doi": "10.1007/s10799-012-0126-9",
        "relevance": "High",
        "finding": "Integer programming model integrating strategic alignment score, expected benefit, cost, and synergy — directly analogous to the CAPEX de-prioritization scoring logic.",
        "category": "SQ2 | SQ3",
    },
    {
        "cluster": "A", "bibkey": "hansen2022seven",
        "authors": "Hansen, L.K.; Svejvig, P.",
        "year": 2022,
        "title": "Seven Decades of Project Portfolio Management Research (1950–2019) and Perspectives for the Future",
        "source": "Project Management Journal (SAGE / PMI)",
        "doi": "10.1177/87569728221089537",
        "relevance": "High",
        "finding": "Resource allocation under constraints and strategic re-alignment identified as the two most persistently unresolved PPM challenges — establishes the research gap context.",
        "category": "Research Design",
    },
    {
        "cluster": "A", "bibkey": "hansen2023principles",
        "authors": "Hansen, L.K.; Svejvig, P.",
        "year": 2023,
        "title": "Principles in Project Portfolio Management: Building Upon What We Know to Prepare for the Future",
        "source": "Project Management Journal (SAGE / PMI)",
        "doi": "10.1177/87569728231178427",
        "relevance": "Medium",
        "finding": "17 PPM principles in 4 semantic categories provide normative criteria for evaluating legitimacy of automated suspension recommendations.",
        "category": "SQ2 | Research Design",
    },
    {
        "cluster": "A", "bibkey": "rodriguezgarcia2025ai",
        "authors": "Rodriguez-Garcia, P.; Juan, A.A.; Martin, J.A.; Lopez-Lopez, D.; Marco, J.M.",
        "year": 2025,
        "title": "AI-driven Optimization of project portfolios in corporate ecosystems with synergies and strategic factors",
        "source": "Expert Systems with Applications (Elsevier)",
        "doi": "10.1016/j.eswa.2025.129593",
        "relevance": "High",
        "finding": "Hybrid ML + MIP framework integrating strategic priorities and synergies outperforms purely financial portfolio selection — closest existing AI-driven portfolio optimisation paper.",
        "category": "SQ2 | SQ3",
    },
    # ── B — Project Preemption / Suspension / De-Prioritization ──────────────
    {
        "cluster": "B", "bibkey": "ballestin2008preemption",
        "authors": "Ballestín, F.; Valls, V.; Quintanilla, S.",
        "year": 2008,
        "title": "Pre-emption in resource-constrained project scheduling",
        "source": "European Journal of Operational Research (Elsevier)",
        "doi": "10.1016/j.ejor.2007.05.055",
        "relevance": "High",
        "finding": "Seminal: single preemption yields significant makespan reductions; defines formal conditions under which suspension releases resources — foundational for the suspension-and-resumption logic.",
        "category": "SQ1 | SQ3",
    },
    {
        "cluster": "B", "bibkey": "vanpeteghem2010genetic",
        "authors": "Van Peteghem, V.; Vanhoucke, M.",
        "year": 2010,
        "title": "A genetic algorithm for the preemptive and non-preemptive multi-mode resource-constrained project scheduling problem",
        "source": "European Journal of Operational Research (Elsevier)",
        "doi": "10.1016/j.ejor.2009.01.056",
        "relevance": "High",
        "finding": "Preemptive flexibility consistently improves schedules when resource availability is uneven — suspension most effective when specialised skill pools are the binding constraint.",
        "category": "SQ1 | SQ3",
    },
    {
        "cluster": "B", "bibkey": "hatamimoghaddam2024robust",
        "authors": "Hatami-Moghaddam, L.; Khalilzadeh, M.; Shahsavari-Pour, N.; Sajadi, S.M.",
        "year": 2024,
        "title": "Developing a Robust Multi-Skill, Multi-Mode RCPSP with Partial Preemption, Resource Leveling, and Time Windows",
        "source": "Mathematics (MDPI)",
        "doi": "10.3390/math12193129",
        "relevance": "High",
        "finding": "Partial preemption triggered by multi-skill unavailability within time windows — closest existing model to the thesis's suspension mechanism and skill-shortage trigger.",
        "category": "SQ1 | SQ3",
    },
    {
        "cluster": "B", "bibkey": "vanhoucke2025impact",
        "authors": "Vanhoucke, M.; Demeulemeester, E.; Herroelen, W.",
        "year": 2025,
        "title": "The impact of the number of preemptions in RCPSP with time-varying resources",
        "source": "Annals of Operations Research (Springer)",
        "doi": "10.1007/s10479-025-06892-2",
        "relevance": "High",
        "finding": "Bi-criteria (makespan vs. interruptions) formulation quantifies optimal suspension frequency — supports data-driven dynamic policy over static threshold rules.",
        "category": "SQ1 | SQ3",
    },
    # ── C — Multi-Skill RCPSP / Multi-Project Scheduling ─────────────────────
    {
        "cluster": "C", "bibkey": "artigues2025fifty",
        "authors": "Artigues, C.; Hartmann, S.; Vanhoucke, M.",
        "year": 2025,
        "title": "Fifty years of research on resource-constrained project scheduling explored from different perspectives",
        "source": "European Journal of Operational Research (Elsevier)",
        "doi": "10.1016/j.ejor.2025.03.024",
        "relevance": "High",
        "finding": "Definitive 50-year review covering exact methods, heuristics, multi-skill, multi-project, and preemption extensions.",
        "category": "SQ1 | Research Design",
    },
    {
        "cluster": "C", "bibkey": "hartmann2022updated",
        "authors": "Hartmann, S.; Briskorn, D.",
        "year": 2022,
        "title": "An updated survey of variants and extensions of the resource-constrained project scheduling problem",
        "source": "European Journal of Operational Research (Elsevier)",
        "doi": "10.1016/j.ejor.2021.05.004",
        "relevance": "High",
        "finding": "Skill-based and multi-project RCPSP variants are the fastest-growing research areas — maps the solution landscape for the thesis.",
        "category": "SQ1 | Research Design",
    },
    {
        "cluster": "C", "bibkey": "gomez2023survey",
        "authors": "Gómez Sánchez, M.; Lalla-Ruiz, E.; Fernández Gil, A.; Castro, C.; Voß, S.",
        "year": 2023,
        "title": "Resource-constrained multi-project scheduling problem: A survey",
        "source": "European Journal of Operational Research (Elsevier)",
        "doi": "10.1016/j.ejor.2022.09.033",
        "relevance": "High",
        "finding": "Shared resource pools across concurrent projects are the primary conflict source — directly mirrors the 300+ project CAPEX portfolio environment.",
        "category": "SQ1 | SQ3",
    },
    {
        "cluster": "C", "bibkey": "snauwaert2021algorithm",
        "authors": "Snauwaert, J.; Vanhoucke, M.",
        "year": 2021,
        "title": "A new algorithm for resource-constrained project scheduling with breadth and depth of skills",
        "source": "European Journal of Operational Research (Elsevier)",
        "doi": "10.1016/j.ejor.2020.10.032",
        "relevance": "High",
        "finding": "Skill 'breadth' and 'depth' as two distinct modelling dimensions improve resource assignment — key methodological input for tiered engineering skill pools.",
        "category": "SQ1",
    },
    {
        "cluster": "C", "bibkey": "snauwaert2023classification",
        "authors": "Snauwaert, J.; Vanhoucke, M.",
        "year": 2023,
        "title": "A classification and new benchmark instances for the multi-skilled resource-constrained project scheduling problem",
        "source": "European Journal of Operational Research (Elsevier)",
        "doi": "10.1016/j.ejor.2022.05.049",
        "relevance": "High",
        "finding": "Standardised classification and benchmark datasets for MSRCPSP — empirical validation infrastructure for the thesis's skill demand modelling approach.",
        "category": "SQ1",
    },
    {
        "cluster": "C", "bibkey": "torba2024reallife",
        "authors": "Torba, R.; Dauzère-Pérès, S.; Yugma, C.; Gallais, C.; Pouzet, J.",
        "year": 2024,
        "title": "Solving a real-life multi-skill resource-constrained multi-project scheduling problem",
        "source": "Annals of Operations Research (Springer)",
        "doi": "10.1007/s10479-023-05784-7",
        "relevance": "High",
        "finding": "Memetic algorithm for real SNCF case (thousands of activities, multiple skill types) — closest real-world analogue to the 300+ project CAPEX context.",
        "category": "SQ1 | SQ3",
    },
    {
        "cluster": "C", "bibkey": "haroune2023multiproject",
        "authors": "Haroune, M.; Dhib, C.; Neron, E. et al.",
        "year": 2023,
        "title": "Multi-project scheduling problem under shared multi-skill resource constraints",
        "source": "TOP — Journal of the Spanish Statistics and Operations Research Society (Springer)",
        "doi": "10.1007/s11750-022-00633-5",
        "relevance": "High",
        "finding": "Multi-project assignment with shared employee skill pools and efficiency levels minimising weighted tardiness — integrates skill constraints with priority weights.",
        "category": "SQ1 | SQ2",
    },
    {
        "cluster": "C", "bibkey": "snauwaert2025hierarchical",
        "authors": "Snauwaert, J.; Vanhoucke, M.",
        "year": 2025,
        "title": "A solution framework for multi-skilled project scheduling problems with hierarchical skills",
        "source": "Journal of Scheduling (Springer)",
        "doi": "10.1007/s10951-025-00836-1",
        "relevance": "High",
        "finding": "Six hierarchical-skill MSRCPSP variants solved by GA with local search — hierarchical structure mirrors tiered engineering competency pools.",
        "category": "SQ1",
    },
    # ── D — Prescriptive Analytics / AI Decision Support ─────────────────────
    {
        "cluster": "D", "bibkey": "lepenioti2020prescriptive",
        "authors": "Lepenioti, K.; Bousdekis, A.; Apostolou, D.; Mentzas, G.",
        "year": 2020,
        "title": "Prescriptive analytics: Literature review and research challenges",
        "source": "International Journal of Information Management (Elsevier)",
        "doi": "10.1016/j.ijinfomgt.2019.04.003",
        "relevance": "High",
        "finding": "Defines prescriptive analytics; maps 56 key papers; identifies resource allocation as the primary application domain — foundational reference for the thesis's prescriptive framework.",
        "category": "SQ2 | Research Design",
    },
    {
        "cluster": "D", "bibkey": "frazzetto2019prescriptive",
        "authors": "Frazzetto, D.; Nielsen, T.D.; Pedersen, T.B.; Šikšnys, L.",
        "year": 2019,
        "title": "Prescriptive analytics: a survey of emerging trends and technologies",
        "source": "The VLDB Journal (Springer)",
        "doi": "10.1007/s00778-019-00539-y",
        "relevance": "High",
        "finding": "Three core PA pillars: constraint satisfaction, optimisation, simulation — maps directly to the technical design components of the proposed framework.",
        "category": "SQ2 | Research Design",
    },
    {
        "cluster": "D", "bibkey": "bertsimas2020predictive",
        "authors": "Bertsimas, D.; Kallus, N.",
        "year": 2020,
        "title": "From Predictive to Prescriptive Analytics",
        "source": "Management Science (INFORMS)",
        "doi": "10.1287/mnsc.2018.3253",
        "relevance": "High",
        "finding": "Theoretical bridge from demand forecasting to prescriptive optimisation; 'coefficient of prescriptiveness' as evaluation metric — mathematical foundation for the analytics chain.",
        "category": "SQ2 | SQ3",
    },
    {
        "cluster": "D", "bibkey": "wissuchek2025prescriptive",
        "authors": "Wissuchek, C.; Zschech, P.",
        "year": 2025,
        "title": "Prescriptive analytics systems revised: a systematic literature review from an IS perspective",
        "source": "Information Systems and e-Business Management (Springer)",
        "doi": "10.1007/s10257-024-00688-w",
        "relevance": "High",
        "finding": "23-component PAS architecture (data pipeline, intervention model, feedback loop) — direct design template for the thesis artifact; identifies multi-project CAPEX as an open gap.",
        "category": "SQ2 | Research Design",
    },
    {
        "cluster": "D", "bibkey": "misic2020analytics",
        "authors": "Mišić, V.V.; Perakis, G.",
        "year": 2020,
        "title": "Data Analytics in Operations Management: A Review",
        "source": "Manufacturing & Service Operations Management (INFORMS)",
        "doi": "10.1287/msom.2019.0805",
        "relevance": "Medium",
        "finding": "Prescriptive methods deliver greatest value when decision windows are tight and resource constraints binding — contextualises the thesis's use case.",
        "category": "SQ2 | Research Design",
    },
    # ── E — MCDM / Project Prioritization ────────────────────────────────────
    {
        "cluster": "E", "bibkey": "kandakoglu2024mcdm",
        "authors": "Kandakoglu, M.; Walther, G.; Ben Amor, S.",
        "year": 2024,
        "title": "The use of MCDM methods in project portfolio selection: a literature review and future research directions",
        "source": "Annals of Operations Research (Springer)",
        "doi": "10.1007/s10479-023-05564-3",
        "relevance": "High",
        "finding": "'Dynamic re-prioritisation under constraint changes' identified as a major open research gap — directly motivates SQ2.",
        "category": "SQ2 | Research Design",
    },
    {
        "cluster": "E", "bibkey": "chatterjee2018fuzzy",
        "authors": "Chatterjee, K.; Hossain, S.A.; Kar, S.",
        "year": 2018,
        "title": "Prioritization of project proposals in portfolio management using fuzzy AHP",
        "source": "OPSEARCH (Springer)",
        "doi": "10.1007/s12597-018-0331-3",
        "relevance": "High",
        "finding": "Fuzzy AHP handles linguistic expert uncertainty in multi-criteria prioritisation — methodological foundation for the dynamic weighting component of the MCDM framework.",
        "category": "SQ2",
    },
    {
        "cluster": "E", "bibkey": "alhashimi2017mcdm",
        "authors": "Al-Hashimi, M.; Chakrabortty, R.K.; Ryan, M.J.",
        "year": 2017,
        "title": "Analysis of MCDM methods output coherence in oil and gas portfolio prioritization",
        "source": "Journal of Petroleum Exploration and Production Technology (Springer)",
        "doi": "10.1007/s13202-017-0344-0",
        "relevance": "High",
        "finding": "AHP, PROMETHEE, TOPSIS compared on CAPEX-intensive oilfield portfolio — validates MCDM for de-prioritisation in capital-intensive industries.",
        "category": "SQ2",
    },
    # ── F — Workforce / Skill Demand Forecasting ──────────────────────────────
    {
        "cluster": "F", "bibkey": "safarishahrbijari2018workforce",
        "authors": "Safarishahrbijari, A.",
        "year": 2018,
        "title": "Workforce forecasting models: A systematic review",
        "source": "Journal of Forecasting (Wiley)",
        "doi": "10.1002/for.2541",
        "relevance": "High",
        "finding": "Taxonomy of time-series, Markov chain, system-dynamics, and simulation-based forecasting models — provides model selection framework for projecting temporal skill pool demand across 300+ projects.",
        "category": "SQ1",
    },
    {
        "cluster": "F", "bibkey": "macedo2022skills",
        "authors": "Macedo, M.M.G.; Clarke, W.; Lucherini, E.; Baldwin, T.; Queiroz, D.; de Paula, R.; Das, S.",
        "year": 2022,
        "title": "Practical Skills Demand Forecasting via Representation Learning of Temporal Dynamics",
        "source": "AAAI/ACM Conference on AI, Ethics, and Society (AIES '22) — ACM",
        "doi": "10.1145/3514094.3534183",
        "relevance": "High",
        "finding": "RNN/LSTM on 10 years of monthly skill demand data; multivariate skill correlation improves multi-step forecast accuracy — directly applicable to SQ1 temporal skill pool modelling.",
        "category": "SQ1",
    },
    # ── G — Design Science Research ───────────────────────────────────────────
    {
        "cluster": "G", "bibkey": "hevner2004design",
        "authors": "Hevner, A.R.; March, S.T.; Park, J.; Ram, S.",
        "year": 2004,
        "title": "Design Science in Information Systems Research",
        "source": "MIS Quarterly (AIS / University of Minnesota)",
        "doi": "10.2307/25148625",
        "relevance": "High",
        "finding": "Foundational DSR paradigm with 7 guidelines for IS artifact design, evaluation, and communication — primary methodological anchor for the thesis.",
        "category": "Research Design",
    },
    {
        "cluster": "G", "bibkey": "peffers2007dsrm",
        "authors": "Peffers, K.; Tuunanen, T.; Rothenberger, M.A.; Chatterjee, S.",
        "year": 2007,
        "title": "A Design Science Research Methodology for Information Systems Research",
        "source": "Journal of Management Information Systems (Taylor & Francis)",
        "doi": "10.2753/MIS0742-1222240302",
        "relevance": "High",
        "finding": "6-phase DSRM process model (problem → objectives → design → demo → evaluation → communication) — standard process cited in thesis chapters describing artifact construction.",
        "category": "Research Design",
    },
]

# ── Styling ───────────────────────────────────────────────────────────────────
cluster_meta = {
    "H": ("Hybrid AI / Symbolic / Subsymbolic / ESCO",   "D9D2E9"),
    "A": ("CAPEX Portfolio / Strategic Alignment",         "DEEAF1"),
    "B": ("Project Preemption / Suspension",               "E2EFDA"),
    "C": ("Multi-Skill RCPSP / Multi-Project Scheduling",  "FFF2CC"),
    "D": ("Prescriptive Analytics / AI Decision Support",  "FCE4D6"),
    "E": ("MCDM / Project Prioritisation",                 "EAD1DC"),
    "F": ("Workforce / Skill Demand Forecasting",          "D9EAD3"),
    "G": ("Design Science Research",                       "CFE2F3"),
}

header_fill  = PatternFill("solid", fgColor="1F4E79")
high_fill    = PatternFill("solid", fgColor="C6EFCE")
medium_fill  = PatternFill("solid", fgColor="FFEB9C")
low_fill     = PatternFill("solid", fgColor="FFC7CE")
alt_fill     = PatternFill("solid", fgColor="F2F2F2")
cluster_fills = {k: PatternFill("solid", fgColor=v[1]) for k, v in cluster_meta.items()}

header_font  = Font(bold=True, color="FFFFFF", size=11)
body_font    = Font(size=10)
mono_font    = Font(name="Courier New", size=9)
wrap_align   = Alignment(wrap_text=True, vertical="top")
center_align = Alignment(horizontal="center", vertical="top", wrap_text=True)
thin_border  = Border(
    left=Side(style="thin"),  right=Side(style="thin"),
    top=Side(style="thin"),   bottom=Side(style="thin"),
)

# ── Workbook ──────────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Literature Screening"

headers    = ["#", "Cluster", "BibTeX Key", "Authors", "Year", "Title",
              "Source / Journal", "DOI", "Relevance", "Key Finding", "Category", "Notes"]
col_widths = [4, 9, 26, 34, 6, 52, 32, 38, 11, 58, 30, 18]

for col_idx, (h, w) in enumerate(zip(headers, col_widths), start=1):
    cell = ws.cell(row=1, column=col_idx, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_align
    cell.border = thin_border
    ws.column_dimensions[get_column_letter(col_idx)].width = w

ws.row_dimensions[1].height = 22
ws.freeze_panes = "A2"

for i, p in enumerate(papers, start=1):
    row = i + 1
    values = [
        i, p["cluster"], p["bibkey"], p["authors"], p["year"],
        p["title"], p["source"], p["doi"], p["relevance"],
        p["finding"], p["category"], "",
    ]
    for col_idx, val in enumerate(values, start=1):
        cell = ws.cell(row=row, column=col_idx, value=val)
        cell.border = thin_border
        if col_idx == 3:          # BibTeX key — monospace
            cell.font = mono_font
            cell.alignment = wrap_align
        elif col_idx in (1, 2, 5, 9):
            cell.font = body_font
            cell.alignment = center_align
        else:
            cell.font = body_font
            cell.alignment = wrap_align

    ws.cell(row=row, column=2).fill = cluster_fills.get(p["cluster"], alt_fill)
    rel = ws.cell(row=row, column=9)
    rel.fill = high_fill if p["relevance"] == "High" else (
               medium_fill if p["relevance"] == "Medium" else low_fill)
    ws.row_dimensions[row].height = 85

ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

# ── Summary sheet ─────────────────────────────────────────────────────────────
ws2 = wb.create_sheet("Summary")
ws2.merge_cells("A1:D1")
ws2["A1"] = "Literature Screening Summary"
ws2["A1"].font = Font(bold=True, size=14, color="FFFFFF")
ws2["A1"].fill = PatternFill("solid", fgColor="1F4E79")

totals = {"High": 0, "Medium": 0, "Low": 0}
cluster_counts = {k: 0 for k in cluster_meta}
for p in papers:
    totals[p["relevance"]] += 1
    cluster_counts[p["cluster"]] += 1

ws2["A3"] = "Relevance"; ws2["B3"] = "Count"
ws2["A3"].font = Font(bold=True); ws2["B3"].font = Font(bold=True)
for r, (rel, fill) in enumerate(
        [("High", high_fill), ("Medium", medium_fill), ("Low", low_fill)], start=4):
    ws2.cell(r, 1, rel).fill = fill
    ws2.cell(r, 2, totals[rel])
ws2.cell(7, 1, "Total").font = Font(bold=True)
ws2.cell(7, 2, sum(totals.values())).font = Font(bold=True)

ws2["A9"] = "Cluster"; ws2["B9"] = "Description"; ws2["C9"] = "Count"
for c in ["A9","B9","C9"]: ws2[c].font = Font(bold=True)
for offset, (k, (name, color)) in enumerate(cluster_meta.items(), start=10):
    ws2.cell(offset, 1, k).fill = cluster_fills[k]
    ws2.cell(offset, 2, name)
    ws2.cell(offset, 3, cluster_counts[k])

ws2["A19"] = "Scopus Keyword Clusters"
ws2["A19"].font = Font(bold=True, size=12)
kw = [
    ("H", "hybrid AI knowledge graph ontology; neuro-symbolic AI; ESCO skill ontology; generative AI enterprise"),
    ("A", "CAPEX portfolio management; capital expenditure portfolio; project portfolio strategic alignment"),
    ("B", "project preemption; preemptive RCPSP; activity splitting; project suspension resumption"),
    ("C", "resource-constrained project scheduling; RCPSP; multi-skill scheduling; multi-project scheduling"),
    ("D", "prescriptive analytics; decision support resource allocation; from predictive to prescriptive"),
    ("E", "MCDM project prioritisation; AHP TOPSIS portfolio; dynamic weighting strategic alignment"),
    ("F", "workforce skill demand forecasting; temporal skill modelling; labour demand machine learning"),
    ("G", "Design Science Research methodology; DSR information systems; Hevner design science"),
]
for offset, (k, terms) in enumerate(kw, start=20):
    ws2.cell(offset, 1, k).fill = cluster_fills[k]
    ws2.cell(offset, 2, terms)

for col, w in [("A",12),("B",70),("C",10)]:
    ws2.column_dimensions[col].width = w
for r in range(20, 28):
    ws2.row_dimensions[r].height = 28
    ws2.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")

out = r"C:\Users\ramon\00_no_sync\masterthesis\proposal\literature_screening.xlsx"
wb.save(out)
print(f"Saved: {out}")
print(f"Total: {sum(totals.values())} papers  |  High: {totals['High']}  Medium: {totals['Medium']}  Low: {totals['Low']}")
for k, (name, _) in cluster_meta.items():
    print(f"  {k} — {name}: {cluster_counts[k]}")
