"""
All slide text/data as structured dicts.
Edit this file to change content without touching layout code in build_pptx.py.

The 12 axes are confirmed by Otávio (from whatsapp info.txt, updated 2026-06-02).
They are strategic orientations/principles, not program categories.
Programs appear inside each axis as empirical proof.
"""

SLIDES = [
    # ── 1. TITLE ─────────────────────────────────────────────────────────────
    {
        "id": "title",
        "type": "title",
        "title": "Rio de Janeiro:\nA City Built for Investment",
        "subtitle": "Municipal strategies for inclusive economic growth\n2021–2028",
        "presenter": "Otávio Bopp",
        "role": "Economist | SMDEIS – Secretaria Municipal de Desenvolvimento Econômico,\nInovação e Serviços | City of Rio de Janeiro",
        "event": "Coleção Estudos Cariocas · IPP · June 2026",
    },

    # ── 2. THE CHALLENGE ─────────────────────────────────────────────────────
    {
        "id": "challenge",
        "type": "content",
        "title": "The Challenge",
        "bullets": [
            "Brazil's 2015–2016 recession: worst in a century — GDP fell ~7%",
            "COVID-19 (2020): another shock to employment and urban activity",
            "Structural constraint: most fiscal levers sit at federal & state level — municipalities must be creative",
            "Adversarial state government: limited coordination on security, transport, regulation",
            "Starting point (2020): unemployment at 15.9% · economy contracting · labor market scarred",
        ],
        "footer_note": "Despite these constraints, Rio's city government designed a structured, data-driven response across 12 strategic axes.",
    },

    # ── 3. THE RESULTS ───────────────────────────────────────────────────────
    {
        "id": "results",
        "type": "stats",
        "title": "The Results (2021–2024)",
        "stats": [
            {"value": "15.9% → 8.0%", "label": "Unemployment rate", "detail": "~7 pp drop · ~210,000 people re-employed"},
            {"value": "+640,000", "label": "Net new jobs vs. 2020", "detail": "Highest employed population in Rio's history (3.4M)"},
            {"value": "+330,000", "label": "New formal jobs (2021–24)", "detail": "Registered with carteira assinada"},
            {"value": "#2 Capital", "label": "Formal job creation rank", "detail": "In Brazil, behind only São Paulo"},
        ],
        "source": "Source: PNAD Contínua / IBGE · Caged · IPP-Rio (IAE-Rio indicator)",
        "highlight": "A city-level response — under federal and state constraints — delivered the highest employed population in Rio's recorded history: 3.4 million people in formal work.",
    },

    # ── 4. CHART: RIO vs BRAZIL (P06) ────────────────────────────────────────
    {
        "id": "chart_p06",
        "type": "chart_slide",
        "title": "Rio Recovered Faster Than Brazil",
        "subtitle": "Economic Activity Index — Rio de Janeiro vs. Brazil (Base: 1Q2003 = 100)",
        "graph": "P06_atividade_economica_br_rj.png",
        "callout": "From 2022 onward: RJ growth consistently above Brazil — first time since the 1990s.\n\nBR index: 166 · RJ index: 142 (2024). The gap is closing after three decades of underperformance.",
        "source": "Source: IBC-Br (BCB) · IBCR-RJ (BCB SGS 25397) · Federal Reserve FRED",
    },

    # ── 5. CHART: LABOR MARKET ───────────────────────────────────────────────
    {
        "id": "chart_labor",
        "type": "chart_slide",
        "title": "A Labor Market Transformed",
        "subtitle": "Unemployment Rate — Rio de Janeiro (%) · PNAD Contínua / IBGE",
        "graph": "C11_desemprego.png",
        "callout": "Unemployment halved in 4 years: 15.9% → 8.0%\n\n~300,000 residents moved out of precarious labor. Rio now at its lowest unemployment rate on record.",
        "source": "Source: PNAD Contínua / IBGE · Tabela SIDRA 4093 · IPP-Rio",
    },

    # ── 6. STRATEGY OVERVIEW: 12 AXES ────────────────────────────────────────
    {
        "id": "axes_overview",
        "type": "axes_grid",
        "title": "The Strategy: 12 Axes of Inclusive Growth",
        "intro": "Rio's recovery was structured around 12 interconnected strategic orientations — guiding all programs across every sector and neighborhood.",
        "axes": [
            {"num": "01", "label": "Driven by\nempirical data"},
            {"num": "02", "label": "Economy\nas a whole"},
            {"num": "03", "label": "Lowering\ninequalities"},
            {"num": "04", "label": "Creating\nhuman capital"},
            {"num": "05", "label": "Reducing\nbureaucracy"},
            {"num": "06", "label": "Supporting\ninnovation"},
            {"num": "07", "label": "Green\neconomy"},
            {"num": "08", "label": "City\nrevitalization"},
            {"num": "09", "label": "Urban\nroutes"},
            {"num": "10", "label": "Culture &\nevents"},
            {"num": "11", "label": "Rio as\neconomic center"},
            {"num": "12", "label": "Rethinking\npublic security"},
        ],
        "note": "Axis 12 is addressed separately — it goes beyond the academic paper.",
    },

    # ── 7. AXES 1–3: DATA / WHOLE ECONOMY / INEQUALITIES ─────────────────────
    {
        "id": "axes_1_3",
        "type": "multi_axis",
        "title": "Axes 1–3: Foundation of the Approach",
        "axes": [
            {
                "num": "01",
                "label": "Driven by empirical data",
                "programs": [
                    "IAE-Rio (Economic Activity Index): municipal real-time indicator developed to track recovery independently of federal data",
                    "All policies evaluated against measurable targets — employment, formalization, income, sector output",
                    "7 Rio-specific economic indicators built in-house and published quarterly — enabling faster, evidence-based decisions without waiting for state or federal releases",
                ],
            },
            {
                "num": "02",
                "label": "Economy as a whole",
                "programs": [
                    "Programs targeted all firm sizes (micro → large) and all sectors (services, industry, tourism, tech, finance, logistics)",
                    "No sector left behind: emergency support in 2020 + long-term structural bets in 2021–2024",
                    "Coverage extended to all 33 districts — geographic equity built into program design, not only downtown and South Zone",
                ],
            },
            {
                "num": "03",
                "label": "Lowering inequalities",
                "programs": [
                    "Auxílio Empresa Carioca: 5,000+ micro/small firms supported · 18,000 jobs preserved during COVID",
                    "Programadores Cariocas: 750 youth trained (70% Black, 40% women) · ~50% job insertion",
                    "Hub Favela Empreendedora (2025–2028): support and training specifically for favela entrepreneurs",
                ],
            },
        ],
    },

    # ── 8. AXES 4–5: HUMAN CAPITAL / BUREAUCRACY ─────────────────────────────
    {
        "id": "axes_4_5",
        "type": "multi_axis",
        "title": "Axes 4–5: People and Process",
        "axes": [
            {
                "num": "04",
                "label": "Creation of human capital",
                "programs": [
                    "Programadores Cariocas: tech training for 750 young people from vulnerable communities",
                    "Digital training platform (2025–2028): professional training, employment and income management",
                    "15,000 people with disabilities (PWDs) trained and placed by 2028",
                    "Biblioteca do Saber (Library of Knowledge): network expansion by 2026",
                ],
            },
            {
                "num": "05",
                "label": "Reducing bureaucracy",
                "programs": [
                    "Lei da Liberdade Econômica: estimated R$ 3.8K GDP per capita impact · 130,000 potential new jobs",
                    "LICIN — digital business licensing: analysis time cut from ~9 months to ~30 days",
                    "100% digital processes: all new business registrations fully digital since January 2022",
                    "Sandbox.Rio: experimental regulatory environment for tech startups",
                ],
            },
        ],
    },

    # ── 9. AXES 6–7: INNOVATION / GREEN ECONOMY ──────────────────────────────
    {
        "id": "axes_6_7",
        "type": "multi_axis",
        "title": "Axes 6–7: Innovation and Green Economy",
        "axes": [
            {
                "num": "06",
                "label": "Supporting innovation",
                "programs": [
                    "Porto Maravalley / IMPA Tech: port region → global innovation hub · 700 students over 10 years · R$ 220M payroll potential",
                    "Web Summit Rio: 1M audience across 8 editions · R$ 1.8B economic impact",
                    "ISS Tech: tax incentive attracting R$ 62.7M in annual tech payroll",
                    "1 GW operational data center capacity target by 2028 — Rio AI City ecosystem",
                ],
            },
            {
                "num": "07",
                "label": "Green economy",
                "programs": [
                    "ISS Neutro (Neutral ISS): voluntary carbon offset · 250,000 tCO₂eq in mitigation projects · R$ 5.5M in tax credits",
                    "B3 / ACX Carbon Credit Platform: Rio as operational hub for Brazil's voluntary carbon market",
                    "Green infrastructure integration in 2025–2028 Strategic Plan",
                    "Zero-carbon events framework: major conferences and mega-events required to offset emissions — Rio as model for sustainable urban hosting",
                ],
            },
        ],
    },

    # ── 10. AXIS 8: CITY REVITALIZATION ──────────────────────────────────────
    {
        "id": "axis_8",
        "type": "axis_detail",
        "axis_num": "08",
        "axis_label": "Revitalizing Different Parts of the City",
        "title": "From Downtown to Porto: Remaking Rio's Urban Fabric",
        "items": [
            {
                "program": "Reviver Centro",
                "detail": "Attract residents back to the city center · reoccupy vacant properties · integrated urban requalification of the historic downtown",
            },
            {
                "program": "Reviver Cultural",
                "detail": "43 contracts approved for cultural projects in the revitalized center",
            },
            {
                "program": "Rua da Carioca / Rua da Cerveja",
                "detail": "Gastronomic, cultural, and tourist hub in historic center · 500 jobs · R$ 41.8M payroll · R$ 222M economic impact (4 years)",
            },
            {
                "program": "Porto Maravalley (port region)",
                "detail": "Transform the abandoned port zone into a global innovation and tech hub — anchored by IMPA Tech",
            },
            {
                "program": "Santa Cruz Industrial District",
                "detail": "Infrastructure improvements in Rio's largest industrial hub — new developments in the west zone",
            },
        ],
    },

    # ── 11. AXIS 9: URBAN ROUTES ─────────────────────────────────────────────
    {
        "id": "axis_9",
        "type": "axis_detail",
        "axis_num": "09",
        "axis_label": "Rethinking Urban Routes",
        "title": "Mobility as Economic Infrastructure",
        "items": [
            {
                "program": "LRT, BRT & Intermodal Stations",
                "detail": "Light rail, bus rapid transit, and intermodal hubs connecting neighborhoods, reducing commute time and boosting labor market access",
            },
            {
                "program": "Cycleways & 'Motorcycleways'",
                "detail": "Expanded dedicated cycling infrastructure; motorcycle lane integration — Rio leads Brazil in urban cycling infrastructure expansion",
            },
            {
                "program": "Overpasses → Underpasses",
                "detail": "Major urban redesign: converting elevated roads into underground passages, reclaiming public space and improving traffic flow",
            },
            {
                "program": "Roundabouts & traffic redesign",
                "detail": "Intersection modernization program reducing congestion in key corridors",
            },
            {
                "program": "Galeão International Airport revitalization",
                "detail": "Strategic coordination with federal and private sectors · R$ 50.6B GDP impact over 10 years · 684,000+ jobs linked · +10% annual passenger growth target (2025–2028)",
            },
        ],
    },

    # ── 12. AXES 10–11: CULTURE & EVENTS / ECONOMIC CENTER ───────────────────
    {
        "id": "axes_10_11",
        "type": "multi_axis",
        "title": "Axes 10–11: Events and Economic Repositioning",
        "axes": [
            {
                "num": "10",
                "label": "Cultural and entrepreneurial events",
                "programs": [
                    "Web Summit Rio: world-class tech conference · 1M audience · R$ 1.8B economic impact · 8 editions",
                    "Carioca Arts Week: 600,000 audience target by 2028",
                    "8 mega-events of 50,000+ participants each (2025–2028)",
                    "+30% increase in authorized filming days for major international productions",
                    "+15% national and international tourists by 2028",
                ],
            },
            {
                "num": "11",
                "label": "Recentering Rio as economic center of Brazil",
                "programs": [
                    "BASE / ATG — New Stock Exchange: Law 8,467/2024 · derivatives and commodities exchange repositioning Rio as national financial capital",
                    "Rio AI City: AI, fintech, and data center ecosystem · formal partnerships with CVM, BCB, ANBIMA",
                    "1 GW data center capacity by 2028",
                    "+10% IAE-Rio economic activity target · 35,000 new construction jobs by 2028",
                ],
            },
        ],
    },

    # ── 13. BREAK SLIDE ──────────────────────────────────────────────────────
    {
        "id": "break",
        "type": "break_slide",
        "title": "Beyond the Paper",
        "subtitle": "What We're Also Building",
        "note": "The following content goes beyond the published academic paper by Lima, Balassiano et al. (2026).\nIt reflects ongoing work at SMDEIS and the Gabinete do Vereador Flávio Valle.",
    },

    # ── 14. AXIS 12: PUBLIC SECURITY ─────────────────────────────────────────
    {
        "id": "security",
        "type": "axis_detail",
        "axis_num": "12",
        "axis_label": "Rethinking Public Security",
        "title": "CompStat Rio: Data-Driven Urban Safety",
        "items": [
            {
                "program": "Força de Segurança Municipal",
                "detail": "Municipal security force launched in 2025 — explicitly modeled on New York's CompStat revolution of the 1990s, which halved homicides in a decade",
            },
            {
                "program": "Weekly accountability meetings",
                "detail": "Every Tuesday: precinct commanders confronted with real-time crime maps · executive presence required · genuine data confrontation, not 'Compstat Lite'",
            },
            {
                "program": "22 Priority Areas",
                "detail": "Data-defined hotspots · integrated with CIVITAS intelligence platform · granular neighborhood-level targeting",
            },
            {
                "program": "Why it matters for investors",
                "detail": "Public safety is a prerequisite for tourism, hospitality, retail, and urban commerce. We are addressing it with the same empirical rigor as our economic programs.",
            },
        ],
        "disclaimer": "⚠ This axis is NOT covered in Lima et al. (2026). It represents Otávio Bopp's work at the Gabinete do Vereador Flávio Valle and SMDEIS.",
    },

    # ── 15. FORWARD 2025–2028 ─────────────────────────────────────────────────
    {
        "id": "forward",
        "type": "forward",
        "title": "The Road Ahead: 2025–2028",
        "subtitle": "Rio Legado e Futuro — Official Strategic Plan",
        "bullets": [
            "88 goals · 134 projects · 30 strategic initiatives — all publicly accountable and monitored quarterly",
            "Economy & Innovation: +10% IAE-Rio activity index · 35,000 new construction jobs by 2028",
            "Tourism & Culture: +15% tourists · 8 mega-events · 100 new cultural spaces · 600,000 Carioca Arts Week audience",
            "Human Capital: 15,000 PWDs trained · favela entrepreneur hub · knowledge library network",
            "Strategic: Rio AI City · 1 GW data center capacity · financial sector consolidation (BASE, CVM/BCB/ANBIMA)",
            "Infrastructure: Galeão Airport revitalization · R$ 50.6B GDP impact over 10 years · 684,000+ jobs linked · intermodal integration with Linha 4 metro",
        ],
        "footer_note": "Source: Prefeitura do Rio — Plano Estratégico 2025–2028 'Rio Legado e Futuro'",
    },

    # ── 16. WHY RIO? ──────────────────────────────────────────────────────────
    {
        "id": "why_rio",
        "type": "why_slide",
        "title": "Why Rio — For German Investors",
        "points": [
            {
                "label": "Proven Recovery",
                "detail": "Outpaced Brazil's GDP growth in 2022 & 2023 — first time in 30 years. Unemployment halved. #2 for formal job creation.",
            },
            {
                "label": "Financial Hub",
                "detail": "New stock exchange (BASE), CVM/BCB/ANBIMA partnerships, AI city — Rio is reclaiming its role as Brazil's financial capital.",
            },
            {
                "label": "ESG-Ready",
                "detail": "Voluntary carbon market (B3/ACX), ISS Neutro, green infrastructure — aligned with European ESG investment standards.",
            },
            {
                "label": "Ease of Business",
                "detail": "Licensing time: 9 months → 30 days. 100% digital processes. Economic Freedom Law. Sandbox for innovation.",
            },
            {
                "label": "Cultural Capital",
                "detail": "Web Summit, Carnaval, film industry growth, +15% tourist target — Rio is Brazil's gateway to the world.",
            },
            {
                "label": "Safety Improving",
                "detail": "CompStat Rio (Força Municipal): data-driven, neighborhood-level security accountability — same model that transformed New York.",
            },
        ],
    },

    # ── 17. BIO ───────────────────────────────────────────────────────────────
    {
        "id": "bio",
        "type": "bio_slide",
        "name": "Otávio Bopp",
        "title": "Economist · SMDEIS",
        "bio": (
            "Otávio Bopp is an economist from Fundação Getúlio Vargas. "
            "He has worked as a researcher at IBRE (FGV) and as an advisor "
            "for municipal councilman Flávio Valle, before joining SMDEIS — "
            "the Secretaria Municipal de Desenvolvimento Econômico, Inovação e Serviços "
            "of the City of Rio de Janeiro.\n\n"
            "Bopp thrives at the intersection of public policy, academic research, "
            "and innovative technologies."
        ),
        "photo": "otavio_headshot.jpeg",
    },
]
