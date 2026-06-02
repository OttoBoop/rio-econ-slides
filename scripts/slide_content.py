"""
All slide text/data as structured dicts.
Edit this file to change content without touching layout code in build_pptx.py.
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
            "COVID-19 (2020): additional shock to employment and urban activity",
            "Structural constraint: most fiscal levers sit at federal & state level — municipalities must be creative",
            "Adversarial state government: limited coordination on security, transport, regulation",
            "Starting point: unemployment at 15.9% · economy contracting · labor market scarred",
        ],
        "footer_note": "Despite these constraints, Rio's city government chose to act — across 12 axes of inclusive economic development.",
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
    },

    # ── 4. CHART: RIO vs BRAZIL (P06) ────────────────────────────────────────
    {
        "id": "chart_p06",
        "type": "chart_slide",
        "title": "Rio Recovered Faster Than Brazil",
        "subtitle": "Economic Activity Index — Rio de Janeiro vs. Brazil (Base: 1Q2003 = 100)",
        "graph": "P06_atividade_economica_br_rj.png",
        "callout": "From 2022 onward: RJ growth rate consistently above Brazil's — first time since the 1990s.",
        "source": "Source: IBC-Br (BCB) · IBCR-RJ (BCB SGS 25397) · Federal Reserve FRED",
    },

    # ── 5. CHART: LABOR MARKET ───────────────────────────────────────────────
    {
        "id": "chart_labor",
        "type": "chart_slide",
        "title": "A Labor Market Transformed",
        "subtitle": "Unemployment Rate — Rio de Janeiro (%) · PNAD Contínua / IBGE",
        "graph": "C11_desemprego.png",
        "callout": "Unemployment halved in 4 years. ~300,000 residents left precarious labor conditions (2021–2024).",
        "source": "Source: PNAD Contínua / IBGE · Tabela SIDRA 4093 · IPP-Rio",
    },

    # ── 6. STRATEGY OVERVIEW: 12 AXES ────────────────────────────────────────
    {
        "id": "axes_overview",
        "type": "axes_grid",
        "title": "The Strategy: 12 Axes of Inclusive Growth",
        "intro": "Rio's recovery was deliberate — structured around 12 interlocking policy axes, helping firms of every size and sector.",
        "axes": [
            {"num": "01", "label": "Innovation &\nTechnology"},
            {"num": "02", "label": "Green Economy\n& Sustainability"},
            {"num": "03", "label": "Regulatory\nSimplification"},
            {"num": "04", "label": "Financial Sector\n& AI"},
            {"num": "05", "label": "Tourism &\nCreative Economy"},
            {"num": "06", "label": "City Center\nRevitalization"},
            {"num": "07", "label": "Workforce &\nHuman Capital"},
            {"num": "08", "label": "Inclusive Growth\n& Inequality"},
            {"num": "09", "label": "Entrepreneurship\n& SMEs"},
            {"num": "10", "label": "Public Mobility\n& Urban Planning"},
            {"num": "11", "label": "Infrastructure\n& Logistics"},
            {"num": "12", "label": "Public Security\n(Força Municipal)"},
        ],
        "note": "Axis 12 is addressed separately — it is not covered in the academic paper.",
    },

    # ── 7. AXIS: INNOVATION & TECH ───────────────────────────────────────────
    {
        "id": "axis_innovation",
        "type": "axis_detail",
        "axis_num": "01",
        "axis_label": "Innovation & Technology",
        "title": "Building an Innovation Ecosystem",
        "items": [
            {
                "program": "Porto Maravalley / IMPA Tech",
                "detail": "Transform the port region into a global innovation hub · 700 students over 10 years · R$ 220M payroll potential",
            },
            {
                "program": "Web Summit Rio",
                "detail": "One of the world's largest tech conferences relocated to Rio · 1M audience across 8 editions · R$ 1.8B economic impact",
            },
            {
                "program": "Programadores Cariocas",
                "detail": "750 young people trained in tech (70% Black, 40% women) · ~50% job insertion rate",
            },
            {
                "program": "ISS Tech",
                "detail": "Tax incentive for tech firms · R$ 62.7M in annual payroll attracted",
            },
            {
                "program": "Sandbox.Rio",
                "detail": "Experimental regulatory sandbox for tech startups — test innovations without standard licensing barriers",
            },
        ],
    },

    # ── 8. AXIS: FINANCIAL SECTOR & AI ───────────────────────────────────────
    {
        "id": "axis_finance",
        "type": "axis_detail",
        "axis_num": "04",
        "axis_label": "Financial Sector & AI",
        "title": "Repositioning Rio as Brazil's Financial Hub",
        "items": [
            {
                "program": "BASE / ATG — New Stock Exchange",
                "detail": "Municipal Law 8,467/2024: new derivatives & commodities exchange in Rio · institutional repositioning of the city as a financial center",
            },
            {
                "program": "Rio AI City",
                "detail": "Ecosystem for AI, financial technology, and data centers · target: 1 GW installed data center capacity by 2028",
            },
            {
                "program": "Partnerships: CVM · BCB · ANBIMA",
                "detail": "Formal institutional partnerships with Brazil's capital markets regulator, central bank, and asset management association",
            },
        ],
        "callout": "Rio is Brazil's historic financial capital — we are reclaiming that position with modern infrastructure.",
    },

    # ── 9. AXIS: TOURISM & CREATIVE ECONOMY ──────────────────────────────────
    {
        "id": "axis_tourism",
        "type": "axis_detail",
        "axis_num": "05",
        "axis_label": "Tourism & Creative Economy",
        "title": "Making Rio a Cultural & Tourism Powerhouse",
        "items": [
            {
                "program": "Tourism growth target",
                "detail": "+15% national and international tourists by 2028",
            },
            {
                "program": "Mega-events strategy",
                "detail": "8 events of 50,000+ participants · Carioca Arts Week (600,000 audience target by 2028)",
            },
            {
                "program": "Galeão International Airport",
                "detail": "+10% annual passenger growth · R$ 50.6B GDP impact over 10 years · 684,000+ jobs linked",
            },
            {
                "program": "Film industry",
                "detail": "+30% increase in authorized filming days for major international productions",
            },
            {
                "program": "Reviver Rua da Carioca / Rua da Cerveja",
                "detail": "Gastronomic & cultural hub in historic center · 500 jobs · R$ 222M economic impact (4 years)",
            },
        ],
    },

    # ── 10. AXIS: GREEN ECONOMY ───────────────────────────────────────────────
    {
        "id": "axis_green",
        "type": "axis_detail",
        "axis_num": "02",
        "axis_label": "Green Economy & Sustainability",
        "title": "Green Growth as Competitive Advantage",
        "items": [
            {
                "program": "ISS Neutro (Neutral ISS)",
                "detail": "Tax credit for firms that offset their carbon footprint · 250,000 tCO₂eq in mitigation projects · R$ 5.5M in tax credit value",
            },
            {
                "program": "Voluntary Carbon Credit Platform (B3 / ACX)",
                "detail": "Rio as the operational hub for Brazil's emerging voluntary carbon market",
            },
            {
                "program": "Green infrastructure targets (2025–2028)",
                "detail": "Integrated sustainable urban development — green public spaces, sustainable transport corridors",
            },
        ],
        "callout": "Rio's green economy agenda aligns with European ESG standards — a natural partner for German capital.",
    },

    # ── 11. AXIS: REGULATORY SIMPLIFICATION ──────────────────────────────────
    {
        "id": "axis_regulatory",
        "type": "axis_detail",
        "axis_num": "03",
        "axis_label": "Regulatory Simplification & Economic Freedom",
        "title": "Cutting Red Tape — Dramatically",
        "items": [
            {
                "program": "Lei da Liberdade Econômica",
                "detail": "Economic Freedom Law: estimated R$ 3.8K GDP per capita impact · potential for 130,000 new jobs",
            },
            {
                "program": "LICIN — Digital Business Licensing",
                "detail": "License analysis time cut from ~9 months to ~30 days",
            },
            {
                "program": "100% Digital Processes",
                "detail": "All new business registrations fully digital since January 2022 — zero physical paperwork",
            },
        ],
        "callout": "Ease of doing business improved dramatically — and this directly benefits foreign investors.",
    },

    # ── 12. REMAINING AXES (QUICK GRID) ──────────────────────────────────────
    {
        "id": "axes_remaining",
        "type": "remaining_axes",
        "title": "The Full Picture: All 12 Axes",
        "remaining": [
            {
                "num": "06",
                "label": "City Center Revitalization",
                "detail": "Reviver Centro: new residents, reoccupied buildings, integrated urban requalification. Reviver Cultural: 43 approved cultural projects.",
            },
            {
                "num": "07",
                "label": "Workforce & Human Capital",
                "detail": "Digital training platform · 15,000 people with disabilities trained · Hub Favela Empreendedora · Biblioteca do Saber.",
            },
            {
                "num": "08",
                "label": "Inclusive Growth",
                "detail": "Emergency COVID measures (Auxílio Empresa Carioca, Crédito Carioca) · 5,000+ micro enterprises supported · 18,000 jobs preserved.",
            },
            {
                "num": "09",
                "label": "Entrepreneurship & SMEs",
                "detail": "Crédito Carioca: R$ 5M in SME credit · Santa Cruz Industrial District infrastructure improvement.",
            },
            {
                "num": "10",
                "label": "Public Mobility",
                "detail": "Urban planning integration: land use, transport, housing, and economic development as a single strategy.",
            },
            {
                "num": "11",
                "label": "Infrastructure & Logistics",
                "detail": "Airport sector coordination · Santa Cruz Industrial District · data center capacity expansion.",
            },
        ],
    },

    # ── 13. BREAK SLIDE ──────────────────────────────────────────────────────
    {
        "id": "break",
        "type": "break_slide",
        "title": "Beyond the Paper",
        "subtitle": "What We're Also Building",
        "note": "The following content goes beyond the published academic paper.\nIt reflects ongoing work at SMDEIS and the Gabinete Flávio Valle.",
    },

    # ── 14. PUBLIC SECURITY ───────────────────────────────────────────────────
    {
        "id": "security",
        "type": "axis_detail",
        "axis_num": "12",
        "axis_label": "Public Security — Força Municipal",
        "title": "CompStat Rio: Data-Driven Safety",
        "items": [
            {
                "program": "Força de Segurança Municipal",
                "detail": "Municipal security force launched in 2025 — explicitly modeled on New York's CompStat revolution of the 1990s",
            },
            {
                "program": "Weekly accountability meetings",
                "detail": "Every Tuesday: precinct commanders confronted with real-time crime maps · executive presence required · no 'Compstat Lite'",
            },
            {
                "program": "22 Priority Areas",
                "detail": "Data-defined hotspots · integrated with CIVITAS intelligence platform · granular neighborhood-level targeting",
            },
            {
                "program": "Why it matters for investors",
                "detail": "Public safety is a prerequisite for tourism, hospitality, and urban commerce investment — we are addressing it with the same empirical rigor as our economic programs.",
            },
        ],
        "disclaimer": "⚠ This axis is not covered in the academic paper by Lima et al. (2026). It represents work in progress at the Gabinete do Vereador Flávio Valle and SMDEIS.",
    },

    # ── 15. FORWARD 2025–2028 ─────────────────────────────────────────────────
    {
        "id": "forward",
        "type": "forward",
        "title": "The Road Ahead: 2025–2028",
        "subtitle": "Rio Legado e Futuro — Strategic Plan",
        "bullets": [
            "88 goals · 134 projects · 30 strategic initiatives",
            "Economy & Innovation: +10% IAE-Rio activity index · 35,000 new construction jobs by 2028",
            "Tourism & Culture: +15% tourists · 8 mega-events · 100 new cultural spaces",
            "Human Capital: 15,000 PWDs trained · favela entrepreneur hub · knowledge library network",
            "Strategic: Rio AI City · 1 GW data center capacity · financial sector consolidation",
        ],
        "footer_note": "All 2025–2028 targets are embedded in the official Strategic Plan — publicly accountable and monitored quarterly.",
    },

    # ── 16. WHY RIO? ──────────────────────────────────────────────────────────
    {
        "id": "why_rio",
        "type": "why_slide",
        "title": "Why Rio — For German Investors",
        "points": [
            {
                "icon": "📈",
                "label": "Proven Recovery",
                "detail": "Outpaced Brazil's GDP growth in 2022 & 2023 — first time in 30 years. Unemployment halved.",
            },
            {
                "icon": "🏛",
                "label": "Financial Hub",
                "detail": "New stock exchange (BASE), CVM/BCB partnerships, AI city — Rio is re-emerging as Brazil's financial capital.",
            },
            {
                "icon": "🌿",
                "label": "ESG-Ready",
                "detail": "Voluntary carbon market (B3/ACX), ISS Neutro, green infrastructure — aligned with European investment standards.",
            },
            {
                "icon": "⚡",
                "label": "Ease of Business",
                "detail": "Licensing time cut from 9 months to 30 days. 100% digital processes. Economic Freedom Law.",
            },
            {
                "icon": "🎭",
                "label": "Cultural Capital",
                "detail": "Web Summit, Carnaval, film industry, tourism growth — Rio is Brazil's gateway to the world.",
            },
            {
                "icon": "🔒",
                "label": "Safety Improving",
                "detail": "CompStat Rio (Força Municipal): data-driven urban security — neighborhood-level accountability.",
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
            "for municipal councilman Flávio Valle, before working at SMDEIS — "
            "the Secretaria Municipal de Desenvolvimento Econômico, Inovação e Serviços "
            "of the City of Rio de Janeiro.\n\n"
            "Bopp thrives at the intersection of public policy, academic research, "
            "and innovative technologies."
        ),
        "photo": "otavio_headshot.jpeg",
    },
]
