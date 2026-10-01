"""
generate_presentation_guide.py
Generates a comprehensive, beautifully-styled Microsoft Word (.docx) document
for the DSA-3 college project presentation: Threat-Intelligence Kill-Chain Correlation Engine.
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_document():
    doc = docx.Document()

    # 1-inch margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # ── Styling Helpers ──
    def set_cell_shading(cell, color_hex):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'''
            <w:tcMar {nsdecls("w")}>
                <w:top w:w="{top}" w:type="dxa"/>
                <w:bottom w:w="{bottom}" w:type="dxa"/>
                <w:left w:w="{left}" w:type="dxa"/>
                <w:right w:w="{right}" w:type="dxa"/>
            </w:tcMar>
        ''')
        tcPr.append(tcMar)

    def set_table_borders(table, color="CBD5E1", sz="4", val="single"):
        tblPr = table._tbl.tblPr
        borders = parse_xml(f'''
            <w:tblBorders {nsdecls("w")}>
                <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:left w:val="none"/>
                <w:right w:val="none"/>
                <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:insideV w:val="none"/>
            </w:tblBorders>
        ''')
        tblPr.append(borders)

    def set_callout_border(cell, color="1E3A8A", sz="24"):
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(f'''
            <w:tcBorders {nsdecls("w")}>
                <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:top w:val="none"/>
                <w:right w:val="none"/>
                <w:bottom w:val="none"/>
            </w:tcBorders>
        ''')
        tcPr.append(borders)

    def add_callout(title, text, callout_type="info"):
        colors = {
            "info":    {"border": "0284C7", "bg": "F0F9FF", "title": "0369A1"},
            "tip":     {"border": "16A34A", "bg": "F0FDF4", "title": "15803D"},
            "warning": {"border": "D97706", "bg": "FFFBEB", "title": "B45309"},
            "viva":    {"border": "1E3A8A", "bg": "F8FAFC", "title": "1E3A8A"},
        }
        c = colors.get(callout_type, colors["info"])

        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)

        set_cell_shading(cell, c["bg"])
        set_callout_border(cell, color=c["border"], sz="24")
        set_cell_margins(cell, top=140, bottom=140, left=180, right=180)

        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        run_t = p.add_run(f"📌 {title.upper()}")
        run_t.font.name = "Calibri"
        run_t.font.size = Pt(10.5)
        run_t.font.bold = True
        run_t.font.color.rgb = RGBColor.from_string(c["title"])

        p2 = cell.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        p2.paragraph_format.line_spacing = 1.15
        run_body = p2.add_run(text)
        run_body.font.name = "Calibri"
        run_body.font.size = Pt(10)
        run_body.font.color.rgb = RGBColor(30, 41, 59)

        doc.add_paragraph()

    def add_code_block(code_text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)

        set_cell_shading(cell, "F1F5F9")
        set_callout_border(cell, color="94A3B8", sz="12")
        set_cell_margins(cell, top=120, bottom=120, left=160, right=160)

        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.05
        run = p.add_run(code_text.strip())
        run.font.name = "Consolas"
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(15, 23, 42)

        doc.add_paragraph()

    def add_styled_heading(text, level):
        h = doc.add_heading(text, level=level)
        h.paragraph_format.keep_with_next = True
        run = h.runs[0]
        run.font.name = "Calibri"

        if level == 1:
            run.font.size = Pt(16.5)
            run.font.bold = True
            run.font.color.rgb = RGBColor(30, 58, 138)
            h.paragraph_format.space_before = Pt(18)
            h.paragraph_format.space_after = Pt(6)
        elif level == 2:
            run.font.size = Pt(13)
            run.font.bold = True
            run.font.color.rgb = RGBColor(2, 132, 199)
            h.paragraph_format.space_before = Pt(14)
            h.paragraph_format.space_after = Pt(4)
        elif level == 3:
            run.font.size = Pt(11)
            run.font.bold = True
            run.font.color.rgb = RGBColor(15, 23, 42)
            h.paragraph_format.space_before = Pt(10)
            h.paragraph_format.space_after = Pt(2)
        return h

    def add_p(text, bold_prefix="", italic_prefix=""):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15

        if bold_prefix:
            r_b = p.add_run(bold_prefix)
            r_b.font.name = "Calibri"
            r_b.font.size = Pt(10.5)
            r_b.font.bold = True
            r_b.font.color.rgb = RGBColor(15, 23, 42)

        if italic_prefix:
            r_i = p.add_run(italic_prefix)
            r_i.font.name = "Calibri"
            r_i.font.size = Pt(10.5)
            r_i.font.italic = True
            r_i.font.color.rgb = RGBColor(71, 85, 105)

        r_text = p.add_run(text)
        r_text.font.name = "Calibri"
        r_text.font.size = Pt(10.5)
        r_text.font.color.rgb = RGBColor(30, 41, 59)
        return p

    def add_bullet(text, bold_prefix=""):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15

        if bold_prefix:
            r_b = p.add_run(bold_prefix)
            r_b.font.name = "Calibri"
            r_b.font.size = Pt(10.5)
            r_b.font.bold = True
            r_b.font.color.rgb = RGBColor(15, 23, 42)

        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(10.5)
        r.font.color.rgb = RGBColor(30, 41, 59)
        return p

    def add_styled_table(headers, data, col_widths=None):
        tbl = doc.add_table(rows=len(data) + 1, cols=len(headers))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(tbl, color="CBD5E1")

        hdr_row = tbl.rows[0]
        trPr = hdr_row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

        for idx, header_text in enumerate(headers):
            cell = hdr_row.cells[idx]
            set_cell_shading(cell, "1E3A8A")
            set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(header_text)
            run.font.name = "Calibri"
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

        for row_idx, row_data in enumerate(data):
            row = tbl.rows[row_idx + 1]
            bg_color = "F8FAFC" if row_idx % 2 == 1 else "FFFFFF"
            for col_idx, cell_value in enumerate(row_data):
                cell = row.cells[col_idx]
                if bg_color != "FFFFFF":
                    set_cell_shading(cell, bg_color)
                set_cell_margins(cell, top=90, bottom=90, left=140, right=140)
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.1
                run = p.add_run(str(cell_value))
                run.font.name = "Calibri"
                run.font.size = Pt(9.5)
                run.font.color.rgb = RGBColor(30, 41, 59)

        if col_widths:
            for row in tbl.rows:
                for idx, width in enumerate(col_widths):
                    row.cells[idx].width = width

        doc.add_paragraph()

    # ═══════════════════════════════════════════════════════════════
    # COVER / TITLE BLOCK
    # ═══════════════════════════════════════════════════════════════
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(10)
    title_p.paragraph_format.space_after = Pt(2)
    run_title = title_p.add_run("Threat-Intelligence Kill-Chain Correlation Engine")
    run_title.font.name = "Calibri"
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(30, 58, 138)

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_before = Pt(0)
    sub_p.paragraph_format.space_after = Pt(12)
    run_sub = sub_p.add_run("Comprehensive Presentation, Architecture & Viva Defense Manual")
    run_sub.font.name = "Calibri"
    run_sub.font.size = Pt(13)
    run_sub.font.color.rgb = RGBColor(2, 132, 199)

    # Metadata Box
    meta_tbl = doc.add_table(rows=4, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(meta_tbl, color="E2E8F0")
    meta_info = [
        ("Course / Context:", "Data Structures and Algorithms III (DSA-3) | College Capstone Project"),
        ("Technology Stack:", "Frontend: React 18 + Vite + Tailwind | Backend: FastAPI (Python 3.14) | DB: MongoDB | DSA: Pure Python"),
        ("Core Algorithms:", "Hashing (O(1)), KMP Pattern Matching (O(n+m)), Graph Adjacency List (O(V+E)), BFS & DFS, Tarjan's Articulation Points (O(V+E)), Greedy Set Cover (O(P*N²)), Max-Heap Priority Queue (O(log n))"),
        ("Project Status:", "64 / 64 Unit Tests Passing (pytest) | Full-Stack End-to-End Pipeline Active"),
    ]
    for r_idx, (k, v) in enumerate(meta_info):
        row = meta_tbl.rows[r_idx]
        set_cell_shading(row.cells[0], "F1F5F9")
        set_cell_shading(row.cells[1], "FFFFFF")
        set_cell_margins(row.cells[0], top=60, bottom=60, left=100, right=100)
        set_cell_margins(row.cells[1], top=60, bottom=60, left=100, right=100)
        p0 = row.cells[0].paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(k)
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        p1 = row.cells[1].paragraphs[0]
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(v)
        r1.font.size = Pt(9.5)

    doc.add_paragraph()

    add_callout(
        "Presenter Emergency Note for Tomorrow",
        "This guide is specifically tailored for presenting and defending this project in front of college faculty and reviewers. "
        "It breaks down the exact verbal pitches, mathematical proofs of complexities, the rationale behind every data structure choice, "
        "step-by-step UI demo instructions, and 30 anticipated viva questions with winning technical answers.",
        callout_type="tip"
    )

    # ═══════════════════════════════════════════════════════════════
    # SECTION 1: EXECUTIVE OVERVIEW & PITCH SCRIPTS
    # ═══════════════════════════════════════════════════════════════
    add_styled_heading("1. Executive Overview & Verbal Pitch Scripts", level=1)

    add_styled_heading("1.1 The 30-Second Elevator Pitch (Memorize or Read Aloud)", level=2)
    add_p(
        '"Good morning, Respected Reviewers. Our project is the Threat-Intelligence Kill-Chain Correlation Engine. '
        'In cybersecurity, modern Security Operations Centers are flooded with thousands of disconnected log events every day. '
        'Our project demonstrates how 8 classical Data Structures and Algorithms—such as KMP string matching, Graph traversal, '
        'Tarjan\'s cut-vertex algorithm, Greedy set cover, and Max-Heap priority queues—can systematically parse raw events, '
        'detect coordinated cyberattacks, reconstruct multi-stage kill chains, and isolate critical network choke points. '
        'We implemented every single algorithm in pure Python from scratch, backed by FastAPI, MongoDB, and an interactive React dashboard."',
        italic_prefix="Script: "
    )

    add_styled_heading("1.2 The 2-Minute Deep Pitch (When Reviewer Asks for a Walkthrough)", level=2)
    add_bullet("1. Ingestion & O(1) Hashing: Raw security logs arrive in CSV format. Using custom hash tables, we classify incoming event types in O(1) time without scanning signature databases.")
    add_bullet("2. KMP Pattern Matching: An attack is rarely a single event; it is an ordered sequence (e.g., three failed logins, or malware followed by privilege escalation). Using Knuth-Morris-Pratt pattern matching with an LPS array, we scan event streams in linear O(n + m) time without backtracking.")
    add_bullet("3. Alert Generation & Hash-Based Deduplication: Matches produce alert objects. Duplicate alerts are pruned using a hash set with composite keys.")
    add_bullet("4. Multi-Stage Correlation: Alerts sharing the same source IP are grouped into coherent attack campaigns in O(n) time.")
    add_bullet("5. Honest Kill Chain Reconstruction: We map correlated stages against the 7 standard MITRE ATT&CK stages. Crucially, our engine only reports 'Observed' stages if backed by real data—never fabricating evidence.")
    add_bullet("6. Graph Modeling (Adjacency List): We model the network topology as a directed graph in O(V + E) space. We run BFS to find the shortest attack paths to databases and DFS to trace all lateral movement branches.")
    add_bullet("7. Tarjan's Articulation Points: Using discovery and low times in a single DFS pass (O(V + E)), we identify network choke points—nodes whose isolation would sever attack paths.")
    add_bullet("8. Greedy Monitoring Optimization: To solve the NP-hard Set Cover problem of sensor placement, our greedy optimizer selects the minimum monitoring nodes to cover all attack paths.")
    add_bullet("9. Max-Heap Alert Prioritization: Alerts are pushed into a Max-Heap priority queue so analysts always triage CRITICAL threats before HIGH, MEDIUM, or LOW in O(log n) time.")

    add_styled_heading("1.3 Academic Prototype Scope & Honesty", level=2)
    add_callout(
        "Reviewer Appreciation Strategy: Academic Honesty",
        "Reviewers frequently deduct marks if students claim their college project is a 'commercial-grade enterprise SOC tool'. "
        "State clearly up front: 'This is an academic prototype designed to prove the utility of Data Structures & Algorithms in cybersecurity. "
        "It uses 30 structured fictional logs and 5 attack signatures to demonstrate pure algorithmic concepts without third-party black-box libraries.' "
        "Professors love this maturity and honesty!",
        callout_type="viva"
    )

    # ═══════════════════════════════════════════════════════════════
    # SECTION 2: END-TO-END WORKFLOW & SYSTEM ARCHITECTURE
    # ═══════════════════════════════════════════════════════════════
    add_styled_heading("2. End-to-End Workflow & System Architecture", level=1)
    add_p(
        "The system follows a clean 4-tier separation of concerns: Presentation (React), REST API (FastAPI), "
        "Algorithm Engine (Pure Python without ML or external graph libraries), and Persistence (MongoDB)."
    )

    add_styled_heading("2.1 High-Level Architecture Diagram", level=2)
    arch_diagram = """
+-----------------------------------------------------------------------------------+
|                            REACT 18 + VITE FRONTEND                               |
|  [Dashboard KPI]   [Logs Table]   [Alerts + Heap Toggle]   [Force Graph 2D]       |
|  [BFS/DFS Controls] [Choke Points] [Kill Chain Visualizer] [Greedy Monitoring UI] |
+------------------------------------------+----------------------------------------+
                                           | Axios REST Calls (HTTP :8000)
                                           v
+-----------------------------------------------------------------------------------+
|                               FASTAPI BACKEND ROUTERS                             |
|  /api/logs           /api/alerts          /api/graph          /api/analysis       |
+------------------------------------------+----------------------------------------+
                                           | Coordinates Services
                                           v
+-----------------------------------------------------------------------------------+
|                               CORE SERVICES LAYER                                 |
|  log_processor.py -> alert_generator.py -> alert_correlator.py -> kill_chain.py   |
|  monitoring_optimizer.py                                                          |
+------------------------------------------+----------------------------------------+
                                           | Calls Pure DSA (No DB dependencies)
                                           v
+-----------------------------------------------------------------------------------+
|                               PURE DSA ALGORITHMS                                 |
|  1. hash_lookup.py         (Hash Table O(1) lookups)                              |
|  2. kmp.py                 (Knuth-Morris-Pratt O(n+m) Pattern Matching)           |
|  3. graph.py               (Adjacency List Directed Network Graph O(V+E))         |
|  4. bfs_dfs.py             (BFS Shortest Path & DFS Path Enumeration O(V+E))      |
|  5. articulation_points.py (Tarjan's Cut-Vertex & Bridge Detection O(V+E))        |
|  6. priority_queue.py      (Binary Max-Heap with tie-breaking O(log n))           |
|  7. monitoring_optimizer   (Greedy Set Cover Approximation O(P*N^2))              |
+------------------------------------------+----------------------------------------+
                                           | Motor Async Driver
                                           v
+-----------------------------------------------------------------------------------+
|                               MONGODB DATABASE                                    |
|  Collections: logs | alerts | correlations | kill_chains | graph_cache            |
+-----------------------------------------------------------------------------------+
"""
    add_code_block(arch_diagram)

    add_styled_heading("2.2 The 11-Step Data Pipeline Walkthrough", level=2)
    pipeline_steps = [
        ("Step 1", "CSV Ingestion", "Parses raw CSV string, validates required headers ('log_id', 'timestamp', 'source_ip', 'destination_ip', 'event_type', 'severity'), strips whitespace, normalizes casing."),
        ("Step 2", "Hash Lookup", "Passes each log event through 'SignatureHashTable' to assign categories ('Malware', 'Authentication') and compute 'is_suspicious' flag in O(1) time."),
        ("Step 3", "KMP Search", "Extracts ordered event types into a temporal list; runs KMP with LPS array for each of the 5 attack signatures, recording start indices and involved log IDs in O(n+m) time."),
        ("Step 4", "Alert Creation", "Converts KMP matches into formal Alert entities with generated IDs ('A001', 'A002'), timestamps, source/destination IPs, and severity."),
        ("Step 5", "Deduplication", "Filters raw alerts through a Python hash set storing tuple keys: (source_ip, attack_type, tuple(sorted(related_log_ids))) to prevent duplicate alerts on re-runs."),
        ("Step 6", "Correlation", "Groups alerts by 'source_ip' using an adjacency-style hash map. Computes max campaign severity and extracts unified attack sequences in O(n) time."),
        ("Step 7", "Kill Chain", "Maps correlated sequences onto the 7 MITRE stages. Flags stages as 'Observed' or 'Not Observed' to preserve forensic integrity without fabricating data."),
        ("Step 8", "Graph Build", "Constructs an 'AttackGraph' using an Adjacency List. Nodes represent IPs with auto-classified roles (workstation, server, database); edges represent traffic."),
        ("Step 9", "BFS / DFS", "BFS explores level-by-level blast radius and calculates unweighted shortest paths to databases. DFS explores all simple paths and records discovery/finish times."),
        ("Step 10", "Tarjan's AP", "Undirected DFS pass calculates disc[u] and low[u] arrays to discover articulation points (graph choke points) and bridge edges in O(V+E) time."),
        ("Step 11", "Greedy Cover", "Extracts all alert paths and iteratively selects the node covering the maximum number of uncovered paths (Greedy Set Cover) to output monitoring recommendations."),
        ("Step 12", "Heap Triage", "Pushes all alerts into a Max-Heap priority queue. Enables analysts to fetch triage queues where CRITICAL alerts always appear before HIGH, MEDIUM, or LOW in O(log n) time."),
    ]
    add_styled_table(["Step", "Pipeline Stage", "Technical Operation & DSA Role"], pipeline_steps, [Inches(0.9), Inches(1.5), Inches(4.1)])

    # ═══════════════════════════════════════════════════════════════
    # SECTION 3: IN-DEPTH DSA CONCEPTS (CORE DEFENSE SECTION)
    # ═══════════════════════════════════════════════════════════════
    add_styled_heading("3. In-Depth DSA Concepts & Implementation Walkthrough", level=1)
    add_p(
        "This is the most critical section for your college viva. College reviewers will test whether you truly "
        "understand the algorithms and wrote them yourselves, or if you simply used libraries. "
        "Every single algorithm below is implemented from scratch in pure Python inside the 'algorithms/' directory."
    )

    # --- CONCEPT 1 ---
    add_styled_heading("Concept 1: Hashing & Hash Table Lookups", level=2)
    add_p("File Location: backend/algorithms/hash_lookup.py | Class: SignatureHashTable", bold_prefix="Source Code: ")
    add_p(
        "In a security system, every incoming event needs instant classification (e.g., Is 'LOGIN_FAILED' an authentication event or normal traffic?). "
        "If we stored signatures in a flat list, querying each event would take O(k) time where k is the number of signatures. "
        "With millions of logs, this would degrade performance. We implemented a dedicated SignatureHashTable utilizing Python dictionaries "
        "(which use an internal hash table with open addressing and perturbation)."
    )
    add_p("Data Structure Details:", bold_prefix="Implementation: ")
    add_bullet("event_to_category: Maps event string -> dict with category and base_severity. E.g., 'LOGIN_FAILED' -> {'category': 'Authentication', 'base_severity': 'LOW'}.")
    add_bullet("attack_type_to_sig: Maps attack type -> full signature object. E.g., 'BRUTE_FORCE' -> SIG001.")
    add_bullet("sig_id_to_sig: Maps signature ID string -> full signature object. E.g., 'SIG001' -> SIG001.")
    add_bullet("is_suspicious_event(event_type): Evaluates if category != 'Normal' in O(1) time.")
    add_p("Complexity Analysis:", bold_prefix="Mathematical Complexity: ")
    add_bullet("Average Query Time: O(1) — Direct bucket calculation via hash(key) % capacity.")
    add_bullet("Worst-case Query Time: O(n) — Occurs only if all keys collide into the same bucket (extremely rare in Python's high-entropy hash function).")
    add_bullet("Build Time: O(n) — One-time cost during application startup lifespan.")
    add_bullet("Space Complexity: O(n) — Linear space proportional to the number of known signatures and event types.")

    add_callout(
        "Viva Trap Question: How does Python resolve hash collisions?",
        "Answer: 'Python dictionaries resolve collisions using Open Addressing with Quadratic Probing and a perturbation technique. "
        "When two keys hash to the same bucket, Python uses a pseudo-random probe sequence based on the high-order bits of the hash code "
        "to search for the next available slot, guaranteeing minimal clustering.'",
        callout_type="viva"
    )

    # --- CONCEPT 2 ---
    add_styled_heading("Concept 2: KMP (Knuth-Morris-Pratt) Pattern Matching", level=2)
    add_p("File Location: backend/algorithms/kmp.py | Functions: build_lps(), kmp_search()", bold_prefix="Source Code: ")
    add_p(
        "Cyberattacks are rarely single isolated events; they manifest as temporal sequences of events. "
        "For example, a Brute Force attack is a sequence of ['LOGIN_FAILED', 'LOGIN_FAILED', 'LOGIN_FAILED']. "
        "A naive brute-force search compares the pattern against the text with two nested loops, taking O(n * m) time in the worst case "
        "because it repeatedly backtracks the text pointer upon encountering a mismatch."
    )
    add_p(
        "KMP eliminates all backtracking in the text stream by precomputing the Longest Proper Prefix which is also a Suffix (LPS array). "
        "When a mismatch occurs at pattern index j, the text index i is never decremented; instead, j falls back to lps[j - 1], "
        "reusing the partial matches already established."
    )
    add_p("1. LPS Array Preprocessing (build_lps):", bold_prefix="Step-by-Step Mechanics: ")
    add_p(
        "Given pattern P of length m, lps[i] stores the length of the longest proper prefix of P[0...i] that is also a suffix of P[0...i].\n"
        "For example, for pattern ['LOGIN_FAILED', 'LOGIN_FAILED', 'LOGIN_FAILED']:\n"
        "- lps[0] = 0 (by definition, proper prefix cannot equal entire string)\n"
        "- lps[1] = 1 (prefix ['LOGIN_FAILED'] matches suffix ['LOGIN_FAILED'])\n"
        "- lps[2] = 2 (prefix of length 2 matches suffix of length 2)\n"
        "Resulting LPS = [0, 1, 2]. Built in O(m) time."
    )
    add_p("2. Search Execution (kmp_search):", bold_prefix="Step-by-Step Mechanics: ")
    add_p(
        "We slide the pattern over the log stream text of length n. If text[i] == pattern[j], increment both i and j. "
        "If j reaches m, a full attack pattern is recorded at index (i - j), and j resets to lps[j - 1]. "
        "If a mismatch occurs and j > 0, j updates to lps[j - 1] without changing i; if j == 0, simply increment i."
    )
    add_p("Complexity Analysis:", bold_prefix="Mathematical Complexity: ")
    add_bullet("LPS Construction: O(m) time, O(m) space.")
    add_bullet("Pattern Search: O(n) time, O(1) additional space.")
    add_bullet("Total KMP Time Complexity: O(n + m) — Strictly linear!")
    add_bullet("Space Complexity: O(m) — For storing the LPS array.")

    add_callout(
        "Viva Trap Question: Why use KMP instead of Rabin-Karp or Boyer-Moore?",
        "Answer: 'Rabin-Karp computes rolling numerical hashes on character substrings. However, in log analysis, our elements are discrete categorical tokens (event strings) rather than characters. Computing polynomial rolling hashes over token arrays introduces unnecessary hash collision overhead. Boyer-Moore works best on large character alphabets where bad-character skips are large. KMP is mathematically clean, guarantees O(n+m) worst-case time, and adapts naturally to arrays of discrete event strings without backtracking.'",
        callout_type="viva"
    )

    # --- CONCEPT 3 ---
    add_styled_heading("Concept 3: Graph Representation — Adjacency List", level=2)
    add_p("File Location: backend/algorithms/graph.py | Class: AttackGraph", bold_prefix="Source Code: ")
    add_p(
        "To understand how an adversary traverses an enterprise infrastructure, we model hosts as graph vertices (V) "
        "and communication/attack hops as directed graph edges (E). "
        "We specifically selected an Adjacency List representation implemented via Python dictionaries."
    )
    add_p("Why Adjacency List over Adjacency Matrix?", bold_prefix="Design Choice: ")
    add_bullet("Adjacency Matrix requires O(V²) space regardless of how many connections exist. In enterprise networks, most workstations never connect directly to one another; the network graph is highly sparse (E << V²).")
    add_bullet("Adjacency List requires only O(V + E) space, which saves enormous memory.")
    add_bullet("Iterating over neighbours of a node in an adjacency list takes O(degree(v)) time, whereas an adjacency matrix requires scanning all V columns (O(V)) every single time.")
    add_p("Graph Implementation Details:", bold_prefix="Implementation: ")
    add_p(
        "The graph is stored as: self.adj_list = { node_id: { 'label': ..., 'node_type': ..., 'neighbours': [ {'to': ..., 'edge_type': ..., 'event_type': ...} ] } }. "
        "When logs are imported, AttackGraph.from_logs() dynamically extracts unique IPs and directed hops, automatically classifying nodes "
        "into 'workstation', 'server', 'database', or 'external' based on network topology heuristics."
    )

    # --- CONCEPT 4 ---
    add_styled_heading("Concept 4: Breadth-First Search (BFS) & Shortest Path", level=2)
    add_p("File Location: backend/algorithms/bfs_dfs.py | Functions: bfs(), bfs_shortest_path()", bold_prefix="Source Code: ")
    add_p(
        "BFS explores the graph layer by layer using a FIFO Queue (implemented with collections.deque for O(1) pops). "
        "In cybersecurity operations, BFS answers two vital tactical questions:\n"
        "1. Blast Radius: How many machines could the attacker reach in 1 hop? 2 hops? 3 hops?\n"
        "2. Fewest Hops to Target: What is the shortest path from the attacker's initial breach point to critical internal assets (e.g., the primary database)?"
    )
    add_p("Mechanics & Parent Backtracking:", bold_prefix="Algorithm: ")
    add_p(
        "We push the start node into the deque and record its level as 0. For each node popped from the front, we examine all its neighbours. "
        "If a neighbour has not been visited, we mark it visited, record level[neighbor] = level[current] + 1, set parent[neighbor] = current, and append it to the queue. "
        "To reconstruct the shortest path to a destination node, we trace parent pointers backwards from destination to source, then reverse the list."
    )
    add_bullet("Time Complexity: O(V + E) — Each vertex is enqueued once, and each edge is examined once.")
    add_bullet("Space Complexity: O(V) — For the visited set, queue, parent map, and level map.")

    # --- CONCEPT 5 ---
    add_styled_heading("Concept 5: Depth-First Search (DFS) & Path Enumeration", level=2)
    add_p("File Location: backend/algorithms/bfs_dfs.py | Functions: dfs(), find_all_paths()", bold_prefix="Source Code: ")
    add_p(
        "DFS explores each branch to its maximum depth before backtracking, utilizing recursion (implicit LIFO stack). "
        "In our engine, DFS serves two primary purposes:\n"
        "1. Exhaustive Path Enumeration (find_all_paths): Discovers all possible alternative attack routes an adversary could take between two hosts. This feeds directly into our Greedy Monitoring Optimizer.\n"
        "2. Discovery & Finish Time Tracking: Records when each node was first visited (discovery_time) and when its entire subtree was fully explored (finish_time). This structure provides the theoretical foundation for Tarjan's algorithm."
    )
    add_bullet("Time Complexity: O(V + E) for standard traversal; bounded DFS for all simple paths.")
    add_bullet("Space Complexity: O(V) for the recursion call stack and visited tracking.")

    # --- CONCEPT 6 ---
    add_styled_heading("Concept 6: Articulation Points & Bridges (Tarjan's Algorithm)", level=2)
    add_p("File Location: backend/algorithms/articulation_points.py | Function: find_articulation_points_and_bridges()", bold_prefix="Source Code: ")
    add_p(
        "Definition: An Articulation Point (or Cut Vertex) is a node whose removal disconnects the graph into two or more components. "
        "A Bridge is an edge whose removal disconnects the graph.\n"
        "In our security model, articulation points represent potential choke points—critical network junction nodes through which "
        "all traffic between separate network zones must pass. Isolating or heavily monitoring these nodes maximizes containment efficiency."
    )
    add_p("Why Tarjan's Algorithm?", bold_prefix="Algorithmic Advantage: ")
    add_p(
        "A naive approach to find cut vertices would test every node: remove node v, run BFS/DFS to check connectivity (O(V+E)), and repeat for all V vertices. "
        "This naive algorithm takes O(V * (V + E)), which is O(V² + V*E) — far too slow. "
        "Tarjan's algorithm finds ALL articulation points and bridges in a SINGLE DFS pass in O(V + E) time."
    )
    add_p("The Two Critical Arrays:", bold_prefix="Core Tarjan Mechanics: ")
    add_bullet("disc[u]: The Discovery Time when node u is first reached during DFS.")
    add_bullet("low[u]: The lowest discovery time reachable from u's subtree via at most one back-edge (ancestor connection).")
    add_p("The Two Articulation Point Rules:", bold_prefix="Evaluation Conditions: ")
    add_bullet("Rule 1 (Root Node): If u is the root of the DFS tree, it is an articulation point if and only if it has two or more distinct children in the DFS tree.")
    add_bullet("Rule 2 (Non-Root Node): If u is NOT the root, it is an articulation point if it has a child v such that: low[v] >= disc[u]. This mathematically proves that child v has no back-edge to any ancestor of u; removing u completely traps v and its subtree!")
    add_bullet("Bridge Rule: Edge (u, v) is a bridge if: low[v] > disc[u].")
    add_bullet("Time Complexity: O(V + E) — Single undirected DFS traversal.")
    add_bullet("Space Complexity: O(V) — For disc, low, parent, and visited arrays.")

    # --- CONCEPT 7 ---
    add_styled_heading("Concept 7: Greedy Approximation for Set Cover (Monitoring Optimizer)", level=2)
    add_p("File Location: backend/services/monitoring_optimizer.py | Function: greedy_monitoring_optimizer()", bold_prefix="Source Code: ")
    add_p(
        "The Problem: An enterprise has limited security monitoring sensors. Given multiple discovered attack paths across the network graph, "
        "which minimum subset of nodes should we equip with sensors so that EVERY attack path passes through at least one monitored node?"
    )
    add_p("Why Greedy? (The NP-Hard Reality):", bold_prefix="Theoretical Justification: ")
    add_p(
        "This problem is formally identical to the classical Set Cover Problem, one of Karp's 21 NP-complete problems. "
        "Finding the guaranteed absolute minimum subset of nodes is NP-hard; an exact algorithm (such as Integer Linear Programming or brute force) "
        "has an exponential worst-case time complexity O(2^N). "
        "Therefore, we implemented Chvátal's Greedy Approximation Algorithm."
    )
    add_p("Greedy Algorithm Mechanics:", bold_prefix="Step-by-Step Logic: ")
    add_bullet("1. Extract all attack paths from alerts using bounded DFS (get_attack_paths).")
    add_bullet("2. Build an inverted index: node -> set of path indices containing that node.")
    add_bullet("3. Loop while uncovered paths remain: Select the node that covers the largest number of currently UNCOVERED paths.")
    add_bullet("4. Add that node to monitoring_nodes and mark its associated paths as covered.")
    add_bullet("5. Record step-by-step reasoning in coverage_steps for full transparency on the frontend.")
    add_p("Complexity & Approximation Bound:", bold_prefix="Mathematical Complexity: ")
    add_bullet("Time Complexity: O(P * N²) in the worst case (where P is total attack paths, N is candidate nodes).")
    add_bullet("Space Complexity: O(P + N).")
    add_bullet("Approximation Ratio: Guaranteed within ln(n) + 1 of the optimal minimum set.")

    add_callout(
        "Academic Honesty Disclaimer on Set Cover",
        "Notice that our code and UI explicitly state: 'Greedy approximation: this is NOT guaranteed to be the mathematically optimal minimum set. "
        "It is an approximation algorithm for the NP-hard Set Cover problem.' Highlighting this theoretical limitation proves to reviewers that "
        "you understand computational complexity theory!",
        callout_type="tip"
    )

    # --- CONCEPT 8 ---
    add_styled_heading("Concept 8: Priority Queue via Binary Max-Heap", level=2)
    add_p("File Location: backend/algorithms/priority_queue.py | Class: AlertPriorityQueue", bold_prefix="Source Code: ")
    add_p(
        "In a live SOC, hundreds of alerts trigger simultaneously. Analysts cannot waste time reading low-severity noise when a ransomware "
        "lateral movement or data exfiltration is underway. Alerts must be processed strictly by urgency: CRITICAL (4) -> HIGH (3) -> MEDIUM (2) -> LOW (1)."
    )
    add_p("Why a Heap over Array Sorting?", bold_prefix="Algorithmic Advantage: ")
    add_p(
        "Sorting an array with QuickSort/MergeSort takes O(n log n) and is a static, one-time operation. "
        "In contrast, security alerts arrive dynamically as a continuous stream. "
        "A Priority Queue supports dynamic real-time operations:\n"
        "- Inserting a newly triggered alert: O(log n)\n"
        "- Peeking at the highest-urgency alert: O(1)\n"
        "- Removing and handling the most critical alert: O(log n)"
    )
    add_p("Max-Heap Implementation & The Min-Heap Negation Trick:", bold_prefix="Custom Implementation: ")
    add_p(
        "We built a custom binary heap stored in a Python list where for any index i:\n"
        "- Parent index = (i - 1) // 2\n"
        "- Left child index = 2*i + 1\n"
        "- Right child index = 2*i + 2\n"
        "Because Python's internal tuple comparison evaluates smaller values first, we negate the priority value: entry = (-priority, counter, alert). "
        "Thus, a CRITICAL alert (priority 4) becomes -4, bubbles to the top, and behaves as a true Max-Heap!"
    )
    add_p("Tie-Breaking Counter Mechanism:", bold_prefix="Stability & FIFO: ")
    add_p(
        "If two alerts have identical priority (e.g., both are CRITICAL), comparing alert dictionaries would cause a TypeError. "
        "We store a monotonic integer counter in the middle of the tuple: (-priority, counter, alert). "
        "This acts as an automatic tie-breaker, guaranteeing that equal-priority alerts are served in FIFO (First-In, First-Out) arrival order!"
    )
    add_bullet("push(alert): Appends to heap and executes _bubble_up() in O(log n) time.")
    add_bullet("pop(): Swaps root with last element, pops last, and executes _heapify_down() in O(log n) time.")
    add_bullet("peek(): Returns root element at index 0 in O(1) time.")
    add_bullet("get_all_sorted(): Drains and restores heap in O(n log n) time.")

    # ═══════════════════════════════════════════════════════════════
    # SECTION 4: PIPELINE INTEGRATION & BUSINESS LOGIC
    # ═══════════════════════════════════════════════════════════════
    add_styled_heading("4. Pipeline Integration & Business Logic Services", level=1)
    add_p("While the 'algorithms/' directory contains pure DSA code, the 'services/' directory coordinates the business logic:")

    add_styled_heading("4.1 CSV Parsing & Validation (services/log_processor.py)", level=2)
    add_p(
        "Robust input validation is critical. parse_logs_from_csv() enforces that incoming files include all required columns: "
        "'log_id', 'timestamp', 'source_ip', 'destination_ip', 'event_type', 'description', and 'severity'. "
        "Any row missing fields or containing malformed values is flagged with informative error strings rather than crashing the server."
    )

    add_styled_heading("4.2 Alert Deduplication via Hash Sets (services/alert_generator.py)", level=2)
    add_p(
        "A common flaw in security engines is alert explosion—running pattern detection multiple times creates duplicate alerts. "
        "deduplicate_alerts() constructs a composite hashable tuple for each alert: (source_ip, attack_type, tuple(sorted(related_log_ids))). "
        "By checking this tuple against a Python set (O(1) lookup), duplicate alerts are pruned in linear O(n) time."
    )

    add_styled_heading("4.3 Campaign Correlation (services/alert_correlator.py)", level=2)
    add_p(
        "A single attack sequence produces multiple alerts across time. correlate_alerts() groups individual alerts by source_ip "
        "using a hash map. Each correlation object synthesizes the campaign's overall timeline, maximum severity, unique attack types, "
        "and ordered alert IDs."
    )

    add_styled_heading("4.4 MITRE ATT&CK Kill Chain Builder (services/kill_chain_builder.py)", level=2)
    add_p(
        "The Lockheed Martin / MITRE ATT&CK Kill Chain defines the standard lifecycle of an advanced cyberattack: "
        "Reconnaissance -> Initial Access -> Execution -> Privilege Escalation -> Lateral Movement -> Collection -> Exfiltration.\n"
        "Our engine inspects the correlated alerts and maps each alert to its corresponding stage. "
        "Crucially, our system maintains forensic integrity: if a stage has no supporting log evidence, it is strictly marked as 'Not Observed'. "
        "We never fabricate intermediate stages."
    )

    # ═══════════════════════════════════════════════════════════════
    # SECTION 5: MASTER TIME & SPACE COMPLEXITY TABLE
    # ═══════════════════════════════════════════════════════════════
    add_styled_heading("5. Master Complexity & Algorithm Reference Table", level=1)
    add_p("This table summarizes the theoretical time and space complexities of every algorithm across the engine. Reviewers frequently inspect this first:")

    complexity_data = [
        ("Hash Table Build", "algorithms/hash_lookup.py", "O(n)", "O(n)", "O(n)", "One-time loading of signatures into Python dict"),
        ("Hash Table Lookup", "algorithms/hash_lookup.py", "O(1) avg", "O(n) worst", "O(1)", "Direct hash-bucket index computation"),
        ("KMP LPS Build", "algorithms/kmp.py", "O(m)", "O(m)", "O(m)", "Prefix-suffix failure array construction"),
        ("KMP Search", "algorithms/kmp.py", "O(n)", "O(n)", "O(1)", "Sliding window without text backtracking"),
        ("KMP Total", "algorithms/kmp.py", "O(n + m)", "O(n + m)", "O(m)", "Optimal linear string/token pattern matching"),
        ("Graph Build", "algorithms/graph.py", "O(V + E)", "O(V + E)", "O(V + E)", "Constructing Adjacency List from parsed logs"),
        ("BFS Traversal", "algorithms/bfs_dfs.py", "O(V + E)", "O(V + E)", "O(V)", "FIFO queue layer-by-layer exploration"),
        ("BFS Shortest Path", "algorithms/bfs_dfs.py", "O(V + E)", "O(V + E)", "O(V)", "Parent pointer backtracking to source"),
        ("DFS Traversal", "algorithms/bfs_dfs.py", "O(V + E)", "O(V + E)", "O(V)", "Recursive stack depth exploration with timers"),
        ("Tarjan's AP & Bridges", "algorithms/articulation_points.py", "O(V + E)", "O(V + E)", "O(V)", "Single DFS pass tracking disc[] and low[]"),
        ("Greedy Set Cover", "services/monitoring_optimizer.py", "O(P * N²)", "O(P * N²)", "O(P + N)", "Iteratively picking node covering max paths"),
        ("Max-Heap Push", "algorithms/priority_queue.py", "O(log n)", "O(log n)", "O(1)", "Bubble-up restoring max-heap property"),
        ("Max-Heap Pop", "algorithms/priority_queue.py", "O(log n)", "O(log n)", "O(1)", "Heapify-down from root to leaves"),
        ("Max-Heap Peek", "algorithms/priority_queue.py", "O(1)", "O(1)", "O(1)", "Immediate access to root at array index 0"),
        ("Alert Correlation", "services/alert_correlator.py", "O(n)", "O(n)", "O(n)", "Hash map grouping by source IP"),
    ]
    add_styled_table(
        ["Operation / Algorithm", "Source File", "Avg Time", "Worst Time", "Space", "Algorithmic Justification"],
        complexity_data,
        [Inches(1.3), Inches(1.3), Inches(0.7), Inches(0.7), Inches(0.6), Inches(1.9)]
    )
    add_p("Notation: n = number of logs/alerts, m = pattern length, V = vertices (hosts), E = edges (hops), P = attack paths, N = candidate nodes.", italic_prefix="Complexity Variables: ")

    # ═══════════════════════════════════════════════════════════════
    # SECTION 6: FULL-STACK ARCHITECTURE & API REFERENCE
    # ═══════════════════════════════════════════════════════════════
    add_styled_heading("6. Full-Stack Architecture & API Reference", level=1)

    add_styled_heading("6.1 Backend API Endpoints (FastAPI)", level=2)
    api_endpoints = [
        ("POST", "/api/logs/import-sample", "Loads the 30 built-in sample logs, executes the full pipeline, and populates MongoDB."),
        ("POST", "/api/logs/import", "Accepts multipart CSV upload, validates headers and rows, and executes pipeline."),
        ("GET", "/api/logs/", "Returns all stored log entries with category and suspicious flags."),
        ("GET", "/api/alerts/", "Returns all generated alerts with related log IDs and severity."),
        ("GET", "/api/alerts/priority", "Returns alerts sorted strictly via the Binary Max-Heap priority queue."),
        ("GET", "/api/graph/", "Returns graph topology (nodes + links), BFS/DFS results, and choke points."),
        ("GET", "/api/graph/choke-points", "Returns Tarjan's cut vertices and bridge edges with discovery/low values."),
        ("GET", "/api/analysis/kill-chains", "Returns reconstructed MITRE kill chains with Observed vs Not Observed stages."),
        ("GET", "/api/analysis/correlations", "Returns alert groups correlated by source IP with severity summaries."),
        ("GET", "/api/analysis/monitoring", "Returns Greedy Set Cover sensor allocation recommendations and steps."),
        ("GET", "/api/analysis/dashboard", "Returns aggregated stats (total logs, alerts, kill chains, choke points) in one call."),
    ]
    add_styled_table(["Method", "Endpoint Route", "Purpose & Execution"], api_endpoints, [Inches(0.8), Inches(2.2), Inches(3.5)])

    add_styled_heading("6.2 Frontend Architecture (React 18 + Vite + Tailwind CSS)", level=2)
    add_p(
        "The frontend is structured into modular page views and reusable components:\n"
        "- Dashboard.jsx: High-level KPI cards, threat activity trends, critical alerts feed, and kill chain overview.\n"
        "- Logs.jsx: Interactive log table with search, severity filtering, and CSV upload / sample loading controls.\n"
        "- Alerts.jsx: Alert management featuring the interactive 'Sort by Priority (Max-Heap)' toggle.\n"
        "- AttackGraph.jsx: Interactive canvas powered by react-force-graph-2d, displaying node relationships, BFS shortest path routes, and raw Adjacency List JSON inspection.\n"
        "- KillChain.jsx: Visual timeline progression through the 7 MITRE stages.\n"
        "- Monitoring.jsx: Step-by-step breakdown of the Greedy Set Cover sensor allocation algorithm.\n"
        "- Settings.jsx: System status and MongoDB connection diagnostics."
    )

    # ═══════════════════════════════════════════════════════════════
    # SECTION 7: STEP-BY-STEP LIVE DEMO PRESENTATION SCRIPT
    # ═══════════════════════════════════════════════════════════════
    add_styled_heading("7. Step-by-Step Live Demo Presentation Script", level=1)
    add_p("Follow this exact sequence during the live presentation. Have your terminal windows ready before the reviewers arrive!")

    add_styled_heading("7.1 Pre-Demo Setup Checklist", level=2)
    add_bullet("Terminal 1 (MongoDB): Verify MongoDB is running (mongod or MongoDB Compass on localhost:27017).")
    add_bullet("Terminal 2 (Backend): cd backend && uvicorn main:app --reload (Runs on http://localhost:8000).")
    add_bullet("Terminal 3 (Frontend): cd frontend && npm run dev (Runs on http://localhost:5173).")
    add_bullet("Browser: Open http://localhost:5173 on your browser, and open http://localhost:8000/docs in a second tab.")

    add_styled_heading("7.2 Screen-by-Screen Presenter Walkthrough", level=2)

    demo_steps = [
        ("Screen 1: Dashboard Overview",
         "1. Point to the top KPI cards: Total Logs (30), Active Alerts, Correlated Attacks, and Choke Points.\n"
         "2. Explain: 'The dashboard gives the security analyst immediate situational awareness.'\n"
         "3. Highlight that when the backend started, it automatically verified MongoDB and loaded sample logs."),

        ("Screen 2: Security Logs (Logs tab)",
         "1. Navigate to 'Logs' from the sidebar.\n"
         "2. Point out the raw tabular format: Log ID, Timestamp, Source IP, Destination IP, Event Type, and Severity.\n"
         "3. Explain: 'Here you see our CSV parsing and O(1) hash categorization in action. Notice how each event has a category like Authentication or Malware, and suspicious events are automatically highlighted.'"),

        ("Screen 3: Alerts & Priority Queue (Alerts tab)",
         "1. Navigate to 'Alerts'.\n"
         "2. Point out the alerts detected by KMP pattern matching (Brute Force, Malware Execution, Lateral Movement).\n"
         "3. CLICK the 'Sort by Priority (Max-Heap)' button!\n"
         "4. Explain: 'When I click this button, the backend invokes our custom binary max-heap. Notice that CRITICAL severity alerts immediately jump to the top, followed by HIGH, MEDIUM, and LOW. In an enterprise SOC, this ensures analysts never miss high-impact intrusions.'"),

        ("Screen 4: Attack Graph & Topology (Attack Graph tab)",
         "1. Navigate to 'Attack Graph'.\n"
         "2. Show the force-directed graph. Drag a node to show interactivity.\n"
         "3. Switch to the 'Adjacency List' view tab on screen.\n"
         "4. Explain: 'Behind this visual graph is our pure Python AttackGraph adjacency list. We chose an adjacency list over an adjacency matrix because enterprise networks are sparse (O(V+E) vs O(V²)).'\n"
         "5. Run the Shortest Path query: Select Attacker IP (192.168.1.10) to Database (192.168.1.100). Show the BFS result: 'BFS found the fewest hops through the compromised workstation and server.'"),

        ("Screen 5: Choke Points / Articulation Points",
         "1. On the Attack Graph screen, toggle 'Highlight Choke Points'.\n"
         "2. Point out the highlighted nodes and bridge edges.\n"
         "3. Explain: 'These critical nodes were calculated by Tarjan's Articulation Point algorithm in O(V+E) time by tracking discovery and low times during DFS. If an attacker's lateral movement must cross these cut vertices, isolating them mathematically severs the attack path.'"),

        ("Screen 6: Kill Chain Reconstruction (Kill Chain tab)",
         "1. Navigate to 'Kill Chain'.\n"
         "2. Select a correlated attack campaign (e.g., CORR001).\n"
         "3. Show the 7 MITRE stages. Point out which are 'Observed' (green) and which are 'Not Observed' (gray).\n"
         "4. Explain: 'This demonstrates our commitment to forensic honesty. A real SOC analyst needs to know what evidence was actually detected. We never invent unobserved stages.'"),

        ("Screen 7: Monitoring Optimization (Monitoring tab)",
         "1. Navigate to 'Monitoring'.\n"
         "2. Show the recommended monitoring nodes and the step-by-step greedy coverage log.\n"
         "3. Explain: 'Because placing sensors everywhere is expensive and Set Cover is NP-Hard, our engine uses a Greedy Set Cover approximation. In Step 1, it picks the node covering the most attack paths, then repeats until 100% of attack paths are covered.'"),
    ]

    for title, script in demo_steps:
        add_p(script, bold_prefix=f"{title}:\n")

    # ═══════════════════════════════════════════════════════════════
    # SECTION 8: 30 COMPREHENSIVE REVIEWER VIVA Q&A
    # ═══════════════════════════════════════════════════════════════
    add_styled_heading("8. Comprehensive Reviewer Viva Q&A (30 Questions & Winning Answers)", level=1)
    add_p(
        "College reviewers test different aspects of a project: some focus purely on theoretical DSA, some on system architecture, "
        "and some on practical edge cases. Below are 30 predicted questions with exact, technically rigorous answers."
    )

    viva_qa = [
        # --- Group A: Core DSA & Algorithms ---
        ("Q1: What is the exact time complexity of KMP and why is it faster than naive search?",
         "Naive pattern matching has a worst-case time complexity of O(n * m) because whenever a mismatch occurs after partial matching, the text pointer i is reset backwards. KMP runs in O(n + m) time because it precomputes the LPS (Longest Proper Prefix-Suffix) array in O(m) time. When a mismatch occurs at index j, the text pointer i never moves backwards; instead, j falls back to lps[j-1], skipping comparisons that are guaranteed to match."),

        ("Q2: Walk me through the mathematical definition of an Articulation Point.",
         "An articulation point (cut vertex) in an undirected graph G is a vertex v such that G - {v} has strictly more connected components than G. In Tarjan's algorithm, during a DFS tree traversal: (1) The root is an articulation point if it has >= 2 children in the DFS tree. (2) Any other vertex u is an articulation point if it has a child v such that low[v] >= disc[u], meaning v has no back-edge to any ancestor of u."),

        ("Q3: What do the disc[] and low[] arrays represent in Tarjan's algorithm?",
         "disc[u] is the discovery time (timestamp) of vertex u when first visited in DFS. low[u] is the minimum discovery time reachable from the subtree rooted at u, utilizing tree edges and at most one back-edge. If low[v] >= disc[u], it proves that subtree v cannot reach any node visited before u without passing through u."),

        ("Q4: Why did you choose an Adjacency List instead of an Adjacency Matrix?",
         "Network graphs are sparse graphs where the number of edges E is typically much less than V². An adjacency matrix requires O(V²) space and O(V) time to find the neighbours of a node, wasting memory and computation on non-existent connections. An adjacency list requires only O(V + E) space and queries neighbours in O(degree(v)) time."),

        ("Q5: What is the Set Cover problem and why did you use a Greedy approach?",
         "Set Cover is one of Karp's 21 NP-complete problems: given a universe of elements (attack paths) and a collection of subsets (nodes covering paths), find the minimum sub-collection that covers all elements. An exact solution requires exponential time O(2^N). We use Chvátal's Greedy approximation, which iteratively chooses the set covering the maximum uncovered elements, achieving a proven ln(n) approximation ratio in polynomial time O(P * N²)."),

        ("Q6: How does your Max-Heap handle two alerts with identical severity?",
         "We store elements as tuples: (-priority, counter, alert). When two alerts have identical priority (e.g., both CRITICAL with priority 4, stored as -4), Python compares the second tuple element. Because counter is a monotonically increasing integer assigned on insertion, the earlier alert has a smaller counter value and is popped first. This prevents TypeErrors from comparing alert dictionaries and guarantees FIFO stability."),

        ("Q7: How does a Binary Heap perform insertions and deletions in O(log n)?",
         "A binary heap is a complete binary tree of height floor(log2(n)). On push, the element is appended to the end of the array and bubble_up() swaps it with its parent ((i-1)//2) at most log(n) times. On pop, the root is replaced with the last element and heapify_down() swaps it with its higher-priority child (2i+1 or 2i+2) at most log(n) times."),

        ("Q8: Why is Hash Table lookup average O(1) but worst-case O(n)?",
         "Average lookup is O(1) because a well-distributed hash function computes hash(key) % capacity and jumps directly to the memory bucket. In the worst case, if multiple keys produce identical bucket indices (hash collisions), open addressing probes sequential slots, degrading to O(n) linear search. Python minimizes this with high-entropy hashing and dynamic resizing when load factor exceeds 2/3."),

        ("Q9: What is the difference between BFS and DFS in this specific application?",
         "BFS explores level-by-level using a FIFO queue, guaranteeing the shortest path (fewest hops) from an attacker to critical assets in unweighted graphs. DFS explores deeply along branches using recursion/stack, making it ideal for enumerating all possible lateral movement attack paths and tracking discovery/finish timestamps for cycle and cut-vertex detection."),

        ("Q10: Can KMP handle overlapping pattern matches in the event stream?",
         "Yes! In our kmp_search() implementation, when a complete match is found at j == m, we record the start index (i - j) and immediately set j = lps[j - 1] rather than resetting j to 0. This allows the algorithm to detect overlapping patterns seamlessly without re-scanning."),

        # --- Group B: Design & Architecture ---
        ("Q11: What is the 4-tier architecture of your project?",
         "Presentation Layer: React 18 SPA with Tailwind CSS and React Force Graph 2D. API Layer: FastAPI with async route handlers and automatic Swagger documentation. Business/DSA Layer: Pure Python algorithms without external graph or ML dependencies. Persistence Layer: MongoDB with Motor async driver."),

        ("Q12: Why did you keep the DSA algorithms in pure Python without NetworkX or heapq?",
         "Because this is a DSA college capstone project! The objective is to demonstrate deep mastery of core computer science fundamentals. Using black-box libraries like NetworkX or heapq obscures the algorithmic mechanics. Writing our own adjacency list, BFS/DFS, Tarjan's algorithm, and binary heap from scratch proves we understand the underlying logic."),

        ("Q13: How does the system deduplicate alerts?",
         "In services/alert_generator.py, deduplicate_alerts() maps each alert to a hashable tuple: (source_ip, attack_type, tuple(sorted(related_log_ids))). It checks this composite key against a Python hash set in O(1) time, ensuring multiple pipeline runs never generate redundant alerts."),

        ("Q14: How does alert correlation work in your engine?",
         "In services/alert_correlator.py, correlate_alerts() groups individual alerts by source_ip using a hash table in O(n) time. It aggregates all related log IDs, extracts unique attack types in chronological sequence, and identifies the maximum severity across the attack campaign."),

        ("Q15: What is the MITRE ATT&CK Kill Chain and how do you reconstruct it?",
         "The MITRE ATT&CK / Lockheed Martin Kill Chain models the sequential phases of an advanced intrusion: Reconnaissance -> Initial Access -> Execution -> Privilege Escalation -> Lateral Movement -> Collection -> Exfiltration. Our engine maps each alert's kill_chain_stage to this order, marking present stages as 'Observed' and missing stages as 'Not Observed'."),

        ("Q16: Why don't you invent missing kill chain stages?",
         "Forensic integrity. In real cybersecurity investigations, fabricating unobserved attack steps creates dangerous false positives and misleads analysts. Our engine is academically and professionally honest: it only reports stages directly substantiated by log evidence."),

        ("Q17: Why did you use FastAPI instead of Flask or Django?",
         "FastAPI is modern, asynchronous (ASGI), and significantly faster than Flask (WSGI). It natively supports Python type hints, automatic Pydantic data validation, and generates interactive OpenAPI / Swagger documentation at /docs out of the box."),

        ("Q18: Why did you choose MongoDB over a relational SQL database like PostgreSQL?",
         "Security logs, graph links, and kill chains have dynamic, semi-structured schemas. A single alert contains nested arrays of related log IDs, and kill chains contain lists of stages. MongoDB's BSON document model stores these nested structures natively without complex multi-table relational joins."),

        ("Q19: How does the backend communicate with MongoDB asynchronously?",
         "We use Motor, the official asynchronous MongoDB driver for Tornado and asyncio, integrated with FastAPI's async/await event loop. This ensures database queries never block the server thread."),

        ("Q20: What is CORS and why did you configure it in FastAPI?",
         "Cross-Origin Resource Sharing (CORS) is a browser security mechanism that restricts scripts from making HTTP requests to a different origin. Because Vite runs on http://localhost:5173 and FastAPI runs on http://localhost:8000, we added CORSMiddleware to allow requests between these two distinct local ports."),

        # --- Group C: Edge Cases & Resilience ---
        ("Q21: What happens if the CSV log file is completely empty or missing columns?",
         "In services/log_processor.py, parse_logs_from_csv() verifies the presence of REQUIRED_COLUMNS. If the header is missing, columns are absent, or the file is empty, it safely returns an empty list and a descriptive error message without crashing."),

        ("Q22: What happens if an unknown event type or invalid severity appears in a log?",
         "Unknown event types return a fallback dictionary {'category': 'Unknown', 'base_severity': 'LOW'} from SignatureHashTable in O(1) time. Invalid severities default safely to 'LOW' and log a validation warning."),

        ("Q23: What happens if the attack graph is disconnected (multiple components)?",
         "Our Tarjan's implementation loops over all vertices: 'for node in nodes: if not visited[node]: dfs_ap(node)'. This ensures every disconnected component is visited, correctly finding articulation points and bridges across the entire forest."),

        ("Q24: What happens if there is no path between the attacker and the database?",
         "In bfs_shortest_path(), if the destination node is not present in the parent dictionary after BFS completes, the function safely returns an empty list [] indicating unreachable targets."),

        ("Q25: What happens if the graph contains cycles?",
         "Both BFS and DFS maintain a visited hash set. When an edge leads to an already-visited vertex, the traversal skips re-enqueueing it. In Tarjan's algorithm, back-edges to visited ancestors are used to update low[u] = min(low[u], disc[v]) without re-running DFS."),

        ("Q26: What is the space complexity of your priority queue during peak alert floods?",
         "Linear space O(N), where N is the number of active alerts. The underlying Python list stores exactly N 3-element tuples."),

        ("Q27: How does the greedy monitoring optimizer handle isolated nodes with no attack paths?",
         "get_attack_paths() only extracts paths that exist between alert source and destination nodes. If a node is not part of any attack path, it does not appear in the inverted index and is simply never selected, avoiding wasted monitoring resources."),

        ("Q28: Why is your graph directed for BFS/DFS but undirected for Tarjan's algorithm?",
         "Network traffic and lateral movement are directed (host A attacks host B). However, classical cut-vertex and bridge connectivity analysis mathematically applies to undirected graphs. To evaluate whether isolating a node disconnects communication between network segments, we consider bidirectional physical reachability."),

        ("Q29: How did you test your DSA implementations?",
         "We wrote a comprehensive test suite of 64 unit tests in backend/tests/test_all_dsa.py using pytest. It tests LPS generation, KMP matching, edge cases, graph construction, BFS levels, DFS times, Tarjan's AP, Heap ordering, Greedy coverage, and CSV parsing. All 64 tests pass with 100% success."),

        ("Q30: What are the main limitations and future scope of this project?",
         "Limitations: (1) Rule-based matching on 5 predefined signatures cannot detect novel zero-day attacks. (2) Fictional sample data. (3) Greedy monitoring is an approximation. Future Scope: (1) Real-time streaming log ingestion via Apache Kafka / WebSockets. (2) Graph Neural Networks (GNNs) or Machine Learning for anomaly detection. (3) Integration with MITRE ATT&CK STIX/TAXII threat feeds."),
    ]

    for q, a in viva_qa:
        add_p(a, bold_prefix=f"{q}\n", italic_prefix="Answer: ")

    # ═══════════════════════════════════════════════════════════════
    # SECTION 9: EMERGENCY 5-MINUTE CHEAT SHEET
    # ═══════════════════════════════════════════════════════════════
    doc.add_page_break()
    add_styled_heading("9. Emergency 5-Minute 'Before Entering the Room' Cheat Sheet", level=1)
    add_p("Scan these quick bullet points 5 minutes before your presentation to keep the core facts fresh in your mind:")

    add_bullet("Project Name: Threat-Intelligence Kill-Chain Correlation Engine (DSA-3 Capstone).")
    add_bullet("Core Purpose: Transforming noisy security logs into correlated attack campaigns and graph choke points using 8 pure DSA algorithms.")
    add_bullet("1. Hashing: O(1) avg lookup in hash_lookup.py. Maps event types to categories.")
    add_bullet("2. KMP: O(n+m) linear search in kmp.py. Uses LPS array to skip backtracking in temporal event streams.")
    add_bullet("3. Graph: Adjacency list in graph.py. O(V+E) space because network graphs are sparse (E << V²).")
    add_bullet("4. BFS: O(V+E) via deque in bfs_dfs.py. Finds blast radius and fewest hops to database.")
    add_bullet("5. DFS: O(V+E) via recursion in bfs_dfs.py. Enumerates all simple attack paths; tracks discovery/finish times.")
    add_bullet("6. Tarjan's AP: O(V+E) in articulation_points.py. Finds cut vertices (low[v] >= disc[u]) and bridges (low[v] > disc[u]) in a single pass.")
    add_bullet("7. Greedy Set Cover: O(P*N²) in monitoring_optimizer.py. Approximates NP-hard sensor placement within ln(n)+1 ratio.")
    add_bullet("8. Max-Heap: O(log n) push/pop, O(1) peek in priority_queue.py. Negates priority (-4 for CRITICAL) with tie-breaker counter.")
    add_bullet("Kill Chain: 7 MITRE stages. Strictly marks unobserved stages as 'Not Observed' to preserve forensic integrity.")
    add_bullet("Tests: 64 unit tests in pytest, all passing 100%.")
    add_bullet("Commands to run: 'mongod' (DB), 'uvicorn main:app --reload' (Backend :8000), 'npm run dev' (Frontend :5173).")

    # Save Document
    output_path = os.path.join(os.path.dirname(__file__), "Threat_Kill_Chain_Engine_Presentation_Guide.docx")
    doc.save(output_path)
    print(f"Document successfully created at: {output_path}")

if __name__ == "__main__":
    create_document()
