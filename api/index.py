import io
import os
from flask import Flask, render_template, request, send_file
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Template path resolution for local and serverless execution
base_dir = os.path.dirname(os.path.abspath(__file__))
if os.path.exists(os.path.join(base_dir, "templates")):
    template_dir = os.path.join(base_dir, "templates")
else:
    template_dir = os.path.abspath(os.path.join(base_dir, "..", "templates"))

app = Flask(__name__, template_folder=template_dir)

# Theme Palette matching reference style
NAVY       = RGBColor(15, 23, 42)
BLUE       = RGBColor(2, 132, 199)
TEAL       = RGBColor(13, 148, 136)
GREEN      = RGBColor(16, 185, 129)
GRAY       = RGBColor(100, 116, 139)
LIGHT_CARD = RGBColor(248, 250, 252)
BORDER     = RGBColor(226, 232, 240)
WHITE      = RGBColor(255, 255, 255)
AMBER_BG   = RGBColor(254, 243, 199)
AMBER_TXT  = RGBColor(180, 83, 9)
DARK_BG    = RGBColor(7, 14, 30)

CLR_30M   = RGBColor(16, 185, 129)
CLR_1H    = RGBColor(245, 158, 11)
CLR_2H    = RGBColor(249, 115, 22)
CLR_3H    = RGBColor(239, 68, 68)
CLR_OVER3 = RGBColor(112, 14, 14)

def add_header(slide, section, title, subtitle=None, tag=None):
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.73), Inches(1.2))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p1 = tf.paragraphs[0]
    p1.text = section.upper()
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = BLUE
    p2 = tf.add_paragraph()
    p2.text = title
    p2.font.size = Pt(22)
    p2.font.bold = True
    p2.font.color.rgb = NAVY
    if subtitle:
        p3 = tf.add_paragraph()
        p3.text = subtitle
        p3.font.size = Pt(12)
        p3.font.color.rgb = GRAY
    if tag:
        tag_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.5), Inches(0.45), Inches(4.0), Inches(0.42))
        tag_box.fill.solid(); tag_box.fill.fore_color.rgb = LIGHT_CARD; tag_box.line.color.rgb = BORDER
        p_t = tag_box.text_frame.paragraphs[0]; p_t.text = tag; p_t.alignment = PP_ALIGN.CENTER; p_t.font.size = Pt(10.5); p_t.font.bold = True; p_t.font.color.rgb = BLUE

def generate_presentation(payload):
    team_name = payload.get("teamName", "Team 1")
    period = payload.get("reportPeriod", "Week 42 (Current Shift)")
    critical_incidents = str(payload.get("criticalIncidents", "0"))
    platform = payload.get("platform", {})
    roles = payload.get("roles", {})

    # Default Team 1 Staff Roster from reference file
    default_staff = [
        {"name": "LONG TOKA",            "m30": 26, "h1": 2, "h2": 1, "h3": 0, "over3": 0, "total": 29},
        {"name": "KORRY MUHAMMADSOFEK",  "m30": 21, "h1": 2, "h2": 1, "h3": 0, "over3": 0, "total": 24},
        {"name": "OUM SENGKRUY",         "m30": 11, "h1": 1, "h2": 1, "h3": 0, "over3": 0, "total": 13},
        {"name": "POV PUNLORK",          "m30": 6,  "h1": 1, "h2": 0, "h3": 0, "over3": 0, "total": 7},
        {"name": "LENG CHANNLY",         "m30": 5,  "h1": 1, "h2": 0, "h3": 0, "over3": 0, "total": 6},
        {"name": "I CHHENGKHEANG",       "m30": 4,  "h1": 0, "h2": 1, "h3": 0, "over3": 1, "total": 6},
        {"name": "NHEAN PISETH",         "m30": 4,  "h1": 0, "h2": 0, "h3": 0, "over3": 0, "total": 4},
        {"name": "CHUM SOVANNVISAL",     "m30": 2,  "h1": 0, "h2": 0, "h3": 0, "over3": 0, "total": 2},
        {"name": "MAK BUN LEANG",        "m30": 1,  "h1": 0, "h2": 0, "h3": 0, "over3": 1, "total": 2},
    ]

    raw_staff = payload.get("staff", [])
    processed_staff = []
    t30 = t1 = t2 = t3 = tover = 0

    if raw_staff and len(raw_staff) > 0:
        for s in raw_staff:
            m30 = int(s.get("m30", 0)); h1 = int(s.get("h1", 0)); h2 = int(s.get("h2", 0)); h3 = int(s.get("h3", 0)); over3 = int(s.get("over3", 0))
            tot = s.get("total")
            tot = int(tot) if tot is not None else (m30 + h1 + h2 + h3 + over3)
            t30 += m30; t1 += h1; t2 += h2; t3 += h3; tover += over3
            processed_staff.append({"name": s.get("name", "Staff"), "m30": m30, "h1": h1, "h2": h2, "h3": h3, "over3": over3, "total": tot})
    else:
        for s in default_staff:
            t30 += s["m30"]; t1 += s["h1"]; t2 += s["h2"]; t3 += s["h3"]; tover += s["over3"]
            processed_staff.append(dict(s))

    processed_staff.sort(key=lambda x: x["total"], reverse=True)
    grand_total = sum(s["total"] for s in processed_staff)
    sla_rate = f"{(t30 / grand_total * 100):.1f}%" if grand_total > 0 else "0.0%"

    # Dynamic role extraction with fallback to Team 1 structure
    main_super = roles.get("mainSuper") or roles.get("main_super") or "MAK BUN LEANG"
    backup_super = roles.get("backupSuper") or roles.get("backup_super") or "POV PUNLORK"
    hq01 = roles.get("hq01") or "MAK BUN LEANG"
    hq02 = roles.get("hq02") or "POV PUNLORK"
    hq03 = roles.get("hq03") or "CHUM SOVANNVISAL"
    hq04 = roles.get("hq04") or "NHEAN P. / OUM S. / I CHHENGKHEANG"
    mon_lead = roles.get("monitoringLead") or roles.get("monitoring_lead") or roles.get("monitoring") or "LONG TOKA"
    mon_backup = roles.get("monitoringBackup") or roles.get("monitoring_backup") or "KORRY MUHAMMADSOFEK"
    amis_rep = roles.get("amisReport") or roles.get("amis_report") or "NHEAN P. / OUM S. / I CHHENGKHEANG"
    int_queue = roles.get("internalQueue") or roles.get("internal_queue") or roles.get("internal") or "LENG CHANNLY / Team Roster"

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # ==========================================
    # SLIDE 1: Cover Slide
    # ==========================================
    s1 = prs.slides.add_slide(blank)
    tb = s1.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.3), Inches(3.0))
    tf = tb.text_frame
    p1 = tf.paragraphs[0]; p1.text = "OPERATIONS & NETWORK SUPPORT"; p1.font.size = Pt(13); p1.font.bold = True; p1.font.color.rgb = BLUE
    p2 = tf.add_paragraph(); p2.text = "Weekly Report & Teamwork Roles"; p2.font.size = Pt(34); p2.font.bold = True; p2.font.color.rgb = NAVY
    p3 = tf.add_paragraph(); p3.text = f"Incident overview, system health, bandwidth alerts, and {team_name} operational assignments"; p3.font.size = Pt(15); p3.font.color.rgb = GRAY
    p4 = tf.add_paragraph(); p4.text = f"{team_name}  |  {period}"; p4.font.size = Pt(12); p4.font.bold = True; p4.font.color.rgb = BLUE

    # 3 Summary chips
    c_data = [("ASSIGNED TEAM", "Team Alpha"), ("TOTAL RESOLVED", f"{grand_total} Cases"), ("ROSTER COVERAGE", f"{len(processed_staff)} Active Engineers")]
    for i, (lbl, val) in enumerate(c_data):
        c_box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0 + i * 3.8), Inches(5.4), Inches(3.5), Inches(1.1))
        c_box.fill.solid(); c_box.fill.fore_color.rgb = LIGHT_CARD; c_box.line.color.rgb = BORDER
        tf_c = c_box.text_frame
        p_c1 = tf_c.paragraphs[0]; p_c1.text = lbl; p_c1.font.size = Pt(10); p_c1.font.bold = True; p_c1.font.color.rgb = BLUE
        p_c2 = tf_c.add_paragraph(); p_c2.text = val; p_c2.font.size = Pt(14); p_c2.font.bold = True; p_c2.font.color.rgb = NAVY

    # ==========================================
    # SLIDE 2: KPI Overview
    # ==========================================
    s2 = prs.slides.add_slide(blank)
    add_header(s2, "SECTION I • Incident Metrics", "Weekly Case Summary Overview", "Key operational performance metrics across customer and internal support queues", "Team Alpha Operations")
    kpis = [
        (str(grand_total), "Total Cases Handled", "Weekly operational volume resolved.", BLUE),
        (critical_incidents, "Critical Incidents", "Zero service-impacting severity.", GREEN),
        (str(len(processed_staff)), "Active Engineers", "Full team dispatch coverage.", BLUE)
    ]
    for i, (val, title, sub, val_clr) in enumerate(kpis):
        box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 4.0), Inches(1.85), Inches(3.7), Inches(3.7))
        box.fill.solid(); box.fill.fore_color.rgb = LIGHT_CARD; box.line.color.rgb = BORDER
        tf = box.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = val; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(46); p.font.bold = True; p.font.color.rgb = val_clr
        p = tf.add_paragraph(); p.text = title; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
        p = tf.add_paragraph(); p.text = sub; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(12); p.font.color.rgb = GRAY

    ft_s2 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.85), Inches(11.73), Inches(0.95))
    ft_s2.fill.solid(); ft_s2.fill.fore_color.rgb = LIGHT_CARD; ft_s2.line.color.rgb = BORDER
    tf_ft2 = ft_s2.text_frame; tf_ft2.word_wrap = True
    top1 = processed_staff[0]["name"] if len(processed_staff) > 0 else "Staff"
    top1_c = processed_staff[0]["total"] if len(processed_staff) > 0 else 0
    top2 = processed_staff[1]["name"] if len(processed_staff) > 1 else "Staff"
    top2_c = processed_staff[1]["total"] if len(processed_staff) > 1 else 0
    top_share = f"{((top1_c + top2_c) / grand_total * 100):.0f}%" if grand_total > 0 else "0%"
    p_ft2 = tf_ft2.paragraphs[0]
    p_ft2.text = f"• {sla_rate} of total volume ({t30}/{grand_total} cases) resolved in under 30 minutes.\n• Top frontline contributors: {top1} ({top1_c}) and {top2} ({top2_c}) handled {top_share} of total queue."
    p_ft2.font.size = Pt(11); p_ft2.font.color.rgb = NAVY

    # ==========================================
    # SLIDE 3: Staff KPI Support Cases (Ranked List)
    # ==========================================
    s3 = prs.slides.add_slide(blank)
    add_header(s3, "SECTION I • Individual Throughput", "Staff KPI Support Cases", f"Total Handled: {grand_total} Cases", f"{team_name} Support Roster")
    max_cases = max([s["total"] for s in processed_staff] + [1])
    y_start = Inches(1.75)
    row_h = Inches(0.52)
    row_gap = Inches(0.08)

    for i, s in enumerate(processed_staff[:9]):
        cur_y = y_start + i * (row_h + row_gap)
        row_card = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), cur_y, Inches(11.73), row_h)
        row_card.fill.solid(); row_card.fill.fore_color.rgb = LIGHT_CARD; row_card.line.color.rgb = BORDER

        # Rank badge
        rank_tb = s3.shapes.add_textbox(Inches(0.9), cur_y + Inches(0.06), Inches(0.6), Inches(0.4))
        p_rk = rank_tb.text_frame.paragraphs[0]; p_rk.text = f"#{i+1}"; p_rk.font.size = Pt(12); p_rk.font.bold = True; p_rk.font.color.rgb = BLUE

        # Name
        tb_n = s3.shapes.add_textbox(Inches(1.6), cur_y + Inches(0.06), Inches(3.2), Inches(0.4))
        p_n = tb_n.text_frame.paragraphs[0]; p_n.text = s["name"]; p_n.font.size = Pt(12); p_n.font.bold = True; p_n.font.color.rgb = NAVY

        # Scaled Bar
        bar_max_w = 4.8
        bar_w = Inches(max(0.2, (s["total"] / max_cases) * bar_max_w))
        bar = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.0), cur_y + Inches(0.13), bar_w, Inches(0.24))
        bar.fill.solid(); bar.fill.fore_color.rgb = BLUE; bar.line.fill.background()

        # Count
        tb_c = s3.shapes.add_textbox(Inches(10.2), cur_y + Inches(0.06), Inches(2.1), Inches(0.4))
        p_c = tb_c.text_frame.paragraphs[0]; p_c.text = f"{s['total']} cases"; p_c.font.size = Pt(12); p_c.font.bold = True; p_c.font.color.rgb = BLUE; p_c.alignment = PP_ALIGN.RIGHT

    # ==========================================
    # SLIDE 4: Resolution Time Breakdown (Chart)
    # ==========================================
    s4 = prs.slides.add_slide(blank)
    add_header(s4, "SECTION I • Resolution Turnaround Time", "Resolution Time Breakdown by Engineer", None)
    lg_x = Inches(8.3)
    for lbl, clr in [("< 30mn", CLR_30M), ("1H", CLR_1H), ("2H", CLR_2H), ("3H", CLR_3H), ("Over 3H", CLR_OVER3)]:
        sq = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, lg_x, Inches(1.15), Inches(0.16), Inches(0.16))
        sq.fill.solid(); sq.fill.fore_color.rgb = clr; sq.line.fill.background()
        tb_lg = s4.shapes.add_textbox(lg_x + Inches(0.2), Inches(1.02), Inches(0.8), Inches(0.35))
        p_lg = tb_lg.text_frame.paragraphs[0]; p_lg.text = lbl; p_lg.font.size = Pt(9.5); p_lg.font.bold = True; p_lg.font.color.rgb = GRAY
        lg_x += Inches(0.95)

    chart_bg = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.6), Inches(11.73), Inches(4.55))
    chart_bg.fill.solid(); chart_bg.fill.fore_color.rgb = WHITE; chart_bg.line.color.rgb = BORDER

    scale = Inches(7.5) / max_cases
    row_step = Inches(4.2 / max(len(processed_staff[:9]), 1))
    for idx, s in enumerate(processed_staff[:9]):
        curr_y = Inches(1.75) + idx * row_step
        tb_name = s4.shapes.add_textbox(Inches(0.9), curr_y - Inches(0.08), Inches(2.4), Inches(0.35))
        p_name = tb_name.text_frame.paragraphs[0]; p_name.text = s["name"]; p_name.font.size = Pt(9.5); p_name.font.bold = True; p_name.font.color.rgb = NAVY; p_name.alignment = PP_ALIGN.RIGHT
        curr_x = Inches(3.4)
        for count, clr in [(s["m30"], CLR_30M), (s["h1"], CLR_1H), (s["h2"], CLR_2H), (s["h3"], CLR_3H), (s["over3"], CLR_OVER3)]:
            if count > 0:
                seg_w = count * scale
                seg = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, curr_x, curr_y, seg_w, Inches(0.20))
                seg.fill.solid(); seg.fill.fore_color.rgb = clr; seg.line.fill.background()
                curr_x += seg_w
        tb_tot = s4.shapes.add_textbox(curr_x + Inches(0.12), curr_y - Inches(0.08), Inches(0.6), Inches(0.35))
        p_tot = tb_tot.text_frame.paragraphs[0]; p_tot.text = str(s["total"]); p_tot.font.size = Pt(10); p_tot.font.bold = True; p_tot.font.color.rgb = NAVY

    ft_box = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.32), Inches(11.73), Inches(0.68))
    ft_box.fill.solid(); ft_box.fill.fore_color.rgb = LIGHT_CARD; ft_box.line.color.rgb = BORDER
    p_ft = ft_box.text_frame.paragraphs[0]
    p_ft.text = f"Fast Turnaround: {t30} cases closed <30m. Complex escalations (>3h) concentrated under Lead {main_super} & I CHHENGKHEANG.   {sla_rate} SLA Sub-30m"
    p_ft.font.size = Pt(10.5); p_ft.font.color.rgb = GRAY; p_ft.alignment = PP_ALIGN.CENTER

    # ==========================================
    # SLIDE 5: Staff Knowledge & Ability
    # ==========================================
    s5 = prs.slides.add_slide(blank)
    add_header(s5, "SECTION II • Continuous Training", "Staff Knowledge & Ability", "Core operational capabilities and continuous technical upskilling roadmap", "Scheduled Sessions Active")
    box1 = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.5))
    box1.fill.solid(); box1.fill.fore_color.rgb = LIGHT_CARD; box1.line.color.rgb = BORDER
    tf1 = box1.text_frame; tf1.word_wrap = True
    p = tf1.paragraphs[0]; p.text = "Support Competencies"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf1.add_paragraph(); p.text = "\n• Customer Support: End-to-end client issue resolution.\n\n• Internal Support: Cross-functional technical assistance."; p.font.size = Pt(13); p.font.color.rgb = NAVY

    box2 = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.5))
    box2.fill.solid(); box2.fill.fore_color.rgb = LIGHT_CARD; box2.line.color.rgb = BORDER
    tf2 = box2.text_frame; tf2.word_wrap = True
    p = tf2.paragraphs[0]; p.text = "Upcoming Internal Training"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf2.add_paragraph(); p.text = "\n• Core Network: Internal infrastructure & topology.\n\n• Monitoring Tools: Alerting thresholds & telemetry.\n\n• Products & Services: Solution portfolio updates."; p.font.size = Pt(13); p.font.color.rgb = NAVY

    # ==========================================
    # SLIDE 6: Tool Supports & System Health
    # ==========================================
    s6 = prs.slides.add_slide(blank)
    add_header(s6, "SECTION III • Infrastructure Health", "Tool Supports & System Health", "Operational health status across internal tools, billing pipelines, and support suites", "Weekly Uptime: 100%")
    tools = [
        ("Portal Support", platform.get("portal", "Helpful & Responsive"), "100% query resolution.", "STABLE"),
        ("AMIS Portal", platform.get("amis", "No Issue"), "Zero system downtime.", "OPTIMAL"),
        ("Billing System", platform.get("billing", "No Issue"), "Normal billing processing.", "OPTIMAL"),
        ("RT Ticketing", platform.get("rt", "No Issue"), "Continuous dispatch flow.", "OPTIMAL")
    ]
    for i, (name, status, sub, badge_txt) in enumerate(tools):
        box = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 2.95), Inches(2.0), Inches(2.8), Inches(4.1))
        box.fill.solid(); box.fill.fore_color.rgb = LIGHT_CARD; box.line.color.rgb = BORDER
        tf = box.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = name; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
        p = tf.add_paragraph(); p.text = "\n" + status; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = TEAL
        p = tf.add_paragraph(); p.text = sub; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(11); p.font.color.rgb = GRAY
        badge = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.35 + i * 2.95), Inches(5.4), Inches(1.7), Inches(0.42))
        badge.fill.solid(); badge.fill.fore_color.rgb = WHITE; badge.line.color.rgb = BORDER
        p_b = badge.text_frame.paragraphs[0]; p_b.text = badge_txt; p_b.alignment = PP_ALIGN.CENTER; p_b.font.size = Pt(10.5); p_b.font.bold = True; p_b.font.color.rgb = BLUE

    ft_s6 = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.32), Inches(11.73), Inches(0.68))
    ft_s6.fill.solid(); ft_s6.fill.fore_color.rgb = LIGHT_CARD; ft_s6.line.color.rgb = BORDER
    p_ft6 = ft_s6.text_frame.paragraphs[0]
    p_ft6.text = "All operational tools maintained 100% availability with zero ticket synchronization lags."
    p_ft6.font.size = Pt(11); p_ft6.font.color.rgb = GRAY; p_ft6.alignment = PP_ALIGN.CENTER

    # ==========================================
    # SLIDE 7: Upstream & IPLC: High Capacity Rules
    # ==========================================
    s7 = prs.slides.add_slide(blank)
    add_header(s7, "SECTION IV • Bandwidth Monitoring", "Upstream & IPLC: High Capacity Rules", "Committed bandwidth trigger rules for high-capacity circuits (≥ 5 Gbps)")
    bw_high = [
        ["Carrier / Circuit", "Link Type", "Alert Trigger Threshold", "Trigger Logic"],
        ["IPLC Viettel", "IPLC Circuit", "Traffic over 20 Gbps", "Alert Over Committed"],
        ["VNPT", "Global Upstream", "Traffic over 20 Gbps", "Alert Over Committed"],
        ["PCCW", "Global Upstream", "Traffic over 10 Gbps", "Alert Over Committed"],
        ["Cogent", "Global Upstream", "Traffic over 10 Gbps", "Alert Over Committed"],
        ["Viettel-BVT / CHT", "Global Upstream", "Traffic over 5 Gbps each", "Alert Over Committed"]
    ]
    tbl_shape = s7.shapes.add_table(6, 4, Inches(0.8), Inches(1.9), Inches(11.73), Inches(4.3))
    tbl = tbl_shape.table
    for i, w in enumerate([Inches(3.2), Inches(2.5), Inches(3.0), Inches(3.03)]): tbl.columns[i].width = w
    for r_idx, row in enumerate(bw_high):
        for c_idx, val in enumerate(row):
            c = tbl.cell(r_idx, c_idx); c.text = val; p = c.text_frame.paragraphs[0]; p.font.size = Pt(11)
            if r_idx == 0:
                c.fill.solid(); c.fill.fore_color.rgb = BLUE
                p.font.bold = True; p.font.color.rgb = WHITE
            else:
                p.font.color.rgb = NAVY

    ft_s7 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.4), Inches(11.73), Inches(0.6))
    ft_s7.fill.solid(); ft_s7.fill.fore_color.rgb = LIGHT_CARD; ft_s7.line.color.rgb = BORDER
    p_ft7 = ft_s7.text_frame.paragraphs[0]
    p_ft7.text = "Automated Telegram & email telemetry alerts trigger immediately when thresholds are exceeded.   5 Main High-Capacity Links Active"
    p_ft7.font.size = Pt(10.5); p_ft7.font.color.rgb = GRAY; p_ft7.alignment = PP_ALIGN.CENTER

    # ==========================================
    # SLIDE 8: Upstream & IPLC: Links Under 5 Gbps
    # ==========================================
    s8 = prs.slides.add_slide(blank)
    add_header(s8, "SECTION IV • Bandwidth Alert Rules", "Upstream & IPLC: Links Under 5 Gbps", "Committed bandwidth alert thresholds for secondary uplinks and regional transit")
    sub_links = [
        ("Telia (Arelion)", "> 3 Gbps"),
        ("NTT", "> 2 Gbps"),
        ("CU (China Unicom)", "> 400 Mbps"),
        ("CT (China Telecom)", "> 300 Mbps"),
        ("CMI (China Mobile)", "> 50 Mbps")
    ]
    for i, (carrier, threshold) in enumerate(sub_links):
        y_pos = Inches(1.85 + i * 0.88)
        card = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), y_pos, Inches(11.73), Inches(0.72))
        card.fill.solid(); card.fill.fore_color.rgb = LIGHT_CARD; card.line.color.rgb = BORDER
        tb_car = s8.shapes.add_textbox(Inches(1.2), y_pos + Inches(0.15), Inches(4.5), Inches(0.4))
        p_car = tb_car.text_frame.paragraphs[0]; p_car.text = carrier; p_car.font.size = Pt(13); p_car.font.bold = True; p_car.font.color.rgb = NAVY
        tb_th = s8.shapes.add_textbox(Inches(6.8), y_pos + Inches(0.15), Inches(2.2), Inches(0.4))
        p_th = tb_th.text_frame.paragraphs[0]; p_th.text = threshold; p_th.font.size = Pt(13); p_th.font.bold = True; p_th.font.color.rgb = BLUE
        badge = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.8), y_pos + Inches(0.14), Inches(2.4), Inches(0.42))
        badge.fill.solid(); badge.fill.fore_color.rgb = AMBER_BG; badge.line.fill.background()
        p_b = badge.text_frame.paragraphs[0]; p_b.text = "Alert over committed"; p_b.alignment = PP_ALIGN.CENTER; p_b.font.size = Pt(10.5); p_b.font.bold = True; p_b.font.color.rgb = AMBER_TXT

    ft_s8 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.4), Inches(11.73), Inches(0.6))
    ft_s8.fill.solid(); ft_s8.fill.fore_color.rgb = LIGHT_CARD; ft_s8.line.color.rgb = BORDER
    p_ft8 = ft_s8.text_frame.paragraphs[0]
    p_ft8.text = "Bandwidth alarms prompt immediate traffic rerouting to avoid overage surcharges.   5 Regional Links Monitored"
    p_ft8.font.size = Pt(10.5); p_ft8.font.color.rgb = GRAY; p_ft8.alignment = PP_ALIGN.CENTER

    # ==========================================
    # SLIDE 9: Team 1: Senior Team Leader & Supervision
    # ==========================================
    s9 = prs.slides.add_slide(blank)
    add_header(s9, "SECTION V • Leadership & Escalation", f"{team_name}: Senior Team Leader & Supervision", "Escalation path and operational supervision command structure", "24/7 Escalation Active")
    lead1 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.5))
    lead1.fill.solid(); lead1.fill.fore_color.rgb = LIGHT_CARD; lead1.line.color.rgb = BORDER
    tf_l1 = lead1.text_frame; tf_l1.word_wrap = True
    p = tf_l1.paragraphs[0]; p.text = "PRIMARY LEAD"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = BLUE
    p = tf_l1.add_paragraph(); p.text = "Main Supervisor"; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = GRAY
    p = tf_l1.add_paragraph(); p.text = main_super; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf_l1.add_paragraph(); p.text = "\nPrimary operational lead for shift supervision, queue balance, and primary escalations."; p.font.size = Pt(12); p.font.color.rgb = GRAY

    lead2 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.5))
    lead2.fill.solid(); lead2.fill.fore_color.rgb = LIGHT_CARD; lead2.line.color.rgb = BORDER
    tf_l2 = lead2.text_frame; tf_l2.word_wrap = True
    p = tf_l2.paragraphs[0]; p.text = "STANDBY LEAD"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = BLUE
    p = tf_l2.add_paragraph(); p.text = "Backup Supervisor"; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = GRAY
    p = tf_l2.add_paragraph(); p.text = backup_super; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf_l2.add_paragraph(); p.text = "\nStandby supervisor to manage queue escalations during supervisor absence or high ticket load."; p.font.size = Pt(12); p.font.color.rgb = GRAY

    # ==========================================
    # SLIDE 10: Team 1: Support Role Stations & Daily Reports
    # ==========================================
    s10 = prs.slides.add_slide(blank)
    add_header(s10, "SECTION V • Dispatch Stations & Roster", f"{team_name}: Support Role Stations & Daily Reports", "Workstation assignments and daily status reporting distribution", "Daily Rotation Active")
    st1 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.5))
    st1.fill.solid(); st1.fill.fore_color.rgb = LIGHT_CARD; st1.line.color.rgb = BORDER
    tf_st1 = st1.text_frame; tf_st1.word_wrap = True
    p = tf_st1.paragraphs[0]; p.text = "External Stations (HQ)"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf_st1.add_paragraph(); p.text = "Direct client-facing response\n"; p.font.size = Pt(11); p.font.color.rgb = GRAY
    p = tf_st1.add_paragraph()
    p.text = f"HQ01: {hq01}\nHQ02: {hq02}\nHQ03: {hq03}\nHQ04: {hq04}\n\nNote: Station HQ04 rotated daily across designated engineers."
    p.font.size = Pt(12.5); p.font.color.rgb = NAVY

    st2 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.5))
    st2.fill.solid(); st2.fill.fore_color.rgb = LIGHT_CARD; st2.line.color.rgb = BORDER
    tf_st2 = st2.text_frame; tf_st2.word_wrap = True
    p = tf_st2.paragraphs[0]; p.text = "Internal Support & Reports"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf_st2.add_paragraph(); p.text = "Shift monitoring & AMIS handling\n"; p.font.size = Pt(11); p.font.color.rgb = GRAY
    p = tf_st2.add_paragraph()
    p.text = f"Monitoring Lead: {mon_lead}\nMonitoring Backup: {mon_backup}\nAMIS Reports: {amis_rep}\nInternal Queue: {int_queue}"
    p.font.size = Pt(12.5); p.font.color.rgb = NAVY

    # ==========================================
    # SLIDE 11: Action Plan & Strategic Checkpoints
    # ==========================================
    s11 = prs.slides.add_slide(blank)
    add_header(s11, "EXECUTIVE SUMMARY • Next Steps", "Action Plan & Strategic Checkpoints", "Key operational achievements and immediate priorities for upcoming cycle")
    actions = [
        ("Zero Critical Incidents", f"Successfully resolved {grand_total} cases with 100% tool availability across AMIS, Billing, and RT suites."),
        ("Active Alert Rules", "Committed threshold rules strictly enforced across all 11 Upstream and IPLC connections."),
        ("Continuous Upskilling", "Scheduled internal training rollouts for core network topology, alert telemetry, and service updates.")
    ]
    for i, (title, desc) in enumerate(actions):
        box = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 4.0), Inches(2.0), Inches(3.7), Inches(4.1))
        box.fill.solid(); box.fill.fore_color.rgb = LIGHT_CARD; box.line.color.rgb = BORDER
        tf = box.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = title; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
        p = tf.add_paragraph(); p.text = "\n" + desc; p.font.size = Pt(12); p.font.color.rgb = GRAY

    ft_s11 = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.32), Inches(11.73), Inches(0.68))
    ft_s11.fill.solid(); ft_s11.fill.fore_color.rgb = LIGHT_CARD; ft_s11.line.color.rgb = BORDER
    p_ft11 = ft_s11.text_frame.paragraphs[0]
    p_ft11.text = "Weekly Review Cadence: Performance benchmarking completed with 100% team compliance.   Operations Team Alpha"
    p_ft11.font.size = Pt(11); p_ft11.font.color.rgb = GRAY; p_ft11.alignment = PP_ALIGN.CENTER

    # ==========================================
    # SLIDE 12: Questions & Discussion
    # ==========================================
    s12 = prs.slides.add_slide(blank)
    bg = s12.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid(); bg.fill.fore_color.rgb = DARK_BG; bg.line.fill.background()

    tb_q = s12.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(2.0))
    tf_q = tb_q.text_frame; tf_q.word_wrap = True
    p = tf_q.paragraphs[0]; p.text = "Questions & Discussion"; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(38); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf_q.add_paragraph(); p.text = f"Weekly Case Summary Report • Operations & Network Support {team_name}"; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(15); p.font.color.rgb = RGBColor(148, 163, 184)

    # 3 Bottom chips
    q_data = [("SUPERVISOR LEAD", main_super), ("BACKUP LEAD", backup_super), ("ASSIGNED TEAM", f"Team Alpha ({len(processed_staff)} Engineers)")]
    for i, (lbl, val) in enumerate(q_data):
        box_q = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2 + i * 3.8), Inches(5.2), Inches(3.4), Inches(1.1))
        box_q.fill.solid(); box_q.fill.fore_color.rgb = RGBColor(15, 23, 42); box_q.line.color.rgb = RGBColor(30, 41, 59)
        tf_bq = box_q.text_frame
        p_bq1 = tf_bq.paragraphs[0]; p_bq1.text = lbl; p_bq1.alignment = PP_ALIGN.CENTER; p_bq1.font.size = Pt(10); p_bq1.font.bold = True; p_bq1.font.color.rgb = BLUE
        p_bq2 = tf_bq.add_paragraph(); p_bq2.text = val; p_bq2.alignment = PP_ALIGN.CENTER; p_bq2.font.size = Pt(13); p_bq2.font.bold = True; p_bq2.font.color.rgb = WHITE

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    filename = f"Weekly_Report_{team_name.replace(' ', '_')}.pptx"
    return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")

@app.route("/", defaults={"path": ""}, methods=["GET", "POST"])
@app.route("/<path:path>", methods=["GET", "POST"])
def catch_all(path):
    if request.method == "POST" or "export" in path:
        payload = request.get_json(silent=True) or {}
        return generate_presentation(payload)
    return render_template("index.html")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
