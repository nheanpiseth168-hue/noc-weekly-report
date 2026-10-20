$code = @'
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

# Theme Palette matching the reference style
NAVY         = RGBColor(15, 23, 42)
BLUE         = RGBColor(2, 132, 199)
TEAL         = RGBColor(13, 148, 136)
GREEN        = RGBColor(16, 185, 129)
GRAY         = RGBColor(100, 116, 139)
LIGHT_CARD   = RGBColor(248, 250, 252)
BORDER       = RGBColor(226, 232, 240)
WHITE        = RGBColor(255, 255, 255)
AMBER_BG     = RGBColor(254, 243, 199)
AMBER_TXT    = RGBColor(180, 83, 9)
DARK_BG      = RGBColor(10, 17, 40)

def add_header(slide, section, title, subtitle=None):
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.2))
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

def generate_presentation(payload):
    team_name = payload.get("teamName", "Team 2")
    period = payload.get("reportPeriod", "Week 42 (Current Shift)")
    critical_incidents = str(payload.get("criticalIncidents", "0"))
    platform = payload.get("platform", {})
    roles = payload.get("roles", {})

    # Staff Data: Use submitted data or fallback to template defaults
    raw_staff = payload.get("staff", [])
    if raw_staff:
        staff_data = []
        for s in raw_staff:
            tot = s.get("total")
            if tot is None:
                tot = int(s.get("m30", 0)) + int(s.get("h1", 0)) + int(s.get("h2", 0)) + int(s.get("h3", 0)) + int(s.get("over3", 0))
            staff_data.append({"name": s.get("name", "Staff"), "total": int(tot)})
    else:
        staff_data = [
            {"name": "Meu Chanleaksmey", "total": 98},
            {"name": "Ly Manith",        "total": 69},
            {"name": "Phy Reaksmey",      "total": 64},
            {"name": "Mak Bunfei",       "total": 54},
            {"name": "Soth Rasal",       "total": 49},
            {"name": "Ouk Ratha",        "total": 45},
            {"name": "Thy Piseth",       "total": 44},
            {"name": "Oung Savorn",      "total": 27}
        ]

    staff_data.sort(key=lambda x: x["total"], reverse=True)
    grand_total = sum(s["total"] for s in staff_data)
    active_engineers = str(len(staff_data))

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # ==========================================
    # SLIDE 1: Cover Slide
    # ==========================================
    s1 = prs.slides.add_slide(blank)
    tb = s1.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.3), Inches(3.5))
    tf = tb.text_frame
    p1 = tf.paragraphs[0]; p1.text = "OPERATIONS & NETWORK SUPPORT"; p1.font.size = Pt(13); p1.font.bold = True; p1.font.color.rgb = BLUE
    p2 = tf.add_paragraph(); p2.text = "Weekly Report & Teamwork Roles"; p2.font.size = Pt(34); p2.font.bold = True; p2.font.color.rgb = NAVY
    p3 = tf.add_paragraph(); p3.text = f"Incident overview, system health, bandwidth alerts, and {team_name} operational assignments"; p3.font.size = Pt(15); p3.font.color.rgb = GRAY
    p4 = tf.add_paragraph(); p4.text = f"{team_name}  |  {period}"; p4.font.size = Pt(12); p4.font.bold = True; p4.font.color.rgb = BLUE

    # ==========================================
    # SLIDE 2: KPI Overview (3 Cards)
    # ==========================================
    s2 = prs.slides.add_slide(blank)
    add_header(s2, "SECTION I", "Weekly Case Summary Overview", "Key incident metrics across customer and internal queues")
    kpis = [
        (str(grand_total), "Total Cases Handled", "Weekly operational volume", BLUE),
        (critical_incidents, "Critical Incidents", "Zero service-impacting severity", GREEN),
        (active_engineers, "Active Engineers", "Full team dispatch coverage", BLUE)
    ]
    for i, (val, title, sub, val_clr) in enumerate(kpis):
        box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 4.0), Inches(2.0), Inches(3.7), Inches(4.3))
        box.fill.solid(); box.fill.fore_color.rgb = LIGHT_CARD; box.line.color.rgb = BORDER
        tf = box.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = val; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(46); p.font.bold = True; p.font.color.rgb = val_clr
        p = tf.add_paragraph(); p.text = title; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
        p = tf.add_paragraph(); p.text = sub; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(12); p.font.color.rgb = GRAY

    # ==========================================
    # SLIDE 3: Staff KPI Support Cases (Cards & Bars)
    # ==========================================
    s3 = prs.slides.add_slide(blank)
    add_header(s3, "SECTION I", "Staff KPI Support Cases", f"Total support cases handled per engineer (Total: {grand_total})")
    max_cases = max(s["total"] for s in staff_data) if staff_data else 1
    y_start = Inches(1.8)
    row_height = Inches(0.55)
    row_gap = Inches(0.1)

    for i, s in enumerate(staff_data[:8]):
        cur_y = y_start + i * (row_height + row_gap)
        # Background card
        row_card = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), cur_y, Inches(11.73), row_height)
        row_card.fill.solid(); row_card.fill.fore_color.rgb = LIGHT_CARD; row_card.line.color.rgb = BORDER

        # Staff Name
        tb_n = s3.shapes.add_textbox(Inches(1.0), cur_y + Inches(0.08), Inches(3.0), Inches(0.4))
        p_n = tb_n.text_frame.paragraphs[0]; p_n.text = s["name"]; p_n.font.size = Pt(12); p_n.font.bold = True; p_n.font.color.rgb = NAVY

        # Progress Bar
        bar_max_w = Inches(5.6)
        bar_w = max(Inches(0.2), (s["total"] / max_cases) * bar_max_w)
        bar = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.2), cur_y + Inches(0.14), bar_w, Inches(0.26))
        bar.fill.solid(); bar.fill.fore_color.rgb = BLUE; bar.line.fill.background()

        # Case badge text
        tb_c = s3.shapes.add_textbox(Inches(10.0), cur_y + Inches(0.08), Inches(2.2), Inches(0.4))
        p_c = tb_c.text_frame.paragraphs[0]; p_c.text = f"{s['total']} cases"; p_c.font.size = Pt(12); p_c.font.bold = True; p_c.font.color.rgb = BLUE; p_c.alignment = PP_ALIGN.RIGHT

    # ==========================================
    # SLIDE 4: Staff Knowledge & Ability
    # ==========================================
    s4 = prs.slides.add_slide(blank)
    add_header(s4, "SECTION II", "Staff Knowledge & Ability", "Core capabilities and scheduled continuous training programs")

    box1 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.5))
    box1.fill.solid(); box1.fill.fore_color.rgb = LIGHT_CARD; box1.line.color.rgb = BORDER
    tf1 = box1.text_frame; tf1.word_wrap = True
    p = tf1.paragraphs[0]; p.text = "Support Competencies"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf1.add_paragraph(); p.text = "\n• Customer Support: End-to-end client issue resolution.\n\n• Internal Support: Cross-functional technical assistance."; p.font.size = Pt(13); p.font.color.rgb = NAVY

    box2 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.5))
    box2.fill.solid(); box2.fill.fore_color.rgb = LIGHT_CARD; box2.line.color.rgb = BORDER
    tf2 = box2.text_frame; tf2.word_wrap = True
    p = tf2.paragraphs[0]; p.text = "Upcoming Internal Training"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf2.add_paragraph(); p.text = "\n• Core Network: Internal infrastructure & topology.\n\n• Monitoring Tools: Alerting thresholds & telemetry.\n\n• Products & Services: Solution portfolio updates."; p.font.size = Pt(13); p.font.color.rgb = NAVY

    # ==========================================
    # SLIDE 5: Tool Supports & System Health
    # ==========================================
    s5 = prs.slides.add_slide(blank)
    add_header(s5, "SECTION III", "Tool Supports & System Health", "Operational health status across internal tools and support suites")
    tools = [
        ("Portal Support", platform.get("portal", "Helpful & Responsive")),
        ("AMIS", platform.get("amis", "No Issue")),
        ("Billing", platform.get("billing", "No Issue")),
        ("RT", platform.get("rt", "No Issue"))
    ]
    for i, (name, status) in enumerate(tools):
        box = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 2.95), Inches(2.0), Inches(2.8), Inches(4.3))
        box.fill.solid(); box.fill.fore_color.rgb = LIGHT_CARD; box.line.color.rgb = BORDER
        tf = box.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = name; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
        p = tf.add_paragraph(); p.text = "\n" + status; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = TEAL

    # ==========================================
    # SLIDE 6: Upstream & IPLC: High Capacity Rules
    # ==========================================
    s6 = prs.slides.add_slide(blank)
    add_header(s6, "SECTION IV", "Upstream & IPLC: Bandwidth Alert Rules (High Capacity)", "Committed bandwidth threshold rules (≥ 5 Gbps)")
    bw_high = [
        ["Carrier / Circuit", "Link Type", "Alert Trigger Threshold", "Trigger Logic"],
        ["IPLC Viettel", "IPLC Circuit", "Traffic over 20 Gbps", "Alert Over Committed Bandwidth"],
        ["VNPT", "Global Upstream", "Traffic over 20 Gbps", "Alert Over Committed Bandwidth"],
        ["PCCW", "Global Upstream", "Traffic over 10 Gbps", "Alert Over Committed Bandwidth"],
        ["Cogent", "Global Upstream", "Traffic over 10 Gbps", "Alert Over Committed Bandwidth"],
        ["Viettel-BVT / Viettel-CHT", "Global Upstream", "Traffic over 5 Gbps each", "Alert Over Committed Bandwidth"]
    ]
    tbl_shape = s6.shapes.add_table(6, 4, Inches(0.8), Inches(1.9), Inches(11.73), Inches(4.6))
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

    # ==========================================
    # SLIDE 7: Upstream & IPLC: Links Under 5 Gbps
    # ==========================================
    s7 = prs.slides.add_slide(blank)
    add_header(s7, "SECTION IV", "Upstream & IPLC: Bandwidth Alert Rules", "Committed bandwidth alert thresholds for links below 5 Gbps")
    sub_links = [
        ("Telia (Arelion)", "> 3 Gbps"),
        ("NTT", "> 2 Gbps"),
        ("CU", "> 400 Mbps"),
        ("CT", "> 300 Mbps"),
        ("CMI", "> 50 Mbps")
    ]
    for i, (carrier, threshold) in enumerate(sub_links):
        y_pos = Inches(1.9 + i * 0.95)
        card = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), y_pos, Inches(11.73), Inches(0.75))
        card.fill.solid(); card.fill.fore_color.rgb = LIGHT_CARD; card.line.color.rgb = BORDER

        # Carrier
        tb_car = s7.shapes.add_textbox(Inches(1.2), y_pos + Inches(0.18), Inches(4.0), Inches(0.4))
        p_car = tb_car.text_frame.paragraphs[0]; p_car.text = carrier; p_car.font.size = Pt(14); p_car.font.bold = True; p_car.font.color.rgb = NAVY

        # Threshold
        tb_th = s7.shapes.add_textbox(Inches(6.8), y_pos + Inches(0.18), Inches(2.2), Inches(0.4))
        p_th = tb_th.text_frame.paragraphs[0]; p_th.text = threshold; p_th.font.size = Pt(14); p_th.font.bold = True; p_th.font.color.rgb = BLUE

        # Alert Pill Badge
        badge = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.8), y_pos + Inches(0.15), Inches(2.4), Inches(0.45))
        badge.fill.solid(); badge.fill.fore_color.rgb = AMBER_BG; badge.line.fill.background()
        p_b = badge.text_frame.paragraphs[0]; p_b.text = "Alert over committed"; p_b.alignment = PP_ALIGN.CENTER; p_b.font.size = Pt(11); p_b.font.bold = True; p_b.font.color.rgb = AMBER_TXT

    # ==========================================
    # SLIDE 8: Team Leader & Supervision
    # ==========================================
    s8 = prs.slides.add_slide(blank)
    add_header(s8, "SECTION V", f"{team_name}: Backup Senior Team Leader & Supervision", "Escalation line and supervisory distribution")
    lead1 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.5))
    lead1.fill.solid(); lead1.fill.fore_color.rgb = LIGHT_CARD; lead1.line.color.rgb = BORDER
    tf_l1 = lead1.text_frame; tf_l1.word_wrap = True
    p = tf_l1.paragraphs[0]; p.text = "Main Supervisor"; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = BLUE
    p = tf_l1.add_paragraph(); p.text = roles.get('mainSuper', 'Ouk Ratha'); p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf_l1.add_paragraph(); p.text = "\nPrimary operational lead for shift supervision, queue balance, and primary escalations."; p.font.size = Pt(12); p.font.color.rgb = GRAY

    lead2 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.5))
    lead2.fill.solid(); lead2.fill.fore_color.rgb = LIGHT_CARD; lead2.line.color.rgb = BORDER
    tf_l2 = lead2.text_frame; tf_l2.word_wrap = True
    p = tf_l2.paragraphs[0]; p.text = "Backup Supervisor"; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = BLUE
    p = tf_l2.add_paragraph(); p.text = roles.get('backupSuper', 'Ly Manith'); p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf_l2.add_paragraph(); p.text = "\nStandby supervisor to manage queue escalations during supervisor absence or high ticket load."; p.font.size = Pt(12); p.font.color.rgb = GRAY

    # ==========================================
    # SLIDE 9: Support Role Stations
    # ==========================================
    s9 = prs.slides.add_slide(blank)
    add_header(s9, "SECTION V", f"{team_name}: Support Role Stations", "Workstation assignments for external desk and internal support")
    st1 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.5))
    st1.fill.solid(); st1.fill.fore_color.rgb = LIGHT_CARD; st1.line.color.rgb = BORDER
    tf_st1 = st1.text_frame; tf_st1.word_wrap = True
    p = tf_st1.paragraphs[0]; p.text = "External Support Roles"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf_st1.add_paragraph(); p.text = "HQ01: Ouk Ratha\nHQ02: Ly Manith\nHQ03: Soth Rasal\nHQ04: Thy Piseth / Phy Reaksmey / Oung Savorn\n\nNote: HQ04 rotated daily among the 3 staff"; p.font.size = Pt(13); p.font.color.rgb = NAVY

    st2 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.5))
    st2.fill.solid(); st2.fill.fore_color.rgb = LIGHT_CARD; st2.line.color.rgb = BORDER
    tf_st2 = st2.text_frame; tf_st2.word_wrap = True
    p = tf_st2.paragraphs[0]; p.text = "Internal Support Roles"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf_st2.add_paragraph(); p.text = "Dedicated internal inquiry and coordination coverage\n\nStaff: Thy Piseth\nStaff: Phy Reaksmey\nStaff: Oung Savorn\nStaff: Meu Chanleaksmey"; p.font.size = Pt(13); p.font.color.rgb = NAVY

    # ==========================================
    # SLIDE 10: Daily Report Handling
    # ==========================================
    s10 = prs.slides.add_slide(blank)
    add_header(s10, "SECTION V", f"{team_name}: Daily Report Handling", "Shift roles for Network Monitoring/email and AMIS reporting")
    rep1 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.5))
    rep1.fill.solid(); rep1.fill.fore_color.rgb = LIGHT_CARD; rep1.line.color.rgb = BORDER
    tf_rep1 = rep1.text_frame; tf_rep1.word_wrap = True
    p = tf_rep1.paragraphs[0]; p.text = "Network Monitoring & Email"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf_rep1.add_paragraph(); p.text = "Daily status tracking and external correspondence\n\nMain: Soth Rasal\nBackup: Thy Piseth"; p.font.size = Pt(13); p.font.color.rgb = NAVY

    rep2 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.5))
    rep2.fill.solid(); rep2.fill.fore_color.rgb = LIGHT_CARD; rep2.line.color.rgb = BORDER
    tf_rep2 = rep2.text_frame; tf_rep2.word_wrap = True
    p = tf_rep2.paragraphs[0]; p.text = "AMIS Report Handling"; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf_rep2.add_paragraph(); p.text = "Assigned on rotating shifts:\n\n• Thy Piseth\n• Phy Reaksmey\n• Oung Savorn\n• Meu Chanleaksmey\n\nNote: Staff assigned rotately for AMIS"; p.font.size = Pt(13); p.font.color.rgb = NAVY

    # ==========================================
    # SLIDE 11: Executive Summary & Action Plan
    # ==========================================
    s11 = prs.slides.add_slide(blank)
    add_header(s11, "Executive Summary & Action Plan", "Key operational review checkpoints")
    actions = [
        ("Zero Critical Incidents Maintained", f"Successfully resolved {grand_total} cases with stable tool performance (AMIS, Billing, RT)."),
        ("Active Bandwidth Alert Thresholds", "Rules enforced across all 11 Upstream and IPLC connections."),
        ("Capability Upskilling", "Scheduled internal training on core network, monitoring tools, products & services.")
    ]
    for i, (title, desc) in enumerate(actions):
        box = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 4.0), Inches(2.0), Inches(3.7), Inches(4.3))
        box.fill.solid(); box.fill.fore_color.rgb = LIGHT_CARD; box.line.color.rgb = BORDER
        tf = box.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = title; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = NAVY
        p = tf.add_paragraph(); p.text = "\n" + desc; p.font.size = Pt(12); p.font.color.rgb = GRAY

    # ==========================================
    # SLIDE 12: Questions & Discussion (Dark Background)
    # ==========================================
    s12 = prs.slides.add_slide(blank)
    bg = s12.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid(); bg.fill.fore_color.rgb = DARK_BG; bg.line.fill.background()

    tb_q = s12.shapes.add_textbox(Inches(1.0), Inches(2.8), Inches(11.333), Inches(2.0))
    tf_q = tb_q.text_frame; tf_q.word_wrap = True
    p = tf_q.paragraphs[0]; p.text = "Questions & Discussion"; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(36); p.font.bold = True; p.font.color.rgb = WHITE
    p = tf_q.add_paragraph(); p.text = f"Weekly Case Summary Report • {team_name} Operations"; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(14); p.font.color.rgb = RGBColor(148, 163, 184)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    filename = f"Weekly_Report_{team_name.replace(' ', '_')}.pptx"
    return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")

# Universal catch-all route for Vercel
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
'@

Set-Content -Path "C:\Downloader\test\app.py" -Value $code -Encoding utf8
Set-Content -Path "C:\Downloader\test\api\index.py" -Value $code -Encoding utf8