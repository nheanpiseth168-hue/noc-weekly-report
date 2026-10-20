import io
import os
from flask import Flask, render_template, request, send_file
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

base_dir = os.path.dirname(os.path.abspath(__file__))
template_dir = os.path.abspath(os.path.join(base_dir, "..", "templates"))
app = Flask(__name__, template_folder=template_dir)

NAVY = RGBColor(15, 23, 42)
BLUE = RGBColor(2, 132, 199)
TEAL = RGBColor(13, 148, 136)
GRAY = RGBColor(100, 116, 139)
LIGHT = RGBColor(248, 250, 252)
BORDER = RGBColor(226, 232, 240)
WHITE = RGBColor(255, 255, 255)

CLR_30M   = RGBColor(16, 185, 129)
CLR_1H    = RGBColor(245, 158, 11)
CLR_2H    = RGBColor(249, 115, 22)
CLR_3H    = RGBColor(239, 68, 68)
CLR_OVER3 = RGBColor(112, 14, 14)

def add_header(slide, section, title, subtitle):
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(1.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]; p1.text = section.upper(); p1.font.size = Pt(11); p1.font.bold = True; p1.font.color.rgb = BLUE
    p2 = tf.add_paragraph(); p2.text = title; p2.font.size = Pt(22); p2.font.bold = True; p2.font.color.rgb = NAVY
    if subtitle:
        p3 = tf.add_paragraph(); p3.text = subtitle; p3.font.size = Pt(12); p3.font.color.rgb = GRAY

def generate_presentation(payload):
    team_name = payload.get("teamName", "Team 1")
    period = payload.get("reportPeriod", "Week 42 (Current Shift)")
    critical_incidents = str(payload.get("criticalIncidents", "0"))
    staff_data = payload.get("staff", [])
    platform = payload.get("platform", {})
    roles = payload.get("roles", {})

    processed_staff = []
    t30 = t1 = t2 = t3 = tover = 0
    for s in staff_data:
        m30 = int(s.get("m30", 0)); h1 = int(s.get("h1", 0)); h2 = int(s.get("h2", 0)); h3 = int(s.get("h3", 0)); over3 = int(s.get("over3", 0))
        total = m30 + h1 + h2 + h3 + over3
        t30 += m30; t1 += h1; t2 += h2; t3 += h3; tover += over3
        processed_staff.append({"name": s.get("name", "Staff"), "m30": m30, "h1": h1, "h2": h2, "h3": h3, "over3": over3, "total": total})

    grand_total = t30 + t1 + t2 + t3 + tover
    processed_staff.sort(key=lambda x: x["total"], reverse=True)
    sla_rate = f"{(t30 / grand_total * 100):.1f}%" if grand_total > 0 else "0.0%"

    prs = Presentation()
    prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Slide 1: Cover
    s1 = prs.slides.add_slide(blank)
    tb = s1.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.3), Inches(3.5))
    tf = tb.text_frame
    p = tf.paragraphs[0]; p.text = "OPERATIONS & NETWORK SUPPORT"; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = BLUE
    p = tf.add_paragraph(); p.text = "Weekly Report & Teamwork Roles"; p.font.size = Pt(32); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf.add_paragraph(); p.text = f"Incident overview, system health, bandwidth alerts, and {team_name} operational assignments"; p.font.size = Pt(15); p.font.color.rgb = GRAY
    p = tf.add_paragraph(); p.text = f"{team_name}  |  {period}"; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = BLUE

    # Slide 2: KPI Overview
    s2 = prs.slides.add_slide(blank)
    add_header(s2, "SECTION I", "Weekly Case Summary Overview", "Key incident metrics across customer and internal queues")
    kpis = [(str(grand_total), "Total Cases Handled", "Weekly operational volume"), (critical_incidents, "Critical Incidents", "Zero service-impacting severity"), (str(len(processed_staff)), "Active Engineers", "Full team dispatch coverage"), (sla_rate, "Sub-30m SLA Rate", f"{t30} of {grand_total} resolved under 30 mins")]
    for i, (val, title, sub) in enumerate(kpis):
        box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 2.95), Inches(2.2), Inches(2.8), Inches(3.5))
        box.fill.solid(); box.fill.fore_color.rgb = LIGHT; box.line.color.rgb = BORDER
        tf = box.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = val; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(36); p.font.bold = True; p.font.color.rgb = BLUE
        p = tf.add_paragraph(); p.text = title; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = NAVY
        p = tf.add_paragraph(); p.text = sub; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(11); p.font.color.rgb = GRAY

    # Slide 3: KPI Table
    s3 = prs.slides.add_slide(blank)
    add_header(s3, "SECTION I", "Staff KPI: Support Cases Handled", f"Total support cases handled per engineer (Total: {grand_total} cases)")
    rows = len(processed_staff) + 2
    tbl = s3.shapes.add_table(rows, 8, Inches(0.8), Inches(1.75), Inches(11.7), Inches(4.8)).table
    for i, w in enumerate([Inches(3.2), Inches(1.1), Inches(1.1), Inches(1.1), Inches(1.1), Inches(1.1), Inches(1.3), Inches(1.7)]): tbl.columns[i].width = w
    for i, h in enumerate(["Engineer Name", "< 30m", "1h", "2h", "3h", "> 3h", "Total", "Share"]):
        c = tbl.cell(0, i); c.text = h; c.fill.solid(); c.fill.fore_color.rgb = BLUE
        p = c.text_frame.paragraphs[0]; p.font.bold = True; p.font.size = Pt(11); p.font.color.rgb = WHITE
        if i > 0: p.alignment = PP_ALIGN.CENTER
    for r_idx, s in enumerate(processed_staff, start=1):
        share_str = f"{(s['total'] / grand_total * 100):.1f}%" if grand_total > 0 else "0.0%"
        for c_idx, val in enumerate([s["name"], str(s["m30"]), str(s["h1"]), str(s["h2"]), str(s["h3"]), str(s["over3"]), str(s["total"]), share_str]):
            c = tbl.cell(r_idx, c_idx); c.text = val
            p = c.text_frame.paragraphs[0]; p.font.size = Pt(10.5); p.font.color.rgb = NAVY
            if c_idx > 0: p.alignment = PP_ALIGN.CENTER
    for c_idx, val in enumerate(["Team Total", str(t30), str(t1), str(t2), str(t3), str(tover), str(grand_total), "100%"]):
        c = tbl.cell(rows - 1, c_idx); c.text = val; c.fill.solid(); c.fill.fore_color.rgb = LIGHT
        p = c.text_frame.paragraphs[0]; p.font.bold = True; p.font.size = Pt(11); p.font.color.rgb = NAVY
        if c_idx > 0: p.alignment = PP_ALIGN.CENTER

    # Slide 4: Chart
    s4 = prs.slides.add_slide(blank)
    add_header(s4, "SECTION I  •  Resolution Turnaround Time", "Resolution Time Breakdown by Engineer", None)
    lg_x = Inches(8.3)
    for lbl, clr in [("< 30mn", CLR_30M), ("1H", CLR_1H), ("2H", CLR_2H), ("3H", CLR_3H), ("Over 3H", CLR_OVER3)]:
        sq = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, lg_x, Inches(1.1), Inches(0.16), Inches(0.16))
        sq.fill.solid(); sq.fill.fore_color.rgb = clr; sq.line.fill.background()
        tb_lg = s4.shapes.add_textbox(lg_x + Inches(0.2), Inches(0.96), Inches(0.8), Inches(0.35))
        p_lg = tb_lg.text_frame.paragraphs[0]; p_lg.text = lbl; p_lg.font.size = Pt(9.5); p_lg.font.bold = True; p_lg.font.color.rgb = GRAY
        lg_x += Inches(0.95)
    chart_bg = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(11.73), Inches(4.7))
    chart_bg.fill.solid(); chart_bg.fill.fore_color.rgb = WHITE; chart_bg.line.color.rgb = BORDER
    max_cases = max([s["total"] for s in processed_staff] + [1])
    scale = Inches(8.0) / max_cases
    row_step = Inches(4.3 / max(len(processed_staff), 1))
    for idx, s in enumerate(processed_staff):
        curr_y = Inches(1.7) + idx * row_step
        tb_name = s4.shapes.add_textbox(Inches(0.9), curr_y - Inches(0.08), Inches(2.2), Inches(0.35))
        p_name = tb_name.text_frame.paragraphs[0]; p_name.text = s["name"]; p_name.font.size = Pt(9.5); p_name.font.bold = True; p_name.font.color.rgb = NAVY; p_name.alignment = PP_ALIGN.RIGHT
        curr_x = Inches(3.2)
        for count, clr in [(s["m30"], CLR_30M), (s["h1"], CLR_1H), (s["h2"], CLR_2H), (s["h3"], CLR_3H), (s["over3"], CLR_OVER3)]:
            if count > 0:
                seg_w = count * scale
                seg = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, curr_x, curr_y, seg_w, Inches(0.20))
                seg.fill.solid(); seg.fill.fore_color.rgb = clr; seg.line.fill.background()
                curr_x += seg_w
        tb_tot = s4.shapes.add_textbox(curr_x + Inches(0.12), curr_y - Inches(0.08), Inches(0.6), Inches(0.35))
        p_tot = tb_tot.text_frame.paragraphs[0]; p_tot.text = str(s["total"]); p_tot.font.size = Pt(10); p_tot.font.bold = True; p_tot.font.color.rgb = NAVY
    ft_box = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.35), Inches(11.73), Inches(0.65))
    ft_box.fill.solid(); ft_box.fill.fore_color.rgb = LIGHT; ft_box.line.color.rgb = BORDER
    p_ft = ft_box.text_frame.paragraphs[0]
    p_ft.text = f"Fast Turnaround: {t30} cases closed <30m. Complex escalations (>3h) concentrated under MAK BUN LEANG & I CHHENGKHEANG.   {sla_rate} SLA Sub-30m"
    p_ft.font.size = Pt(10); p_ft.font.color.rgb = GRAY; p_ft.alignment = PP_ALIGN.CENTER

    # Slide 5: Health
    s5 = prs.slides.add_slide(blank)
    add_header(s5, "SECTION III", "Tool Supports & System Health", "Operational health status across internal tools and support suites")
    tools = [("Portal Support", platform.get("portal", "Helpful & Responsive"), "No delays reported. Queue handling responsive."), ("AMIS Platform", platform.get("amis", "No Issue"), "Customer database and records synchronized normally."), ("Billing System", platform.get("billing", "No Issue"), "Customer billing and account payments operating normally."), ("Request Tracker", platform.get("rt", "No Issue"), "Internal ticketing queue operating without latency.")]
    for i, (name, status, desc) in enumerate(tools):
        box = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 2.95), Inches(2.2), Inches(2.8), Inches(3.5))
        box.fill.solid(); box.fill.fore_color.rgb = LIGHT; box.line.color.rgb = BORDER
        tf = box.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = name; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = NAVY
        p = tf.add_paragraph(); p.text = status; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = TEAL
        p = tf.add_paragraph(); p.text = desc; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(11); p.font.color.rgb = GRAY

    # Slide 6: BW Alerts
    s6 = prs.slides.add_slide(blank)
    add_header(s6, "SECTION IV", "Upstream & IPLC: Bandwidth Alert Rules (High Capacity)", "Committed bandwidth threshold rules (≥ 5 Gbps)")
    bw_high = [["Carrier / Circuit", "Link Type", "Alert Trigger Threshold", "Trigger Logic"], ["IPLC Viettel", "IPLC Circuit", "Traffic over 20 Gbps", "Alert Over Committed Bandwidth"], ["VNPT", "Global Upstream", "Traffic over 20 Gbps", "Alert Over Committed Bandwidth"], ["PCCW", "Global Upstream", "Traffic over 10 Gbps", "Alert Over Committed Bandwidth"], ["Cogent", "Global Upstream", "Traffic over 10 Gbps", "Alert Over Committed Bandwidth"], ["Viettel-BVT / Viettel-CHT", "Global Upstream", "Traffic over 5 Gbps each", "Alert Over Committed Bandwidth"]]
    tbl_high = s6.shapes.add_table(6, 4, Inches(0.8), Inches(1.9), Inches(11.7), Inches(3.8)).table
    for i, w in enumerate([Inches(3.2), Inches(2.5), Inches(3.0), Inches(3.0)]): tbl_high.columns[i].width = w
    for r_idx, row in enumerate(bw_high):
        for c_idx, val in enumerate(row):
            cell = tbl_high.cell(r_idx, c_idx); cell.text = val; p = cell.text_frame.paragraphs[0]; p.font.size = Pt(11)
            if r_idx == 0: cell.fill.solid(); cell.fill.fore_color.rgb = BLUE; p.font.bold = True; p.font.color.rgb = WHITE
            else: p.font.color.rgb = NAVY

    # Slide 7: Roles
    s7 = prs.slides.add_slide(blank)
    add_header(s7, "SECTION V", f"{team_name}: Supervision & Workstation Assignments", "Shift roles, external desk seating, and internal inquiry coverage")
    box1 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.9), Inches(5.6), Inches(4.5))
    box1.fill.solid(); box1.fill.fore_color.rgb = LIGHT; box1.line.color.rgb = BORDER
    tf1 = box1.text_frame; tf1.word_wrap = True
    p = tf1.paragraphs[0]; p.text = "External Support & Supervision"; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf1.add_paragraph()
    p.text = f"• Main Supervisor: {roles.get('mainSuper', '')}\n• Backup Supervisor: {roles.get('backupSuper', '')}\n• {roles.get('hqFixed', '')}\n• HQ04 Station (Daily Rotation):\n   {roles.get('hq04', '')}"
    p.font.size = Pt(12); p.font.color.rgb = NAVY
    box2 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(1.9), Inches(5.6), Inches(4.5))
    box2.fill.solid(); box2.fill.fore_color.rgb = LIGHT; box2.line.color.rgb = BORDER
    tf2 = box2.text_frame; tf2.word_wrap = True
    p = tf2.paragraphs[0]; p.text = "Internal Roles & Daily Reporting"; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = NAVY
    p = tf2.add_paragraph()
    p.text = f"• Dedicated Internal Support:\n   {roles.get('internal', '')}\n\n• Network Monitoring & Email Shift:\n   {roles.get('monitoring', '')}\n\n• AMIS Daily Report Handling (Rotating):\n   {roles.get('amisReport', '')}"
    p.font.size = Pt(12); p.font.color.rgb = NAVY

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    filename = f"Weekly_Report_{team_name.replace(' ', '_')}.pptx"
    return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")

# Universal Catch-All Route (Handles both direct visits and Vercel rewrites)
@app.route("/", defaults={"path": ""}, methods=["GET", "POST"])
@app.route("/<path:path>", methods=["GET", "POST"])
def catch_all(path):
    if request.method == "POST" or "export" in path:
        payload = request.get_json(silent=True) or {}
        return generate_presentation(payload)
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True, port=5000)
