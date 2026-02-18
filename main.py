from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def create_resume():
    doc = Document()
    
    # --- PAGE & STYLE SETUP ---
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.25)
    section.left_margin = section.right_margin = Inches(0.25)

    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    pf = style.paragraph_format
    pf.line_spacing = 1.0
    pf.space_before = pf.space_after = Pt(0)

    # Helper: Full-row Underline with Padding
    def add_section_underline(paragraph):
        pBr = paragraph._element.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        # Setting 'w:space' to 4 creates the vertical gap between text and line
        bottom.set(qn('w:space'), '4') 
        bottom.set(qn('w:color'), 'auto')
        pBdr.append(bottom)
        pBr.append(pBdr)

    # --- CONDENSED CENTERED HEADER --- 
    name_p = doc.add_paragraph()
    name_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name_p.paragraph_format.space_after = Pt(1)
    name_run = name_p.add_run("MICHAEL KEN LOUIE")
    name_run.bold = True
    name_run.font.size = Pt(28)
    

    contact_p = doc.add_paragraph()
    contact_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact_p.paragraph_format.space_after = Pt(2)
    contact_run = contact_p.add_run("✉ michael.louie10@gmail.com  | ✆ 540-645-1914  | 📍 McLean, VA")
    contact_run.font.size = Pt(12)

    link_p = doc.add_paragraph()
    link_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    link_run = link_p.add_run("</> github.com/Michaelis105  | [in] linkedin.com/in/louiemichael")
    link_run.font.size = Pt(12)    

    # --- SECTION: WORK EXPERIENCE ---
    hw = doc.add_paragraph("WORK EXPERIENCE")
    hw.runs[0].bold = True
    hw.paragraph_format.space_before = Pt(12)
    add_section_underline(hw)

    def add_job(co, title, loc, dates, bullets, level2_range=None):
        p1 = doc.add_paragraph()
        p1.paragraph_format.space_before = Pt(10)
        p1.paragraph_format.tab_stops.add_tab_stop(Inches(8.0), WD_TAB_ALIGNMENT.RIGHT)
        p1.add_run(co).bold = True
        p1.add_run(f"\t{dates}")
        
        p2 = doc.add_paragraph()
        p2.paragraph_format.space_after = Pt(4)
        p2.paragraph_format.tab_stops.add_tab_stop(Inches(8.0), WD_TAB_ALIGNMENT.RIGHT)
        p2.add_run(title).italic = True
        p2.add_run(f"\t{loc}")
        
        for idx, b in enumerate(bullets):
            bp = doc.add_paragraph()
            # Handle both plain strings and tuples with formatting
            if isinstance(b, tuple):
                bold_text, is_bold, rest = b
                run = bp.add_run(f"• ")
                run = bp.add_run(f"{bold_text}")
                run.bold = is_bold
                bp.add_run(f" {rest}")
            else:
                bp.add_run(f"• {b}")
            # Indenting bullets slightly from the left margin
            if level2_range and level2_range[0] <= idx <= level2_range[1]:
                bp.paragraph_format.left_indent = Inches(0.5) # Deeper nested indent
                bp.paragraph_format.first_line_indent = Inches(-0.2)
            else:
                bp.paragraph_format.left_indent = Inches(0.3) # Standard slight indent
                bp.paragraph_format.first_line_indent = Inches(-0.2)

    # Capital One 
    c1_bullets = [
        "Tech Lead launching U.S. industry first self-service cashier’s check kiosk.",
        ("Patent", True, "(US11636464B2) granted on innovative kiosk experience."),
        "Supporting high-stakes transactions totaling >$20 million in customer cashier’s checks across money markets.",
        "Scaling AWS serverless microservice stack via Terraform-like infra-as-code to production, 99.9% uptime.",
        "Introducing standard React, RESTful Node.js API, NoSQL tech stack across multiple self-service platforms.",#
        "Steering engineers via pair programming, code reviews, and resiliency on-call playbooks.",
        "Conveying technical intent with product owners and engineers via architecture/dataflow/API design diagrams.",
        "Centralizing transaction and kiosk state at single source data lake via Kafka for real-time monitoring.",
        "Reducing kiosk deployment times by 80% by pioneering fleet management pub-sub operation code mechanism.",
        "Tech Lead for owning technical direction of green-field self-service instant payment issuance card kiosk.",
        "Designing and developing ATM fleet managing/monitoring distributed system serving real-time operations/auditing.",
        "Building MSI to automate ATM software platform lifecycle management, reducing per kiosk downtime by 60%."
    ]
    add_job("Capital One Financial", "Bank Tech – Consumer Self-Servicing, Lead Software Engineer", "Tysons, VA", "July 2019 – Present", c1_bullets, level2_range=(1,7))

    bloomberg_bullets = ["Minimized customers' secure access outage via preemptive SAML certificate expiration notifications detection.", "Improved platform observability by implementing health check web service using Spring, Vue.js, and Vuetify."]
    add_job("Bloomberg Industry Group", "Subscription Management and Customer Support Platform, Software Engineer", "Arlington, VA", "August 2018 – July 2019", bloomberg_bullets)
    
    vs_bullets = [
        "Optimized DotGov portal domain management web service UX collaborating with customer service and GSA users.",
        "Developed internal code dependency analysis reporting tool to analyze and report project security vulnerabilities."
    ]
    add_job("Verisign, Inc.", "Consolidated Top-Level Domain - Infrastructure Services, Software Engineer I-II", "Reston, VA", "February 2017 – August 2018", vs_bullets)
    
    lm_bullets = [
        "Enabled sonar applications task clustering scaling and management via Mesos/Marathon."
    ]
    add_job("Lockheed Martin", "Acoustic Rapid COTS Insertion System Services, Software Engineer Associate", "Manassas, VA", "June 2016 – February 2017", lm_bullets)

    # --- SECTION: TECHNICAL EXPERTISE ---
    ht = doc.add_paragraph("TECHNICAL EXPERTISE")
    ht.runs[0].bold = True
    ht.paragraph_format.space_before = Pt(12)
    ht.paragraph_format.space_after = Pt(8)
    add_section_underline(ht)

    skills = [
        ("Languages & Core Tech", "Python + Flask, Node.js + Express, Bash, Shell Scripting, Typescript/JavaScript"),
        ("Frontend", "React / VueJS, HTML / CSS"),
        ("Cloud & Infrastructure", "Amazon Web Services (AWS), Unix / Linux, Docker, Terraform"),
        ("Data Engineering", "SQL / NoSQL, MySQL, PostgreSQL, Snowflake, Kafka"),
        ("Observability/SRE Tools", "Splunk, New Relic, Cloudwatch, PagerDuty, Git / GitHub"),
        ("AI-Assisted Development", "Windsurf / Copilot / Gemini")
    ]
    for skill in skills:
        p = doc.add_paragraph()
        category, items = skill
        run = p.add_run(f"• {category}")
        run.bold = True
        p.add_run(f": {items}")
        p.paragraph_format.left_indent = Inches(0.3)

    # --- SECTION: CERTIFICATIONS --- 
    hc = doc.add_paragraph("CERTIFICATIONS")
    hc.runs[0].bold = True
    hc.paragraph_format.space_before = Pt(12)
    hc.paragraph_format.space_after = Pt(8)
    add_section_underline(hc)
    
    certs = ["AWS Certified Developer Associate and Cloud Practitioner", "AWS Certified Generative AI Developer – Professional and Solutions Architect (scheduled Q2 2026)", "CompTIA Network+ Certification N10-006"]
    for cert in certs:
        p = doc.add_paragraph(f"• {cert}")
        p.paragraph_format.left_indent = Inches(0.3)

    # --- SECTION: EDUCATION --- 
    he = doc.add_paragraph("EDUCATION")
    he.runs[0].bold = True
    he.paragraph_format.space_before = Pt(12)
    add_section_underline(he)

    # Georgia Tech 
    pe1 = doc.add_paragraph()
    pe1.paragraph_format.space_before = Pt(8)
    pe1.paragraph_format.tab_stops.add_tab_stop(Inches(8.0), WD_TAB_ALIGNMENT.RIGHT)
    pe1.add_run("Georgia Institute of Technology").bold = True
    pe1.add_run("\tSpring 2021")
    pe1_sub = doc.add_paragraph()
    pe1_sub.paragraph_format.tab_stops.add_tab_stop(Inches(8.0), WD_TAB_ALIGNMENT.RIGHT)
    pe1_sub.add_run("M.S. in Computer Science (Computing Systems)").italic = True
    pe1_sub.add_run("\tAtlanta, GA")
    pe1_sub.paragraph_format.space_after = Pt(4)

    # Virginia Tech 
    pe2 = doc.add_paragraph()
    pe2.paragraph_format.space_before = Pt(8)
    pe2.paragraph_format.tab_stops.add_tab_stop(Inches(8.0), WD_TAB_ALIGNMENT.RIGHT)
    pe2.add_run("Virginia Polytechnic Institute and State University").bold = True
    pe2.add_run("\tMay 2016")
    pe2_sub = doc.add_paragraph()
    pe2_sub.paragraph_format.tab_stops.add_tab_stop(Inches(8.0), WD_TAB_ALIGNMENT.RIGHT)
    pe2_sub.add_run("B.S. in Computer Science").italic = True
    pe2_sub.add_run("\tBlacksburg, VA")
    pe2_sub.paragraph_format.space_after = Pt(4)

    doc.save('michael-louie-resume.docx')

create_resume()