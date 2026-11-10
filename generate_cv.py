"""Generate Tusher_Mondal_CV.xlsx — Traditional CV reimagined in Microsoft Excel.

Reads structured portfolio and resume data and produces a polished,
print-ready 4-tab Excel workbook:
  1. CV                     - Traditional monochrome editorial CV layout
  2. Career Dashboard       - Formula-driven KPIs and native Excel charts
  3. Experience & Projects  - Filterable career and project ledger
  4. Education & Skills     - Academic record, achievements, and skills matrix
"""

from __future__ import annotations

import datetime as dt
import math
from pathlib import Path
from typing import Iterable

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break
from PIL import ImageFont

# ---------------------------------------------------------------------------
# Palette & Typography (Editorial Monochrome + Subtle Excel Emerald Accent)
# ---------------------------------------------------------------------------
INK = "111827"          # Primary near-black
SLATE = "374151"        # Dark slate body
MUTED = "4B5563"        # Secondary metadata
CAPTION = "6B7280"      # Subtle caption / footer
ACCENT = "1B5E3A"       # Classic Excel forest emerald for links & highlights
ACCENT_SOFT = "ECFDF5"  # Soft emerald tint for callout / KPI highlights
CARD_BG = "F8FAFC"      # Subtle card / alternating tint
ZEBRA_BG = "F9FAFB"     # Zebra row background
HEADER_BG = "111827"    # Dark editorial table header
HEADER_FG = "FFFFFF"    # Header white text
BORDER_LIGHT = "E5E7EB" # Hairline divider
BORDER_MED = "CBD5E1"   # Medium card border
BORDER_DARK = "111827"  # Strong section rule
WHITE = "FFFFFF"

FONT_DISPLAY = "Georgia"
FONT_BODY = "Calibri"

# Font metrics loader for accurate wrapped-row height estimation
_FONT_CACHE: dict[tuple[str, int, bool], ImageFont.FreeTypeFont] = {}


def _get_pil_font(family: str, size_pt: int, bold: bool = False) -> ImageFont.FreeTypeFont | None:
    key = (family, size_pt, bold)
    if key in _FONT_CACHE:
        return _FONT_CACHE[key]
    candidates = {
        ("Calibri", False): r"C:\Windows\Fonts\calibri.ttf",
        ("Calibri", True): r"C:\Windows\Fonts\calibrib.ttf",
        ("Georgia", False): r"C:\Windows\Fonts\georgia.ttf",
        ("Georgia", True): r"C:\Windows\Fonts\georgiab.ttf",
    }
    path = Path(candidates.get((family, bold), r"C:\Windows\Fonts\calibri.ttf"))
    if not path.exists():
        return None
    # PIL renders at 96 DPI -> 1 pt = 96/72 px
    px = max(8, int(round(size_pt * 96 / 72)))
    font = ImageFont.truetype(str(path), px)
    _FONT_CACHE[key] = font
    return font


def estimate_wrapped_lines(text: str, width_chars: float, font_family: str = FONT_BODY, size_pt: int = 10, bold: bool = False) -> int:
    """Estimate line count when Excel wraps `text` inside `width_chars`."""
    if not text:
        return 1
    # In Excel, 1 character width unit at 11pt Calibri ~ 7 px + 5 px cell padding.
    # Conservative 0.90 factor prevents last-word clipping across zoom levels.
    max_px = max(40.0, (width_chars * 7.0 - 6.0) * 0.90)
    pil_font = _get_pil_font(font_family, size_pt, bold)
    total_lines = 0
    for para in str(text).split("\n"):
        if not para:
            total_lines += 1
            continue
        words = para.split(" ")
        lines = 1
        cur = ""
        for word in words:
            candidate = word if not cur else f"{cur} {word}"
            if pil_font is not None:
                w = pil_font.getlength(candidate)
            else:
                w = len(candidate) * (size_pt * 0.56)
            if w <= max_px:
                cur = candidate
            else:
                lines += 1
                cur = word
        total_lines += lines
    return max(1, total_lines)


def calc_row_height(lines: int, size_pt: int = 10, extra_pad: float = 4.0) -> float:
    line_h = size_pt * 1.38
    return round(max(15.0, lines * line_h + extra_pad), 1)


# ---------------------------------------------------------------------------
# Reusable Style Helpers
# ---------------------------------------------------------------------------
def mk_font(
    size: int = 10,
    bold: bool = False,
    italic: bool = False,
    color: str = INK,
    family: str = FONT_BODY,
    underline: str | None = None,
) -> Font:
    return Font(name=family, size=size, bold=bold, italic=italic, color=color, underline=underline)


def mk_fill(color: str) -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=color, bgColor=color)


def mk_border(
    left: tuple[str, str] | None = None,
    right: tuple[str, str] | None = None,
    top: tuple[str, str] | None = None,
    bottom: tuple[str, str] | None = None,
) -> Border:
    return Border(
        left=Side(style=left[0], color=left[1]) if left else Side(),
        right=Side(style=right[0], color=right[1]) if right else Side(),
        top=Side(style=top[0], color=top[1]) if top else Side(),
        bottom=Side(style=bottom[0], color=bottom[1]) if bottom else Side(),
    )


def style_range(ws, min_col: int, min_row: int, max_col: int, max_row: int, *, font=None, fill=None, border=None, alignment=None):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if font is not None:
                cell.font = font
            if fill is not None:
                cell.fill = fill
            if border is not None:
                cell.border = border
            if alignment is not None:
                cell.alignment = alignment


def apply_box_border(ws, min_col: int, min_row: int, max_col: int, max_row: int, style: str = "thin", color: str = BORDER_MED, top_accent: tuple[str, str] | None = None):
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cur = cell.border
            top_side = (Side(style=top_accent[0], color=top_accent[1]) if top_accent else Side(style=style, color=color)) if r == min_row else cur.top
            bottom_side = Side(style=style, color=color) if r == max_row else cur.bottom
            left_side = Side(style=style, color=color) if c == min_col else cur.left
            right_side = Side(style=style, color=color) if c == max_col else cur.right
            cell.border = Border(left=left_side, right=right_side, top=top_side, bottom=bottom_side)


# ---------------------------------------------------------------------------
# Portfolio & Resume Data (Reconciled from Tusher Mondal CV.tex + profile.js)
# ---------------------------------------------------------------------------
PROFILE = {
    "name": "Tusher Mondal",
    "role": "Manager — Management Reporting — CIB",
    "company": "HSBC",
    "location": "Bengaluru - 560001, Karnataka, India",
    "phone": "+91 91230 90573",
    "email": "tusher9073mondal@gmail.com",
    "linkedin_label": "linkedin.com/in/tusher-mondal",
    "linkedin_url": "https://www.linkedin.com/in/tusher-mondal/",
    "github_label": "github.com/Hurley2017",
    "github_url": "https://github.com/Hurley2017",
    "web_label": "www.tusher.in",
    "web_url": "https://www.tusher.in",
    "leetcode_url": "https://leetcode.com/Hurley2017/",
    "summary": (
        "Adept at learning new technologies, I bring a diverse skillset across Data Analytics, "
        "Enterprise Reporting, Web Development, and Software Engineering — turning messy data into "
        "decisions at HSBC and building local-AI and full-stack software."
    ),
}

EXPERIENCE = [
    {
        "company": "HSBC",
        "role": "Manager — Management Reporting — CIB",
        "type": "Full-Time",
        "location": "Bengaluru, India",
        "start": dt.date(2024, 7, 1),
        "end": "Present",
        "period": "Jul 2024 — Present",
        "bullets": [
            "Own management reporting for the Corporate & Investment Banking (CIB) division, turning raw numbers into actionable insights for enterprise stakeholders.",
            "Work across the enterprise analytics stack with Microsoft Excel (including VBA), Power BI, TM1, and Alteryx.",
            "Automate repetitive tasks by building custom internal tools, streamlining workflows and maximizing time efficiency and productivity.",
        ],
        "tech": "Excel, VBA, Power BI, Power Query, TM1, Alteryx, SQL, Python",
    },
    {
        "company": "HSBC",
        "role": "Analyst Intern (converted to full-time)",
        "type": "Internship",
        "location": "Chennai, India",
        "start": dt.date(2024, 1, 1),
        "end": dt.date(2024, 7, 1),
        "period": "Jan 2024 — Jul 2024",
        "bullets": [
            "Built an automated email reporting system monitoring data quality across specified businesses and regions for designated recipients.",
            "Automated repetitive tasks with custom internal tools, streamlining workflows and maximizing team productivity.",
            "Uncovered business trends through exploratory data analysis and stakeholder-ready visualizations.",
        ],
        "tech": "Python (pandas, openpyxl, xlwings, pywin32), Power Query, Microsoft Excel, VS Code, GitHub",
    },
    {
        "company": "Celebal Technologies",
        "role": "Data Science Intern · CSI Summer Internship 2023",
        "type": "Internship",
        "location": "Remote",
        "start": dt.date(2023, 6, 1),
        "end": dt.date(2023, 8, 31),
        "period": "Jun 2023 — Aug 2023",
        "bullets": [
            "Created a forecasting model utilizing user behavioral traits to predict future bandwidth consumption.",
            "Developed an interactive full-stack dashboard panel visualizing usage reports across key metrics.",
        ],
        "tech": "Python (pandas, scikit-learn), Data Space, NoSQL, VS Code",
    },
    {
        "company": "CodeClause",
        "role": "Python Developer Intern",
        "type": "Internship",
        "location": "Kolkata, India · Remote",
        "start": dt.date(2023, 5, 1),
        "end": dt.date(2023, 6, 30),
        "period": "May 2023 — Jun 2023",
        "bullets": [
            "Created a web-based URL shortener enabling sharing of long URLs using Flask and MongoDB database mapping.",
            "Engineered a Voice Assistant CLI leveraging Google's Gemini (formerly Bard) API, speech recognition, and pyttsx3, plus a Tkinter GUI alarm clock.",
        ],
        "tech": "Python (Flask, Tkinter, speech_recognition, Bard API, pyttsx3), MongoDB, GitHub",
    },
    {
        "company": "Phoenix India Research & Development Group",
        "role": "Software Developer Intern",
        "type": "Internship",
        "location": "Kolkata, India",
        "start": dt.date(2022, 5, 1),
        "end": dt.date(2022, 10, 31),
        "period": "May 2022 — Oct 2022",
        "bullets": [
            "Built an Automatic Commercial Authentication System using OpenCV and Dlib facial landmark pipelines.",
            "Implemented a multi-threaded verification architecture to concurrently execute feature checks, boosting stability and throughput.",
        ],
        "tech": "Python, OpenCV, Dlib, Threading, Git",
    },
]

PROJECTS = [
    {
        "name": "Writer",
        "year": 2026,
        "domain": "Local AI",
        "lead_tech": "Python",
        "tagline": "A local AI agent that writes full stories",
        "stack": "Python, LLM, SDXL, TTS",
        "github": "https://github.com/Hurley2017/Writer",
        "demo": None,
        "bullets": [
            "3-stage on-device AI pipeline drafting stories with a local LLM (LM Studio), generating photorealistic RealVisXL (SDXL) covers with CPU offload, and narrating dual-voice audiobooks via Orpheus 3B TTS.",
            "Strict-JSON outline generation with schema-aware retries, plain-prose fallback, and an automated A5 PDF book builder.",
        ],
        "show_on_cv": True,
    },
    {
        "name": "Piston-And-Sound",
        "year": 2026,
        "domain": "Audio / Web",
        "lead_tech": "JavaScript",
        "tagline": "Physics-based engine sound in the browser",
        "stack": "JavaScript, Web Audio API",
        "github": "https://github.com/Hurley2017/Piston-And-Sound",
        "demo": None,
        "bullets": [
            "Procedural in-browser audio synthesizer driven by engine physics — cylinder count, RPM, firing order, pulse attack shaping, idle sub-bass, and mechanical vibration rumble.",
        ],
        "show_on_cv": True,
    },
    {
        "name": "Rusty Radio",
        "year": 2026,
        "domain": "Systems",
        "lead_tech": "Rust",
        "tagline": "An ultra-light music player in Rust",
        "stack": "Rust",
        "github": "https://github.com/Hurley2017/Rusty_Radio",
        "demo": None,
        "bullets": [
            "Intentionally minimal background music player consuming bare-minimum system resources with an old-school, distraction-free design.",
        ],
        "show_on_cv": True,
    },
    {
        "name": "HandMouse",
        "year": 2026,
        "domain": "Computer Vision",
        "lead_tech": "Python",
        "tagline": "Your webcam is your mouse",
        "stack": "Python, OpenCV",
        "github": "https://github.com/Hurley2017/HandMouse",
        "demo": None,
        "bullets": [
            "Real-time hand tracking and gesture-to-mouse-action mapping using only a standard webcam and OpenCV computer vision.",
        ],
        "show_on_cv": True,
    },
    {
        "name": "Perspective Pivot",
        "year": 2025,
        "domain": "Data Tools",
        "lead_tech": "JavaScript",
        "tagline": "Binary files as a pivot table",
        "stack": "HTML, JavaScript",
        "github": "https://github.com/Hurley2017/Perspective_Pivot",
        "demo": None,
        "bullets": [
            "Web view framework that parses and renders raw binary files into an interactive pivot-table exploration UI in the browser.",
        ],
        "show_on_cv": False,
    },
    {
        "name": "AGG Web",
        "year": 2024,
        "domain": "Security / Web",
        "lead_tech": "Python",
        "tagline": "Attack graph generator (thesis project)",
        "stack": "Python, JavaScript, MulVAL, Graph",
        "github": "https://github.com/Hurley2017/AGG_WEB",
        "demo": "https://attack-graph-generator.vercel.app",
        "bullets": [
            "Full-stack web interface for generating and visualizing MulVAL attack graphs supporting network-simulated digital forensics (MCA thesis).",
        ],
        "show_on_cv": True,
    },
    {
        "name": "Email Spam Classifier",
        "year": 2023,
        "domain": "Machine Learning",
        "lead_tech": "Python",
        "tagline": "ML-powered spam detection",
        "stack": "Python, scikit-learn, ML",
        "github": "https://github.com/Hurley2017/Email_Spam",
        "demo": "https://email-spam-tusher.vercel.app",
        "bullets": [
            "Machine-learning text classification pipeline built with scikit-learn during the Celebal Technologies internship and deployed live on Vercel.",
        ],
        "show_on_cv": False,
    },
    {
        "name": "Socket Transfer",
        "year": 2023,
        "domain": "Networking",
        "lead_tech": "Python",
        "tagline": "Cross-platform P2P file transfer",
        "stack": "Python, Sockets, TCP/IP",
        "github": "https://github.com/Hurley2017/Socket_Streaming",
        "demo": None,
        "bullets": [
            "Server-client TCP/IP file-transfer system utilizing full NIC capacity with automatic chunk signing for fast, reliable reassembly.",
        ],
        "show_on_cv": True,
    },
    {
        "name": "Railway Reservation System",
        "year": 2022,
        "domain": "Full-Stack Web",
        "lead_tech": "Python",
        "tagline": "Multi-level train booking & admin platform",
        "stack": "Python, Flask, HTML/CSS, JavaScript, MongoDB",
        "github": "https://github.com/Hurley2017",
        "demo": None,
        "bullets": [
            "Multi-level access train booking platform allowing passengers to reserve seats while admins manage schedules, trains, and transactions.",
        ],
        "show_on_cv": True,
    },
]

EDUCATION = [
    {
        "school": "Jadavpur University",
        "degree": "Master of Computer Applications (MCA) — Computer Engineering",
        "period": "Sep 2022 — Jul 2024",
        "score": 8.8,
        "scale": "10.0 CGPA",
        "location": "Jadavpur, West Bengal",
        "notes": "Thesis: Customizing MulVAL for Attack Graph Generation · Network-Simulated Digital Forensics",
        "show_on_cv": True,
    },
    {
        "school": "West Bengal State University (Acharya Prafulla Chandra College)",
        "degree": "B.Sc (Hons) — Computer Science",
        "period": "Jul 2019 — Aug 2022",
        "score": 9.4,
        "scale": "10.0 CGPA",
        "location": "Barasat, West Bengal",
        "notes": "Core CS foundations, competitive programming, Railway Reservation & Socket Transfer",
        "show_on_cv": True,
    },
    {
        "school": "Nebadhai High School (WBCHSE)",
        "degree": "Higher Secondary — Mathematics & Computer Science",
        "period": "Apr 2016 — May 2019",
        "score": 80,
        "scale": "Percentage (%)",
        "location": "West Bengal",
        "notes": "Mathematics, logic, and first formal exposure to programming",
        "show_on_cv": False,
    },
    {
        "school": "Duttapukur Adarsha Vidyapith (WBBSE)",
        "degree": "Secondary — General Studies",
        "period": "Jan 2010 — Jan 2016",
        "score": 75,
        "scale": "Percentage (%)",
        "location": "West Bengal",
        "notes": "Discipline and problem-solving fundamentals",
        "show_on_cv": False,
    },
]

ACHIEVEMENTS_ROWS = [
    ("LeetCode — Easy", "Competitive Programming", "Problems Solved", 72, "Problems", "leetcode.com/Hurley2017 · Foundational DSA"),
    ("LeetCode — Medium", "Competitive Programming", "Problems Solved", 101, "Problems", "Dynamic programming, hash tables, binary search, graphs"),
    ("LeetCode — Hard", "Competitive Programming", "Problems Solved", 5, "Problems", "Complex algorithmic optimization problems"),
    ("HackerRank & Other Judges", "Competitive Programming", "Problems Solved", 25, "Problems", "200+ combined problems solved across LeetCode & HackerRank"),
    ("HackerRank Certifications", "Certifications", "Verified Credentials", 3, "Certs", "Certified in Python (Basic), SQL (Basic), and Java (Basic)"),
    ("LeetCode 50 Days Badge 2022", "Consistency", "Published Solutions", 30, "Solutions", "Consistent daily problem-solving streak · 19.1K profile views"),
    ("Final Year Thesis (MCA)", "Academic Research", "Jadavpur University", 1, "Thesis", "Customizing MulVAL for Attack Graph Generation & Digital Forensics"),
    ("CSI Summer Internship 2023", "Industry Selection", "Celebal Technologies", 1, "Award", "Selected for mentored summer data science internship"),
]

SKILLS_MATRIX = [
    # Data & Analytics (9)
    ("Microsoft Excel", "Data & Analytics", "Expert", 4.5, "Enterprise Reporting", "CIB management reporting, pivot modeling, automation at HSBC"),
    ("VBA", "Data & Analytics", "Advanced", 2.5, "Office Automation", "Custom macros and reporting workflow tools at HSBC"),
    ("Power BI", "Data & Analytics", "Advanced", 2.5, "BI Dashboards", "Executive management reporting & KPI visualization"),
    ("Power Query", "Data & Analytics", "Advanced", 2.5, "ETL & Shaping", "Automated data ingestion and transformation pipelines"),
    ("IBM TM1", "Data & Analytics", "Proficient", 2.0, "OLAP / Planning", "Enterprise financial & management reporting cubes at HSBC"),
    ("Alteryx", "Data & Analytics", "Proficient", 2.0, "Analytics Workflows", "Automated data blending and quality checks at HSBC"),
    ("Pandas", "Data & Analytics", "Advanced", 3.5, "Data Engineering", "Automated Excel/email reporting & ML preprocessing"),
    ("NumPy", "Data & Analytics", "Advanced", 3.5, "Numerical Computing", "Array computation and forecasting pipelines"),
    ("SQL", "Data & Analytics", "Advanced", 4.0, "Querying & Modeling", "HackerRank certified · Enterprise data extraction"),
    # Local AI & Models (6)
    ("LM Studio", "Local AI & Models", "Advanced", 1.5, "Local LLM Inference", "On-device LLM orchestration in Writer pipeline"),
    ("Local LLMs", "Local AI & Models", "Advanced", 1.5, "AI Agents", "Structured-JSON drafting, retries & prose generation"),
    ("Stable Diffusion (SDXL)", "Local AI & Models", "Advanced", 1.5, "Image Generation", "Embedded SDXL pipeline with VRAM CPU offload"),
    ("RealVisXL", "Local AI & Models", "Advanced", 1.0, "Photorealistic Covers", "Automated cover illustration in Writer V4.0"),
    ("Orpheus TTS", "Local AI & Models", "Advanced", 1.0, "Speech Synthesis", "Dual-voice audiobook narration with inline emotion tags"),
    ("Embedded ML", "Local AI & Models", "Proficient", 2.0, "On-Device AI", "Local-first AI workflows without cloud dependency"),
    # Languages (10)
    ("Python", "Languages", "Expert", 6.0, "Primary Language", "Automation, AI agents, Flask backends, OpenCV, ML"),
    ("JavaScript", "Languages", "Advanced", 4.5, "Web & Audio", "React UIs, Web Audio API synthesis, Perspective Pivot"),
    ("Java", "Languages", "Proficient", 4.0, "OOP & DSA", "HackerRank certified · Academic & DSA foundations"),
    ("C", "Languages", "Proficient", 5.0, "Systems & DSA", "Core computer science & memory fundamentals"),
    ("C++", "Languages", "Proficient", 4.5, "Competitive Coding", "Algorithms & data structures problem solving"),
    ("Rust", "Languages", "Proficient", 1.5, "Systems Programming", "Ultra-lightweight Rusty Radio background music player"),
    ("R", "Languages", "Familiar", 1.5, "Statistical Analysis", "Academic statistics & exploratory analysis"),
    ("Bash", "Languages", "Proficient", 3.5, "CLI & Automation", "Linux scripting, tooling, and server administration"),
    ("HTML/CSS", "Languages", "Advanced", 5.0, "Frontend UI", "Responsive editorial layouts and custom design systems"),
    ("OpenEdge ABL", "Languages", "Familiar", 1.0, "Enterprise Systems", "Business logic exposure"),
    # Frameworks & Libraries (9)
    ("React", "Frameworks & Libraries", "Advanced", 2.5, "Frontend SPA", "Portfolio website (www.tusher.in) & interactive web UIs"),
    ("Node.js", "Frameworks & Libraries", "Proficient", 2.5, "Serverless / Backend", "API endpoints and Vite tooling"),
    ("Flask", "Frameworks & Libraries", "Advanced", 3.5, "Python Web Apps", "URL shortener, Railway Reservation, AGG Web"),
    ("Django", "Frameworks & Libraries", "Proficient", 2.0, "Web Backends", "Full-stack web application architecture"),
    ("Express", "Frameworks & Libraries", "Proficient", 2.0, "REST APIs", "Lightweight Node.js backend services"),
    ("scikit-learn", "Frameworks & Libraries", "Advanced", 2.5, "Machine Learning", "Bandwidth forecasting & Email Spam Classifier"),
    ("OpenCV", "Frameworks & Libraries", "Advanced", 3.0, "Computer Vision", "Biometric auth (Phoenix R&D) & HandMouse tracking"),
    ("WebSockets", "Frameworks & Libraries", "Proficient", 2.0, "Real-Time Comms", "Low-latency streaming & interactive dashboards"),
    ("Web Audio API", "Frameworks & Libraries", "Proficient", 1.0, "Procedural Audio", "Physics-based engine sound synthesis in Piston-And-Sound"),
    # Databases (6)
    ("MongoDB", "Databases", "Advanced", 3.5, "NoSQL Documents", "Railway Reservation, URL Shortener, Celebal internships"),
    ("Redis", "Databases", "Proficient", 1.5, "KV / Caching", "Vercel KV / Upstash Redis analytics store on www.tusher.in"),
    ("MySQL", "Databases", "Advanced", 4.0, "Relational DB", "Structured querying, schema design & reporting"),
    ("PostgreSQL", "Databases", "Proficient", 2.5, "Relational DB", "Analytical SQL queries and application storage"),
    ("Data Space", "Databases", "Proficient", 1.5, "Analytics Storage", "Data science forecasting project at Celebal"),
    ("Microsoft Access", "Databases", "Proficient", 2.5, "Desktop DB", "Legacy enterprise data workflows & reporting"),
    # Tools & Platforms (10)
    ("Git", "Tools & Platforms", "Advanced", 5.0, "Version Control", "Daily collaborative engineering workflows"),
    ("GitHub", "Tools & Platforms", "Advanced", 5.0, "Repos & CI", "28+ public repositories (Hurley2017)"),
    ("Vercel", "Tools & Platforms", "Advanced", 2.5, "Cloud Hosting", "Hosting tusher.in, AGG Web, and Email Spam Classifier"),
    ("Linux", "Tools & Platforms", "Advanced", 4.0, "OS & Forensics", "MulVAL thesis environment, tooling & servers"),
    ("Docker", "Tools & Platforms", "Proficient", 2.0, "Containers", "Reproducible dev and deployment environments"),
    ("Postman", "Tools & Platforms", "Advanced", 3.0, "API Testing", "REST endpoint verification & debugging"),
    ("VS Code", "Tools & Platforms", "Expert", 5.0, "Primary IDE", "Python, JS, Rust, and LaTeX development"),
    ("Nmap", "Tools & Platforms", "Proficient", 2.0, "Network Security", "Network-simulated digital forensics & MulVAL thesis"),
    ("Nessus", "Tools & Platforms", "Proficient", 1.5, "Vulnerability Scan", "Attack graph input generation for MulVAL"),
    ("Oval", "Tools & Platforms", "Proficient", 1.5, "Security Modeling", "Vulnerability assessment definitions in thesis"),
]

CV_SKILLS_SUMMARY = [
    ("Data & Analytics", "Microsoft Excel, VBA, Power BI, Power Query, IBM TM1, Alteryx, Pandas, NumPy, SQL"),
    ("Local AI & Models", "LM Studio, Local LLMs, Stable Diffusion (SDXL / RealVisXL), Orpheus TTS, Embedded ML"),
    ("Languages", "Python, JavaScript, Java, C, C++, Rust, SQL, R, Bash, HTML/CSS, OpenEdge ABL"),
    ("Frameworks & Libraries", "React, Node.js, Flask, Django, Express, scikit-learn, OpenCV, WebSockets, Web Audio API"),
    ("Databases & Platforms", "MongoDB, Redis, MySQL, PostgreSQL, Data Space, Access, Git, GitHub, Vercel, Linux, Docker, Nmap"),
    ("Relevant Coursework", "Computer Networks, Machine Learning, Data Structures & Algorithms, Advanced Python & Java"),
]


# ---------------------------------------------------------------------------
# Sheet Builders
# ---------------------------------------------------------------------------
def build_education_skills_sheet(wb: openpyxl.Workbook) -> None:
    ws = wb.create_sheet("Education & Skills")
    ws.sheet_properties.tabColor = "475569"
    ws.views.sheetView[0].showGridLines = True

    widths = {1: 5.0, 2: 32.0, 3: 32.0, 4: 20.0, 5: 11.0, 6: 16.0, 7: 48.0}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    # Banner
    ws.merge_cells("A1:G1")
    ws["A1"] = "EDUCATION, ACHIEVEMENTS & TECHNICAL SKILLS MATRIX"
    ws["A1"].font = mk_font(15, bold=True, color=HEADER_FG, family=FONT_DISPLAY)
    ws["A1"].fill = mk_fill(HEADER_BG)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:G2")
    ws["A2"] = "Structured academic credentials, competitive programming metrics, and filterable 50-skill technical inventory"
    ws["A2"].font = mk_font(9.5, italic=True, color=MUTED)
    ws["A2"].fill = mk_fill(CARD_BG)
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 8

    # Section 1: Academic Record (Rows 4-9)
    ws.merge_cells("A4:G4")
    ws["A4"] = "1. ACADEMIC RECORD"
    ws["A4"].font = mk_font(10.5, bold=True, color=INK, family=FONT_DISPLAY)
    ws["A4"].fill = mk_fill(ACCENT_SOFT)
    ws["A4"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    apply_box_border(ws, 1, 4, 7, 4, style="thin", color=BORDER_MED)
    ws.row_dimensions[4].height = 22

    edu_headers = ["#", "Institution", "Degree / Program", "Period", "Score", "Scale / Unit", "Location & Highlights"]
    for idx, h in enumerate(edu_headers, 1):
        c = ws.cell(row=5, column=idx, value=h)
        c.font = mk_font(9.5, bold=True, color=HEADER_FG)
        c.fill = mk_fill("334155")
        c.alignment = Alignment(horizontal="center" if idx in (1, 5, 6) else "left", vertical="center", indent=0 if idx in (1, 5, 6) else 1)
    ws.row_dimensions[5].height = 21

    for idx, edu in enumerate(EDUCATION, 1):
        r = 5 + idx  # Rows 6..9
        vals = [idx, edu["school"], edu["degree"], edu["period"], edu["score"], edu["scale"], f"{edu['location']} · {edu['notes']}"]
        bg = ZEBRA_BG if idx % 2 == 0 else WHITE
        for c_idx, val in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c_idx, value=val)
            cell.font = mk_font(9.5, bold=(c_idx == 2), color=INK)
            cell.fill = mk_fill(bg)
            cell.border = mk_border(bottom=("thin", BORDER_LIGHT), left=("thin", BORDER_LIGHT), right=("thin", BORDER_LIGHT))
            horiz = "center" if c_idx in (1, 5, 6) else "left"
            cell.alignment = Alignment(horizontal=horiz, vertical="center", wrap_text=True)
            if c_idx == 5:
                cell.number_format = "0.0" if isinstance(val, float) else "0"
        lines = max(
            estimate_wrapped_lines(edu["school"], widths[2], size_pt=9.5, bold=True),
            estimate_wrapped_lines(edu["degree"], widths[3], size_pt=9.5),
            estimate_wrapped_lines(vals[6], widths[7], size_pt=9.5),
        )
        ws.row_dimensions[r].height = calc_row_height(lines, 9.5, 6)

    ws.row_dimensions[10].height = 10

    # Section 2: Competitive Programming & Achievements (Rows 11-20)
    ws.merge_cells("A11:G11")
    ws["A11"] = "2. COMPETITIVE PROGRAMMING, CERTIFICATIONS & RESEARCH"
    ws["A11"].font = mk_font(10.5, bold=True, color=INK, family=FONT_DISPLAY)
    ws["A11"].fill = mk_fill(ACCENT_SOFT)
    ws["A11"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    apply_box_border(ws, 1, 11, 7, 11, style="thin", color=BORDER_MED)
    ws.row_dimensions[11].height = 22

    ach_headers = ["#", "Achievement / Platform", "Category", "Metric", "Count", "Unit", "Details"]
    for idx, h in enumerate(ach_headers, 1):
        c = ws.cell(row=12, column=idx, value=h)
        c.font = mk_font(9.5, bold=True, color=HEADER_FG)
        c.fill = mk_fill("334155")
        c.alignment = Alignment(horizontal="center" if idx in (1, 5, 6) else "left", vertical="center", indent=0 if idx in (1, 5, 6) else 1)
    ws.row_dimensions[12].height = 21

    for idx, row_data in enumerate(ACHIEVEMENTS_ROWS, 1):
        r = 12 + idx  # Rows 13..20
        vals = [idx, *row_data]
        bg = ZEBRA_BG if idx % 2 == 0 else WHITE
        for c_idx, val in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c_idx, value=val)
            cell.font = mk_font(9.5, bold=(c_idx == 2), color=INK)
            cell.fill = mk_fill(bg)
            cell.border = mk_border(bottom=("thin", BORDER_LIGHT), left=("thin", BORDER_LIGHT), right=("thin", BORDER_LIGHT))
            horiz = "center" if c_idx in (1, 5, 6) else "left"
            cell.alignment = Alignment(horizontal=horiz, vertical="center", wrap_text=True)
            if c_idx == 5:
                cell.number_format = "#,##0"
        ws.row_dimensions[r].height = 20

    ws.row_dimensions[21].height = 10

    # Section 3: Technical Skills Matrix (Rows 22..73, AutoFilter A23:G73)
    ws.merge_cells("A22:G22")
    ws["A22"] = "3. TECHNICAL SKILLS INVENTORY (FILTERABLE)"
    ws["A22"].font = mk_font(10.5, bold=True, color=INK, family=FONT_DISPLAY)
    ws["A22"].fill = mk_fill(ACCENT_SOFT)
    ws["A22"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    apply_box_border(ws, 1, 22, 7, 22, style="thin", color=BORDER_MED)
    ws.row_dimensions[22].height = 22

    skill_headers = ["#", "Skill / Technology", "Domain", "Proficiency", "Years", "Primary Usage", "Applied Context"]
    for idx, h in enumerate(skill_headers, 1):
        c = ws.cell(row=23, column=idx, value=h)
        c.font = mk_font(9.5, bold=True, color=HEADER_FG)
        c.fill = mk_fill(HEADER_BG)
        c.alignment = Alignment(horizontal="center" if idx in (1, 4, 5) else "left", vertical="center", indent=0 if idx in (1, 4, 5) else 1)
    ws.row_dimensions[23].height = 22

    for idx, (skill, domain, prof, yrs, usage, ctx) in enumerate(SKILLS_MATRIX, 1):
        r = 23 + idx  # Rows 24..73
        vals = [idx, skill, domain, prof, yrs, usage, ctx]
        bg = ZEBRA_BG if idx % 2 == 0 else WHITE
        for c_idx, val in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c_idx, value=val)
            cell.font = mk_font(9.5, bold=(c_idx == 2), color=ACCENT if (c_idx == 4 and val == "Expert") else INK)
            cell.fill = mk_fill(bg)
            cell.border = mk_border(bottom=("thin", BORDER_LIGHT), left=("thin", BORDER_LIGHT), right=("thin", BORDER_LIGHT))
            horiz = "center" if c_idx in (1, 4, 5) else "left"
            cell.alignment = Alignment(horizontal=horiz, vertical="center", wrap_text=True)
            if c_idx == 5:
                cell.number_format = "0.0"
        lines = max(
            estimate_wrapped_lines(skill, widths[2], size_pt=9.5, bold=True),
            estimate_wrapped_lines(ctx, widths[7], size_pt=9.5),
        )
        ws.row_dimensions[r].height = calc_row_height(lines, 9.5, 5)

    # Summary Row 74
    ws.cell(row=74, column=1, value="").fill = mk_fill(CARD_BG)
    ws.cell(row=74, column=2, value='=COUNTA(B24:B73)&" Technical Skills"').font = mk_font(9.5, bold=True, color=INK)
    ws.cell(row=74, column=3, value="Across 6 Core Domains").font = mk_font(9.5, italic=True, color=MUTED)
    ws.cell(row=74, column=4, value='=COUNTIF(D24:D73,"Expert")&" Expert / "&COUNTIF(D24:D73,"Advanced")&" Adv"').font = mk_font(9.5, bold=True, color=ACCENT)
    ws.cell(row=74, column=4).alignment = Alignment(horizontal="center", vertical="center")
    avg_yrs = ws.cell(row=74, column=5, value="=AVERAGE(E24:E73)")
    avg_yrs.font = mk_font(9.5, bold=True, color=INK)
    avg_yrs.number_format = '0.0" avg"'
    avg_yrs.alignment = Alignment(horizontal="center", vertical="center")
    style_range(ws, 1, 74, 7, 74, fill=mk_fill(CARD_BG), border=mk_border(top=("medium", BORDER_DARK), bottom=("medium", BORDER_DARK)))
    ws.row_dimensions[74].height = 22

    # Enable AutoFilter on Skills Inventory
    ws.auto_filter.ref = "A23:G73"

    # Print setup (Landscape, break cleanly before Technical Skills Inventory)
    ws.row_breaks.append(Break(id=21))
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def build_experience_projects_sheet(wb: openpyxl.Workbook) -> None:
    ws = wb.create_sheet("Experience & Projects")
    ws.sheet_properties.tabColor = "334155"
    ws.views.sheetView[0].showGridLines = True

    widths = {1: 4.5, 2: 24.0, 3: 30.0, 4: 15.0, 5: 15.0, 6: 13.0, 7: 15.0, 8: 12.0, 9: 46.0, 10: 28.0}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    # Banner
    ws.merge_cells("A1:J1")
    ws["A1"] = "PROFESSIONAL EXPERIENCE & ENGINEERING PROJECTS"
    ws["A1"].font = mk_font(15, bold=True, color=HEADER_FG, family=FONT_DISPLAY)
    ws["A1"].fill = mk_fill(HEADER_BG)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:J2")
    ws["A2"] = "Full career timeline with live dynamic tenure formulas and complete 9-project engineering portfolio"
    ws["A2"].font = mk_font(9.5, italic=True, color=MUTED)
    ws["A2"].fill = mk_fill(CARD_BG)
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 8

    # Table 1: Experience (Rows 4..11)
    ws.merge_cells("A4:J4")
    ws["A4"] = "1. PROFESSIONAL EXPERIENCE (DYNAMIC TENURE)"
    ws["A4"].font = mk_font(10.5, bold=True, color=INK, family=FONT_DISPLAY)
    ws["A4"].fill = mk_fill(ACCENT_SOFT)
    ws["A4"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    apply_box_border(ws, 1, 4, 10, 4, style="thin", color=BORDER_MED)
    ws.row_dimensions[4].height = 22

    exp_headers = ["#", "Organization", "Role", "Type", "Location", "Start", "End", "Months", "Key Contributions & Impact", "Tech Stack"]
    for idx, h in enumerate(exp_headers, 1):
        c = ws.cell(row=5, column=idx, value=h)
        c.font = mk_font(9.5, bold=True, color=HEADER_FG)
        c.fill = mk_fill("334155")
        c.alignment = Alignment(horizontal="center" if idx in (1, 4, 6, 7, 8) else "left", vertical="center", indent=0 if idx in (1, 4, 6, 7, 8) else 1)
    ws.row_dimensions[5].height = 22

    for idx, job in enumerate(EXPERIENCE, 1):
        r = 5 + idx  # Rows 6..10
        impact = " • ".join(job["bullets"])
        end_val = job["end"]
        months_formula = f"=ROUND((IF(ISNUMBER(G{r}),G{r},TODAY())-F{r})/30.4375,1)"
        vals = [idx, job["company"], job["role"], job["type"], job["location"], job["start"], end_val, months_formula, impact, job["tech"]]
        bg = ZEBRA_BG if idx % 2 == 0 else WHITE
        for c_idx, val in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c_idx, value=val)
            cell.font = mk_font(9.5, bold=(c_idx in (2, 8)), color=ACCENT if (c_idx == 7 and val == "Present") else INK)
            cell.fill = mk_fill(bg)
            cell.border = mk_border(bottom=("thin", BORDER_LIGHT), left=("thin", BORDER_LIGHT), right=("thin", BORDER_LIGHT))
            horiz = "center" if c_idx in (1, 4, 6, 7, 8) else "left"
            cell.alignment = Alignment(horizontal=horiz, vertical="center", wrap_text=True)
            if c_idx == 6 or (c_idx == 7 and isinstance(val, dt.date)):
                cell.number_format = "mmm yyyy"
            elif c_idx == 8:
                cell.number_format = "0.0"
        lines = max(
            estimate_wrapped_lines(job["company"], widths[2], size_pt=9.5, bold=True),
            estimate_wrapped_lines(job["role"], widths[3], size_pt=9.5),
            estimate_wrapped_lines(impact, widths[9], size_pt=9.5),
            estimate_wrapped_lines(job["tech"], widths[10], size_pt=9.5),
        )
        ws.row_dimensions[r].height = calc_row_height(lines, 9.5, 13)

    # Row 11: Experience Totals
    ws.cell(row=11, column=2, value="Total Experience").font = mk_font(9.5, bold=True, color=INK)
    ws.cell(row=11, column=3, value='=COUNTA(B6:B10)&" Roles & Internships"').font = mk_font(9.5, bold=True, color=INK)
    ws.cell(row=11, column=4, value='=COUNTIF(D6:D10,"Full-Time")&" FT / "&COUNTIF(D6:D10,"Internship")&" Intern"').font = mk_font(9.5, bold=True, color=MUTED)
    ws.cell(row=11, column=4).alignment = Alignment(horizontal="center", vertical="center")
    tot_mo = ws.cell(row=11, column=8, value="=SUM(H6:H10)")
    tot_mo.font = mk_font(9.5, bold=True, color=ACCENT)
    tot_mo.number_format = '0.0" mo"'
    tot_mo.alignment = Alignment(horizontal="center", vertical="center")
    c_note = ws.cell(row=11, column=9, value='="Including "&ROUND((H6+H7)/12,1)&" yrs combined at HSBC (Intern → Manager)"')
    c_note.font = mk_font(9.5, italic=True, color=MUTED)
    c_note.alignment = Alignment(horizontal="left", vertical="center")
    style_range(ws, 1, 11, 10, 11, fill=mk_fill(CARD_BG), border=mk_border(top=("medium", BORDER_DARK), bottom=("medium", BORDER_DARK)))
    ws.row_dimensions[11].height = 24

    ws.row_dimensions[12].height = 12

    # Table 2: Projects (Rows 13..23, AutoFilter A14:J23)
    ws.merge_cells("A13:J13")
    ws["A13"] = "2. ENGINEERING & AI PROJECTS PORTFOLIO (FILTERABLE)"
    ws["A13"].font = mk_font(10.5, bold=True, color=INK, family=FONT_DISPLAY)
    ws["A13"].fill = mk_fill(ACCENT_SOFT)
    ws["A13"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    apply_box_border(ws, 1, 13, 10, 13, style="thin", color=BORDER_MED)
    ws.row_dimensions[13].height = 22

    proj_headers = ["#", "Project Name", "Tagline", "Domain", "Year", "Lead Tech", "GitHub Repo", "Live Demo", "Highlights & Architecture", "Stack Used"]
    for idx, h in enumerate(proj_headers, 1):
        c = ws.cell(row=14, column=idx, value=h)
        c.font = mk_font(9.5, bold=True, color=HEADER_FG)
        c.fill = mk_fill(HEADER_BG)
        c.alignment = Alignment(horizontal="center" if idx in (1, 4, 5, 6, 7, 8) else "left", vertical="center", indent=0 if idx in (1, 4, 5, 6, 7, 8) else 1)
    ws.row_dimensions[14].height = 22

    for idx, proj in enumerate(PROJECTS, 1):
        r = 14 + idx  # Rows 15..23
        desc = " • ".join(proj["bullets"])
        vals = [
            idx,
            proj["name"],
            proj["tagline"],
            proj["domain"],
            proj["year"],
            proj["lead_tech"],
            "View Repo ↗",
            "Live App ↗" if proj["demo"] else "—",
            desc,
            proj["stack"],
        ]
        bg = ZEBRA_BG if idx % 2 == 0 else WHITE
        for c_idx, val in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c_idx, value=val)
            is_link = (c_idx == 7) or (c_idx == 8 and proj["demo"])
            cell.font = mk_font(
                9.5,
                bold=(c_idx == 2),
                color=ACCENT if is_link else INK,
                underline="single" if is_link else None,
            )
            cell.fill = mk_fill(bg)
            cell.border = mk_border(bottom=("thin", BORDER_LIGHT), left=("thin", BORDER_LIGHT), right=("thin", BORDER_LIGHT))
            horiz = "center" if c_idx in (1, 4, 5, 6, 7, 8) else "left"
            cell.alignment = Alignment(horizontal=horiz, vertical="center", wrap_text=True)
            if c_idx == 5:
                cell.number_format = "0"
            if c_idx == 7:
                cell.hyperlink = proj["github"]
            elif c_idx == 8 and proj["demo"]:
                cell.hyperlink = proj["demo"]
        lines = max(
            estimate_wrapped_lines(proj["name"], widths[2], size_pt=9.5, bold=True),
            estimate_wrapped_lines(proj["tagline"], widths[3], size_pt=9.5),
            estimate_wrapped_lines(desc, widths[9], size_pt=9.5),
            estimate_wrapped_lines(proj["stack"], widths[10], size_pt=9.5),
        )
        ws.row_dimensions[r].height = calc_row_height(lines, 9.5, 13)

    # Row 24: Projects Summary
    ws.cell(row=24, column=2, value='=COUNTA(B15:B23)&" Projects Shipped"').font = mk_font(9.5, bold=True, color=INK)
    ws.cell(row=24, column=3, value='=COUNTIF(E15:E23,2026)&" shipped in 2026"').font = mk_font(9.5, bold=True, color=ACCENT)
    ws.cell(row=24, column=6, value='=COUNTIF(F15:F23,"Python")&" Python"').font = mk_font(9.5, bold=True, color=MUTED)
    ws.cell(row=24, column=6).alignment = Alignment(horizontal="center", vertical="center")
    style_range(ws, 1, 24, 10, 24, fill=mk_fill(CARD_BG), border=mk_border(top=("medium", BORDER_DARK), bottom=("medium", BORDER_DARK)))
    ws.row_dimensions[24].height = 22

    ws.auto_filter.ref = "A14:J23"
    ws.freeze_panes = "A6"

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1


def build_dashboard_sheet(wb: openpyxl.Workbook) -> None:
    ws = wb.create_sheet("Career Dashboard")
    ws.sheet_properties.tabColor = ACCENT
    ws.views.sheetView[0].showGridLines = False

    # 3-zone balanced grid: B:D (Zone 1), E (gutter), F:H (Zone 2), I (gutter), J:L (Zone 3)
    widths = {
        1: 2.5,
        2: 24.0, 3: 11.0, 4: 11.0,
        5: 2.5,
        6: 24.0, 7: 11.0, 8: 11.0,
        9: 2.5,
        10: 20.0, 11: 12.0, 12: 12.0,
        13: 2.5,
    }
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.row_dimensions[1].height = 8

    # Header Banner B2:L3
    ws.merge_cells("B2:L2")
    ws["B2"] = "TUSHER MONDAL — CAREER & SKILLS ANALYTICS"
    ws["B2"].font = mk_font(18, bold=True, color=INK, family=FONT_DISPLAY)
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 28

    ws.merge_cells("B3:L3")
    ws["B3"] = '="Executive management-reporting view powered by live cross-sheet formulas  ·  Snapshot as of "&TEXT(TODAY(),"dd mmm yyyy")'
    ws["B3"].font = mk_font(10, italic=True, color=MUTED)
    ws["B3"].alignment = Alignment(horizontal="left", vertical="center")
    style_range(ws, 2, 3, 12, 3, border=mk_border(bottom=("medium", BORDER_DARK)))
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 10

    # Helper to draw a crisp 2-row KPI card spanning 3 columns
    def draw_kpi(col_start: int, row_top: int, label: str, formula: str, num_fmt: str = "0", sub_formula: str | None = None):
        c1, c2 = col_start, col_start + 2
        ws.merge_cells(start_row=row_top, start_column=c1, end_row=row_top, end_column=c2)
        ws.merge_cells(start_row=row_top + 1, start_column=c1, end_row=row_top + 1, end_column=c2)
        lbl_cell = ws.cell(row=row_top, column=c1, value=label)
        lbl_cell.font = mk_font(8.5, bold=True, color=MUTED)
        lbl_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        val_cell = ws.cell(row=row_top + 1, column=c1, value=formula)
        val_cell.font = mk_font(20, bold=True, color=INK, family=FONT_DISPLAY)
        val_cell.number_format = num_fmt
        val_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        style_range(ws, c1, row_top, c2, row_top + 1, fill=mk_fill(CARD_BG))
        apply_box_border(ws, c1, row_top, c2, row_top + 1, style="thin", color=BORDER_MED, top_accent=("medium", ACCENT))
        if sub_formula:
            ws.merge_cells(start_row=row_top + 2, start_column=c1, end_row=row_top + 2, end_column=c2)
            sc = ws.cell(row=row_top + 2, column=c1, value=sub_formula)
            sc.font = mk_font(8.5, italic=True, color=ACCENT)
            sc.fill = mk_fill(WHITE)
            sc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            apply_box_border(ws, c1, row_top + 2, c2, row_top + 2, style="thin", color=BORDER_LIGHT)

    ws.row_dimensions[5].height = 18
    ws.row_dimensions[6].height = 30
    ws.row_dimensions[7].height = 18
    draw_kpi(2, 5, "EXPERIENCE TENURE (MONTHS)", "='Experience & Projects'!H11", '0.0" mos"', '="HSBC: "&ROUND((\'Experience & Projects\'!H6+\'Experience & Projects\'!H7)/12,1)&" yrs (Intern → Manager)"')
    draw_kpi(6, 5, "ROLES & INTERNSHIPS", "=COUNTA('Experience & Projects'!B6:B10)", "0", '="1 Full-Time · "&COUNTIF(\'Experience & Projects\'!D6:D10,"Internship")&" Internships"')
    draw_kpi(10, 5, "ENGINEERING & AI PROJECTS", "=COUNTA('Experience & Projects'!B15:B23)", "0", '="2026 releases: "&COUNTIF(\'Experience & Projects\'!E15:E23,2026)&" · 28+ GitHub repos"')

    ws.row_dimensions[8].height = 10
    ws.row_dimensions[9].height = 18
    ws.row_dimensions[10].height = 30
    ws.row_dimensions[11].height = 18
    draw_kpi(2, 9, "TECHNICAL SKILLS TRACKED", "=COUNTA('Education & Skills'!B24:B73)", "0", '="Expert: "&COUNTIF(\'Education & Skills\'!D24:D73,"Expert")&" · Advanced: "&COUNTIF(\'Education & Skills\'!D24:D73,"Advanced")')
    draw_kpi(6, 9, "DSA PROBLEMS SOLVED", "=SUM('Education & Skills'!E13:E16)", '#,##0"+"', '="LeetCode: "&SUM(\'Education & Skills\'!E13:E15)&" · 50 Days Badge"')
    draw_kpi(10, 9, "HIGHER-ED AVG CGPA", "=AVERAGE('Education & Skills'!E6:E7)", '0.00" / 10"', '="MCA: "&\'Education & Skills\'!E6&" · B.Sc (Hons): "&\'Education & Skills\'!E7')

    ws.row_dimensions[12].height = 12

    # Breakdown Tables Row 13..21
    # Table A (B13:D21): Skills by Domain
    ws.merge_cells("B13:D13")
    ws["B13"] = "TECHNICAL SKILLS BY DOMAIN"
    ws["B13"].font = mk_font(10, bold=True, color=HEADER_FG, family=FONT_DISPLAY)
    ws["B13"].fill = mk_fill(HEADER_BG)
    ws["B13"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    for col, txt, align in [(2, "Domain", "left"), (3, "Skills", "center"), (4, "Avg Yrs", "center")]:
        c = ws.cell(row=14, column=col, value=txt)
        c.font = mk_font(9, bold=True, color=INK)
        c.fill = mk_fill(CARD_BG)
        c.border = mk_border(bottom=("thin", BORDER_DARK))
        c.alignment = Alignment(horizontal=align, vertical="center", indent=1 if align == "left" else 0)

    domains = [
        "Data & Analytics",
        "Local AI & Models",
        "Languages",
        "Frameworks & Libraries",
        "Databases",
        "Tools & Platforms",
    ]
    for i, dom in enumerate(domains):
        r = 15 + i  # 15..20
        bg = ZEBRA_BG if i % 2 == 1 else WHITE
        c_dom = ws.cell(row=r, column=2, value=dom)
        c_cnt = ws.cell(row=r, column=3, value=f"=COUNTIF('Education & Skills'!$C$24:$C$73, B{r})")
        c_yrs = ws.cell(row=r, column=4, value=f"=AVERAGEIF('Education & Skills'!$C$24:$C$73, B{r}, 'Education & Skills'!$E$24:$E$73)")
        c_dom.font = mk_font(9.5, color=INK)
        c_cnt.font = mk_font(9.5, bold=True, color=INK)
        c_yrs.font = mk_font(9.5, color=MUTED)
        c_cnt.number_format = "0"
        c_yrs.number_format = "0.0"
        c_dom.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c_cnt.alignment = Alignment(horizontal="center", vertical="center")
        c_yrs.alignment = Alignment(horizontal="center", vertical="center")
        style_range(ws, 2, r, 4, r, fill=mk_fill(bg), border=mk_border(bottom=("thin", BORDER_LIGHT)))

    ws.cell(row=21, column=2, value="Total / Overall").font = mk_font(9.5, bold=True, color=INK)
    ws.cell(row=21, column=2).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.cell(row=21, column=3, value="=SUM(C15:C20)").font = mk_font(9.5, bold=True, color=ACCENT)
    ws.cell(row=21, column=3).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=21, column=4, value="=AVERAGE('Education & Skills'!E24:E73)").font = mk_font(9.5, bold=True, color=INK)
    ws.cell(row=21, column=4).number_format = "0.0"
    ws.cell(row=21, column=4).alignment = Alignment(horizontal="center", vertical="center")
    style_range(ws, 2, 21, 4, 21, fill=mk_fill(ACCENT_SOFT), border=mk_border(top=("thin", BORDER_DARK), bottom=("medium", BORDER_DARK)))

    # Table B (F13:H21): Competitive Coding & Certs
    ws.merge_cells("F13:H13")
    ws["F13"] = "PROBLEM SOLVING & CREDENTIALS"
    ws["F13"].font = mk_font(10, bold=True, color=HEADER_FG, family=FONT_DISPLAY)
    ws["F13"].fill = mk_fill(HEADER_BG)
    ws["F13"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    for col, txt, align in [(6, "Category", "left"), (7, "Count", "center"), (8, "LC Share", "center")]:
        c = ws.cell(row=14, column=col, value=txt)
        c.font = mk_font(9, bold=True, color=INK)
        c.fill = mk_fill(CARD_BG)
        c.border = mk_border(bottom=("thin", BORDER_DARK))
        c.alignment = Alignment(horizontal=align, vertical="center", indent=1 if align == "left" else 0)

    cp_rows = [
        ("LeetCode — Easy", "='Education & Skills'!E13", "=G15/$G$18"),
        ("LeetCode — Medium", "='Education & Skills'!E14", "=G16/$G$18"),
        ("LeetCode — Hard", "='Education & Skills'!E15", "=G17/$G$18"),
        ("LeetCode Subtotal", "=SUM(G15:G17)", "=SUM(H15:H17)"),
        ("HackerRank & Other", "='Education & Skills'!E16", "—"),
        ("HackerRank Certs", "='Education & Skills'!E17", "Py/SQL/Java"),
    ]
    for i, (lbl, f_cnt, f_shr) in enumerate(cp_rows):
        r = 15 + i  # 15..20
        is_sub = lbl == "LeetCode Subtotal"
        bg = CARD_BG if is_sub else (ZEBRA_BG if i % 2 == 1 else WHITE)
        c1 = ws.cell(row=r, column=6, value=lbl)
        c2 = ws.cell(row=r, column=7, value=f_cnt)
        c3 = ws.cell(row=r, column=8, value=f_shr)
        c1.font = mk_font(9.5, bold=is_sub, color=INK)
        c2.font = mk_font(9.5, bold=True, color=ACCENT if is_sub else INK)
        c3.font = mk_font(9.5, bold=is_sub, color=MUTED)
        c1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c2.number_format = "#,##0"
        if isinstance(f_shr, str) and f_shr.startswith("="):
            c3.number_format = "0.0%"
        bdr = mk_border(top=("thin", BORDER_MED), bottom=("thin", BORDER_MED)) if is_sub else mk_border(bottom=("thin", BORDER_LIGHT))
        style_range(ws, 6, r, 8, r, fill=mk_fill(bg), border=bdr)

    ws.cell(row=21, column=6, value="Combined Problems").font = mk_font(9.5, bold=True, color=INK)
    ws.cell(row=21, column=6).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.cell(row=21, column=7, value="=G18+G19").font = mk_font(9.5, bold=True, color=ACCENT)
    ws.cell(row=21, column=7).number_format = '#,##0"+"'
    ws.cell(row=21, column=7).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=21, column=8, value="19.1K views").font = mk_font(9, italic=True, color=MUTED)
    ws.cell(row=21, column=8).alignment = Alignment(horizontal="center", vertical="center")
    style_range(ws, 6, 21, 8, 21, fill=mk_fill(ACCENT_SOFT), border=mk_border(top=("thin", BORDER_DARK), bottom=("medium", BORDER_DARK)))

    # Table C (J13:L21): Projects by Year
    ws.merge_cells("J13:L13")
    ws["J13"] = "PROJECTS BY RELEASE YEAR"
    ws["J13"].font = mk_font(10, bold=True, color=HEADER_FG, family=FONT_DISPLAY)
    ws["J13"].fill = mk_fill(HEADER_BG)
    ws["J13"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    for col, txt, align in [(10, "Year", "center"), (11, "Projects", "center"), (12, "Share", "center")]:
        c = ws.cell(row=14, column=col, value=txt)
        c.font = mk_font(9, bold=True, color=INK)
        c.fill = mk_fill(CARD_BG)
        c.border = mk_border(bottom=("thin", BORDER_DARK))
        c.alignment = Alignment(horizontal=align, vertical="center")

    years = [2026, 2025, 2024, 2023, 2022]
    for i, yr in enumerate(years):
        r = 15 + i  # 15..19
        bg = ZEBRA_BG if i % 2 == 1 else WHITE
        c_yr = ws.cell(row=r, column=10, value=yr)
        c_cnt = ws.cell(row=r, column=11, value=f"=COUNTIF('Experience & Projects'!$E$15:$E$23, J{r})")
        c_shr = ws.cell(row=r, column=12, value=f"=K{r}/$K$21")
        c_yr.font = mk_font(9.5, bold=(yr == 2026), color=INK)
        c_cnt.font = mk_font(9.5, bold=True, color=INK)
        c_shr.font = mk_font(9.5, color=MUTED)
        c_yr.alignment = Alignment(horizontal="center", vertical="center")
        c_cnt.alignment = Alignment(horizontal="center", vertical="center")
        c_shr.alignment = Alignment(horizontal="center", vertical="center")
        c_yr.number_format = "0"
        c_cnt.number_format = "0"
        c_shr.number_format = "0.0%"
        style_range(ws, 10, r, 12, r, fill=mk_fill(bg), border=mk_border(bottom=("thin", BORDER_LIGHT)))

    # Spacer row 20 in Table C
    ws.cell(row=20, column=10, value="Python Projects").font = mk_font(9, italic=True, color=MUTED)
    ws.cell(row=20, column=10).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=20, column=11, value='=COUNTIF(\'Experience & Projects\'!$F$15:$F$23,"Python")').font = mk_font(9.5, bold=True, color=INK)
    ws.cell(row=20, column=11).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=20, column=12, value="=K20/$K$21").font = mk_font(9, color=MUTED)
    ws.cell(row=20, column=12).number_format = "0.0%"
    ws.cell(row=20, column=12).alignment = Alignment(horizontal="center", vertical="center")
    style_range(ws, 10, 20, 12, 20, fill=mk_fill(ZEBRA_BG), border=mk_border(bottom=("thin", BORDER_LIGHT)))

    ws.cell(row=21, column=10, value="Total Projects").font = mk_font(9.5, bold=True, color=INK)
    ws.cell(row=21, column=10).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=21, column=11, value="=SUM(K15:K19)").font = mk_font(9.5, bold=True, color=ACCENT)
    ws.cell(row=21, column=11).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=21, column=12, value="=SUM(L15:L19)").font = mk_font(9.5, bold=True, color=INK)
    ws.cell(row=21, column=12).number_format = "0%"
    ws.cell(row=21, column=12).alignment = Alignment(horizontal="center", vertical="center")
    style_range(ws, 10, 21, 12, 21, fill=mk_fill(ACCENT_SOFT), border=mk_border(top=("thin", BORDER_DARK), bottom=("medium", BORDER_DARK)))

    for r in range(13, 22):
        ws.row_dimensions[r].height = 22 if r == 13 else 20

    ws.row_dimensions[22].height = 12

    # Native Excel Charts (Rows 23..38)
    # Zone widths: B:F ~ 75 chars (~14.5cm), H:L ~ 70 chars (~13.8cm) with gutter G/I
    # Chart 1: Horizontal Bar Chart — Skills by Domain
    chart1 = BarChart()
    chart1.type = "bar"
    chart1.style = 10
    chart1.title = "Technical Skills Distribution by Domain"
    chart1.legend = None
    data1 = Reference(ws, min_col=3, min_row=14, max_row=20)
    cats1 = Reference(ws, min_col=2, min_row=15, max_row=20)
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats1)
    chart1.series[0].graphicalProperties.solidFill = ACCENT
    chart1.series[0].graphicalProperties.line.solidFill = ACCENT
    chart1.dataLabels = DataLabelList(showVal=True, showCatName=False, showSerName=False, showPercent=False, showLegendKey=False)
    chart1.x_axis.delete = False
    chart1.y_axis.delete = False
    chart1.y_axis.scaling.min = 0
    chart1.y_axis.scaling.max = 12
    chart1.y_axis.majorUnit = 2
    chart1.y_axis.number_format = "0"
    chart1.y_axis.majorGridlines = None
    chart1.width = 13.8
    chart1.height = 7.4
    ws.add_chart(chart1, "B23")

    # Chart 2: Vertical Column Chart — Projects by Year
    chart2 = BarChart()
    chart2.type = "col"
    chart2.style = 10
    chart2.title = "Engineering & AI Projects Shipped by Year"
    chart2.legend = None
    data2 = Reference(ws, min_col=11, min_row=14, max_row=19)
    cats2 = Reference(ws, min_col=10, min_row=15, max_row=19)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.series[0].graphicalProperties.solidFill = INK
    chart2.series[0].graphicalProperties.line.solidFill = INK
    chart2.dataLabels = DataLabelList(showVal=True, showCatName=False, showSerName=False, showPercent=False, showLegendKey=False)
    chart2.x_axis.delete = False
    chart2.y_axis.delete = False
    chart2.y_axis.scaling.min = 0
    chart2.y_axis.scaling.max = 5
    chart2.y_axis.majorUnit = 1
    chart2.y_axis.number_format = "0"
    chart2.y_axis.majorGridlines = None
    chart2.width = 13.8
    chart2.height = 7.4
    ws.add_chart(chart2, "G23")

    # Footer strip well below charts (Row 40)
    for spacer_r in range(23, 40):
        ws.row_dimensions[spacer_r].height = 15
    ws.row_dimensions[39].height = 10
    ws.merge_cells("B40:L40")
    ws["B40"] = '="Academic Progression:  WBBSE "&\'Education & Skills\'!E9&"%  →  WBCHSE "&\'Education & Skills\'!E8&"%  →  B.Sc (Hons) "&\'Education & Skills\'!E7&" CGPA  →  MCA "&\'Education & Skills\'!E6&" CGPA"'
    ws["B40"].font = mk_font(9.5, bold=True, color=INK)
    ws["B40"].fill = mk_fill(CARD_BG)
    ws["B40"].alignment = Alignment(horizontal="center", vertical="center")
    apply_box_border(ws, 2, 40, 12, 40, style="thin", color=BORDER_MED)
    ws.row_dimensions[40].height = 24

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1


def build_cv_sheet(wb: openpyxl.Workbook) -> None:
    ws = wb.active
    ws.title = "CV"
    ws.sheet_properties.tabColor = INK
    ws.views.sheetView[0].showGridLines = False

    # Column layout designed for clean A4/Letter Portrait print
    # A: margin, B: bullet, C: label/contact, D: body, E: body, F: right dates/loc, G: margin
    widths = {1: 2.2, 2: 2.8, 3: 22.0, 4: 26.0, 5: 23.0, 6: 24.0, 7: 2.2}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    body_total_w = widths[3] + widths[4] + widths[5] + widths[6]  # C..F merged width (95 chars)
    left_header_w = widths[2] + widths[3] + widths[4] + widths[5]  # B..E merged width (73.8 chars)

    r = 1
    ws.row_dimensions[r].height = 6
    r += 1

    # --- Header Block ---
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    c_name = ws.cell(row=r, column=2, value="TUSHER MONDAL")
    c_name.font = mk_font(22, bold=True, color=INK, family=FONT_DISPLAY)
    c_name.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 30
    r += 1

    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    c_sub = ws.cell(
        row=r,
        column=2,
        value=f"{PROFILE['role']} ({PROFILE['company']})  ·  {PROFILE['location']}  ·  {PROFILE['phone']}",
    )
    c_sub.font = mk_font(9.5, color=MUTED)
    c_sub.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 17
    r += 1

    # Contact Strip with 4 live clickable hyperlinks
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    links = [
        (2, PROFILE["email"], f"mailto:{PROFILE['email']}"),
        (4, PROFILE["linkedin_label"], PROFILE["linkedin_url"]),
        (5, PROFILE["github_label"], PROFILE["github_url"]),
        (6, PROFILE["web_label"], PROFILE["web_url"]),
    ]
    for col_idx, label, url in links:
        cell = ws.cell(row=r, column=col_idx, value=label)
        cell.hyperlink = url
        cell.font = mk_font(9.5, bold=False, color=ACCENT, underline="single")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    style_range(ws, 2, r, 6, r, fill=mk_fill(WHITE), border=mk_border(bottom=("thin", BORDER_MED)))
    ws.row_dimensions[r].height = 18
    r += 1

    # Profile Summary
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    c_sum = ws.cell(row=r, column=2, value=PROFILE["summary"])
    c_sum.font = mk_font(9.5, color=SLATE)
    c_sum.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    sum_lines = estimate_wrapped_lines(PROFILE["summary"], widths[2] + body_total_w, size_pt=9.5)
    ws.row_dimensions[r].height = calc_row_height(sum_lines, 9.5, 6)
    r += 1

    # Subtle Formula-Powered Snapshot Strip (4 cells across B:F)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    snap_formulas = [
        (2, '="HSBC Tenure: "&ROUND((\'Experience & Projects\'!H6+\'Experience & Projects\'!H7)/12,1)&" yrs"'),
        (4, '="Roles & Internships: "&COUNTA(\'Experience & Projects\'!B6:B10)'),
        (5, '="Projects Shipped: "&COUNTA(\'Experience & Projects\'!B15:B23)'),
        (6, '="DSA Solved: "&SUM(\'Education & Skills\'!E13:E16)&"+  ·  MCA 8.8"'),
    ]
    for col_idx, fml in snap_formulas:
        cell = ws.cell(row=r, column=col_idx, value=fml)
        cell.font = mk_font(8.5, bold=True, color=INK)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    style_range(ws, 2, r, 6, r, fill=mk_fill(CARD_BG), border=mk_border(top=("thin", BORDER_LIGHT), bottom=("thin", BORDER_LIGHT)))
    ws.row_dimensions[r].height = 19
    r += 1

    def add_section_header(title: str) -> int:
        nonlocal r
        ws.row_dimensions[r].height = 4
        r += 1
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        sec_cell = ws.cell(row=r, column=2, value=title)
        sec_cell.font = mk_font(11, bold=True, color=INK, family=FONT_DISPLAY)
        sec_cell.alignment = Alignment(horizontal="left", vertical="bottom")
        style_range(ws, 2, r, 6, r, border=mk_border(bottom=("medium", BORDER_DARK)))
        ws.row_dimensions[r].height = 19
        r += 1
        return r

    def add_bullet(text: str, is_tech: bool = False) -> None:
        nonlocal r
        b_cell = ws.cell(row=r, column=2, value="•")
        b_cell.font = mk_font(9.5, bold=True, color=ACCENT if is_tech else MUTED)
        b_cell.alignment = Alignment(horizontal="center", vertical="top")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        t_cell = ws.cell(row=r, column=3, value=text)
        t_cell.font = mk_font(9.0 if is_tech else 9.5, italic=is_tech, color=MUTED if is_tech else SLATE)
        t_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        lines = estimate_wrapped_lines(text, body_total_w, size_pt=9.0 if is_tech else 9.5)
        ws.row_dimensions[r].height = calc_row_height(lines, 9.5, 2.0)
        r += 1

    # --- Section: EXPERIENCE ---
    add_section_header("EXPERIENCE")
    for idx, job in enumerate(EXPERIENCE):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        c_co = ws.cell(row=r, column=2, value=job["company"])
        c_co.font = mk_font(10, bold=True, color=INK)
        c_co.alignment = Alignment(horizontal="left", vertical="bottom")

        c_per = ws.cell(row=r, column=6, value=job["period"])
        c_per.font = mk_font(9.5, bold=True, color=INK)
        c_per.alignment = Alignment(horizontal="right", vertical="bottom")
        ws.row_dimensions[r].height = 17
        r += 1

        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        c_role = ws.cell(row=r, column=2, value=job["role"])
        c_role.font = mk_font(9.5, italic=True, color=MUTED)
        c_role.alignment = Alignment(horizontal="left", vertical="top")

        c_loc = ws.cell(row=r, column=6, value=job["location"])
        c_loc.font = mk_font(9, italic=True, color=MUTED)
        c_loc.alignment = Alignment(horizontal="right", vertical="top")
        ws.row_dimensions[r].height = 15.5
        r += 1

        for b in job["bullets"]:
            add_bullet(b)
        add_bullet(f"Technologies: {job['tech']}", is_tech=True)

        if idx < len(EXPERIENCE) - 1:
            ws.row_dimensions[r].height = 4
            r += 1

    # --- Section: PROJECTS ---
    # Break right after the last row of Experience (r - 1) so PROJECTS starts cleanly on Page 2.
    ws.row_breaks.append(Break(id=r - 1))

    add_section_header("PROJECTS")
    cv_projects = [p for p in PROJECTS if p["show_on_cv"]]
    for idx, proj in enumerate(cv_projects):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        c_title = ws.cell(row=r, column=2, value=f"{proj['name']}  |  {proj['stack']}")
        c_title.font = mk_font(10, bold=True, color=INK)
        c_title.alignment = Alignment(horizontal="left", vertical="bottom")

        link_label = f"{proj['year']}  ·  GitHub ↗"
        c_link = ws.cell(row=r, column=6, value=link_label)
        c_link.hyperlink = proj["github"]
        c_link.font = mk_font(9, color=ACCENT, underline="single")
        c_link.alignment = Alignment(horizontal="right", vertical="bottom")
        ws.row_dimensions[r].height = 17
        r += 1

        for b in proj["bullets"]:
            add_bullet(b)

        if idx < len(cv_projects) - 1:
            ws.row_dimensions[r].height = 2.0
            r += 1

    # --- Section: EDUCATION ---
    add_section_header("EDUCATION")
    for idx, edu in enumerate([e for e in EDUCATION if e["show_on_cv"]]):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        c_sch = ws.cell(row=r, column=2, value=edu["school"])
        c_sch.font = mk_font(10, bold=True, color=INK)
        c_sch.alignment = Alignment(horizontal="left", vertical="bottom")

        c_per = ws.cell(row=r, column=6, value=edu["period"])
        c_per.font = mk_font(9.5, bold=True, color=INK)
        c_per.alignment = Alignment(horizontal="right", vertical="bottom")
        ws.row_dimensions[r].height = 17
        r += 1

        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        c_deg = ws.cell(row=r, column=2, value=f"{edu['degree']}  ·  {edu['score']} CGPA")
        c_deg.font = mk_font(9.5, italic=True, color=SLATE)
        c_deg.alignment = Alignment(horizontal="left", vertical="top")

        c_loc = ws.cell(row=r, column=6, value=edu["location"])
        c_loc.font = mk_font(9, italic=True, color=MUTED)
        c_loc.alignment = Alignment(horizontal="right", vertical="top")
        ws.row_dimensions[r].height = 15.5
        r += 1

        if idx == 0:
            ws.row_dimensions[r].height = 3
            r += 1

    # --- Section: CO-CURRICULAR ACTIVITIES & ACHIEVEMENTS ---
    add_section_header("CO-CURRICULAR ACTIVITIES & ACHIEVEMENTS")
    cv_achievements = [
        "Final Year Thesis: Customizing MulVAL for Attack Graph Generation with Network-Simulated Digital Forensics.",
        "Certifications: HackerRank Certified in Python (Basic), SQL (Basic), and Java (Basic).",
        "Competitive Programming: 200+ problems solved across LeetCode (178 solved: 72 Easy, 101 Medium, 5 Hard; 50 Days Badge 2022) and HackerRank.",
    ]
    for item in cv_achievements:
        add_bullet(item)

    # --- Section: TECHNICAL SKILLS ---
    add_section_header("TECHNICAL SKILLS")
    for idx, (category, items) in enumerate(CV_SKILLS_SUMMARY):
        bg = ZEBRA_BG if idx % 2 == 0 else WHITE
        c_cat = ws.cell(row=r, column=3, value=category)
        c_cat.font = mk_font(9.5, bold=True, color=INK)
        c_cat.alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
        c_items = ws.cell(row=r, column=4, value=items)
        c_items.font = mk_font(9.5, color=SLATE)
        c_items.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        style_range(ws, 2, r, 6, r, fill=mk_fill(bg), border=mk_border(bottom=("thin", BORDER_LIGHT)))
        skill_w = widths[4] + widths[5] + widths[6]
        lines = estimate_wrapped_lines(items, skill_w, size_pt=9.5)
        ws.row_dimensions[r].height = calc_row_height(lines, 9.5, 2.0)
        r += 1

    # Subtle Dynamic Footer
    ws.row_dimensions[r].height = 4
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    c_foot = ws.cell(
        row=r,
        column=2,
        value='="Traditional CV in Excel  ·  Explore Career Dashboard, Experience & Projects, and Education & Skills tabs  ·  "&TEXT(TODAY(),"mmm yyyy")',
    )
    c_foot.font = mk_font(8.5, italic=True, color=CAPTION)
    c_foot.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 16

    # Page Setup (Portrait A4, 1 page wide, respect manual break after Experience)
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins.left = 0.35
    ws.page_margins.right = 0.35
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35
    ws.page_margins.header = 0.15
    ws.page_margins.footer = 0.15
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def generate_workbook(output_path: Path) -> Path:
    wb = openpyxl.Workbook()
    build_cv_sheet(wb)
    build_dashboard_sheet(wb)
    build_experience_projects_sheet(wb)
    build_education_skills_sheet(wb)

    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    out = Path(__file__).resolve().parent / "Tusher_Mondal_CV.xlsx"
    generate_workbook(out)
    print(f"Generated {out}")
