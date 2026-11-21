import io
import os
import re
import struct
import zipfile
import urllib.parse
from collections import Counter
from datetime import datetime
from email.message import EmailMessage
from email.policy import SMTP as SMTP_POLICY
from email.utils import formatdate

import pandas as pd
import streamlit as st
from PIL import Image as PILImage

# ReportLab imports for dynamic PDF generation
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import inch, cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth, registerFontFamily
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak, KeepInFrame
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_JUSTIFY, TA_LEFT


# ==========================================
# PATHS & MULTI-COLLEGE CONFIGURATION
# ==========================================

try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    BASE_DIR = os.getcwd()

ASSET_DIR = os.path.join(BASE_DIR, "logos")
SVKM_LOGO_PATH = os.path.join(ASSET_DIR, "svkm_logo.png")

SIGNATORY = {
    "name": "Alka Shukla",
    "title": "Controller of Examination",
    "signature_image": os.path.join(ASSET_DIR, "signature.png"),  # 0.86" x 0.43"
}



def _register_bookman_fonts():
    """
    Registers 'BookmanOldStyle' / 'BookmanOldStyle-Bold' for DJSCE letterhead.
    Prefers installed Bookman Old Style (.ttf/.otf) and falls back to system URW Bookman (.otf)
    or Times-Roman / Times-Bold if no font file is present.
    """
    reg_candidates = [
        os.path.join(BASE_DIR, "fonts", "BOOKOS.TTF"),
        os.path.join(BASE_DIR, "BOOKOS.TTF"),
        r"C:\Windows\Fonts\BOOKOS.TTF",
        "/Library/Fonts/Bookman Old Style.ttf",
        "/System/Library/Fonts/Supplemental/Bookman Old Style.ttf",
        "/usr/share/fonts/truetype/msttcorefonts/Bookman_Old_Style.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
    ]
    bold_candidates = [
        os.path.join(BASE_DIR, "fonts", "BOOKOSB.TTF"),
        os.path.join(BASE_DIR, "BOOKOSB.TTF"),
        r"C:\Windows\Fonts\BOOKOSB.TTF",
        r"C:\Windows\Fonts\BOOKOSBI.TTF",
        "/Library/Fonts/Bookman Old Style Bold.ttf",
        "/System/Library/Fonts/Supplemental/Bookman Old Style Bold.ttf",
        "/usr/share/fonts/truetype/msttcorefonts/Bookman_Old_Style_Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf",
    ]
    reg_path = next((p for p in reg_candidates if os.path.isfile(p)), None)
    bold_path = next((p for p in bold_candidates if os.path.isfile(p)), None)
    if reg_path and bold_path:
        try:
            pdfmetrics.registerFont(TTFont("BookmanOldStyle", reg_path))
            pdfmetrics.registerFont(TTFont("BookmanOldStyle-Bold", bold_path))
            registerFontFamily(
                "BookmanOldStyle",
                normal="BookmanOldStyle",
                bold="BookmanOldStyle-Bold",
                italic="BookmanOldStyle",
                boldItalic="BookmanOldStyle-Bold",
            )
            return "BookmanOldStyle", "BookmanOldStyle-Bold"
        except Exception:
            pass
    return "Times-Roman", "Times-Bold"


BOOKMAN_REGULAR, BOOKMAN_BOLD = _register_bookman_fonts()

COLLEGE_CONFIGS = {
    "UPG": {
        "code": "UPG",
        "ui_label": "UPG — Usha Pravin Gandhi College of Arts, Science & Commerce",
        "short_name": "Usha Pravin Gandhi College of Arts, Science & Commerce",
        "docx_candidates": [
            "UPG Appointment Letter Format.docx",
            "UPG_Appointment_Letter_Format.docx",
            "Appointment Letter Format.docx",
            "Appointment_Letter_Format.docx",
        ],
        "right_logo": os.path.join(ASSET_DIR, "upg_logo.jpeg"),
        "right_logo_size": (0.74 * inch, 0.87 * inch),
        "header_lines": [
            {
                "text": "Shri Vile Parle Kelavani Mandal's",
                "font": "Times-Bold",
                "size": 10,
                "leading": 12.5,
                "color": "#1F4E79",
            },
            {
                "text": "USHA PRAVIN GANDHI COLLEGE OF<br/>ARTS, SCIENCE &amp; COMMERCE (AUTONOMOUS)",
                "font": "Times-Bold",
                "size": 15,
                "leading": 17.5,
                "color": "#000000",
            },
            {
                "text": "Affiliated to University of Mumbai",
                "font": "Times-Roman",
                "size": 10,
                "leading": 12.5,
                "color": "#000000",
            },
            {
                "text": "NAAC Re-Accredited \"A+\" Grade with CGPA: 3.27",
                "font": "Times-Roman",
                "size": 10,
                "leading": 12.5,
                "color": "#EE0000",
                "spaceAfter": 4,
            },
        ],
    },
    "MBC": {
        "code": "MBC",
        "ui_label": "MBC — Mithibai College of Arts, Chauhan Institute of Science & A.J. College of Commerce & Economics",
        "short_name": (
            "Mithibai College of Arts, Chauhan Institute of Science and "
            "Amrutben Jivanlal College of Commerce and Economics"
        ),
        "docx_candidates": [
            "MBC Appointment Letter Format.docx",
            "MBC_Appointment_Letter_Format.docx",
        ],
        "right_logo": os.path.join(ASSET_DIR, "mbc_logo.png"),
        "right_logo_size": (0.51 * inch, 0.87 * inch),
        "header_lines": [
            {
                "text": "Shri Vile Parle Kelavani Mandal's",
                "font": "Times-Bold",
                "size": 10,
                "leading": 12.5,
                "color": "#1F4E79",
            },
            {
                "text": (
                    "MITHIBAI COLLEGE OF ARTS, CHAUHAN INSTITUTE OF SCIENCE AND<br/>"
                    "AMRUTBEN JIVANLAL COLLEGE OF COMMERCE AND ECONOMICS"
                ),
                "font": "Times-Bold",
                "size": 11.5,
                "leading": 14,
                "color": "#000000",
            },
            {
                "text": "(Empowered Autonomous)",
                "font": "Times-Bold",
                "size": 10,
                "leading": 12.5,
                "color": "#000000",
            },
            {
                "text": "Affiliated to University of Mumbai",
                "font": "Times-Roman",
                "size": 10,
                "leading": 12.5,
                "color": "#000000",
            },
            {
                "text": "NAAC Reaccredited \"A++\" Grade with CGPA: 3.55 (November 2024)",
                "font": "Times-Roman",
                "size": 10,
                "leading": 12.5,
                "color": "#EE0000",
                "spaceAfter": 4,
            },
        ],
    },
    "NMC": {
        "code": "NMC",
        "ui_label": "NMC — Narsee Monjee College of Commerce & Economics",
        "short_name": "Narsee Monjee College of Commerce & Economics",
        "docx_candidates": [
            "NMC Appointment Letter Format.docx",
            "NMC_Appointment_Letter_Format.docx",
        ],
        "right_logo": os.path.join(ASSET_DIR, "nmc_logo.png"),
        "right_logo_size": (0.79 * inch, 0.79 * inch),
        "header_lines": [
            {
                "text": "Shri Vile Parle Kelavani Mandal's",
                "font": "Times-Bold",
                "size": 10,
                "leading": 12.5,
                "color": "#1F4E79",
            },
            {
                "text": "NARSEE MONJEE COLLEGE OF COMMERCE &amp; ECONOMICS",
                "font": "Times-Bold",
                "size": 13.2,
                "leading": 15.5,
                "color": "#000000",
            },
            {
                "text": "(Empowered Autonomous)",
                "font": "Times-Bold",
                "size": 12,
                "leading": 14.5,
                "color": "#000000",
            },
            {
                "text": "Affiliated to University of Mumbai",
                "font": "Times-Roman",
                "size": 10,
                "leading": 12.5,
                "color": "#000000",
            },
            {
                "text": "NAAC Reaccredited \"A\" Grade with CGPA: 3.42",
                "font": "Times-Roman",
                "size": 10,
                "leading": 12.5,
                "color": "#EE0000",
                "spaceAfter": 4,
            },
        ],
    },
    "DJSCE": {
        "code": "DJSCE",
        "ui_label": "DJSCE — Dwarkadas J. Sanghvi College of Engineering",
        "short_name": "Dwarkadas J. Sanghvi College of Engineering",
        "docx_candidates": [
            "DJSCE Letter Head.docx",
            "DJSCE_Letter_Head.docx",
            "DJSCE Appointment Letter Format.docx",
        ],
        "left_logo_size": (0.6785 * inch, 0.8661 * inch),
        "right_logo": os.path.join(ASSET_DIR, "djsce_logo.png"),
        "right_logo_size": (0.7874 * inch, 0.7874 * inch),
        "left_col_w": 56.45,
        "right_col_w": 55.8,
        "text_align": TA_LEFT,
        "center_line_below": False,
        "right_valign": "BOTTOM",
        "right_align": "RIGHT",
        "header_lines": [
            {
                "text": "Shri Vile Parle Kelavani Mandal's",
                "font": BOOKMAN_REGULAR,
                "size": 11,
                "leading": 13.0,
                "color": "#1F4E79",
                "spaceAfter": 1.0,
            },
            {
                "text": "DWARKADAS J. SANGHVI COLLEGE OF ENGINEERING",
                "font": BOOKMAN_BOLD,
                "size": 13.0 if BOOKMAN_BOLD == "BookmanOldStyle-Bold" else 14,
                "leading": 16.0,
                "color": "#000000",
                "underline_border": True,
                "spaceAfter": 2.0,
            },
            {
                "text": "(Empowered Autonomous College Affiliated to University of Mumbai)",
                "font": BOOKMAN_REGULAR,
                "size": 11,
                "leading": 13.0,
                "color": "#1F4E79",
                "spaceAfter": 1.0,
            },
            {
                "text": "NAAC Accredited with “A” Grade (CGPA: 3.18)",
                "font": BOOKMAN_REGULAR,
                "size": 10,
                "leading": 12.0,
                "color": "#1F4E79",
                "spaceAfter": 1.0,
            },
        ],
    },
}

# Letters are generated ONLY for these roles (case-insensitive substring match)
ELIGIBLE_ROLE_KEYWORDS = ("moderator", "reval")

DESIGNATION_OPTIONS = [
    "(Leave Blank)",
    "Assistant Professor",
    "Associate Professor",
    "Professor",
    "Vice Principal",
    "Principal",
    "Custom (User Defined)",
]


# ------------------------------------------------------------------
# COLLEGE AUTO-DETECTION: 5-digit course codes found in the "Course" column
# ------------------------------------------------------------------
COLLEGE_COURSE_CODES = {
    "UPG": {
        "53011", "53012", "53013", "53014", "53015", "53016",
        "53021", "53022", "53023", "53024", "53026",
    },
    "NMC": {
        "45611", "45612", "45613", "45614", "45615", "45616", "45617", "45618",
        "45619", "45620", "45621", "45701", "45702", "45703",
    },
    "MBC": {
        "40322", "40711", "40712", "40713", "40714", "40715", "40716", "40717",
        "40718", "40719", "40720", "40721", "40722", "40723", "40724", "40725",
        "40726", "40727", "40761", "40762", "40763", "40764", "40765", "40766",
        "40767", "40768", "40769", "40771", "40772", "40774", "40775", "40777",
        "40778", "40779",
    },
    "DJSCE": {
        "60001", "60011", "60002", "60003", "60004", "60005", "60012", "60013",
        "60009", "60017", "60018", "60019", "60006", "60007", "60008",
    },
}
COURSE_CODE_TO_COLLEGE = {
    code: college for college, codes in COLLEGE_COURSE_CODES.items() for code in codes
}
_COURSE_CODE_RE = re.compile(r"(?<!\d)\d{5}(?!\d)")


def detect_college(course_series):
    """
    Identifies the college from the 5-digit course codes present in the raw Course column.
    Returns (college_key_or_None, hits_per_college, unknown_codes).
    college_key is None when no code matched OR when codes of more than one college are mixed
    (the caller tells the two situations apart using hits_per_college).
    """
    hits = {k: Counter() for k in COLLEGE_COURSE_CODES}
    unknown = Counter()
    for value in course_series.fillna("").astype(str):
        for code in _COURSE_CODE_RE.findall(value):
            college = COURSE_CODE_TO_COLLEGE.get(code)
            if college:
                hits[college][code] += 1
            else:
                unknown[code] += 1
    matched = [k for k, c in hits.items() if c]
    return (matched[0] if len(matched) == 1 else None), hits, unknown


# ==========================================
# AUTOMATIC ASSET EXTRACTION FROM WORD TEMPLATES
# ==========================================

def _extract_png_from_emf(emf_bytes):
    """Extracts the embedded DIB bitmap from a Windows EMF file and converts it to PNG."""
    pos = 0
    total_len = len(emf_bytes)
    while pos + 8 <= total_len:
        rec_type, rec_size = struct.unpack("<II", emf_bytes[pos:pos + 8])
        if rec_size < 8 or pos + rec_size > total_len:
            break
        if rec_type == 0x51 and rec_size >= 64:
            off_bmi, cb_bmi, off_bits, cb_bits = struct.unpack("<IIII", emf_bytes[pos + 48:pos + 64])
            if cb_bmi > 0 and cb_bits > 0 and (pos + off_bits + cb_bits) <= total_len:
                bmi_data = emf_bytes[pos + off_bmi:pos + off_bmi + cb_bmi]
                bits_data = emf_bytes[pos + off_bits:pos + off_bits + cb_bits]
                bf_size = 14 + len(bmi_data) + len(bits_data)
                bf_off_bits = 14 + len(bmi_data)
                bmp_header = struct.pack("<2sIHHI", b"BM", bf_size, 0, 0, bf_off_bits)
                im = PILImage.open(io.BytesIO(bmp_header + bmi_data + bits_data))
                out_buf = io.BytesIO()
                im.save(out_buf, format="PNG", optimize=True)
                return out_buf.getvalue()
        pos += rec_size
    return None


def _save_image_for_reportlab(img_bytes, out_path, crop_svkm=False):
    """Saves an image flattened onto a white RGB background so ReportLab renders it cleanly."""
    im = PILImage.open(io.BytesIO(img_bytes))
    if crop_svkm:
        w, h = im.size
        im = im.crop((int(round(w * 0.12989)), 0, int(round(w * (1.0 - 0.13116))), h))
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        im_rgba = im.convert("RGBA")
        bg = PILImage.new("RGB", im_rgba.size, (255, 255, 255))
        bg.paste(im_rgba, mask=im_rgba.split()[3])
        im = bg
    elif im.mode != "RGB":
        im = im.convert("RGB")
    ext = os.path.splitext(out_path)[1].lower()
    fmt = "JPEG" if ext in (".jpg", ".jpeg") else "PNG"
    im.save(out_path, format=fmt, quality=95, optimize=True)


def extract_college_assets_from_docx(college_key, docx_file_or_path=None):
    """
    Ensures the SVKM logo, the selected college's right-side logo, and the Controller's
    signature exist in ./logos/. Automatically extracts them from the college's .docx file.
    """
    os.makedirs(ASSET_DIR, exist_ok=True)
    cfg = COLLEGE_CONFIGS[college_key]
    right_logo_path = cfg["right_logo"]
    sig_path = SIGNATORY["signature_image"]

    if (
        docx_file_or_path is None
        and os.path.isfile(SVKM_LOGO_PATH)
        and os.path.isfile(right_logo_path)
    ):
        return True

    source = docx_file_or_path
    if source is None:
        for fname in cfg["docx_candidates"]:
            for directory in (BASE_DIR, os.getcwd()):
                cand = os.path.join(directory, fname)
                if os.path.isfile(cand):
                    source = cand
                    break
            if source:
                break

    if source is not None:
        try:
            with zipfile.ZipFile(source, "r") as z:
                namelist = z.namelist()
                if college_key == "DJSCE":
                    if "word/media/image1.png" in namelist and (
                        docx_file_or_path is not None or not os.path.isfile(SVKM_LOGO_PATH)
                    ):
                        _save_image_for_reportlab(z.read("word/media/image1.png"), SVKM_LOGO_PATH, crop_svkm=True)
                    if "word/media/image2.png" in namelist and (
                        docx_file_or_path is not None or not os.path.isfile(right_logo_path)
                    ):
                        _save_image_for_reportlab(z.read("word/media/image2.png"), right_logo_path, crop_svkm=False)
                else:
                    if "word/media/image2.png" in namelist and not os.path.isfile(SVKM_LOGO_PATH):
                        _save_image_for_reportlab(z.read("word/media/image2.png"), SVKM_LOGO_PATH, crop_svkm=True)

                    for media_item in ("word/media/image3.jpeg", "word/media/image3.png"):
                        if media_item in namelist and (docx_file_or_path is not None or not os.path.isfile(right_logo_path)):
                            _save_image_for_reportlab(z.read(media_item), right_logo_path, crop_svkm=False)
                            break

                    if "word/media/image1.emf" in namelist and not os.path.isfile(sig_path):
                        png_data = _extract_png_from_emf(z.read("word/media/image1.emf"))
                        if png_data:
                            with open(sig_path, "wb") as f:
                                f.write(png_data)
                    elif "word/media/image1.png" in namelist and not os.path.isfile(sig_path):
                        with open(sig_path, "wb") as f:
                            f.write(z.read("word/media/image1.png"))
        except Exception:
            pass

    return os.path.isfile(SVKM_LOGO_PATH) and os.path.isfile(right_logo_path)


# Pre-extract assets for all colleges on startup if .docx files are present
for _ckey in COLLEGE_CONFIGS:
    extract_college_assets_from_docx(_ckey)


# ==========================================
# DATA CLEANING & DATE FORMATTING HELPERS
# ==========================================

_BRACKET_WITH_DIGITS = re.compile(r"[\(\[\{][^\)\]\}]*\d[^\)\]\}]*[\)\]\}]")

_NAME_PREFIXES = r"(?:Dr|Prof|Mr|Mrs|Ms|Miss|Shri|Smt)"

def clean_text(value, restore_ampersand=False):
    """Strips bracketed numeric codes and stray numbers; optionally restores ERP-stripped '&'."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    s = str(value).replace(" ", " ")
    s = _BRACKET_WITH_DIGITS.sub(" ", s)
    s = re.sub(r"\d+", " ", s)
    s = re.sub(r"\(\s*\)|\[\s*\]|\{\s*\}", " ", s)
    if restore_ampersand:
        s = re.sub(r"([A-Za-z,])  +([A-Z])", r"\1 & \2", s)
    s = re.sub(r"(?:^|\s)[.,;:]+(?=\s|$)", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"\s+([,.;:])", r"\1", s)
    s = s.strip(" -–,;:")
    return s


def tidy_person_name(value):
    """Cleans examiner name, normalises prefixes ('Dr. Harshali Patil'), and splits CamelCase."""
    s = clean_text(value, restore_ampersand=False)
    if not s:
        return ""
    s = re.sub(rf"^({_NAME_PREFIXES})\.?(?=[A-Z])", r"\1. ", s, flags=re.I)
    s = re.sub(rf"^({_NAME_PREFIXES})\.?\s+", r"\1. ", s, flags=re.I)
    s = re.sub(r"([a-z])([A-Z])", r"\1 \2", s)
    s = re.sub(r"\s+", " ", s).strip()

    words = []
    for w in s.split():
        if re.fullmatch(rf"{_NAME_PREFIXES}\.", w, flags=re.I):
            words.append(w[0].upper() + w[1:].lower())
        elif w and w[0].islower():
            words.append(w[0].upper() + w[1:])
        else:
            words.append(w)
    return " ".join(words)


_FEMALE_NAMES = {
    "aarti", "aditi", "alka", "amrita", "anita", "anjali", "archana", "ashwini",
    "beena", "bhavana", "bhakti", "bulbul", "chhaya", "deepa", "deepali", "divya",
    "farzana", "gayatri", "geeta", "harshali", "hema", "jyoti", "kavita", "komal",
    "kranti", "kubra", "lata", "madhuri", "manisha", "manju", "meena", "megha",
    "monica", "naira", "namrata", "neelam", "neeta", "neha", "nikita", "nisha",
    "pallavi", "pooja", "poonam", "prajakta", "prachi", "pratibha", "preeti",
    "priya", "priyanka", "rachana", "radhika", "rajashree", "rakhi", "rashmi",
    "rasika", "rekha", "renuka", "richa", "ritu", "rohini", "rukhsar", "rupali",
    "sampada", "sandhya", "sangeeta", "sanjana", "sapna", "sarika", "savita",
    "sayali", "seema", "shabana", "shamali", "shamim", "sheetal", "shilpa",
    "shilpi", "shraddha", "shreeja", "shreya", "shriya", "shweta", "sidni",
    "smita", "sneha", "snehal", "sonal", "sonali", "suchitra", "sudha", "sujata",
    "sunita", "supriya", "supritha", "swati", "tanvi", "tejal", "tejasvi",
    "trupti", "ujwala", "uma", "urvashi", "usha", "vaishali", "vandana",
    "varsha", "vidhi", "vidya", "vinita", "vrushali", "yaquta", "yogita", "rama",
}

_MALE_NAMES = {
    "abhijeet", "abhishek", "aditya", "ajay", "akash", "akhilesh", "amit", "amol",
    "anand", "anil", "ankit", "anup", "arjun", "arun", "ashish", "ashok", "atul",
    "avinash", "bharat", "bhushan", "chandan", "chandrakant", "chayan", "chetan",
    "chirag", "darshak", "deepak", "devendra", "dhananjay", "dhiraj", "dilip",
    "dinesh", "farhan", "gajanan", "gajendra", "ganesh", "gaurav", "girish",
    "gopal", "harish", "hemant", "hitesh", "imran", "jagdish", "jay", "jayesh",
    "jaymin", "jaynarayan", "jitendra", "kailash", "kamlesh", "kapil", "karan",
    "kartik", "kishor", "krishna", "kunal", "mahendra", "mahesh", "mandar",
    "manish", "manoj", "mayur", "milind", "mohammad", "mohammed", "mohan",
    "mukesh", "nadim", "naresh", "navin", "nilesh", "ninad", "niranjan", "nitin",
    "omkar", "pankaj", "paresh", "parth", "piyush", "pradeep", "prafulla",
    "prakash", "pramod", "pranav", "prasad", "prashant", "pratik", "pravin",
    "rahul", "raj", "rajan", "rajeev", "rajendra", "rajesh", "rakesh", "ram",
    "ramesh", "ramsagar", "ravi", "ravindra", "rohan", "rohit", "rupendra",
    "sachin", "sagar", "sameer", "sandeep", "sanjay", "santosh", "sarvesh",
    "satish", "saurabh", "shahid", "shailesh", "shankar", "shantaram", "sharad",
    "shashikant", "shrikant", "shyam", "siddharth", "subhash", "subodh", "sudhir",
    "suhas", "sujit", "sumit", "sunil", "suraj", "surajkumar", "suresh", "swapnil",
    "tushar", "uday", "ulhas", "umesh", "vahid", "vaibhav", "vidur", "vijay",
    "vikas", "vikram", "vinay", "vinayak", "vinod", "vipul", "vishal", "vitthal",
    "vivek", "yogesh",
}


def guess_salutation(name):
    """Determines 'Sir', 'Madam', or 'Sir/Madam' from title or given name."""
    n = (name or "").strip()
    if not n:
        return "Sir/Madam"
    low = n.lower()
    if re.match(r"^(mr|shri)\.?\b", low):
        return "Sir"
    if re.match(r"^(mrs|ms|miss|smt)\.?\b", low):
        return "Madam"
    rem = re.sub(rf"^({_NAME_PREFIXES})\.?\s+", "", low, flags=re.I).strip()
    first_token = re.split(r"[\s.]+", rem)[0] if rem else ""
    if first_token in _FEMALE_NAMES:
        return "Madam"
    if first_token in _MALE_NAMES:
        return "Sir"
    return "Sir/Madam"


def to_roman(n):
    vals = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
        (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    ]
    out = ""
    for v, sym in vals:
        while n >= v:
            out += sym
            n -= v
    return out


def format_semester(value):
    """'Semester 5' -> 'Semester V', 'Trimester 2' -> 'Trimester II'."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    s = str(value).strip()
    m = re.search(r"(semester|trimester|sem)\s*[-:]?\s*(\d{1,2})", s, re.I)
    if m:
        label = "Trimester" if m.group(1).lower() == "trimester" else "Semester"
        return f"{label} {to_roman(int(m.group(2)))}"
    if re.fullmatch(r"\d{1,2}", s):
        return f"Semester {to_roman(int(s))}"
    return clean_text(s)


def format_academic_year(value):
    """'2026' -> '2026-27'; '2026-27' stays as is."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    s = str(value).strip()
    m = re.search(r"(\d{4})\s*[-/–]\s*(\d{2,4})", s)
    if m:
        return f"{m.group(1)}-{m.group(2)[-2:]}"
    m = re.search(r"\d{4}", s)
    if m:
        y = int(m.group(0))
        return f"{y}-{str(y + 1)[-2:]}"
    return ""


def format_role(value):
    """
    Normalizes role names for letters:
    - Any role containing 'reval' (e.g. 'Reval 1', 'Reval 2', 'Revaluation') becomes 'Examiner (Revaluation)'.
    - 'Moderator' stays 'Moderator'.
    """
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    s = re.sub(r"\s+", " ", str(value)).strip()
    if "reval" in s.lower():
        return "Examiner (Revaluation)"
    if "moderator" in s.lower():
        return "Moderator"
    return s.title()


def is_eligible_role(value):
    r = str(value or "").lower()
    return any(k in r for k in ELIGIBLE_ROLE_KEYWORDS)


def parse_date(value):
    if value is None or (not isinstance(value, (str, datetime)) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value
    s = str(value).strip()
    if not s or s == "-":
        return None
    try:
        ts = pd.to_datetime(s, dayfirst=True, errors="coerce")
    except Exception:
        return None
    return None if pd.isna(ts) else ts.to_pydatetime()


def parse_int_count(value):
    """Parses answer book count (e.g. '29', '29.0', 29) into an integer >= 0."""
    if value is None or (not isinstance(value, (str, int, float)) and pd.isna(value)):
        return 0
    try:
        return max(0, int(round(float(str(value).strip()))))
    except Exception:
        return 0


def ordinal_suffix(day):
    """Returns 'st', 'nd', 'rd', or 'th' for a day of the month."""
    if 11 <= (day % 100) <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")


def fmt_date_plain(dt):
    """Plain text preview: e.g. '28th October, 2026'."""
    if not dt:
        return ""
    suf = ordinal_suffix(dt.day)
    return f"{dt.day}{suf} {dt.strftime('%B, %Y')}"


def fmt_date_xml(dt):
    """ReportLab XML markup with superscript ordinal: e.g. '28<super>th</super> October, 2026'."""
    if not dt:
        return ""
    suf = ordinal_suffix(dt.day)
    return f"{dt.day}<super>{suf}</super> {dt.strftime('%B, %Y')}"


def _to_dt(v):
    """Safely converts None / NaT / pandas Timestamp / datetime into datetime or None."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    return v if isinstance(v, datetime) else None


def sort_by_assigned(records):
    """Ascending order of Assigned Date (rows without a date go last; Excel order kept on ties)."""
    return sorted(
        records,
        key=lambda r: (_to_dt(r.get("assigned_dt")) is None,
                       _to_dt(r.get("assigned_dt")) or datetime.min),
    )


def semester_roman_only(value):
    """'Semester IV' / 'Trimester II' -> 'IV' / 'II' (for the Thanking Letter table)."""
    s = re.sub(r"(?i)\b(semester|trimester|sem)\b\.?\s*[-:]?\s*", "", str(value or ""))
    return s.strip()


def build_address_lines(name, details):
    """
    Addressee block lines. Comma after Institute Name if Address Line 1 or Address Line 2 is entered;
    comma after Address Line 1 only if Address Line 2 is entered.
    """
    def _c(v):
        return str(v or "").strip()

    def _with_comma(text, add):
        text = _c(text)
        if text and add and not text.endswith(","):
            text += ","
        return text

    institute, addr1, addr2 = _c(details.get("institute")), _c(details.get("addr1")), _c(details.get("addr2"))
    return [
        _c(name),
        _c(details.get("designation")),
        _with_comma(institute, bool(addr1 or addr2)),
        _with_comma(addr1, bool(addr2)),
        addr2,
    ]


# ==========================================
# EXCEL LOADING ("EVALUATION DASHBOARD DETAILS" FORMAT)
# ==========================================

def _norm_header(s):
    """'In Progress Count' / 'in-progress count' / 'InProgressCount' -> 'inprogresscount'."""
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def _count_aliases(base):
    return [base, f"{base} count", f"{base} answer books", f"no. of {base}", f"total {base}"]


COLUMN_ALIASES = {
    "name": ["examiner name", "examinername"],
    "email": ["examiner email", "examineremail", "email"],
    "role": ["role", "rolename"],
    "academic_year": ["academic year", "academicyear"],
    "course": ["course", "coursename"],
    "semester": ["semester/trimester", "semester", "trimester"],
    "subject": ["subject", "categoryname"],
    "exam_date": ["exam date", "examdate"],
    "assigned": ["assigned date time", "assigned date", "assigned datetime"],
    "eval_last_date": ["evaluation last date", "evaluationlastdate", "eval last date"],
    # Evaluation-progress counts (used to decide whether the Thanking Letter may be generated)
    "unchecked": _count_aliases("unchecked"),
    "rejected": _count_aliases("rejected"),
    "in_progress": _count_aliases("in progress"),
    "present": _count_aliases("present"),
    "uploaded": _count_aliases("uploaded"),
    "checked": _count_aliases("checked") + ["no. of answer books"],
}
REQUIRED_FIELDS = ["name", "role", "course", "semester", "subject"]

COMPLETION_FIELDS = ["unchecked", "rejected", "in_progress", "present", "uploaded", "checked"]
COMPLETION_LABELS = {
    "unchecked": "Unchecked",
    "rejected": "Rejected",
    "in_progress": "In Progress",
    "present": "Present",
    "uploaded": "Uploaded",
    "checked": "Checked",
}


def completion_status(rec):
    """
    A Moderator / Revaluator has completed the assignment only when
    Unchecked = Rejected = In Progress = 0  AND  Present = Uploaded = Checked (and not all zero).
    Returns (is_complete, human_readable_reason_if_not).
    """
    reasons = []
    for key in ("unchecked", "rejected", "in_progress"):
        if rec[key] != 0:
            reasons.append(f"{COMPLETION_LABELS[key]}: {rec[key]}")
    present, uploaded, checked = rec["present"], rec["uploaded"], rec["checked"]
    if not (present == uploaded == checked):
        reasons.append(f"Present ({present}) / Uploaded ({uploaded}) / Checked ({checked}) do not match")
    elif present == 0:
        reasons.append("no answer books recorded")
    return (not reasons), "; ".join(reasons)


def load_dashboard_excel(file_obj):
    """Returns (dataframe, missing_required_fields, missing_completion_fields)."""
    raw = pd.read_excel(file_obj, header=None, dtype=str)

    header_idx = 0
    for i in range(min(len(raw), 15)):
        cells = [str(c).strip().lower() for c in raw.iloc[i].tolist()]
        if "examiner name" in cells or "examinername" in cells:
            header_idx = i
            break

    headers = [str(c).strip() for c in raw.iloc[header_idx].tolist()]
    df = raw.iloc[header_idx + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    norm_to_actual = {}
    for h in headers:
        n = _norm_header(h)
        if n and n != "nan" and n not in norm_to_actual:
            norm_to_actual[n] = h
    mapping = {}
    for field, aliases in COLUMN_ALIASES.items():
        for a in aliases:
            if _norm_header(a) in norm_to_actual:
                mapping[field] = norm_to_actual[_norm_header(a)]
                break

    missing = [f for f in REQUIRED_FIELDS if f not in mapping]
    missing_completion = [f for f in COMPLETION_FIELDS if f not in mapping]
    if missing:
        return df, missing, missing_completion

    out = pd.DataFrame({field: df[col] for field, col in mapping.items()})
    for field in COLUMN_ALIASES:
        if field not in out.columns:
            out[field] = None

    has_name = out["name"].fillna("").astype(str).str.strip() != ""
    has_role = out["role"].fillna("").astype(str).str.strip() != ""
    out = out[has_name & has_role].reset_index(drop=True)
    return out, [], missing_completion


def build_records(df):
    recs = pd.DataFrame()
    recs["examiner_key"] = df["name"].fillna("").astype(str).str.strip()
    recs["examiner_name"] = df["name"].apply(tidy_person_name)
    recs["examiner_email"] = df["email"].apply(
        lambda v: str(v).strip().lower() if v is not None and not pd.isna(v) else ""
    )
    recs["raw_role"] = df["role"].fillna("").astype(str).str.strip()
    recs["role"] = df["role"].apply(format_role)
    recs["course"] = df["course"].apply(lambda v: clean_text(v, restore_ampersand=True))
    recs["semester"] = df["semester"].apply(format_semester)
    recs["subject"] = df["subject"].apply(lambda v: clean_text(v, restore_ampersand=True))
    recs["academic_year"] = df["academic_year"].apply(format_academic_year)
    recs["exam_dt"] = df["exam_date"].apply(parse_date)
    recs["assigned_dt"] = df["assigned"].apply(parse_date)
    recs["eval_last_dt"] = df["eval_last_date"].apply(parse_date)
    for field in COMPLETION_FIELDS:                      # includes "checked"
        recs[field] = df[field].apply(parse_int_count)
    statuses = [completion_status(r) for _, r in recs.iterrows()]
    recs["complete"] = [s[0] for s in statuses]
    recs["pending_reason"] = [s[1] for s in statuses]
    return recs


def infer_exam_month_year(df_rows):
    """Infers 'Month Year' (e.g. 'October 2026') from Exam Date or Assigned Date."""
    for col in ("exam_dt", "assigned_dt", "eval_last_dt"):
        if col in df_rows.columns:
            valid = df_rows[col].dropna()
            if not valid.empty:
                months = valid.apply(lambda dt: dt.strftime("%B %Y"))
                return months.mode().iloc[0]
    return datetime.now().strftime("%B %Y")


# ==========================================
# EMAIL DRAFT (.EML) & PDF GENERATOR FUNCTIONS
# ==========================================

def build_eml_draft(to_addr, subject, body, pdf_bytes, pdf_filename):
    """
    Builds a standards-compliant unsent email draft (.eml) with CRLF line endings
    and the PDF already attached so Outlook Desktop / New Outlook / Mail apps open it cleanly.
    """
    msg = EmailMessage()
    msg["To"] = (to_addr or "").strip()
    msg["Subject"] = re.sub(r"[\r\n]+", " ", str(subject or "")).strip()
    msg["Date"] = formatdate(localtime=True)
    msg["X-Unsent"] = "1"          # Opens as an editable compose window in Outlook / mail clients
    msg.set_content(body)
    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_filename,
    )
    return msg.as_bytes(policy=SMTP_POLICY)


def _esc(text):
    """Escapes special characters for ReportLab Paragraph XML markup."""
    return (
        str(text or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _img(path, w, h, align="CENTER"):
    if path and os.path.isfile(path):
        try:
            im = Image(path, width=w, height=h)
            im.hAlign = align
            return im
        except Exception:
            return Spacer(1, 1)
    return Spacer(1, 1)


def build_pdf_letterhead(college_key, styles, page_width):
    """
    Replicates the 3-column header table from each college's Word file with exact
    alignment, font family, font sizes, colors, and border placement.
    """
    cfg = COLLEGE_CONFIGS[college_key]

    left_col_w = cfg.get("left_col_w", 0.785 * inch)
    right_col_w = cfg.get("right_col_w", 0.775 * inch)
    center_col_w = page_width - left_col_w - right_col_w
    text_align = cfg.get("text_align", TA_CENTER)

    center_flowables = []
    for idx, line_cfg in enumerate(cfg["header_lines"]):
        st_line = ParagraphStyle(
            f"LH_{college_key}_{idx}",
            parent=styles["Normal"],
            fontName=line_cfg["font"],
            fontSize=line_cfg["size"],
            leading=line_cfg["leading"],
            textColor=colors.HexColor(line_cfg["color"]),
            alignment=text_align,
            spaceBefore=0,
            spaceAfter=line_cfg.get("spaceAfter", 1),
        )
        p = Paragraph(line_cfg["text"], st_line)
        if line_cfg.get("underline_border"):
            # Replicates Word paragraph bottom border (<w:pBdr><w:bottom w:val="single" w:sz="12"/></w:pBdr>)
            bdr_tbl = Table([[p]], colWidths=[center_col_w - 3.0])
            bdr_tbl.setStyle(TableStyle([
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
                ("LINEBELOW", (0, 0), (0, 0), 1.5, colors.black),
            ]))
            center_flowables.append(bdr_tbl)
            center_flowables.append(Spacer(1, line_cfg.get("spaceAfter", 1)))
        else:
            center_flowables.append(p)

    lw, lh = cfg.get("left_logo_size", (0.68 * inch, 0.87 * inch))
    left_logo = _img(SVKM_LOGO_PATH, lw, lh, "CENTER")
    rw, rh = cfg["right_logo_size"]
    right_align = cfg.get("right_align", "CENTER")
    right_valign = cfg.get("right_valign", "MIDDLE")
    right_logo = _img(cfg["right_logo"], rw, rh, right_align)

    tbl_cmds = [
        ("VALIGN", (0, 0), (1, 0), "MIDDLE"),
        ("VALIGN", (2, 0), (2, 0), right_valign),
        ("ALIGN", (0, 0), (0, 0), "CENTER"),
        ("ALIGN", (2, 0), (2, 0), right_align),
        ("LEFTPADDING", (0, 0), (-1, -1), 1.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1.5),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    if cfg.get("center_line_below", True):
        tbl_cmds.append(("LINEBELOW", (1, 0), (1, 0), 1.25, colors.black))

    t = Table(
        [[left_logo, center_flowables, right_logo]],
        colWidths=[left_col_w, center_col_w, right_col_w],
    )
    t.setStyle(TableStyle(tbl_cmds))
    return [t, Spacer(1, 0.18 * inch)]


def _build_signature_block(styles, page_width, container_inset):
    """Builds the Controller of Examination signature block aligned identically to the letter body."""
    sig_text_style = ParagraphStyle(
        "SigText",
        parent=styles["Normal"],
        fontName="Times-Roman",
        fontSize=11,
        leading=LINE_LEADING,
        leftIndent=container_inset,
        rightIndent=container_inset,
        spaceBefore=0,
        spaceAfter=0,
        alignment=TA_LEFT,
    )
    sig_items = []
    if os.path.isfile(SIGNATORY["signature_image"]):
        sig_img = _img(SIGNATORY["signature_image"], 0.86 * inch, 0.43 * inch, "LEFT")
        # Indent image to match body left margin by wrapping in a full-width table with container_inset left padding
        img_tbl = Table([[sig_img]], colWidths=[page_width])
        img_tbl.setStyle(TableStyle([
            ("LEFTPADDING", (0, 0), (-1, -1), container_inset),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        sig_items.append(img_tbl)
    else:
        sig_items.append(Spacer(1, 0.43 * inch))
    sig_items.append(Paragraph(_esc(SIGNATORY["name"]), sig_text_style))
    sig_items.append(Paragraph(_esc(SIGNATORY["title"]), sig_text_style))
    return sig_items


LINE_LEADING = 14     # one uniform line spacing (pt) for every text line in both letters
LETTER_GAP = 11.5       # one uniform vertical gap (pt) between text blocks (Appointment Letter)


def _letter_styles(styles, container_inset, prefix):
    """
    Flat paragraph styles (no space before/after, one fixed leading). All vertical rhythm is
    created explicitly with Spacer(LETTER_GAP) between blocks, so lines inside a block
    (address lines, Program/Semester/... lines) stay tight and every block gap is identical.
    """
    base = ParagraphStyle(
        f"{prefix}Body",
        parent=styles["Normal"],
        fontName="Times-Roman",
        fontSize=11,
        leading=LINE_LEADING,
        leftIndent=container_inset,
        rightIndent=container_inset,
        spaceBefore=0,
        spaceAfter=0,
        alignment=TA_LEFT,
    )
    return {
        "base": base,
        "justified": ParagraphStyle(f"{prefix}BodyJ", parent=base, alignment=TA_JUSTIFY),
        "bold_centre": ParagraphStyle(f"{prefix}Sub", parent=base, fontName="Times-Bold", alignment=TA_CENTER),
        "date": ParagraphStyle(f"{prefix}Date", parent=base, alignment=TA_RIGHT),
    }


def _assignment_club_key(rec):
    """
    Key for clubbing Present count when an examiner has the same Subject,
    Course, Semester/Trimester, Exam Date, and Academic Year (and Role).
    """
    exam_dt = _to_dt(rec.get("exam_dt"))
    exam_key = exam_dt.strftime("%Y-%m-%d") if exam_dt else str(rec.get("exam_date_text") or "").strip().lower()
    return (
        str(rec.get("role") or "").strip().lower(),
        str(rec.get("course") or "").strip().lower(),
        str(rec.get("semester") or "").strip().lower(),
        str(rec.get("subject") or "").strip().lower(),
        exam_key,
        str(rec.get("academic_year") or "").strip().lower(),
    )


def club_duplicate_records(records):
    """
    Clubs Present count (and Checked/Uploaded/Unchecked/Rejected/In Progress counts where applicable)
    for records sharing the same Subject, Course, Semester/Trimester, Exam Date, and Academic Year.
    Preserves chronological order (by assigned_dt) and keeps completion status accurate.
    """
    if not records:
        return []
    ordered = sort_by_assigned(records)
    grouped = {}
    order_keys = []
    for r in ordered:
        k = _assignment_club_key(r)
        if k not in grouped:
            grouped[k] = []
            order_keys.append(k)
        grouped[k].append(dict(r))

    clubbed = []
    for k in order_keys:
        bucket = grouped[k]
        if len(bucket) == 1:
            base_r = bucket[0]
            base_r.setdefault("present_clubbed", parse_int_count(base_r.get("present", base_r.get("checked", 0))))
            clubbed.append(base_r)
            continue

        lead = dict(bucket[0])
        total_present = sum(parse_int_count(item.get("present", 0)) for item in bucket)
        total_checked = sum(parse_int_count(item.get("checked", 0)) for item in bucket)
        total_uploaded = sum(parse_int_count(item.get("uploaded", 0)) for item in bucket)
        total_unchecked = sum(parse_int_count(item.get("unchecked", 0)) for item in bucket)
        total_rejected = sum(parse_int_count(item.get("rejected", 0)) for item in bucket)
        total_in_progress = sum(parse_int_count(item.get("in_progress", 0)) for item in bucket)

        # Earliest assigned date for appointment letter if available
        dts = [_to_dt(item.get("assigned_dt")) for item in bucket if _to_dt(item.get("assigned_dt")) is not None]
        if dts:
            earliest_dt = min(dts)
            lead["assigned_dt"] = earliest_dt
            lead["date_text"] = fmt_date_plain(earliest_dt)
            lead["date_xml"] = fmt_date_xml(earliest_dt)

        lead["present"] = total_present
        lead["present_clubbed"] = total_present
        lead["checked"] = total_checked
        lead["uploaded"] = total_uploaded
        lead["unchecked"] = total_unchecked
        lead["rejected"] = total_rejected
        lead["in_progress"] = total_in_progress

        all_complete = all(bool(item.get("complete", False)) for item in bucket)
        reasons = [str(item.get("pending_reason") or "").strip() for item in bucket if not item.get("complete", False)]
        lead["complete"] = all_complete
        lead["pending_reason"] = "; ".join(dict.fromkeys(r for r in reasons if r))
        clubbed.append(lead)

    return clubbed


def build_appointment_letter_story(college_key, rec, details, exam_month_year, styles, page_width):
    """Builds flowables for ONE Appointment Letter page (one subject per page)."""
    container_inset = 26.0
    gap = LETTER_GAP

    s = _letter_styles(styles, container_inset, "Appt")
    base, justified = s["base"], s["justified"]
    # Double line spacing (2.0 x 11pt = 22pt) strictly for the body paragraphs of the Appointment Letter
    body_double = ParagraphStyle("ApptBodyDoubleJ", parent=justified, leading=22.0)

    story = list(build_pdf_letterhead(college_key, styles, page_width))

    # 1. Date (right-aligned with superscript ordinal)
    story.append(Paragraph(rec["date_xml"], s["date"]))
    story.append(Spacer(1, gap))

    # 2. Addressee block - tight lines, no gaps between them
    story.append(Paragraph("To,", base))
    for line in build_address_lines(rec["examiner_name"], details):
        if line and line.strip():
            story.append(Paragraph(_esc(line.strip()), base))
    story.append(Spacer(1, 24.0))

    # 3. Subject line
    role_raw = str(rec.get("role") or "").strip()
    is_revaluator = "reval" in role_raw.lower()
    role_label = (
        "Examiner (Revaluation)"
        if is_revaluator
        else ("Moderator" if "mod" in role_raw.lower() else (role_raw.title() or "Moderator"))
    )
    role_in_body = "Examiner (Revaluation)" if is_revaluator else role_label.lower()
    action_cap = "Revaluation" if is_revaluator else "Moderation"
    action_low = action_cap.lower()

    exam_my_esc = _esc(exam_month_year)
    program_esc = f"<b>{_esc(rec.get('course') or '')}</b>"
    semester_esc = f"<b>{_esc(rec.get('semester') or '')}</b>"
    acad_yr_raw = str(rec.get("academic_year") or "").strip()
    acad_yr_part = f" <b>({_esc(acad_yr_raw)})</b>" if acad_yr_raw else ""
    subject_esc = f"<b>{_esc(rec.get('subject') or '')}</b>"
    college_esc = _esc(COLLEGE_CONFIGS[college_key]["short_name"])
    present_cnt = parse_int_count(rec.get("present_clubbed", rec.get("present", rec.get("checked", 0))))
    present_cnt_bold = f"<b>{present_cnt}</b>"

    story.append(Paragraph(
        f"Sub: Appointment as {_esc(role_label)} for End Semester Examination, {exam_my_esc}",
        s["bold_centre"],
    ))
    story.append(Spacer(1, 24.0))

    # 4. Salutation & body (body paragraphs use 2.0 line spacing)
    story.append(Paragraph(f"Dear {_esc(rec['salutation'])},", body_double))
    story.append(Spacer(1, gap))

    para1 = (
        f"We are pleased to appoint you as {_esc(role_in_body)} for the End Semester Examination "
        f"of the {program_esc}, {semester_esc}{acad_yr_part} for the subject {subject_esc} "
        f"at {college_esc}."
    )
    story.append(Paragraph(para1, body_double))
    story.append(Spacer(1, gap))

    if not is_revaluator:
        para2 = (
            f"The answer books have already been assessed and are required to be moderated in "
            f"accordance with the examination guidelines of the College. "
            f"{_esc(action_cap)} of {present_cnt_bold} answer books is required to be done as per the rules."
        )
    else:
        para2 = (
            f"{_esc(action_cap)} of {present_cnt_bold} answer books is required to be done as per the rules."
        )
    story.append(Paragraph(para2, body_double))
    story.append(Spacer(1, gap))

    story.append(Paragraph(
        f"The {_esc(action_low)} will be in online mode (Onscreen Marking) for which you will "
        f"receive an email containing the access details upon assignment of the {_esc(action_low)} task.",
        body_double,
    ))
    story.append(Spacer(1, gap))

    story.append(Paragraph(
        f"An honorarium, as per the rules will be paid towards the {_esc(action_low)}.",
        body_double,
    ))
    story.append(Spacer(1, gap))

    story.append(Paragraph("Thank you.", body_double))
    story.append(Spacer(1, gap))

    # 5. Signature block
    story.extend(_build_signature_block(styles, page_width, container_inset))
    return story


def build_thanking_letter_story(
    college_key, records, details, date_xml, exam_month_year, styles, page_width, page_height
):
    """
    Builds a single-page Thanking Letter combining all selected subjects for the examiner
    into a bordered table (Program | Semester | Subject | No. of Answer books),
    with no Subject (Sub:) line and no Total row at the bottom of the table.
    Wrapped in KeepInFrame(..., mode='shrink') to guarantee a single page.
    """
    container_inset = 26.0
    n_rows = len(records)

    # Slightly tighten the block gap and table when an examiner has many subjects (> 5)
    compact = n_rows > 5
    gap = 6.0 if compact else LETTER_GAP
    tbl_font_size = 9.5 if compact else 10.5
    tbl_leading = 11.5 if compact else 13.0
    tbl_pad = 2 if compact else 3

    s = _letter_styles(styles, container_inset, "Thank")
    base, justified = s["base"], s["justified"]

    story = list(build_pdf_letterhead(college_key, styles, page_width))

    # 1. Date
    story.append(Paragraph(date_xml, s["date"]))
    story.append(Spacer(1, gap))

    # 2. Addressee block - tight lines, no gaps between them
    story.append(Paragraph("To,", base))
    first_rec = records[0]
    for line in build_address_lines(first_rec["examiner_name"], details):
        if line and line.strip():
            story.append(Paragraph(_esc(line.strip()), base))
    story.append(Spacer(1, 24.0))

    # 3. Salutation & opening paragraphs
    roles_txt = " / ".join(dict.fromkeys(r["role"] for r in records))
    role_esc = _esc(roles_txt)
    exam_my_esc = _esc(exam_month_year)

    story.append(Paragraph(f"Dear {_esc(first_rec['salutation'])},", justified))
    story.append(Spacer(1, gap))
    story.append(Paragraph(
        f"We are thankful to you for agreeing to be the {role_esc} at our college for "
        f"the End Semester Examinations held in {exam_my_esc}.",
        justified,
    ))
    story.append(Spacer(1, gap))
    story.append(Paragraph(
        "The total number of answer books are as mentioned below:",
        justified,
    ))
    story.append(Spacer(1, gap))

    # 4. Table: Program | Semester | Subject | No. of Answer books (no Total row)
    # splitLongWords=0 -> a header word is never broken in the middle, whatever the column width.
    th_style = ParagraphStyle(
        "TblHeader",
        parent=styles["Normal"],
        fontName="Times-Bold",
        fontSize=tbl_font_size,
        leading=tbl_leading,
        alignment=TA_CENTER,
        splitLongWords=0,
    )
    td_left = ParagraphStyle(
        "TblCellLeft",
        parent=styles["Normal"],
        fontName="Times-Roman",
        fontSize=tbl_font_size,
        leading=tbl_leading,
        alignment=TA_LEFT,
    )
    td_center = ParagraphStyle("TblCellCenter", parent=td_left, alignment=TA_CENTER)

    # Header lines: any wrapping happens only at the explicit <br/>, never inside a word
    header_lines = [["Program"], ["Semester"], ["Subject"], ["No. of", "Answer books"]]
    table_data = [[
        Paragraph("<b>" + "<br/>".join(lines) + "</b>", th_style) for lines in header_lines
    ]]
    for r in records:
        ans_cnt = parse_int_count(r.get("present_clubbed", r.get("present", r.get("checked", 0))))
        table_data.append([
            Paragraph(_esc(r["course"]), td_left),
            Paragraph(_esc(semester_roman_only(r["semester"])), td_center),
            Paragraph(_esc(r["subject"]), td_left),
            Paragraph(str(ans_cnt), td_center),
        ])

    tbl_width = page_width - 2 * container_inset
    pad_lr = 0.1 * cm
    pad_tb = 0.05 * cm
    cell_pad = 2 * pad_lr
    # Reduce Semester width and transfer that width to Subject column
    sem_min_w = max(stringWidth("Semester", "Times-Bold", tbl_font_size) + cell_pad + 4, 0.52 * inch)
    other_mins = [
        max(stringWidth(line, "Times-Bold", tbl_font_size) for line in lines) + cell_pad + 6
        for idx, lines in enumerate(header_lines) if idx != 1
    ]
    min_widths = [other_mins[0], sem_min_w, other_mins[1], other_mins[2]]
    col_widths = [
        max(tbl_width * pct, mn) for pct, mn in zip((0.37, 0.08, 0.39, 0.16), min_widths)
    ]
    # Transfer any width above sem_min_w from Semester (col 1) to Subject (col 2)
    if col_widths[1] > sem_min_w:
        shift_w = col_widths[1] - sem_min_w
        col_widths[1] = sem_min_w
        col_widths[2] += shift_w

    excess = sum(col_widths) - tbl_width
    if excess > 0:   # take the surplus from Program and Subject (keep Semester compact)
        flex_total = col_widths[0] + col_widths[2]
        col_widths[0] -= excess * col_widths[0] / flex_total
        col_widths[2] -= excess * col_widths[2] / flex_total
    elif excess < 0:  # distribute any remaining slack proportionally so table width matches body width exactly
        flex_total = col_widths[0] + col_widths[2]
        col_widths[0] += (-excess) * col_widths[0] / flex_total
        col_widths[2] += (-excess) * col_widths[2] / flex_total

    # Compute uniform/equal row height across all rows in the Thanking Letter table
    max_lines_per_row = [1, 2]  # header has 2 lines in "No. of / Answer books"
    for r in records:
        row_texts = [
            (str(r.get("course") or ""), col_widths[0]),
            (semester_roman_only(r.get("semester") or ""), col_widths[1]),
            (str(r.get("subject") or ""), col_widths[2]),
            (str(parse_int_count(r.get("present_clubbed", r.get("present", r.get("checked", 0))))), col_widths[3]),
        ]
        line_counts = []
        for txt, cw in row_texts:
            avail_w = max(cw - cell_pad, 10.0)
            words = txt.split()
            if not words:
                line_counts.append(1)
                continue
            lines_cnt = 1
            cur = ""
            for w in words:
                cand = f"{cur} {w}".strip()
                if stringWidth(cand, "Times-Roman", tbl_font_size) <= avail_w:
                    cur = cand
                else:
                    lines_cnt += 1
                    cur = w
            line_counts.append(lines_cnt)
        max_lines_per_row.append(max(line_counts))

    max_lines = max(max_lines_per_row) if max_lines_per_row else 1
    equal_row_h = max_lines * tbl_leading + 2 * pad_tb + 2.0
    row_heights = [equal_row_h] * len(table_data)

    subj_table = Table(table_data, colWidths=col_widths, rowHeights=row_heights, repeatRows=1, hAlign="LEFT")
    subj_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), pad_tb),
        ("BOTTOMPADDING", (0, 0), (-1, -1), pad_tb),
        ("LEFTPADDING", (0, 0), (-1, -1), pad_lr),
        ("RIGHTPADDING", (0, 0), (-1, -1), pad_lr),
    ]))

    wrapper_table = Table([[subj_table]], colWidths=[page_width], hAlign="LEFT")
    wrapper_table.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), container_inset),
        ("RIGHTPADDING", (0, 0), (-1, -1), container_inset),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(wrapper_table)
    story.append(Spacer(1, gap))

    # 5. Closing paragraphs
    story.append(Paragraph(
        "The efforts that you have taken and the time you have spared for this task are truly appreciated.",
        justified,
    ))
    story.append(Spacer(1, gap))
    story.append(Paragraph(
        "Thank you once again and looking forward to your continued support.",
        justified,
    ))
    story.append(Spacer(1, gap))

    # 6. Signature block
    story.extend(_build_signature_block(styles, page_width, container_inset))

    # Guarantee everything stays on a single page even for 12+ rows
    kif = KeepInFrame(page_width - 12, page_height - 12, story, mode="shrink")
    return [kif]


def generate_letter_pdf(
    college_key,
    letter_type,
    records,
    details,
    thank_date_xml="",
    exam_month_year="",
):
    """
    Generates:
    - Appointment Letter PDF (one page per subject assignment),
    - Thanking Letter PDF (single page with all subjects in a table), or
    - All Letters PDF (all Appointment Letters followed by the Thanking Letter on the last page).
    """
    if letter_type in ("Thanking Letter", "All Letters"):
        not_done = [r for r in records if not r.get("complete", False)]
        if not_done:
            raise ValueError(
                "Thanking Letter cannot be generated - evaluation not completed for: "
                + "; ".join(f"{r['subject']} ({r.get('pending_reason') or 'status unknown'})" for r in not_done)
            )

    extract_college_assets_from_docx(college_key)
    cfg = COLLEGE_CONFIGS[college_key]

    buffer = io.BytesIO()
    margin = 0.5 * inch
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=margin,
        leftMargin=margin,
        topMargin=margin,
        bottomMargin=margin,
        title=f"{cfg['code']} {letter_type}",
        author=cfg["short_name"],
    )
    page_width = A4[0] - 2 * margin
    page_height = A4[1] - 2 * margin
    styles = getSampleStyleSheet()

    records = club_duplicate_records(records)   # club matching assignments & sort by Assigned Date

    story = []
    if letter_type in ("Appointment Letter", "All Letters"):
        for i, rec in enumerate(records):
            if i > 0:
                story.append(PageBreak())
            story.extend(
                build_appointment_letter_story(
                    college_key, rec, details, exam_month_year, styles, page_width
                )
            )

    if letter_type in ("Thanking Letter", "All Letters"):
        if story:
            story.append(PageBreak())   # Thanking Letter always on the last page
        story.extend(
            build_thanking_letter_story(
                college_key, records, details, thank_date_xml,
                exam_month_year, styles, page_width, page_height,
            )
        )

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# ==========================================
# COMPACT STREAMLIT UI MODULE
# ==========================================

def show():
    st.markdown(
        """
        <style>
        div[data-testid="stVerticalBlock"] > div { padding-bottom: 0.15rem; }
        .stTextInput, .stSelectbox, .stMultiSelect, .stDateInput { margin-bottom: -0.25rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("✉️ Letters to External Examiners")

    # ------------------------------------------------------------------
    # Upload Evaluation Dashboard Excel - the college is identified automatically
    # ------------------------------------------------------------------
    excel_file = st.file_uploader(
        "Upload Evaluation Dashboard (.xlsx)",
        type=["xlsx", "xls"],
        key="dashboard_excel_uploader",
    )

    if excel_file is None:
        st.info("\U0001F446 Upload the Evaluation Dashboard Excel file — the college is identified automatically "
                "from the 5-digit codes in the Course column.")
        return

    try:
        df_raw, missing, missing_completion = load_dashboard_excel(excel_file)
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return

    if missing:
        nice = {
            "name": "Examiner Name",
            "role": "Role",
            "course": "Course",
            "semester": "Semester/Trimester",
            "subject": "Subject",
        }
        st.error("Missing required column(s): " + ", ".join(nice[m] for m in missing))
        return

    # ------------------------------------------------------------------
    # College auto-detection from Course codes
    # ------------------------------------------------------------------
    detected_key, code_hits, unknown_codes = detect_college(df_raw["course"])
    matched_colleges = [k for k, c in code_hits.items() if c]

    if len(matched_colleges) > 1:
        st.error(
            "This file contains Course codes of more than one college ("
            + ", ".join(f"{k}: {sum(code_hits[k].values())} row(s)" for k in matched_colleges)
            + "). Please upload a file that belongs to a single college."
        )
        return

    if detected_key is None:
        st.warning(
            "Could not identify the college from the Course codes"
            + (f" (codes found: {', '.join(sorted(unknown_codes))})" if unknown_codes else " (no 5-digit code found)")
            + ". Please select the college manually."
        )
        college_key = st.selectbox(
            "Select College",
            options=list(COLLEGE_CONFIGS.keys()),
            format_func=lambda k: COLLEGE_CONFIGS[k]["ui_label"],
            key="manual_college",
        )
    else:
        college_key = detected_key
        st.success(f"\U0001F3EB College identified from Course codes: **{COLLEGE_CONFIGS[college_key]['ui_label']}**")
        if unknown_codes:
            st.caption(f"Unrecognised Course code(s) ignored: {', '.join(sorted(unknown_codes))}")
    college_cfg = COLLEGE_CONFIGS[college_key]

    # Ensure logos & signature for the detected college are available
    assets_ok = extract_college_assets_from_docx(college_key)
    if not assets_ok:
        with st.expander(
            f"⚠️ Missing {college_key} Logo / Signature — Upload '{college_cfg['docx_candidates'][0]}' once",
            expanded=True,
        ):
            docx_up = st.file_uploader(
                f"Upload {college_cfg['docx_candidates'][0]}",
                type=["docx"],
                key=f"docx_uploader_{college_key}",
            )
            if docx_up is not None:
                if extract_college_assets_from_docx(college_key, docx_up):
                    st.success(f"✅ {college_key} letterhead logos and signature extracted to `./logos/`!")
                    st.rerun()
                else:
                    st.error("Could not extract all required images from the uploaded Word file.")

    recs = build_records(df_raw)
    recs["eligible"] = recs["raw_role"].apply(is_eligible_role)
    eligible = recs[recs["eligible"]].reset_index(drop=True)

    if eligible.empty:
        st.warning("No Moderator / Reval rows found in this file.")
        return

    counts = eligible.groupby("examiner_key").size()
    ex_keys = sorted(
        counts.index.tolist(),
        key=lambda k: eligible.loc[eligible["examiner_key"] == k, "examiner_name"].iloc[0].lower(),
    )
    name_of = {k: eligible.loc[eligible["examiner_key"] == k, "examiner_name"].iloc[0] for k in ex_keys}

    st.caption(
        f"**{college_cfg['code']} Format Active** • Loaded **{len(recs)}** rows • "
        f"**{len(eligible)}** Moderator / Examiner (Revaluation) assignments across **{len(ex_keys)}** examiners."
    )

    # ------------------------------------------------------------------
    # Compact Panel 1: Examiner, Date Options, Academic Year, Exam Month & Year, Subjects
    # ------------------------------------------------------------------
    with st.container(border=True):
        r1_c1, r1_c2, r1_c3, r1_c4, r1_c5 = st.columns([2.0, 1.35, 1.1, 0.9, 1.1])

        sel_key = r1_c1.selectbox(
            "Examiner",
            ex_keys,
            format_func=lambda k: f"{name_of[k]} ({counts[k]} subject{'s' if counts[k] > 1 else ''})",
            key="sel_examiner",
        )
        ex_rows_raw = eligible[eligible["examiner_key"] == sel_key].to_dict("records")
        ex_rows = pd.DataFrame(club_duplicate_records(ex_rows_raw)).reset_index(drop=True)

        def subj_label(i):
            r = ex_rows.iloc[i]
            status = "✅" if r["complete"] else "⏳"
            p_cnt = parse_int_count(r.get("present_clubbed", r.get("present", r.get("checked", 0))))
            return f"{status} {r['role']} | {r['semester']} | {r['subject']} ({r['course']}) — Present: {p_cnt}"

        chosen_idx = st.multiselect(
            "Subjects to include   (✅ evaluation completed · ⏳ pending)",
            list(range(len(ex_rows))),
            default=list(range(len(ex_rows))),
            format_func=subj_label,
            key=f"subjects_{sel_key}",
        )

        selected_rows = ex_rows.iloc[chosen_idx] if chosen_idx else ex_rows

        valid_eval_dates = selected_rows["eval_last_dt"].dropna()
        valid_assigned_dates = selected_rows["assigned_dt"].dropna()
        if not valid_eval_dates.empty:
            highest_eval_dt = valid_eval_dates.max()
        elif not valid_assigned_dates.empty:
            highest_eval_dt = valid_assigned_dates.max()
        else:
            highest_eval_dt = datetime.now()

        date_mode = r1_c2.selectbox(
            "Letter Date Source",
            ["Excel Date (Assigned / Max Eval)", "Today's date", "Custom date"],
            key=f"datemode_{sel_key}",
        )
        if date_mode == "Custom date":
            custom_date = r1_c3.date_input(
                "Pick Date (dd-mm-yyyy)",
                value=highest_eval_dt.date(),
                format="DD-MM-YYYY",
                key=f"cdate_{sel_key}",
            )
        else:
            custom_date = None
            preview_dt = highest_eval_dt if date_mode.startswith("Excel") else datetime.now()
            r1_c3.text_input(
                "Max Eval / Today",
                value=fmt_date_plain(preview_dt),
                disabled=True,
                key=f"sample_dt_{sel_key}_{date_mode}_{len(chosen_idx)}",
            )

        default_ay = selected_rows["academic_year"].iloc[0] if len(selected_rows) else "2026-27"
        ay_override = r1_c4.text_input(
            "Academic Year",
            value=default_ay,
            placeholder="e.g. 2026-27",
            key=f"ay_{sel_key}",
        ).strip()

        default_my = infer_exam_month_year(selected_rows)
        exam_month_year = r1_c5.text_input(
            "Exam Month & Year",
            value=default_my,
            placeholder="e.g. October 2026",
            key=f"exam_my_{sel_key}",
        ).strip() or default_my

    if not chosen_idx:
        st.warning("Select at least one subject.")
        return

    # ------------------------------------------------------------------
    # Compact Panel 2: Addressee & Institute Details (with Designation Dropdown)
    # ------------------------------------------------------------------
    default_name = name_of[sel_key]
    with st.container(border=True):
        st.caption("**Addressee Details** *(optional lines left blank are automatically omitted from the letter)*")
        a_c1, a_c2, a_c3 = st.columns([1.7, 0.85, 1.45])
        letter_name = a_c1.text_input(
            "Examiner Name (printed)",
            value=default_name,
            key=f"name_{sel_key}",
        )
        sal_options = ["Sir/Madam", "Sir", "Madam"]
        sal_default = guess_salutation(letter_name)
        salutation = a_c2.selectbox(
            "Salutation",
            sal_options,
            index=sal_options.index(sal_default),
            key=f"sal_{sel_key}",
        )
        desig_choice = a_c3.selectbox(
            "Designation",
            DESIGNATION_OPTIONS,
            index=0,
            key=f"desig_sel_{sel_key}",
        )

        if desig_choice == "Custom (User Defined)":
            b_c0, b_c1, b_c2, b_c3 = st.columns([1.2, 1.4, 1.1, 1.1])
            custom_desig = b_c0.text_input(
                "Custom Designation",
                placeholder="Enter Designation",
                key=f"desig_custom_{sel_key}",
            )
            designation = custom_desig.strip()
        else:
            b_c1, b_c2, b_c3 = st.columns([1.6, 1.3, 1.3])
            designation = "" if desig_choice == "(Leave Blank)" else desig_choice

        institute = b_c1.text_input(
            "Institute Name",
            placeholder="e.g. K. J. Somaiya College",
            key=f"inst_{sel_key}",
        )
        addr1 = b_c2.text_input(
            "Address Line 1",
            placeholder="e.g. Vidyavihar",
            key=f"addr1_{sel_key}",
        )
        addr2 = b_c3.text_input(
            "Address Line 2",
            placeholder="e.g. Mumbai - 400077",
            key=f"addr2_{sel_key}",
        )

    # ------------------------------------------------------------------
    # Build Final Records for PDF Generation
    # ------------------------------------------------------------------
    if date_mode == "Custom date":
        thank_dt = datetime.combine(custom_date, datetime.min.time())
    elif date_mode == "Today's date":
        thank_dt = datetime.now()
    else:
        thank_dt = highest_eval_dt

    thank_date_plain = fmt_date_plain(thank_dt)
    thank_date_xml = fmt_date_xml(thank_dt)

    final_records = []
    for i in chosen_idx:
        r = ex_rows.iloc[i]
        if date_mode.startswith("Excel"):
            appt_dt = _to_dt(r["assigned_dt"]) or datetime.now()
        elif date_mode == "Today's date":
            appt_dt = datetime.now()
        else:
            appt_dt = datetime.combine(custom_date, datetime.min.time())

        final_records.append({
            "examiner_name": letter_name.strip() or default_name,
            "role": r["role"],
            "course": r["course"],
            "semester": r["semester"],
            "subject": r["subject"],
            "academic_year": ay_override or r["academic_year"],
            "exam_dt": _to_dt(r.get("exam_dt")),
            "present": parse_int_count(r.get("present", r.get("checked", 0))),
            "present_clubbed": parse_int_count(r.get("present_clubbed", r.get("present", r.get("checked", 0)))),
            "checked": r["checked"],
            "complete": bool(r["complete"]),
            "pending_reason": r["pending_reason"],
            "salutation": salutation,
            "date_text": fmt_date_plain(appt_dt),
            "date_xml": fmt_date_xml(appt_dt),
            "thank_date_text": thank_date_plain,
            "assigned_dt": _to_dt(r["assigned_dt"]),
        })

    final_records = sort_by_assigned(final_records)

    details = {
        "designation": designation,
        "institute": institute.strip(),
        "addr1": addr1.strip(),
        "addr2": addr2.strip(),
    }

    with st.expander(f"\U0001F4CB Preview Selected Data ({len(final_records)} subject(s))", expanded=False):
        preview_df = pd.DataFrame(final_records)
        preview_df["status"] = preview_df.apply(
            lambda x: "✅ Completed" if x["complete"] else f"⏳ {x['pending_reason']}", axis=1
        )
        preview_df = preview_df[
            ["date_text", "thank_date_text", "examiner_name", "role", "course", "semester",
             "subject", "present_clubbed", "checked", "status"]
        ].rename(columns={
            "date_text": "Appt. Date",
            "thank_date_text": "Thanking Date",
            "examiner_name": "Examiner",
            "role": "Role",
            "course": "Program",
            "semester": "Semester",
            "subject": "Subject",
            "present_clubbed": "Present (Clubbed)",
            "checked": "Checked",
            "status": "Evaluation Status",
        })
        st.dataframe(preview_df, use_container_width=True, hide_index=True)

    # ------------------------------------------------------------------
    # Which letters? Thanking Letter only when ALL selected evaluations are completed
    # ------------------------------------------------------------------
    pending_records = [r for r in final_records if not r["complete"]]
    thanking_ok = (not missing_completion) and (not pending_records)

    with st.container(border=True):
        st.caption("**Letters to generate**")
        l_c1, l_c2 = st.columns(2)
        want_appt = l_c1.checkbox(
            "Appointment Letter(s)",
            value=True,
            key=f"want_appt_{sel_key}",
        )
        want_thank = l_c2.checkbox(
            "Thanking Letter",
            value=thanking_ok,
            disabled=not thanking_ok,
            key=f"want_thank_{sel_key}_{int(thanking_ok)}",
        )
        want_thank = want_thank and thanking_ok

        if missing_completion:
            st.warning(
                "Thanking Letter is unavailable: the Excel file has no column for "
                + ", ".join(COMPLETION_LABELS[f] for f in missing_completion) + "."
            )
        elif pending_records:
            st.warning(
                "Thanking Letter is unavailable until the selected evaluation(s) are completed "
                "(Unchecked, Rejected and In Progress must be 0 and Present = Uploaded = Checked):\n\n"
                + "\n".join(
                    f"- {r['role']} | {r['semester']} | {r['subject']} — {r['pending_reason']}"
                    for r in pending_records
                )
            )

    if want_appt and want_thank:
        letter_type, file_tag, letters_label = "All Letters", "Letters", "Appointment + Thanking Letters"
    elif want_appt:
        letter_type, file_tag, letters_label = "Appointment Letter", "Appointment_Letters", "Appointment Letter(s)"
    elif want_thank:
        letter_type, file_tag, letters_label = "Thanking Letter", "Thanking_Letter", "Thanking Letter"
    else:
        st.warning("Select at least one letter to generate.")
        return

    # ------------------------------------------------------------------
    # Generation button
    # ------------------------------------------------------------------
    base_sig = repr((college_key, sel_key, letter_type, final_records, details, thank_date_xml, exam_month_year))
    clean_name = re.sub(r"[^A-Za-z0-9]+", "_", letter_name.strip() or default_name).strip("_")

    gen_btn = st.button(
        f"⚡ Generate {letters_label} — PDF",
        type="primary",
        use_container_width=True,
    )

    if gen_btn:
        try:
            st.session_state["appt_pdf"] = generate_letter_pdf(
                college_key=college_key,
                letter_type=letter_type,
                records=final_records,
                details=details,
                thank_date_xml=thank_date_xml,
                exam_month_year=exam_month_year,
            )
            st.session_state["appt_sig"] = base_sig
            st.session_state["pdf_filename"] = f"{college_cfg['code']}_{file_tag}_{clean_name}.pdf"
        except Exception as e:
            st.error(f"PDF generation failed: {e}")

    pdf_ready = ("appt_pdf" in st.session_state and st.session_state.get("appt_sig") == base_sig)
    if not pdf_ready:
        if "appt_pdf" in st.session_state:
            st.caption("ℹ️ Inputs changed since last generation — click the generate button above.")
        return

    active_filename = st.session_state.get("pdf_filename", f"{college_cfg['code']}_{file_tag}_{clean_name}.pdf")
    n_appt_pages = len(final_records) if want_appt else 0
    n_thank_pages = 1 if want_thank else 0
    st.caption(
        f"{n_appt_pages} appointment letter page(s) + {n_thank_pages} thanking letter page(s) "
        f"= {n_appt_pages + n_thank_pages} page(s)."
    )

    st.download_button(
        label=f"\U0001F4C4 Download {active_filename}",
        data=st.session_state["appt_pdf"],
        file_name=active_filename,
        mime="application/pdf",
        use_container_width=True,
    )

    # ------------------------------------------------------------------
    # Compact Email Section (always matches the letters contained in the generated PDF)
    # ------------------------------------------------------------------
    with st.container(border=True):
        st.caption(f"**Send {letters_label} via Email** — use the .eml draft for an auto-attached PDF, or the buttons (PDF must be attached manually).")
        excel_email = ex_rows["examiner_email"].iloc[0] if len(ex_rows) else ""
        roles_txt = " / ".join(dict.fromkeys(r["role"] for r in final_records))
        n_subj = len(final_records)
        subj_word = "subject" if n_subj == 1 else "subjects"
        plural_s = "" if n_subj == 1 else "s"

        if letter_type == "All Letters":
            email_subject = (
                f"Appointment & Thanking Letters ({college_cfg['code']}) - {roles_txt} - "
                f"End Semester Examinations ({exam_month_year})"
            )
            attach_para = (
                f"Please find attached your appointment letter(s) as {roles_txt} for the End Semester "
                f"Examination ({n_subj} {subj_word}), followed by our letter of appreciation "
                f"for your valuable contribution to the End Semester Examinations held in {exam_month_year}."
            )
        elif letter_type == "Appointment Letter":
            email_subject = (
                f"Appointment Letter{plural_s} ({college_cfg['code']}) - {roles_txt} - "
                f"End Semester Examinations ({exam_month_year})"
            )
            attach_para = (
                f"Please find attached your appointment letter{plural_s} as {roles_txt} for the End Semester "
                f"Examination ({n_subj} {subj_word})."
            )
        else:
            email_subject = (
                f"Thanking Letter ({college_cfg['code']}) - {roles_txt} - "
                f"End Semester Examinations ({exam_month_year})"
            )
            attach_para = (
                f"Please find attached our letter of appreciation for your valuable contribution as "
                f"{roles_txt} to the End Semester Examinations held in {exam_month_year}."
            )

        default_body = (
            f"Dear {salutation},\n\n"
            f"Greetings from the Examination Department, {college_cfg['short_name']}.\n\n"
            f"{attach_para}\n\n"
            f"If you have any queries, please feel free to reach out to the Examination Cell.\n\n"
            f"Warm regards.\n"            
        )

        email_to = st.text_input("Recipient Email", value=excel_email, key=f"email_{sel_key}")
        email_body = st.text_area(
            "Email Message",
            value=default_body,
            height=155,
            key=f"emailbody_{college_key}_{sel_key}_{salutation.replace('/', '_')}_{file_tag}",
        )

        encoded_subject = urllib.parse.quote(email_subject)
        encoded_body = urllib.parse.quote(email_body)
        mailto_url = f"mailto:{email_to}?subject={encoded_subject}&body={encoded_body}"
        owa_url = (
            f"https://outlook.office.com/mail/deeplink/compose?"
            f"to={urllib.parse.quote(email_to)}&subject={encoded_subject}&body={encoded_body}"
        )
        m_c1, m_c2 = st.columns(2)
        m_c1.link_button("\U0001F310 Open in Outlook Web", owa_url, use_container_width=True)
        m_c2.link_button("✉️ Open in Default Mail App", mailto_url, use_container_width=True)

        st.download_button(
            "\U0001F4CE Download Email Draft (.eml) — PDF already attached",
            data=build_eml_draft(email_to, email_subject, email_body, st.session_state["appt_pdf"], active_filename),
            file_name=f"{college_cfg['code']}_Email_Draft_{clean_name}.eml",
            mime="message/rfc822",
            type="primary",
            use_container_width=True,
            help="Double-click the downloaded file to open a ready-to-send draft with the PDF attached "
                 "(Outlook desktop, Apple Mail, Thunderbird).",
        )


if __name__ == "__main__":
    show()