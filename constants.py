APP_VERSION = "1.0.0"

CLR_SIDEBAR = "#1E2235"
CLR_CARD = "#FFFFFF"
CLR_ACCENT = "#4A90E2"
CLR_BG = "#F4F6F9"
CLR_TEXT = "#2C3E50"
CLR_HEADER = "#343A56"
CLR_ROW_ALT = "#F8F9FC"
CLR_BORDER = "#E0E4EF"

TEMPLATE_DIR = r"C:\sr_templates"

EXCEL_MAP = {
    "SG": rf"{TEMPLATE_DIR}\sg.xlsx",
    "IGWFW": rf"{TEMPLATE_DIR}\igwfw.xlsx",
    "VM": rf"{TEMPLATE_DIR}\vm.xlsx",
    "LB": rf"{TEMPLATE_DIR}\lb.xlsx",
}

CATEGORIES = list(EXCEL_MAP.keys())
