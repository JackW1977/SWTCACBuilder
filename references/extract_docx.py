import zipfile
import sys
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def tag(local):
    return "{%s}%s" % (W, local)

def get_para_style(para):
    pPr = para.find(tag("pPr"))
    if pPr is not None:
        pStyle = pPr.find(tag("pStyle"))
        if pStyle is not None:
            return pStyle.get(tag("val"), "")
    return ""

def get_num_info(para):
    pPr = para.find(tag("pPr"))
    if pPr is not None:
        numPr = pPr.find(tag("numPr"))
        if numPr is not None:
            ilvl = numPr.find(tag("ilvl"))
            numId = numPr.find(tag("numId"))
            level = int(ilvl.get(tag("val"), 0)) if ilvl is not None else 0
            nid = numId.get(tag("val"), "0") if numId is not None else "0"
            return level, nid
    return None, None

def para_text(para):
    parts = []
    for elem in para.iter():
        if elem.tag == tag("t"):
            if elem.text:
                parts.append(elem.text)
        elif elem.tag == tag("tab"):
            parts.append("\t")
        elif elem.tag == tag("br"):
            parts.append("\n")
    return "".join(parts)

def extract_tables(root):
    """Extract all tables as text blocks."""
    tables_text = []
    for tbl in root.iter(tag("tbl")):
        rows = []
        for row in tbl.iter(tag("tr")):
            cells = []
            for cell in row.iter(tag("tc")):
                cell_parts = []
                for para in cell.iter(tag("p")):
                    cell_parts.append(para_text(para))
                cells.append(" | ".join(filter(None, cell_parts)))
            rows.append(" | ".join(cells))
        tables_text.append("\n".join(rows))
    return tables_text

def extract_docx(path):
    with zipfile.ZipFile(path, "r") as z:
        with z.open("word/document.xml") as f:
            content = f.read()

    root = ET.fromstring(content)
    body = root.find(tag("body"))

    output_lines = []

    for child in body:
        if child.tag == tag("p"):
            style = get_para_style(child)
            text = para_text(child)
            level, nid = get_num_info(child)

            if not text.strip() and style == "":
                output_lines.append("")
                continue

            if style.startswith("Heading") or style.startswith("heading"):
                # Extract heading level number
                lvl_char = style.replace("Heading", "").replace("heading", "").strip()
                try:
                    lvl = int(lvl_char)
                    prefix = "#" * lvl
                except:
                    prefix = "##"
                output_lines.append("\n%s %s" % (prefix, text))
            elif level is not None:
                indent = "  " * level
                output_lines.append("%s- %s" % (indent, text))
            else:
                output_lines.append(text)

        elif child.tag == tag("tbl"):
            output_lines.append("\n[TABLE START]")
            for row in child.iter(tag("tr")):
                cells = []
                for cell in row.iter(tag("tc")):
                    cell_parts = []
                    for para in cell.iter(tag("p")):
                        pt = para_text(para)
                        if pt.strip():
                            cell_parts.append(pt)
                    cells.append(" / ".join(cell_parts) if cell_parts else "")
                output_lines.append("| " + " | ".join(cells) + " |")
            output_lines.append("[TABLE END]\n")

    return "\n".join(output_lines)

if __name__ == "__main__":
    path = sys.argv[1]
    result = extract_docx(path)
    print(result)
