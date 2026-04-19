from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

SENSITIVE_DATA ={
    "ScreenXchange": "[APPLICATION]",
    "Neeyamo Enterprise Solutions":"[COMPANY]"
}

def extractDoc():
  documnet = Document('template/Vendor initiation date and time should be captured in the checklevel report.docx')

  full_text =""

  def iter_block_items(parent):
    parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
      if isinstance(child, CT_P):
        yield Paragraph(child, parent)
      elif isinstance(child, CT_Tbl):
        yield Table(child, parent)

  for block in iter_block_items(documnet):
    if isinstance(block, Paragraph):
      if block.text.strip():
        full_text += block.text + "\n"
    elif isinstance(block, Table):
      for row in block.rows:
        row_data = [cell.text.strip() for cell in row.cells]
        full_text += "\t".join(row_data) + "\n"

  print(full_text)

  sanitized_text = full_text
  for original,placeholder in SENSITIVE_DATA.items():
    sanitized_text = sanitized_text.replace(original, placeholder)

  print("============================")
  print(sanitized_text)
  return sanitized_text
