import os
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pdfplumber

def parse_ebill(pdf_path):
    """
    Intelligently extracts metadata and line items from PHED Assam Running Account Bill-C.
    Handles column shifts and extra empty cells gracefully.
    """
    metadata = {
        'contractor': '',
        'scheme': '',
        'bill_serial': '',
        'work_order_no': '',
        'bill_date': '',
        'total_upto_date': 0.0,
        'deduct_previous': 0.0,
        'net_this_bill': 0.0
    }
    
    items = []
    
    def parse_float(val_str):
        if not val_str:
            return 0.0
        cleaned = re.sub(r'[^\d.-]', '', val_str)
        try:
            return float(cleaned)
        except ValueError:
            return 0.0

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for p_idx, page in enumerate(pdf.pages):
            full_text += (page.extract_text() or "") + "\n"
            tables = page.extract_tables()
            for tbl in tables:
                for row in tbl:
                    if not row or not any(row):
                        continue
                    clean_row = [c.replace('\n', ' ').strip() if c else '' for c in row]
                    
                    # Look for item row where first column is a sequence number
                    if clean_row[0].isdigit():
                        sl = int(clean_row[0])
                        unit = clean_row[1]
                        qty = parse_float(clean_row[2])
                        desc = clean_row[3]
                        
                        # Find currency/amount tokens in remaining columns
                        num_cells = []
                        for cell in clean_row[4:]:
                            if cell:
                                # Cell has numbers or currency
                                val = parse_float(cell)
                                num_cells.append(val)
                        
                        # We expect at least: rate, upto_amt, since_amt
                        rate = num_cells[0] if len(num_cells) > 0 else 0.0
                        upto_amt = num_cells[1] if len(num_cells) > 1 else 0.0
                        since_amt = num_cells[2] if len(num_cells) > 2 else 0.0
                        
                        items.append({
                            'sl': sl,
                            'unit': unit,
                            'qty': qty,
                            'desc': desc,
                            'rate': rate,
                            'upto_amt': upto_amt,
                            'since_amt': since_amt
                        })
                            
        # Extract metadata via regex
        m_wo = re.search(r'Formal Work Order Number:\s*([^\n\r]+)', full_text)
        if m_wo: metadata['work_order_no'] = m_wo.group(1).strip()
        
        m_sch = re.search(r'Name of work-\s*([^\n\r]+)', full_text)
        if m_sch: metadata['scheme'] = m_sch.group(1).strip()
        
        m_cont = re.search(r'Name of Work Contractor or supplier\s*:\s*([^\n\r]+)', full_text)
        if m_cont: metadata['contractor'] = m_cont.group(1).strip()
        
        m_ser = re.search(r'Serial No of this Bill\s*:\s*([^\n\r]+)', full_text)
        if m_ser: metadata['bill_serial'] = m_ser.group(1).strip()
        
        m_prev = re.search(r'Deduct value of work or supplies shown on previous bill\s*₹\s*([0-9.,]+)/-', full_text)
        if m_prev: metadata['deduct_previous'] = parse_float(m_prev.group(1))
        
        m_tot = re.search(r'Total value of work’ done or supplies made to do date-.*?₹\s*([0-9.,]+)/-', full_text)
        if m_tot: metadata['total_upto_date'] = parse_float(m_tot.group(1))
        
        m_net = re.search(r'Net value of work supplies since previous bill.*?₹\s*([0-9.,]+)/-', full_text)
        if m_net: metadata['net_this_bill'] = parse_float(m_net.group(1))
        
    return metadata, items

if __name__ == '__main__':
    pdf_file = 'E-Bill__ PHED, ASSAM bengsatra.pdf'
    meta, items = parse_ebill(pdf_file)
    print("Extracted Metadata:")
    for k, v in meta.items():
        print(f"  {k}: {v}")
    print(f"\nExtracted {len(items)} items from E-bill.")
    print("\nSample Extracted Items:")
    for it in items[:6]:
        print(f"  Sl {it['sl']:2d}: Qty={it['qty']} {it['unit']} | Rate={it['rate']} | UptoAmt={it['upto_amt']} | SincePrevAmt={it['since_amt']}")
    print("...")
    for it in items[-5:]:
        print(f"  Sl {it['sl']:2d}: Qty={it['qty']} {it['unit']} | Rate={it['rate']} | UptoAmt={it['upto_amt']} | SincePrevAmt={it['since_amt']}")
