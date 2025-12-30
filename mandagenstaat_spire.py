"""
SPIRE.XLS PDF GENERATOR voor Mandagenstaat
Gebruikt Spire.XLS voor Excel naar PDF conversie zonder watermark
"""

import io
from spire.xls import Workbook, FileFormat


def create_pdf_with_spire(project: dict, user_week_data: dict, start_date, end_date) -> io.BytesIO:
    """
    Maak PDF met Spire.XLS
    - Laadt Excel template
    - Vult data in
    - Converteert naar PDF zonder watermark (free tier)
    """
    # Eerst Excel maken met onze bestaande functie
    from mandagenstaat_template_based import create_from_template
    
    excel_file = create_from_template(project, user_week_data, start_date, end_date)
    
    # Sla Excel tijdelijk op
    import tempfile
    import os
    
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx', mode='wb') as tmp:
        tmp.write(excel_file.read())
        excel_path = tmp.name
    
    try:
        # Laad Excel met Spire.XLS
        workbook = Workbook()
        workbook.LoadFromFile(excel_path)
        
        # PDF conversion - fit to 1 page
        worksheet = workbook.Worksheets[0]
        worksheet.PageSetup.FitToPagesTall = 1
        worksheet.PageSetup.FitToPagesWide = 1
        
        # Save to PDF
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf', mode='wb') as tmp_pdf:
            pdf_path = tmp_pdf.name
        
        workbook.SaveToFile(pdf_path, FileFormat.PDF)
        
        # Lees PDF
        with open(pdf_path, 'rb') as f:
            pdf_bytes = io.BytesIO(f.read())
        
        # Cleanup PDF
        os.unlink(pdf_path)
        
        return pdf_bytes
        
    finally:
        # Cleanup Excel
        try:
            os.unlink(excel_path)
        except:
            pass
