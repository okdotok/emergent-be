"""
ASPOSE.CELLS PDF GENERATOR voor Mandagenstaat
Gebruikt Aspose.Cells voor professionele Excel naar PDF conversie
Met hoge kwaliteit layout behoud
"""

import io
from datetime import datetime
import asposecells

# Start JVM
import jpype
if not jpype.isJVMStarted():
    jpype.startJVM()

from asposecells.api import Workbook, PdfSaveOptions, PdfOptimizationType


def create_pdf_with_aspose(project: dict, user_week_data: dict, start_date, end_date) -> io.BytesIO:
    """
    Maak PDF met Aspose.Cells
    - Laadt Excel template
    - Vult data in
    - Converteert naar PDF met hoge kwaliteit
    """
    # Eerst Excel maken met onze bestaande functie
    from mandagenstaat_template_based import create_from_template
    
    excel_file = create_from_template(project, user_week_data, start_date, end_date)
    
    # Sla Excel tijdelijk op
    import tempfile
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx', mode='wb') as tmp:
        tmp.write(excel_file.read())
        excel_path = tmp.name
    
    try:
        # Laad Excel met Aspose.Cells
        workbook = Workbook(excel_path)
        
        # PDF save options
        pdf_options = PdfSaveOptions()
        pdf_options.setAllColumnsInOnePagePerSheet(True)  # Fit alle kolommen op 1 pagina
        pdf_options.setOptimizationType(PdfOptimizationType.STANDARD)  # Standaard kwaliteit
        
        # Convert naar PDF
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf', mode='wb') as tmp_pdf:
            pdf_path = tmp_pdf.name
        
        workbook.save(pdf_path, pdf_options)
        
        # Lees PDF
        with open(pdf_path, 'rb') as f:
            pdf_bytes = io.BytesIO(f.read())
        
        # Cleanup PDF
        import os
        os.unlink(pdf_path)
        
        return pdf_bytes
        
    finally:
        # Cleanup Excel
        import os
        try:
            os.unlink(excel_path)
        except:
            pass
