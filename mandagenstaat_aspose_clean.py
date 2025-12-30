"""
ASPOSE.CELLS met WATERMARK REMOVER
Genereert PDF met Aspose.Cells en verwijdert daarna het watermark
"""

import io
import fitz  # PyMuPDF

# Start JVM for Aspose
try:
    import jpype
    if not jpype.isJVMStarted():
        jpype.startJVM()
    import asposecells
    from asposecells.api import Workbook, PdfSaveOptions, PdfOptimizationType
except Exception as e:
    print(f"JVM startup error: {e}")


def remove_watermark_from_pdf(pdf_bytes: bytes) -> bytes:
    """
    Verwijder watermark uit PDF
    - Zoekt naar "Evaluation Only" tekst
    - Verwijdert deze uit de PDF
    """
    # Open PDF from bytes
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        # Zoek naar "Evaluation Only" tekst en verwijder deze
        text_instances = page.search_for("Evaluation Only")
        
        for inst in text_instances:
            # Maak een wit rechthoek over de tekst
            page.add_redact_annot(inst, fill=(1, 1, 1))
        
        # Pas redactions toe
        page.apply_redactions()
        
        # Alternatieve methode: verwijder uit content stream
        try:
            # Get page content
            cont = page.read_contents()
            
            # Zoek en verwijder watermark gerelateerde tekst
            watermark_strings = [
                b"Evaluation Only",
                b"Created with Aspose.Cells",
                b"Copyright 2003 - 2025 Aspose Pty Ltd"
            ]
            
            for watermark_str in watermark_strings:
                if watermark_str in cont:
                    # Vervang door lege string
                    cont = cont.replace(watermark_str, b"")
            
            # Update page content
            page.set_contents(cont)
        except Exception as e:
            print(f"Warning: Could not clean content stream: {e}")
    
    # Save cleaned PDF
    output_bytes = io.BytesIO()
    doc.save(output_bytes)
    output_bytes.seek(0)
    doc.close()
    
    return output_bytes.getvalue()


def create_pdf_with_aspose_clean(project: dict, user_week_data: dict, start_date, end_date) -> io.BytesIO:
    """
    Maak PDF met Aspose.Cells en verwijder watermark
    """
    # Eerst Excel maken
    from mandagenstaat_template_based import create_from_template
    
    excel_file = create_from_template(project, user_week_data, start_date, end_date)
    
    # Sla Excel tijdelijk op
    import tempfile
    import os
    import subprocess
    
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx', mode='wb') as tmp:
        tmp.write(excel_file.read())
        excel_path = tmp.name
    
    try:
        # Find Java home
        java_cmd = subprocess.check_output(['which', 'java']).decode().strip()
        java_real = subprocess.check_output(['readlink', '-f', java_cmd]).decode().strip()
        java_home = os.path.dirname(os.path.dirname(java_real))
        os.environ['JAVA_HOME'] = java_home
        
        # Start JVM with explicit path
        import jpype
        if not jpype.isJVMStarted():
            # Find libjvm.so
            libjvm_paths = [
                f"{java_home}/lib/server/libjvm.so",
                f"{java_home}/lib/aarch64/server/libjvm.so",  # ARM
                f"{java_home}/jre/lib/amd64/server/libjvm.so"
            ]
            
            libjvm = None
            for path in libjvm_paths:
                if os.path.exists(path):
                    libjvm = path
                    break
            
            if libjvm:
                jpype.startJVM(libjvm)
            else:
                jpype.startJVM()  # Try default
        
        import asposecells
        from asposecells.api import Workbook, PdfSaveOptions, PdfOptimizationType
        
        # Laad Excel met Aspose.Cells
        workbook = Workbook(excel_path)
        
        # PDF save options
        pdf_options = PdfSaveOptions()
        pdf_options.setAllColumnsInOnePagePerSheet(True)
        pdf_options.setOptimizationType(PdfOptimizationType.STANDARD)
        
        # Convert naar PDF (met watermark)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf', mode='wb') as tmp_pdf:
            pdf_path = tmp_pdf.name
        
        workbook.save(pdf_path, pdf_options)
        
        # Lees PDF
        with open(pdf_path, 'rb') as f:
            pdf_with_watermark = f.read()
        
        # Cleanup temp PDF
        os.unlink(pdf_path)
        
        # VERWIJDER WATERMARK
        clean_pdf_bytes = remove_watermark_from_pdf(pdf_with_watermark)
        
        return io.BytesIO(clean_pdf_bytes)
        
    finally:
        # Cleanup Excel
        try:
            os.unlink(excel_path)
        except:
            pass
