"""
PLAYWRIGHT PDF GENERATOR voor Mandagenstaat
Genereert PDF met scale=1, geen auto-shrink, exact A4 formaat
"""

import asyncio
from playwright.async_api import async_playwright
import os
import base64
from datetime import datetime


async def generate_pdf_from_html(html_content: str, output_path: str = None) -> bytes:
    """
    Genereer PDF vanuit HTML met Playwright
    - scale=1 (geen verkleining)
    - preferCSSPageSize=true
    - printBackground=true
    - A4 format met exacte marges
    """
    async with async_playwright() as p:
        # Launch browser met expliciete executable path
        browser = await p.chromium.launch(
            headless=True,
            executable_path='/pw-browsers/chromium_headless_shell-1187/chrome-linux/headless_shell'
        )
        page = await browser.new_page()
        
        # Set HTML content
        await page.set_content(html_content, wait_until='networkidle')
        
        # Generate PDF met exacte settings
        pdf_bytes = await page.pdf(
            format='A4',  # A4 formaat
            print_background=True,  # Print achtergronden/borders
            prefer_css_page_size=True,  # Respecteer @page CSS
            scale=1.0,  # GEEN scaling - 100%
            margin={
                'top': '15mm',
                'bottom': '15mm', 
                'left': '12mm',
                'right': '12mm'
            },
            display_header_footer=False
        )
        
        await browser.close()
        
        if output_path:
            with open(output_path, 'wb') as f:
                f.write(pdf_bytes)
        
        return pdf_bytes


def create_mandagenstaat_html(project: dict, user_week_data: dict) -> str:
    """
    Genereer HTML voor Mandagenstaat met exacte styling
    - Lichtgrijze borders #D9D9D9
    - Logo embedded als base64
    - Exacte kolom breedtes
    - Datum dd-mm-yyyy, Plaats Utrecht
    """
    
    # Week info
    now = datetime.now()
    week_num = now.isocalendar()[1]
    year = now.year
    runtime_date = now.strftime("%d-%m-%Y")
    
    # Logo inladen als base64
    logo_path = '/app/backend/logo.png'
    logo_b64 = ""
    if os.path.exists(logo_path):
        with open(logo_path, 'rb') as f:
            logo_b64 = base64.b64encode(f.read()).decode()
    
    # Bereken totalen
    day_totals = [0, 0, 0, 0, 0, 0, 0]
    for user_name, data in user_week_data.items():
        for i, hours in enumerate(data['days']):
            day_totals[i] += int(round(hours)) if hours > 0 else 0
    
    grand_total = sum(day_totals)
    
    # HTML genereren
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            @page {{
                size: A4 portrait;
                margin: 15mm 12mm;
            }}
            
            * {{
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }}
            
            body {{
                font-family: Arial, sans-serif;
                font-size: 10pt;
                line-height: 1.2;
                color: #000;
            }}
            
            .container {{
                width: 100%;
                max-width: 210mm;
            }}
            
            /* Header met logo en company name */
            .header {{
                display: flex;
                justify-content: space-between;
                align-items: flex-start;
                margin-bottom: 8mm;
            }}
            
            .logo {{
                width: 60mm;
                height: auto;
            }}
            
            .company-name {{
                font-size: 11pt;
                font-weight: bold;
                text-align: right;
                line-height: 1.3;
            }}
            
            /* Project info tabel */
            .project-info {{
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 8mm;
            }}
            
            .project-info td {{
                border: 0.5pt solid #D9D9D9;
                padding: 3mm 2mm;
                font-size: 9pt;
            }}
            
            .project-info td:first-child {{
                width: 35%;
                font-weight: 500;
            }}
            
            /* Uren tabel */
            .time-table {{
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 6mm;
            }}
            
            .time-table th,
            .time-table td {{
                border: 0.5pt solid #D9D9D9;
                padding: 2mm 1.5mm;
                text-align: center;
                font-size: 9pt;
            }}
            
            .time-table th {{
                font-weight: 600;
                background-color: #fff;
            }}
            
            .time-table td:first-child {{
                text-align: left;
                width: 50mm;
            }}
            
            .time-table td:nth-child(2) {{
                width: 25mm;
            }}
            
            .time-table td:nth-child(3) {{
                width: 18mm;
            }}
            
            .time-table .day-col {{
                width: 12mm;
            }}
            
            .time-table .total-row {{
                font-weight: 600;
            }}
            
            /* Footer info */
            .footer-info {{
                margin-bottom: 8mm;
                font-size: 9pt;
            }}
            
            .footer-info div {{
                margin-bottom: 2mm;
            }}
            
            .footer-info span:first-child {{
                font-weight: 600;
                margin-right: 3mm;
            }}
            
            /* Signatures */
            .signatures {{
                display: flex;
                justify-content: space-between;
                margin-top: 10mm;
            }}
            
            .signature-block {{
                width: 48%;
            }}
            
            .signature-label {{
                font-size: 9pt;
                font-weight: 600;
                margin-bottom: 2mm;
            }}
            
            .signature-box {{
                border: 0.5pt solid #D9D9D9;
                height: 20mm;
                background-color: #fff;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <!-- Header -->
            <div class="header">
                <img src="data:image/png;base64,{logo_b64}" class="logo" alt="Logo">
                <div class="company-name">
                    The Global Bedrijfsdiensten BV
                </div>
            </div>
            
            <!-- Project Info -->
            <table class="project-info">
                <tr>
                    <td>Naam opdrachtgever</td>
                    <td>{project.get('company', '')}</td>
                </tr>
                <tr>
                    <td>Project</td>
                    <td>{project.get('name', '')}</td>
                </tr>
                <tr>
                    <td>Weeknummer/ jaar</td>
                    <td>W{week_num}/{year}</td>
                </tr>
                <tr>
                    <td>Soort werkzaamheden</td>
                    <td>{project.get('description', '')}</td>
                </tr>
            </table>
            
            <!-- Time Table -->
            <table class="time-table">
                <thead>
                    <tr>
                        <th>Naam</th>
                        <th>BSN</th>
                        <th>week nummer</th>
                        <th class="day-col">ma</th>
                        <th class="day-col">di</th>
                        <th class="day-col">wo</th>
                        <th class="day-col">do</th>
                        <th class="day-col">vrij</th>
                        <th class="day-col">zat</th>
                        <th class="day-col">zon</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    # Data rijen
    for user_name, data in sorted(user_week_data.items()):
        # Afkorten naam
        name_parts = user_name.split(' ', 1)
        if len(name_parts) > 1:
            abbreviated_name = f"{name_parts[0][0]}. {name_parts[1]}"
        else:
            abbreviated_name = user_name
        
        html += f"""
                    <tr>
                        <td>{abbreviated_name}</td>
                        <td>{data.get('bsn', '')}</td>
                        <td>{week_num}</td>
        """
        
        for hours in data['days']:
            html += f"<td class='day-col'>{int(round(hours)) if hours > 0 else 0}</td>"
        
        html += """
                    </tr>
        """
    
    # Lege rijen (max 15 rows total)
    current_rows = len(user_week_data)
    for _ in range(15 - current_rows):
        html += """
                    <tr>
                        <td>&nbsp;</td>
                        <td></td>
                        <td></td>
                        <td class='day-col'></td>
                        <td class='day-col'></td>
                        <td class='day-col'></td>
                        <td class='day-col'></td>
                        <td class='day-col'></td>
                        <td class='day-col'></td>
                        <td class='day-col'></td>
                    </tr>
        """
    
    # Totalen rij
    html += f"""
                    <tr class="total-row">
                        <td colspan="3">TOTAAL</td>
    """
    
    for total in day_totals:
        html += f"<td class='day-col'>{total}</td>"
    
    html += f"""
                    </tr>
                </tbody>
            </table>
            
            <!-- Footer Info -->
            <div class="footer-info">
                <div><span>Datum:</span> {runtime_date}</div>
                <div><span>Plaats:</span> Utrecht</div>
            </div>
            
            <!-- Signatures -->
            <div class="signatures">
                <div class="signature-block">
                    <div class="signature-label">Accoord Uitvoerder</div>
                    <div class="signature-box"></div>
                </div>
                <div class="signature-block">
                    <div class="signature-label">Accoord The Global</div>
                    <div class="signature-box"></div>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    
    return html


def create_pdf_playwright(project: dict, user_week_data: dict) -> bytes:
    """
    Wrapper functie - gebruik nest_asyncio voor FastAPI compatibility
    """
    import nest_asyncio
    nest_asyncio.apply()
    
    html = create_mandagenstaat_html(project, user_week_data)
    
    # Run async functie in current event loop
    try:
        loop = asyncio.get_event_loop()
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
    
    pdf_bytes = loop.run_until_complete(
        generate_pdf_from_html(html)
    )
    
    return pdf_bytes
