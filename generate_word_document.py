from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK
from pathlib import Path

OUTPUT = Path(__file__).parent / 'Tactical_IoMT_Soldier_Monitoring_System.docx'
NAVY = '122B45'
BLUE = '1E5E87'
GREEN = '25815F'
AMBER = 'BD761C'
ORANGE = 'D87825'
RED = 'C54A4A'
GRAY = '718091'
LIGHT = 'F4F6F8'
LINE = 'DDE5EA'


def shade(cell, color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tc_pr.append(shd)
    shd.set(qn('w:fill'), color)


def set_cell_text(cell, text, bold=False, color=None, size=9):
    cell.text = ''
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def set_cell_border(cell, color=LINE, size='4'):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in('w:tcBorders')
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tc_pr.append(borders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        tag = 'w:' + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn('w:val'), 'single')
        element.set(qn('w:sz'), size)
        element.set(qn('w:color'), color)


def add_table(doc, headers, rows, widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    for index, header in enumerate(headers):
        cell = table.rows[0].cells[index]
        shade(cell, NAVY)
        set_cell_text(cell, header, True, 'FFFFFF', 8)
    for row in rows:
        cells = table.add_row().cells
        for index, value in enumerate(row):
            set_cell_text(cells[index], value, False, None, 8)
            if index == 0:
                for run in cells[index].paragraphs[0].runs:
                    run.bold = True
                    run.font.color.rgb = RGBColor.from_string(NAVY)
            set_cell_border(cells[index])
    if widths:
        for row in table.rows:
            for index, width in enumerate(widths):
                row.cells[index].width = Inches(width)
    doc.add_paragraph()
    return table


def add_badge_text(paragraph, label, color):
    run = paragraph.add_run(label)
    run.bold = True
    run.font.color.rgb = RGBColor.from_string(color)
    return run


def add_heading(doc, text, level=1):
    paragraph = doc.add_heading(text, level=level)
    if level == 1:
        paragraph.runs[0].font.color.rgb = RGBColor.from_string(NAVY)
    elif level == 2:
        paragraph.runs[0].font.color.rgb = RGBColor.from_string(BLUE)
    return paragraph


def add_bullet(doc, text):
    paragraph = doc.add_paragraph(style='List Bullet')
    paragraph.add_run(text)
    return paragraph


def add_callout(doc, title, text, color=BLUE):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    shade(cell, 'F5F8FA')
    set_cell_border(cell, color=color, size='12')
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(title + '\n')
    run.bold = True
    run.font.color.rgb = RGBColor.from_string(color)
    run.font.size = Pt(10)
    run = paragraph.add_run(text)
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor.from_string(GRAY)
    doc.add_paragraph()


def add_page_break(doc):
    doc.add_page_break()


def configure_document(doc):
    section = doc.sections[0]
    section.top_margin = Inches(.65)
    section.bottom_margin = Inches(.65)
    section.left_margin = Inches(.7)
    section.right_margin = Inches(.7)
    styles = doc.styles
    styles['Normal'].font.name = 'Aptos'
    styles['Normal'].font.size = Pt(9)
    styles['Normal'].font.color.rgb = RGBColor.from_string('283846')
    for style_name, size, color in [('Title', 28, NAVY), ('Heading 1', 19, NAVY), ('Heading 2', 13, BLUE), ('Heading 3', 10, NAVY)]:
        style = styles[style_name]
        style.font.name = 'Aptos Display' if style_name in ('Title', 'Heading 1', 'Heading 2') else 'Aptos'
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run('TACTICAL IoMT | Soldier Monitoring System | Frontend Demonstration').font.size = Pt(8)


def build_document():
    doc = Document()
    configure_document(doc)

    # Cover
    cover = doc.add_paragraph()
    cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover.add_run('\n\n').font.size = Pt(8)
    title = cover.add_run('TACTICAL IoMT')
    title.bold = True
    title.font.name = 'Aptos Display'
    title.font.size = Pt(34)
    title.font.color.rgb = RGBColor.from_string(NAVY)
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run('Soldier Monitoring System')
    run.font.size = Pt(19)
    run.bold = True
    run.font.color.rgb = RGBColor.from_string(BLUE)
    line = doc.add_paragraph()
    line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line.add_run('Real-Time Battlefield Monitoring and Alert System').italic = True
    doc.add_paragraph('\n')
    cover_table = doc.add_table(rows=4, cols=2)
    cover_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cover_data = [('Project type', 'Frontend demonstration'), ('Technology', 'HTML5, CSS3, JavaScript'), ('Communication path', 'ESP32 Edge Processing / LoRa / MQTT'), ('Status', 'Simulated command-center interface')]
    for index, (key, value) in enumerate(cover_data):
        set_cell_text(cover_table.cell(index, 0), key, True, NAVY, 9)
        set_cell_text(cover_table.cell(index, 1), value, False, None, 9)
        shade(cover_table.cell(index, 0), 'E7F0F5')
        set_cell_border(cover_table.cell(index, 0)); set_cell_border(cover_table.cell(index, 1))
    doc.add_paragraph('\n')
    add_callout(doc, 'DEMO ENVIRONMENT', 'This document describes the working frontend and its simulated data. It does not claim live field measurements or backend connectivity.', AMBER)
    doc.add_paragraph('Prepared: 13 September 2026').alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_break(doc)

    # Contents and architecture
    add_heading(doc, '1. Project Overview', 1)
    doc.add_paragraph('Tactical IoMT is a smart Internet of Military Things monitoring concept. Wearable sensor nodes collect soldier and environmental information. An ESP32 performs edge-level processing, LoRa provides long-range low-power communication, and MQTT supports real-time data exchange with the command center.')
    add_heading(doc, 'System Architecture', 2)
    add_table(doc, ['Stage', 'Role', 'Technology'], [
        ('1', 'Wearable Sensor Node', 'MLX90614, MAX30102, MPU6050, MQ-2, NEO-6M'),
        ('2', 'Edge Processing', 'ESP32 local processing and filtering'),
        ('3', 'Long-range communication', 'LoRa'),
        ('4', 'Real-time data exchange', 'MQTT'),
        ('5', 'Monitoring and alerts', 'Command center frontend')
    ], [0.5, 2.4, 3.8])
    add_heading(doc, 'Frontend Navigation', 2)
    for item in ['Dashboard', 'Soldier Details', 'Live Monitoring', 'Sensor Data', 'Alerts', 'Hazardous Gas Database', 'Network Status', 'Reports']:
        add_bullet(doc, item)
    add_callout(doc, 'Implementation note', 'The application is a single-page frontend. Navigation changes the active view without refreshing the page. JavaScript arrays provide demo data that can later be replaced by an API or backend.', BLUE)
    add_page_break(doc)

    # Dashboard
    add_heading(doc, '2. Dashboard', 1)
    doc.add_paragraph('The Dashboard is the command-center overview. It presents the current system state, severity summary, live soldier telemetry, live alerts, and the two-way notification concept.')
    add_heading(doc, 'Severity Summary', 2)
    add_table(doc, ['Category', 'Demo count', 'Meaning', 'Color'], [
        ('Normal', '08', 'Routine conditions', 'Green'),
        ('Warning', '03', 'Soldier can respond', 'Amber / yellow'),
        ('High', '02', 'Requires attention', 'Orange'),
        ('Critical Alerts', '01', 'Immediate intervention', 'Red')
    ], [1.2, 1, 2.7, 1.8])
    add_heading(doc, 'Overview Metrics', 2)
    add_table(doc, ['Metric', 'Value', 'Context'], [
        ('Active Soldiers', '12', 'Demo roster'), ('Online Nodes', '18 / 20', 'Simulated node state'), ('Average Latency', '~200 ms', 'Project documentation value'), ('Average Temperature', '37.2 deg C', 'Simulated telemetry'), ('Average Battery', '82%', 'Simulated telemetry')
    ], [1.8, 1.5, 3.4])
    add_heading(doc, 'Live Soldier Status', 2)
    add_table(doc, ['Soldier ID', 'Location', 'Temperature', 'Heart Rate', 'SpO2', 'Gas', 'Battery', 'Risk', 'Status'], [
        ('S-001', 'Zone A', '36.8 deg C', '78 BPM', '98%', 'Normal', '87%', 'Low', 'Active'),
        ('S-002', 'Zone B', '37.1 deg C', '82 BPM', '97%', 'Normal', '72%', 'Low', 'Active'),
        ('S-003', 'Zone C', '39.2 deg C', '92 BPM', '96%', 'High', '64%', 'High', 'Warning'),
        ('S-004', 'Zone A', '36.7 deg C', '76 BPM', '99%', 'Normal', '91%', 'Low', 'Active')
    ])
    add_heading(doc, 'Live Alerts and Notifications', 2)
    add_table(doc, ['Severity', 'Soldier', 'Alert', 'Time', 'Notification'], [
        ('Critical', 'S-007', 'Hazardous Gas Detected', '09:38:21', 'Soldier + Command Centre'),
        ('High', 'S-002', 'Motion / Unusual Movement', '09:31:20', 'Soldier + Command Centre'),
        ('Warning', 'S-003', 'Higher Temperature', '09:41:12', 'Soldier + Command Centre')
    ])
    add_callout(doc, 'Two-way alert flow', 'Sensor Detection -> Severity Classification -> Soldier Device (Immediate Alert) and Command Centre (Monitoring Alert). Critical alerts are represented as being sent to both destinations.', RED)
    add_page_break(doc)

    # Soldier details
    add_heading(doc, '3. Soldier Details', 1)
    doc.add_paragraph('The Soldier Details page provides a searchable roster and a View Details action for inspecting an individual soldier. The selected detail panel includes sensor readings, connection states, recent chart data, and alert channel state.')
    add_table(doc, ['Soldier ID', 'Name', 'Location', 'Battery', 'GPS', 'Risk', 'Status'], [
        ('S-001', 'Arjun Rao', 'Zone A', '87%', 'Connected', 'Low', 'Active'), ('S-002', 'Meera Singh', 'Zone B', '72%', 'Connected', 'Low', 'Active'), ('S-003', 'Kabir Khan', 'Zone C', '64%', 'Connected', 'High', 'Warning'), ('S-004', 'Nisha Menon', 'Zone A', '91%', 'Connected', 'Low', 'Active')
    ])
    add_heading(doc, 'Selected Soldier Example: S-003', 2)
    add_table(doc, ['Field', 'Value'], [
        ('Location', 'Zone C'), ('GPS', 'Connected'), ('Temperature', '39.2 deg C'), ('Heart Rate', '92 BPM'), ('SpO2', '96%'), ('Motion', 'Detected'), ('Gas Level', 'High'), ('Battery', '64%'), ('LoRa', 'Connected'), ('MQTT', 'Connected'), ('Risk Level', 'High'), ('Current Status', 'Warning')
    ], [2.5, 3.8])
    add_heading(doc, 'Alert Channel Status', 2)
    add_table(doc, ['Channel', 'State'], [('Soldier Device', 'Connected'), ('Command Centre', 'Connected'), ('Last Alert', '09:38:21'), ('Last Notification', 'Delivered')])
    add_page_break(doc)

    # Live monitoring
    add_heading(doc, '4. Live Monitoring', 1)
    doc.add_paragraph('The Live Monitoring page combines a Leaflet map with simulated soldier markers and a live telemetry list. Coordinates are demonstration values only.')
    add_table(doc, ['Soldier', 'Location', 'Temperature', 'Motion', 'Gas', 'Battery', 'Status'], [
        ('S-001', 'Zone A', '36.8 deg C', 'Stable', 'Normal', '87%', 'Active'), ('S-002', 'Zone B', '37.1 deg C', 'Detected', 'Normal', '72%', 'Active'), ('S-003', 'Zone C', '39.2 deg C', 'Detected', 'High', '64%', 'Warning'), ('S-004', 'Zone A', '36.7 deg C', 'Stable', 'Normal', '91%', 'Active')
    ])
    add_callout(doc, 'Map note', 'The map uses simulated coordinates and OpenStreetMap tiles for demonstration. It does not represent live military or field positions.', BLUE)
    add_page_break(doc)

    # Sensors
    add_heading(doc, '5. Sensor Data', 1)
    doc.add_paragraph('The Sensor Data page presents current readings for each project sensor and charts for temperature, heart rate, and battery trends.')
    add_table(doc, ['Sensor', 'Purpose', 'Demo reading', 'Status'], [
        ('MLX90614', 'Body temperature', '37.2 deg C', 'Normal'), ('MAX30102', 'Heart rate and SpO2', '82 BPM / 97%', 'Normal'), ('MPU6050', 'Motion / IMU data', 'Stable', 'Normal'), ('MQ-2', 'Gas detection', 'Normal', 'Safe'), ('NEO-6M', 'GPS location', 'Connected / Zone A', 'Connected')
    ])
    add_heading(doc, 'Charts', 2)
    add_bullet(doc, 'Temperature over time')
    add_bullet(doc, 'Heart rate over time')
    add_bullet(doc, 'Battery level over time')
    add_callout(doc, 'MQ-2 technical accuracy', 'Gas identification shown in demonstration mode. Actual MQ-2 sensor readings indicate gas/smoke presence and do not by themselves provide definitive gas-species identification.', AMBER)
    add_page_break(doc)

    # Alerts
    add_heading(doc, '6. Alert Center', 1)
    doc.add_paragraph('The Alert Center uses four distinct severity levels. Alert records include the event, location, description, action, soldier notification, command-centre notification, and current status.')
    add_table(doc, ['Severity', 'Color', 'Example condition', 'Response meaning'], [
        ('Normal', 'Green', 'Routine condition', 'No intervention required'), ('Warning', 'Amber / yellow', 'Higher Temperature', 'Soldier can check and respond'), ('High', 'Orange', 'Motion or Geofence Breach', 'Requires attention'), ('Critical', 'Red', 'Hazardous Gas Detection', 'Immediate protective action')
    ])
    add_heading(doc, 'Current Demo Alert Records', 2)
    add_table(doc, ['ID', 'Soldier', 'Alert Type', 'Severity', 'Location', 'Time', 'Status'], [
        ('A021', 'S-003', 'Higher Temperature', 'Warning', 'Zone C', '09:41:12', 'Active'), ('A020', 'S-007', 'Hazardous Gas Detected', 'Critical', 'Zone C', '09:38:21', 'Active'), ('A019', 'S-002', 'Motion / Unusual Movement', 'High', 'Zone B', '09:31:20', 'Resolved'), ('A018', 'S-004', 'Geofence Breach', 'High', 'Zone A', '09:27:51', 'Resolved')
    ])
    add_heading(doc, 'Alert Detail Example: A021', 2)
    add_table(doc, ['Field', 'Value'], [
        ('Alert ID', 'A021'), ('Soldier', 'S-003'), ('Alert', 'Higher Temperature'), ('Severity', 'WARNING'), ('Location', 'Zone C'), ('Description', 'Body temperature is above the normal operating range.'), ('Recommended Action', 'Soldier should check condition and move to a safer environment if required.'), ('Soldier Notification', 'SENT'), ('Command Centre Notification', 'SENT'), ('Status', 'ACTIVE')
    ], [2.3, 4])
    add_heading(doc, 'Alert Center Controls', 2)
    for item in ['Filter by All, Normal, Warning, High, Critical, Active, or Resolved.', 'Search by Soldier ID.', 'Search by Alert Type.', 'View an alert in a detail modal.', 'Acknowledge Alert and Mark as Resolved actions are shown in the modal.']:
        add_bullet(doc, item)
    add_page_break(doc)

    # Gas database
    add_heading(doc, '7. Hazardous Gas Database', 1)
    doc.add_paragraph('The Hazardous Gas Database is a new page in the sidebar. It provides a searchable demonstration classification list and a form for adding local gas records.')
    add_table(doc, ['Gas Name', 'Severity', 'Description', 'Recommended Action', 'Status'], [
        ('Carbon Monoxide', 'Critical', 'Toxic gas exposure', 'Immediate protective action', 'Monitoring'), ('Methane', 'Critical', 'Potentially hazardous gas', 'Immediate protective action', 'Monitoring'), ('LPG', 'Critical', 'Flammable gas', 'Move away from source', 'Monitoring'), ('Smoke', 'High', 'Possible fire or combustion', 'Investigate immediately', 'Monitoring'), ('Unknown Gas', 'Unknown', 'Gas detected but not identified', 'Treat as potentially hazardous', 'Monitoring')
    ])
    add_heading(doc, 'Demonstration Matching Logic', 2)
    for item in ['The simulated gas is checked against the database.', 'A match displays the gas name, severity, and recommended action.', 'A critical classification creates a critical alert and sends status to both Soldier Device and Command Centre.', 'An unknown result displays UNKNOWN GAS DETECTED and UNKNOWN / POTENTIALLY HAZARDOUS.', 'Gas evaluation creates a corresponding alert record in the Alert Center.']:
        add_bullet(doc, item)
    add_heading(doc, 'Critical Gas Alert Example', 2)
    add_table(doc, ['Field', 'Value'], [
        ('Alert state', 'CRITICAL ALERT'), ('Condition', 'HAZARDOUS GAS DETECTED'), ('Soldier', 'S-003'), ('Location', 'Zone C'), ('Detection Time', '10:24:18'), ('Gas', 'Carbon Monoxide (demonstration classification)'), ('Severity', 'CRITICAL'), ('Recommended Action', 'Immediate protective action required.'), ('Soldier Device', 'SENT'), ('Command Centre', 'SENT'), ('Status', 'ACTIVE')
    ], [2.4, 3.9])
    add_callout(doc, 'Important technical limitation', 'The MQ-2 sensor is used for gas or smoke presence detection. It cannot, by itself, definitively identify a gas species. Named gases in this interface are demonstration classifications only.', RED)
    add_heading(doc, 'Add Hazardous Gas Form', 2)
    add_table(doc, ['Field', 'Purpose'], [('Gas Name', 'Name used in the demonstration database'), ('Severity', 'Normal, Warning, High, Critical, or Unknown'), ('Description', 'Short explanation of the detected condition'), ('Recommended Action', 'Suggested response for the condition')])
    add_page_break(doc)

    # Network
    add_heading(doc, '8. Network Status', 1)
    doc.add_paragraph('The Network Status page presents the communication path and operational state of core components.')
    add_table(doc, ['Metric', 'Value'], [
        ('LoRa Communication', 'Connected'), ('MQTT Communication', 'Connected'), ('Sensor Nodes', '18 / 20 Online'), ('Data Transmission', 'Active'), ('Average Latency', '~200 ms'), ('Communication', 'Long Range'), ('Power Usage', 'Efficient'), ('Encryption', 'AES-128')
    ])
    add_heading(doc, 'Communication Architecture', 2)
    add_table(doc, ['From', 'Link', 'To', 'Purpose'], [('Wearable Sensor Node', 'LoRa', 'Edge Device / ESP32', 'Long-range low-power transfer'), ('Edge Device / ESP32', 'MQTT', 'Command Centre', 'Real-time data exchange and monitoring')])
    add_page_break(doc)

    # Reports
    add_heading(doc, '9. Reports', 1)
    doc.add_paragraph('The Reports page provides a printable report-style view based on the current demo data and supported project observations.')
    add_table(doc, ['Performance area', 'Reported value', 'Evidence level'], [
        ('Packet Delivery Ratio', 'Reliable', 'Qualitative project observation'), ('Latency', '~200 ms', 'Approximate project documentation value'), ('Power Consumption', 'Efficient', 'Qualitative low-power communication description'), ('Communication Range', 'Long range', 'Qualitative project capability')
    ])
    add_callout(doc, 'Accuracy note', 'The report does not invent an exact packet delivery percentage or exact communication distance. The interface uses only the values and qualitative observations supported by the project information.', BLUE)
    add_heading(doc, 'Generate Report', 2)
    doc.add_paragraph('The Generate Report button opens the browser print dialog, allowing the current report view to be printed or saved as a PDF.')
    add_page_break(doc)

    # Verification
    add_heading(doc, '10. Functionality Verification', 1)
    add_table(doc, ['Check', 'Result'], [
        ('All navigation routes', 'Passed'), ('Severity filters', 'Passed'), ('Soldier ID and alert type search', 'Passed'), ('Hazardous gas search', 'Passed'), ('Add hazardous gas form', 'Passed'), ('Critical gas evaluation', 'Passed'), ('Two-way notification display', 'Passed'), ('Alert detail modal', 'Passed'), ('Responsive mobile layout', 'Passed'), ('Browser console errors', 'None detected'), ('Source diagnostics', 'No errors found')
    ])
    add_heading(doc, 'Demonstration Notes', 2)
    for item in ['All values in the interface are simulated unless explicitly described as project documentation observations.', 'Critical alerts are visually reserved for immediate-intervention conditions such as hazardous gas detection.', 'Higher Temperature is intentionally classified as Warning rather than Critical.', 'Motion and Geofence events are classified as High.', 'The interface is designed for a college major-project presentation and can later connect to real backend or MQTT data.']:
        add_bullet(doc, item)
    add_callout(doc, 'Final system concept', 'Sensor Detection -> Severity Classification -> Gas / Threat Evaluation -> Soldier Device + Command Centre', NAVY)

    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == '__main__':
    build_document()
