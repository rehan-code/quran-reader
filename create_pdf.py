import os
import tempfile
import docx
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, PageBreak, Frame, PageTemplate
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import arabic_reshaper
from bidi.algorithm import get_display

# Create a fonts directory if it doesn't exist
font_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
os.makedirs(font_dir, exist_ok=True)

# Path to Al Majeed Quranic Font (if manually installed)
# al_majeed_path = os.path.join(font_dir, "Al Majeed Quranic Font_shiped.ttf")

# Try to use Al Majeed Quranic Font, fall back to Arial if not available
try:
    if al_majeed_path and os.path.exists(al_majeed_path):
        pdfmetrics.registerFont(TTFont('AlMajeedQuranicFont', al_majeed_path))
        arabic_font = 'AlMajeedQuranicFont'
        print("Using Al Majeed Quranic Font for Quranic text...")
    else:
        # Fall back directly to Arial as requested
        arial_path = r'C:\Windows\Fonts\arial.ttf'
        if os.path.exists(arial_path):
            pdfmetrics.registerFont(TTFont('Arial', arial_path))
            arabic_font = 'Arial'
            print("Using Arial font for text...")
except Exception as e:
    print(f"Font registration error: {e}")
    arial_path = r'C:\Windows\Fonts\arial.ttf'
    if os.path.exists(arial_path):
        pdfmetrics.registerFont(TTFont('Arial', arial_path))
        arabic_font = 'Arial'
        print("Using Arial font for text...")


def is_arabic_numeral(char):
    """Check if a character is an Arabic numeral (0-9)"""
    return char.isdigit()

def extract_text_from_docx(docx_path):
    """Extract text from a .docx file"""
    doc = docx.Document(docx_path)
    text = []
    
    for para in doc.paragraphs:
        if para.text.strip():
            text.append(para.text.strip())
    
    return text

# Function to draw decorative frame on each page
def draw_quran_frame(canvas, doc):
    """Draw a decorative frame on each page to make it look like a Quran"""
    width, height = A4
    
    # Fill background with a light cream color for parchment effect
    canvas.setFillColorRGB(1.0, 0.98, 0.94)  # Very light cream
    canvas.rect(0, 0, width, height, fill=1, stroke=0)
    
    # Set the frame color to a gold/brown tone
    canvas.setStrokeColorRGB(0.6, 0.4, 0.2)  # Brown/gold color
    canvas.setLineWidth(2)
    
    # Draw outer frame with rounded corners
    canvas.roundRect(1*cm, 1*cm, width-2*cm, height-2*cm, radius=10, stroke=1, fill=0)
    
    # Draw inner frame
    canvas.roundRect(1.5*cm, 1.5*cm, width-3*cm, height-3*cm, radius=8, stroke=1, fill=0)
    
    # Draw decorative corners
    corner_size = 0.8*cm
    # Top left
    canvas.line(1*cm, 2*cm, 1*cm+corner_size, 2*cm)
    canvas.line(2*cm, 1*cm, 2*cm, 1*cm+corner_size)
    # Top right
    canvas.line(width-1*cm, 2*cm, width-1*cm-corner_size, 2*cm)
    canvas.line(width-2*cm, 1*cm, width-2*cm, 1*cm+corner_size)
    # Bottom left
    canvas.line(1*cm, height-2*cm, 1*cm+corner_size, height-2*cm)
    canvas.line(2*cm, height-1*cm, 2*cm, height-1*cm-corner_size)
    # Bottom right
    canvas.line(width-1*cm, height-2*cm, width-1*cm-corner_size, height-2*cm)
    canvas.line(width-2*cm, height-1*cm, width-2*cm, height-1*cm-corner_size)
    
    # Draw decorative divider at the top
    canvas.setStrokeColorRGB(0.6, 0.4, 0.2)
    canvas.setLineWidth(1)
    canvas.line(width/2 - 4*cm, height-2.5*cm, width/2 + 4*cm, height-2.5*cm)
    
    # Add Bismillah text at the top of each page
    bismillah = "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ"
    reshaped_bismillah = arabic_reshaper.reshape(bismillah)
    bidi_bismillah = get_display(reshaped_bismillah)
    
    canvas.setFont(arabic_font, 14)
    canvas.setFillColorRGB(0.6, 0.4, 0.2)  # Brown/gold color
    canvas.drawCentredString(width/2, height-2.8*cm, bidi_bismillah)
    

def create_quran_pdf(pages_folder, output_pdf):
    """
    Create a PDF from all .docx files in the pages folder.
    Format the text to look like a Quran with centered text and continuous lines.
    """
    # Get all .docx files
    docx_files = [f for f in os.listdir(pages_folder) if f.endswith('.docx')]
    docx_files.sort()  # Sort files to maintain order
    
    if not docx_files:
        print("No .docx files found in the pages folder.")
        return None
    
    # Create PDF document with custom page template
    doc = SimpleDocTemplate(
        output_pdf,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )
    
    # Create page template with our frame drawing function
    frame = Frame(doc.leftMargin, doc.bottomMargin, 
                 doc.width, doc.height, 
                 id='normal')
    template = PageTemplate(id='quran_template', frames=frame, onPage=draw_quran_frame)
    doc.addPageTemplates([template])
    
    # Create a centered style for Arabic text
    styles = getSampleStyleSheet()
    quran_style = ParagraphStyle(
        'QuranStyle',
        parent=styles['Normal'],
        fontName=arabic_font,
        fontSize=20,  # Larger font size for better readability
        leading=30,   # Increased line spacing
        alignment=TA_CENTER,
        spaceAfter=12,
        textColor=(0, 0, 0.4),  # Dark blue color for traditional Quran text
        allowWidows=0,
        allowOrphans=0,
        wordWrap='CJK',
        allowMarkup=1  # Enable HTML-like markup for inline styling
    )
    
    # Create a style for the page number
    page_number_style = ParagraphStyle(
        'PageNumberStyle',
        parent=styles['Normal'],
        fontName=arabic_font,
        fontSize=10,
        alignment=TA_CENTER
    )
    
    # Flowable elements to add to the PDF
    elements = []
    
    # Process each docx file
    for i, docx_file in enumerate(docx_files):
        docx_path = os.path.join(pages_folder, docx_file)
        print(f"Processing {docx_file}...")
        
        # Extract text from the docx file
        text_lines = extract_text_from_docx(docx_path)
        
        # Add a page break if not the first file
        if i > 0:
            elements.append(PageBreak())
        
        elements.append(Spacer(1, 20))
        
        # Process each line of text
        for line in text_lines:
            # Process the line to add decorative elements around ayah numbers
            processed_line = ""
            i = 0
            while i < len(line):
                if is_arabic_numeral(line[i]):
                    # Collect the full number
                    num_start = i
                    while i < len(line) and is_arabic_numeral(line[i]):
                        i += 1
                    number = line[num_start:i]
                    
                    # Add decorative element around the number in traditional Quran style
                    # Using Unicode ornate parentheses specifically designed for Quranic text
                    processed_line += f" ﴿{number}﴾ "
                else:
                    processed_line += line[i]
                    i += 1
            
            # Reshape Arabic text for proper display
                reshaped_text = arabic_reshaper.reshape(processed_line)
                bidi_text = get_display(reshaped_text)
            
            # Add the line to the PDF
            elements.append(Paragraph(bidi_text, quran_style))
    
    # Build the PDF
    doc.build(elements)
    
    print(f"PDF created successfully: {output_pdf}")
    return output_pdf

if __name__ == "__main__":
    # Get the absolute path to the script's directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Define paths
    pages_folder = os.path.join(script_dir, "pages")
    output_pdf = os.path.join(script_dir, "quran.pdf")
    
    print(f"Creating PDF from files in {pages_folder}...")
    create_quran_pdf(pages_folder, output_pdf)
