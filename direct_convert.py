import os
import tempfile
import shutil
from docx2pdf import convert
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm

def draw_quran_frame(canvas, width, height):
    """Draw a decorative frame on a page to make it look like a Quran"""
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
    canvas.setFillColorRGB(0.6, 0.4, 0.2)  # Brown/gold color
    canvas.setFont("Helvetica", 14)
    # We'll use the Unicode directly for Bismillah
    bismillah = "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ"
    canvas.drawCentredString(width/2, height-2.8*cm, bismillah)

def create_decorated_pdf(input_pdf_path, output_pdf_path):
    """Add decorative frame to each page of the PDF"""
    # Read the input PDF
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    
    # Process each page
    for page_num in range(len(reader.pages)):
        # Get the page
        page = reader.pages[page_num]
        
        # Create a new PDF with just the decorative frame
        temp_frame_path = os.path.join(tempfile.gettempdir(), f"frame_{page_num}.pdf")
        c = canvas.Canvas(temp_frame_path, pagesize=A4)
        width, height = A4
        
        # Draw the decorative frame
        draw_quran_frame(c, width, height)
        c.save()
        
        # Create a new PDF with the frame
        frame_reader = PdfReader(temp_frame_path)
        frame_page = frame_reader.pages[0]
        
        # Merge the original content with the frame
        frame_page.merge_page(page)
        writer.add_page(frame_page)
        
        # Clean up the temporary file
        os.remove(temp_frame_path)
    
    # Save the output PDF
    with open(output_pdf_path, "wb") as output_file:
        writer.write(output_file)
    
    return output_pdf_path

def create_quran_pdf(pages_folder, output_pdf):
    """
    Create a PDF from all .docx files in the pages folder.
    Each .docx file will be converted to a PDF directly to preserve all formatting and harakat.
    Then add decorative frames to make it look like a Quran.
    """
    # Get all .docx files
    docx_files = [f for f in os.listdir(pages_folder) if f.endswith('.docx')]
    docx_files.sort()  # Sort files to maintain order
    
    if not docx_files:
        print("No .docx files found in the pages folder.")
        return None
    
    # Create temporary folder for individual PDFs
    temp_dir = tempfile.mkdtemp()
    pdf_paths = []
    
    try:
        # Convert each .docx to PDF directly to preserve all formatting and harakat
        for docx_file in docx_files:
            docx_path = os.path.join(pages_folder, docx_file)
            pdf_path = os.path.join(temp_dir, f"{os.path.splitext(docx_file)[0]}.pdf")
            
            print(f"Converting {docx_file} to PDF...")
            convert(docx_path, pdf_path)
            pdf_paths.append(pdf_path)
        
        # Merge all PDFs into one
        merged_pdf_path = os.path.join(temp_dir, "merged.pdf")
        merger = PdfMerger()
        
        for pdf_path in pdf_paths:
            merger.append(pdf_path)
        
        # Write to the merged file
        print(f"Merging PDFs...")
        merger.write(merged_pdf_path)
        merger.close()
        
        # Add decorative frames to the merged PDF
        print("Adding decorative Quran frames...")
        create_decorated_pdf(merged_pdf_path, output_pdf)
        
        print(f"PDF created successfully: {output_pdf}")
        return output_pdf
        
    except Exception as e:
        print(f"Error creating PDF: {str(e)}")
        return None
    
    finally:
        # Clean up temporary files
        shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == "__main__":
    # Get the absolute path to the script's directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Define paths
    pages_folder = os.path.join(script_dir, "pages")
    output_pdf = os.path.join(script_dir, "quran.pdf")
    
    print(f"Creating PDF from files in {pages_folder}...")
    create_quran_pdf(pages_folder, output_pdf)
