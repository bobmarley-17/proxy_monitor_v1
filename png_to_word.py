#!/usr/bin/env python3
"""
PNG to Text to Word Document Converter (Fixed)
Extracts text from PNG images using OCR and saves to Word
Handles special characters and NULL bytes
"""

import os
import re
import pytesseract
from PIL import Image
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


def clean_text(text):
    """
    Remove NULL bytes and control characters from text.
    Keep only printable characters.
    """
    if not text:
        return ""
    
    # Remove NULL bytes
    text = text.replace('\x00', '')
    
    # Remove control characters except newline, tab, carriage return
    # Keep only printable ASCII and extended Unicode
    cleaned = ""
    for char in text:
        if char in '\n\t\r':
            cleaned += char
        elif ord(char) >= 32:  # Printable characters
            cleaned += char
    
    # Alternative: Use regex to remove control characters
    # cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)
    
    return cleaned.strip()


def png_to_text_to_word(folder_path, output_file="extracted_text.docx"):
    """
    Extract text from PNG files and save to Word document.
    
    Args:
        folder_path: Path to folder containing PNG files
        output_file: Output Word document filename
    """
    
    # Create a new Word document
    doc = Document()
    
    # Add title
    title = doc.add_heading('Extracted Text from Screenshots', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Get all PNG files and sort them
    png_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.png')]
    png_files.sort()
    
    if not png_files:
        print("No PNG files found in the specified folder.")
        return
    
    print(f"Found {len(png_files)} PNG files.")
    print("Extracting text (this may take a while)...\n")
    
    successful = 0
    failed = 0
    
    # Process each image
    for i, png_file in enumerate(png_files, 1):
        image_path = os.path.join(folder_path, png_file)
        
        print(f"Processing: {png_file} ({i}/{len(png_files)})", end=" ")
        
        # Add filename as heading
        doc.add_heading(png_file, level=2)
        
        try:
            # Open image and extract text
            image = Image.open(image_path)
            extracted_text = pytesseract.image_to_string(image)
            
            # Clean the text to remove problematic characters
            cleaned_text = clean_text(extracted_text)
            
            if cleaned_text:
                # Add extracted text
                paragraph = doc.add_paragraph(cleaned_text)
                paragraph.style.font.size = Pt(11)
                print(f"✓ ({len(cleaned_text)} chars)")
                successful += 1
            else:
                doc.add_paragraph("[No text detected in this image]")
                print("- (no text)")
                successful += 1
                
        except Exception as e:
            error_msg = str(e)
            # Clean error message too
            error_msg = clean_text(error_msg)
            doc.add_paragraph(f"[Error processing image: {error_msg}]")
            print(f"✗ Error: {error_msg}")
            failed += 1
        
        # Add separator
        doc.add_paragraph("─" * 50)
        doc.add_paragraph()
    
    # Save the document
    doc.save(output_file)
    print(f"\n{'='*50}")
    print(f"Document saved as: {output_file}")
    print(f"Successful: {successful}/{len(png_files)}")
    print(f"Failed: {failed}/{len(png_files)}")


if __name__ == "__main__":
    folder_path = input("Enter the folder path containing PNG files: ").strip()
    
    if not os.path.isdir(folder_path):
        print("Invalid folder path!")
    else:
        output_file = input("Enter output filename (default: extracted_text.docx): ").strip()
        if not output_file:
            output_file = "extracted_text.docx"
        if not output_file.endswith('.docx'):
            output_file += '.docx'
        
        png_to_text_to_word(folder_path, output_file)
