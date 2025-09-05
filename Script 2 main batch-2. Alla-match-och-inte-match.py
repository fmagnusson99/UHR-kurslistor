import pytesseract
from PIL import Image
import csv
import re
import os
import fitz  # PyMuPDF
import glob

# Define input and output file paths
pdf_folder = "."  # Current directory - you can change this to a specific folder
csv_path = "All-emails-extracted.csv"
images_folder = "images"  # Folder to store all PNG images

def create_images_folder():
    """Create the images folder if it doesn't exist"""
    if not os.path.exists(images_folder):
        os.makedirs(images_folder)
        print(f"Created images folder: {images_folder}")

def extract_images_from_pdf(pdf_path):
    """Extract images from PDF and save them as PNG files in the images folder"""
    doc = fitz.open(pdf_path)
    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
    print(f"Processing PDF: {pdf_name} ({len(doc)} pages)")
    
    image_files = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        print(f"  Processing page {page_num + 1}...")
        
        # Get the whole page as image
        try:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            image_path = os.path.join(images_folder, f"{pdf_name}_page_{page_num + 1}.png")
            pix.save(image_path)
            image_files.append(image_path)
            print(f"    Saved: {image_path}")
        except Exception as e:
            print(f"    Error saving page {page_num + 1}: {e}")
    
    doc.close()
    return image_files

def extract_text_from_image(image_path):
    """Extract text from an image using OCR"""
    try:
        image = Image.open(image_path)
        
        # Configure OCR for better table recognition
        custom_config = r'--oem 3 --psm 6 -c preserve_interword_spaces=1'
        text = pytesseract.image_to_string(image, lang='eng', config=custom_config)
        
        # Debug: Show the raw OCR output
        lines = text.split('\n')
        print(f"    OCR extracted {len(lines)} lines")
        if lines:
            print(f"    First 5 OCR lines:")
            for i, line in enumerate(lines[:5]):
                print(f"      {i+1}: '{line}'")
        
        return text
    except Exception as e:
        print(f"Error processing {image_path}: {e}")
        return ""

def parse_table_data(text):
    """Parse table-structured data from OCR text - extract university, first name, and email"""
    lines = text.split('\n')
    
    # Patterns for matching
    university_pattern = r'\b[A-Z]{2,3}\b'  # Two or three uppercase letters (GU, JU, LNU, UMU etc.)
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    
    print(f"    OCR extracted {len(lines)} lines")
    
    # Try to reconstruct table structure from OCR lines
    extracted_rows = []
    
    # Look for lines that contain both university code and email
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Check if line contains both university code and email
        university_match = re.search(university_pattern, line)
        email_match = re.search(email_pattern, line)
        
        if university_match and email_match:
            university = university_match.group()
            email = email_match.group()
            
            # Extract text between university and email - this should be the first name
            start_pos = university_match.end()
            end_pos = email_match.start()
            
            if start_pos < end_pos:
                name_text = line[start_pos:end_pos].strip()
                
                # Clean up the name text
                name_text = re.sub(r'[^\w\s-]', '', name_text)  # Remove special chars except spaces and hyphens
                name_text = re.sub(r'\s+', ' ', name_text)  # Normalize spaces
                name_text = name_text.strip()
                
                # Split and take first part as first name
                name_parts = name_text.split()
                if name_parts:
                    first_name = name_parts[0]
                    
                    # Validate first name (should be reasonable length and not look like email)
                    if (len(first_name) >= 2 and 
                        len(first_name) <= 20 and 
                        '@' not in first_name and
                        '.' not in first_name):
                        
                        extracted_rows.append({
                            'university': university,
                            'first_name': first_name,
                            'email': email
                        })
                        print(f"    Matched: {university}, {first_name}, {email}")
                    else:
                        print(f"    Invalid name format: {university}, '{first_name}', {email}")
                else:
                    print(f"    No name found between university and email: {university}, {email}")
            else:
                print(f"    University and email too close: {university}, {email}")
    
    # If no structured rows found, try alternative approach
    if not extracted_rows:
        print(f"    No structured rows found, trying alternative parsing...")
        
        # Extract all universities and emails
        universities = []
        emails = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Find universities
            uni_matches = re.findall(university_pattern, line)
            universities.extend(uni_matches)
            
            # Find emails
            email_matches = re.findall(email_pattern, line)
            emails.extend(email_matches)
        
        print(f"    Found {len(universities)} university codes and {len(emails)} emails")
        
        # Use the most common university
        university = universities[0] if universities else "GU"
        
        # For each email, try to find a corresponding name in nearby lines
        for i, email in enumerate(emails):
            # Look for a name in the lines around this email
            name_found = False
            
            # Search in a window of lines around the email
            search_start = max(0, i - 3)
            search_end = min(len(lines), i + 4)
            
            for j in range(search_start, search_end):
                if j < len(lines):
                    line = lines[j].strip()
                    if not line or '@' in line or re.search(university_pattern, line):
                        continue
                    
                    # Check if this line looks like a name
                    clean_line = re.sub(r'[^\w\s-]', '', line)
                    if (len(clean_line) >= 2 and 
                        len(clean_line) <= 20 and 
                        not clean_line.isupper() and
                        not re.match(r'^[0-9]', clean_line)):
                        
                        # Verify the name appears in the email
                        if clean_line.lower() in email.lower():
                            extracted_rows.append({
                                'university': university,
                                'first_name': clean_line,
                                'email': email
                            })
                            print(f"    Alternative match: {university}, {clean_line}, {email}")
                            name_found = True
                            break
            
            if not name_found:
                print(f"    No name found for email: {email}")
    
    print(f"    Created {len(extracted_rows)} matched rows")
    return extracted_rows

def find_pdf_files():
    """Find all PDF files in the current directory"""
    pdf_files = glob.glob(os.path.join(pdf_folder, "*.pdf"))
    return pdf_files

def main():
    # Create images folder
    create_images_folder()
    
    # Find all PDF files
    pdf_files = find_pdf_files()
    
    if not pdf_files:
        print(f"No PDF files found in {pdf_folder}")
        return
    
    print(f"Found {len(pdf_files)} PDF files to process:")
    for pdf_file in pdf_files:
        print(f"  - {os.path.basename(pdf_file)}")
    
    # Process all PDFs
    all_rows = []
    total_pages = 0
    
    for pdf_file in pdf_files:
        try:
            # Extract images from PDF
            print(f"\n{'='*50}")
            print(f"Processing: {os.path.basename(pdf_file)}")
            print(f"{'='*50}")
            
            image_files = extract_images_from_pdf(pdf_file)
            total_pages += len(image_files)
            
            if not image_files:
                print(f"No images extracted from {pdf_file}")
                continue
            
            # Process each image and extract data
            pdf_rows = []
            for image_path in image_files:
                print(f"\n  Processing image: {os.path.basename(image_path)}")
                
                # Extract text from image
                text = extract_text_from_image(image_path)
                
                if not text.strip():
                    print(f"    No text extracted from {os.path.basename(image_path)}")
                    continue
                
                # Parse the table data
                page_rows = parse_table_data(text)
                pdf_rows.extend(page_rows)
            
            all_rows.extend(pdf_rows)
            print(f"\n  Total rows from {os.path.basename(pdf_file)}: {len(pdf_rows)}")
        
        except Exception as e:
            print(f"Error processing {pdf_file}: {e}")
            continue
    
    # Write the extracted data to the CSV file
    try:
        with open(csv_path, mode='w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            # Write the header row with the specified column names
            writer.writerow(["university", "first_name", "email"])
            # Write all the collected data rows
            for row in all_rows:
                writer.writerow([
                    row['university'],
                    row['first_name'],
                    row['email']
                ])
        
        print(f"\n{'='*50}")
        print(f"PROCESSING COMPLETE")
        print(f"{'='*50}")
        print(f"✅ Extracted {len(all_rows)} total rows to {csv_path}")
        print(f"📄 Processed {len(pdf_files)} PDF files")
        print(f"🖼️  Created {total_pages} image files in {images_folder}/")
        
        # Show statistics
        unique_emails = len(set(row['email'] for row in all_rows))
        universities = set(row['university'] for row in all_rows)
        
        print(f"\n📊 Statistics:")
        print(f"  Total rows: {len(all_rows)}")
        print(f"  Unique emails: {unique_emails}")
        print(f"  Universities found: {', '.join(sorted(universities))}")
        
        # Show first few rows as preview
        if all_rows:
            print(f"\n📋 First few rows:")
            for i, row in enumerate(all_rows[:5]):
                print(f"  {i+1}: {row['university']}, {row['first_name']}, {row['email']}")
        
        print(f"\n💡 Tip: You can safely delete the {images_folder}/ folder after processing if you don't need the images anymore.")
    
    except IOError as e:
        print(f"Error: Could not write to CSV file '{csv_path}': {e}")
    except Exception as e:
        print(f"An unexpected error occurred while writing the CSV: {e}")

if __name__ == "__main__":
    main() 