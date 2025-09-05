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
        text = pytesseract.image_to_string(image, lang='eng')
        return text
    except Exception as e:
        print(f"Error processing {image_path}: {e}")
        return ""

def parse_table_data(text):
    """Parse table-structured data from OCR text - extract ALL emails and match names when possible"""
    lines = text.split('\n')
    
    # Find lines that contain course codes (5-digit numbers)
    course_code_pattern = r'\b\d{5}\b'
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    
    # Collect all lines with course codes and their content
    course_entries = []
    email_entries = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Check for course codes
        if re.search(course_code_pattern, line):
            course_entries.append(line)
        
        # Check for emails
        if re.search(email_pattern, line):
            email_entries.append(line)
    
    print(f"    Found {len(course_entries)} course entries and {len(email_entries)} email entries")
    
    # Process course entries to extract names
    name_entries = []
    for entry in course_entries:
        # Extract course code
        course_match = re.search(course_code_pattern, entry)
        if not course_match:
            continue
        
        course_code = course_match.group()
        
        # Split the entry and find the course code position
        parts = entry.split()
        course_idx = -1
        for i, part in enumerate(parts):
            if course_code in part:
                course_idx = i
                break
        
        if course_idx == -1:
            continue
        
        # Extract name parts (everything after course code)
        name_parts = parts[course_idx + 1:]
        
        # Clean name parts (remove any email-like content)
        clean_name_parts = []
        for part in name_parts:
            if not re.search(email_pattern, part) and not part.startswith('@'):
                clean_name_parts.append(part)
        
        if clean_name_parts:
            name_entries.append({
                'course_code': course_code,
                'name_parts': clean_name_parts
            })
    
    # Process email entries - extract ALL emails
    all_emails = []
    for entry in email_entries:
        email_matches = re.findall(email_pattern, entry)
        all_emails.extend(email_matches)
    
    # Remove duplicates from emails while preserving order
    unique_emails = []
    seen_emails = set()
    for email in all_emails:
        if email not in seen_emails:
            unique_emails.append(email)
            seen_emails.add(email)
    
    print(f"    Found {len(unique_emails)} unique emails")
    
    # Create rows for ALL emails
    complete_rows = []
    used_emails = set()
    
    # First pass: try to match emails with names
    for i, name_entry in enumerate(name_entries):
        course_code = name_entry['course_code']
        name_parts = name_entry['name_parts']
        
        if len(name_parts) < 1:
            continue
        
        first_name = name_parts[0]
        last_name = ""  # Always leave last name empty
        
        # Try to find a matching email by name similarity
        email = ""
        for email_addr in unique_emails:
            if email_addr in used_emails:
                continue
                
            # Check if email contains parts of the name
            email_lower = email_addr.lower()
            name_lower = first_name.lower()
            
            # More strict matching logic
            if (name_lower in email_lower or 
                any(part.lower() in email_lower for part in name_parts if len(part) > 2)):
                email = email_addr
                used_emails.add(email_addr)
                break
        
        # Determine university based on course code
        if course_code.startswith('11') or course_code.startswith('12'):
            university = "JU"  # Jönköping University
        else:
            university = "JU"  # Default
        
        # Add row with matched email
        if email:
            complete_rows.append({
                'university': university,
                'course_code': course_code,
                'first_name': first_name,
                'last_name': last_name,
                'email': email
            })
            print(f"    Matched by name: {university}, {course_code}, {first_name}, {last_name}, {email}")
    
    # Second pass: add ALL remaining unmatched emails as rows with empty fields
    remaining_emails = [email for email in unique_emails if email not in used_emails]
    
    for email in remaining_emails:
        complete_rows.append({
            'university': "",
            'course_code': "",
            'first_name': "",
            'last_name': "",
            'email': email
        })
        print(f"    Unmatched email: {email}")
    
    print(f"    Created {len(complete_rows)} total rows (all emails included)")
    return complete_rows

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
            writer.writerow(["university", "course_code", "first_name", "last_name", "email"])
            # Write all the collected data rows (including unmatched emails)
            for row in all_rows:
                writer.writerow([
                    row['university'],
                    row['course_code'],
                    row['first_name'],
                    row['last_name'],
                    row['email']
                ])
        
        print(f"\n{'='*50}")
        print(f"PROCESSING COMPLETE")
        print(f"{'='*50}")
        print(f"✅ Extracted {len(all_rows)} total rows to {csv_path}")
        print(f"📄 Processed {len(pdf_files)} PDF files")
        print(f"🖼️  Created {total_pages} image files in {images_folder}/")
        
        # Show statistics
        matched_rows = len([row for row in all_rows if row['university'] and row['course_code'] and row['first_name']])
        unmatched_emails = len([row for row in all_rows if not row['university'] and not row['course_code'] and not row['first_name']])
        unique_emails = len(set(row['email'] for row in all_rows))
        
        print(f"\n📊 Statistics:")
        print(f"  Total rows: {len(all_rows)}")
        print(f"  Rows with complete data: {matched_rows}")
        print(f"  Rows with email only: {unmatched_emails}")
        print(f"  Unique emails: {unique_emails}")
        
        # Show first few rows as preview
        if all_rows:
            print(f"\n📋 First few rows:")
            for i, row in enumerate(all_rows[:5]):
                if row['university']:
                    print(f"  {i+1}: {row['university']}, {row['course_code']}, {row['first_name']}, {row['last_name']}, {row['email']}")
                else:
                    print(f"  {i+1}: [empty], [empty], [empty], [empty], {row['email']}")
        
        print(f"\n💡 Tip: You can safely delete the {images_folder}/ folder after processing if you don't need the images anymore.")
    
    except IOError as e:
        print(f"Error: Could not write to CSV file '{csv_path}': {e}")
    except Exception as e:
        print(f"An unexpected error occurred while writing the CSV: {e}")

if __name__ == "__main__":
    main() 