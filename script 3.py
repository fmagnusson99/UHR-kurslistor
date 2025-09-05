import pytesseract
from PIL import Image
import csv
import re
import os
import fitz  # PyMuPDF
import glob

# Define input and output file paths
pdf_folder = "."  # Current directory - you can change this to a specific folder
csv_path = "course-emails-extracted.csv"
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

def parse_course_data(text):
    """Parse course data from OCR text - extract course_code, first_name, and email"""
    lines = text.split('\n')
    
    # Patterns for matching
    course_code_pattern = r'\b[A-Z]{2,3}\s*-\s*\d{5}\b'  # Pattern like GU-74007, GU -03048, etc.
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    
    print(f"    OCR extracted {len(lines)} lines")
    
    # Debug: Show all lines for analysis
    print(f"    All OCR lines:")
    for i, line in enumerate(lines):
        if line.strip():
            print(f"      {i+1}: '{line}'")
    
    # Try to reconstruct table structure from OCR lines
    extracted_rows = []
    
    # Method 1: Look for lines that contain both course code and email
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Check if line contains both course code and email
        course_match = re.search(course_code_pattern, line)
        email_match = re.search(email_pattern, line)
        
        if course_match and email_match:
            course_code = course_match.group()
            email = email_match.group()
            
            # Extract text between course code and email - this should be the first name
            start_pos = course_match.end()
            end_pos = email_match.start()
            
            if start_pos < end_pos:
                name_text = line[start_pos:end_pos].strip()
                
                # Clean up the name text
                name_text = re.sub(r'[^\w\s-]', '', name_text)  # Remove special chars except spaces and hyphens
                name_text = re.sub(r'\s+', ' ', name_text)  # Normalize spaces
                name_text = name_text.strip()
                
                # Split and take all parts as first name (could be multiple words)
                name_parts = name_text.split()
                if name_parts:
                    first_name = ' '.join(name_parts)  # Join all parts for multi-word names
                    
                    # Validate first name (should be reasonable length and not look like email)
                    if (len(first_name) >= 2 and 
                        len(first_name) <= 50 and  # Allow longer names
                        '@' not in first_name and
                        '.' not in first_name and
                        not re.match(r'^[0-9]', first_name)):
                        
                        # Clean up course code by removing spaces around dash
                        clean_course_code = re.sub(r'\s*-\s*', '-', course_code)
                        
                        extracted_rows.append({
                            'course_code': clean_course_code,
                            'first_name': first_name,
                            'email': email
                        })
                        print(f"    Method 1 - Matched: {course_code}, {first_name}, {email}")
                    else:
                        print(f"    Method 1 - Invalid name format: {course_code}, '{first_name}', {email}")
                else:
                    print(f"    Method 1 - No name found between course code and email: {course_code}, {email}")
            else:
                print(f"    Method 1 - Course code and email too close: {course_code}, {email}")
    
    # Method 2: Extract all course codes and emails separately, then try to match them
    if len(extracted_rows) < 5:  # If we didn't get many matches, try alternative approach
        print(f"    Method 1 found {len(extracted_rows)} rows, trying Method 2...")
        
        # Extract all course codes and emails
        course_codes = []
        emails = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Find course codes
            course_matches = re.findall(course_code_pattern, line)
            course_codes.extend(course_matches)
            
            # Find emails
            email_matches = re.findall(email_pattern, line)
            emails.extend(email_matches)
        
        print(f"    Found {len(course_codes)} course codes and {len(emails)} emails")
        
        # For each email, try to find a corresponding course code and name
        for i, email in enumerate(emails):
            # Look for a course code in the lines around this email
            course_found = None
            name_found = None
            
            # Search in a window of lines around the email
            search_start = max(0, i - 5)
            search_end = min(len(lines), i + 6)
            
            for j in range(search_start, search_end):
                if j < len(lines):
                    line = lines[j].strip()
                    if not line:
                        continue
                    
                    # Check for course code
                    course_match = re.search(course_code_pattern, line)
                    if course_match:
                        course_found = course_match.group()
                    
                    # Check if this line looks like a name (not email, not course code)
                    if '@' not in line and not re.search(course_code_pattern, line):
                        clean_line = re.sub(r'[^\w\s-]', '', line)
                        if (len(clean_line) >= 2 and 
                            len(clean_line) <= 50 and  # Allow longer names
                            not re.match(r'^[0-9]', clean_line)):
                            
                            # Verify any part of the name appears in the email
                            name_words = clean_line.lower().split()
                            email_lower = email.lower()
                            if any(word in email_lower for word in name_words if len(word) >= 2):
                                name_found = clean_line
                                break
            
            if course_found and name_found:
                # Check if this combination already exists
                exists = any(row['course_code'] == course_found and row['email'] == email 
                           for row in extracted_rows)
                if not exists:
                    # Clean up course code by removing spaces around dash
                    clean_course_code = re.sub(r'\s*-\s*', '-', course_found)
                    
                    extracted_rows.append({
                        'course_code': clean_course_code,
                        'first_name': name_found,
                        'email': email
                    })
                    print(f"    Method 2 - Matched: {course_found}, {name_found}, {email}")
            else:
                print(f"    Method 2 - No complete match found for email: {email}")
    
    # Method 3: More aggressive parsing - look for any pattern that might be a row
    if len(extracted_rows) < 5:
        print(f"    Method 2 found {len(extracted_rows)} rows, trying Method 3...")
        
        # Look for any line that contains an email and try to extract data from it
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            email_match = re.search(email_pattern, line)
            if email_match:
                email = email_match.group()
                
                # Skip if we already have this email
                if any(row['email'] == email for row in extracted_rows):
                    continue
                
                # Try to find course code and name in the same line or nearby lines
                course_found = None
                name_found = None
                
                # First, check the same line
                course_match = re.search(course_code_pattern, line)
                if course_match:
                    course_found = course_match.group()
                
                # Extract potential name from the line
                email_pos = email_match.start()
                before_email = line[:email_pos].strip()
                
                # Clean up and extract name
                name_text = re.sub(r'[^\w\s-]', '', before_email)
                name_text = re.sub(r'\s+', ' ', name_text).strip()
                
                # Remove course code if present (handle both formats)
                if course_found:
                    # Remove both "GU-03048" and "GU -03048" formats
                    name_text = name_text.replace(course_found, '').strip()
                    name_text = name_text.replace(course_found.replace(' ', ''), '').strip()
                    name_text = name_text.replace(course_found.replace('-', ' - '), '').strip()
                
                name_parts = name_text.split()
                if name_parts:
                    # Try to get the full name (could be multiple words like "ERIC CHRISTOPHER")
                    potential_name = ' '.join(name_parts)  # Take all parts as the name
                    
                    # Validate the name
                    if (len(potential_name) >= 2 and 
                        len(potential_name) <= 50 and  # Allow longer names
                        '@' not in potential_name and
                        '.' not in potential_name and
                        not re.match(r'^[0-9]', potential_name)):
                        
                        # Check if any part of the name appears in the email
                        name_words = potential_name.lower().split()
                        email_lower = email.lower()
                        if any(word in email_lower for word in name_words if len(word) >= 2):
                            name_found = potential_name
                
                # If we found both, add the row
                if course_found and name_found:
                    # Clean up course code by removing spaces around dash
                    clean_course_code = re.sub(r'\s*-\s*', '-', course_found)
                    
                    extracted_rows.append({
                        'course_code': clean_course_code,
                        'first_name': name_found,
                        'email': email
                    })
                    print(f"    Method 3 - Matched: {course_found}, {name_found}, {email}")
                else:
                    print(f"    Method 3 - Partial match for email: {email} (course: {course_found}, name: {name_found})")
    
    print(f"    Created {len(extracted_rows)} matched rows total")
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
                
                # Parse the course data
                page_rows = parse_course_data(text)
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
            writer.writerow(["course_code", "first_name", "email"])
            # Write all the collected data rows
            for row in all_rows:
                writer.writerow([
                    row['course_code'],
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
        course_codes = set(row['course_code'] for row in all_rows)
        
        print(f"\n📊 Statistics:")
        print(f"  Total rows: {len(all_rows)}")
        print(f"  Unique emails: {unique_emails}")
        print(f"  Course codes found: {', '.join(sorted(course_codes))}")
        
        # Show first few rows as preview
        if all_rows:
            print(f"\n📋 First few rows:")
            for i, row in enumerate(all_rows[:5]):
                print(f"  {i+1}: {row['course_code']}, {row['first_name']}, {row['email']}")
        
        print(f"\n💡 Tip: You can safely delete the {images_folder}/ folder after processing if you don't need the images anymore.")
    
    except IOError as e:
        print(f"Error: Could not write to CSV file '{csv_path}': {e}")
    except Exception as e:
        print(f"An unexpected error occurred while writing the CSV: {e}")

if __name__ == "__main__":
    main()
