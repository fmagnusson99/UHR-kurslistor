import pytesseract
from PIL import Image
import csv
import re
import os
import fitz  # PyMuPDF
import glob

# Define input and output file paths
pdf_folder = "."  # Current directory - you can change this to a specific folder
csv_path = "course-emails-all-extracted.csv"
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
        
        # Get the whole page as image with high resolution for better OCR
        try:
            # Use 3x scaling for better OCR quality
            pix = page.get_pixmap(matrix=fitz.Matrix(3, 3))
            image_path = os.path.join(images_folder, f"{pdf_name}_page_{page_num + 1}.png")
            pix.save(image_path)
            image_files.append(image_path)
            print(f"    💾 Saved: {os.path.basename(image_path)}")
        except Exception as e:
            print(f"    ❌ Error saving page {page_num + 1}: {e}")
    
    doc.close()
    return image_files

def extract_text_from_image(image_path):
    """Extract text from an image using OCR with optimized settings for structured data"""
    try:
        image = Image.open(image_path)
        
        # Try multiple OCR configurations to get the best results
        configs = [
            r'--oem 3 --psm 6 -c preserve_interword_spaces=1',  # Uniform block of text
            r'--oem 3 --psm 4 -c preserve_interword_spaces=1',  # Single column of text
            r'--oem 3 --psm 3 -c preserve_interword_spaces=1',  # Fully automatic page segmentation
            r'--oem 1 --psm 6 -c preserve_interword_spaces=1',  # Legacy engine with uniform block
            r'--oem 3 --psm 1 -c preserve_interword_spaces=1',  # Automatic page segmentation with OSD
            r'--oem 3 --psm 12 -c preserve_interword_spaces=1', # Sparse text
        ]
        
        best_text = ""
        best_score = 0
        all_configs_results = []
        
        for i, config in enumerate(configs):
            try:
                text = pytesseract.image_to_string(image, lang='eng', config=config)
                lines = [line.strip() for line in text.split('\n') if line.strip()]
                
                print(f"    Config {i+1} extracted {len(lines)} non-empty lines")
                
                # Score based on finding course codes and emails
                course_code_pattern = r'\b[A-Z]{2,3}\s*-\s*\d{5}\b'
                email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
                
                valid_lines = 0
                total_emails = 0
                total_course_codes = 0
                
                for line in lines:
                    has_course = bool(re.search(course_code_pattern, line))
                    has_email = bool(re.search(email_pattern, line))
                    
                    if has_course and has_email:
                        valid_lines += 3  # Both found - highest score
                    elif has_course or has_email:
                        valid_lines += 1  # One found
                    
                    if has_email:
                        total_emails += 1
                    if has_course:
                        total_course_codes += 1
                
                # Bonus points for having many emails (target ~46 lines)
                if total_emails >= 40:
                    valid_lines += 10  # Bonus for high email count
                elif total_emails >= 30:
                    valid_lines += 5
                
                score = valid_lines
                all_configs_results.append({
                    'config': config,
                    'text': text,
                    'lines': lines,
                    'score': score,
                    'emails': total_emails,
                    'course_codes': total_course_codes
                })
                
                print(f"    Config {i+1}: {len(lines)} lines, {total_emails} emails, {total_course_codes} course codes, score: {score}")
                
                if score > best_score:
                    best_text = text
                    best_score = score
                    print(f"    ✅ Config {i+1} is currently the best")
                    
            except Exception as e:
                print(f"    ❌ Config {i+1} failed: {e}")
                continue
        
        # If we have multiple good results, try to combine them
        if len(all_configs_results) > 1:
            # Sort by score and take top 3
            sorted_results = sorted(all_configs_results, key=lambda x: x['score'], reverse=True)
            top_results = sorted_results[:3]
            
            # Try to combine results from top configurations
            combined_emails = set()
            combined_course_codes = set()
            combined_lines = []
            
            for result in top_results:
                if result['score'] > 0:
                    # Extract emails and course codes from this result
                    for line in result['lines']:
                        emails_in_line = re.findall(email_pattern, line)
                        course_codes_in_line = re.findall(course_code_pattern, line)
                        
                        # Add unique emails and course codes
                        for email in emails_in_line:
                            combined_emails.add(email)
                        for course_code in course_codes_in_line:
                            combined_course_codes.add(course_code)
                        
                        # Add line if it has either email or course code
                        if emails_in_line or course_codes_in_line:
                            combined_lines.append(line)
            
            print(f"    🔄 Combined approach: {len(combined_emails)} unique emails, {len(combined_course_codes)} unique course codes")
            
            # If combined approach gives us more emails, use it
            if len(combined_emails) > best_score // 3:  # Rough estimate
                best_text = '\n'.join(combined_lines)
                print(f"    ✅ Using combined approach with {len(combined_emails)} emails")
        
        # Debug: Show the raw OCR output from best config
        lines = best_text.split('\n')
        print(f"    📊 Best OCR extracted {len(lines)} total lines")
        if lines:
            print(f"    📋 Sample lines:")
            for i, line in enumerate(lines[:10]):
                print(f"      {i+1}: '{line}'")
        
        return best_text
    except Exception as e:
        print(f"❌ Error processing {image_path}: {e}")
        return ""

def parse_course_data(text):
    """Parse course data from OCR text - extract ALL course_code, first_name, and email rows"""
    lines = text.split('\n')
    
    # Patterns for matching
    course_code_pattern = r'\b[A-Z]{2,3}\s*-\s*\d{5}\b'  # Pattern like GU-74007, GU -03048, etc.
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    
    print(f"    🔍 OCR extracted {len(lines)} lines")
    
    # Debug: Show all lines for analysis
    print(f"    📋 All OCR lines:")
    for i, line in enumerate(lines):
        if line.strip():
            print(f"      {i+1}: '{line}'")
    
    extracted_rows = []
    processed_emails = set()
    
    # METHOD 1: Process every line that contains both course code and email
    print(f"    🎯 Method 1: Looking for lines with both course code and email...")
    for line_num, line in enumerate(lines, 1):
        line = line.strip()
        if not line:
            continue
        
        # Check if line contains both course code and email
        course_match = re.search(course_code_pattern, line)
        email_match = re.search(email_pattern, line)
        
        if course_match and email_match:
            course_code = course_match.group()
            email = email_match.group()
            
            # Skip if we already processed this email
            if email in processed_emails:
                continue
            
            # Extract ALL text between course code and email as the name
            start_pos = course_match.end()
            end_pos = email_match.start()
            
            if start_pos < end_pos:
                name_text = line[start_pos:end_pos].strip()
                
                # Clean up the name text - keep everything except special chars
                name_text = re.sub(r'[^\w\s-]', '', name_text)  # Remove special chars except spaces and hyphens
                name_text = re.sub(r'\s+', ' ', name_text)  # Normalize spaces
                name_text = name_text.strip()
                
                # Take ALL text as the name (no validation)
                if name_text:
                    # Clean up course code by removing spaces around dash
                    clean_course_code = re.sub(r'\s*-\s*', '-', course_code)
                    
                    extracted_rows.append({
                        'course_code': clean_course_code,
                        'first_name': name_text,
                        'email': email
                    })
                    processed_emails.add(email)
                    print(f"    ✅ Method 1 - Line {line_num}: {clean_course_code}, {name_text}, {email}")
                else:
                    print(f"    ⚠️  Method 1 - Line {line_num}: Empty name between {course_code} and {email}")
            else:
                print(f"    ⚠️  Method 1 - Line {line_num}: Course code and email too close: {course_code}, {email}")
    
    # METHOD 2: More aggressive - find all emails and try to match with course codes
    print(f"    🔄 Method 2: Aggressive email matching (found {len(extracted_rows)} so far)...")
    
    # Find all emails in the text
    all_emails = []
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
        
        email_matches = re.findall(email_pattern, line)
        for email in email_matches:
            if email not in processed_emails:
                all_emails.append((i, line, email))
    
    print(f"    📧 Found {len(all_emails)} unprocessed emails")
    
    for line_index, line, email in all_emails:
        course_found = None
        name_found = None
        
        # First check the same line
        course_match = re.search(course_code_pattern, line)
        if course_match:
            course_found = course_match.group()
            
            # Extract name from the same line
            email_pos = line.find(email)
            course_pos = course_match.end()
            
            if course_pos < email_pos:
                name_text = line[course_pos:email_pos].strip()
                name_text = re.sub(r'[^\w\s-]', '', name_text)
                name_text = re.sub(r'\s+', ' ', name_text).strip()
                if name_text:
                    name_found = name_text
        
        # If not found in same line, check nearby lines (expanded search)
        if not course_found or not name_found:
            for j in range(max(0, line_index-3), min(len(lines), line_index+4)):
                if j != line_index:
                    nearby_line = lines[j].strip()
                    if not nearby_line:
                        continue
                    
                    # Check for course code in nearby line
                    nearby_course = re.search(course_code_pattern, nearby_line)
                    if nearby_course and not course_found:
                        course_found = nearby_course.group()
                    
                    # Check for name in nearby line (not email, not course code)
                    if '@' not in nearby_line and not re.search(course_code_pattern, nearby_line):
                        clean_nearby = re.sub(r'[^\w\s-]', '', nearby_line)
                        clean_nearby = re.sub(r'\s+', ' ', clean_nearby).strip()
                        if clean_nearby and len(clean_nearby) >= 2 and not name_found:
                            name_found = clean_nearby
        
        # Add the row if we found both course and name
        if course_found and name_found:
            clean_course_code = re.sub(r'\s*-\s*', '-', course_found)
            
            extracted_rows.append({
                'course_code': clean_course_code,
                'first_name': name_found,
                'email': email
            })
            processed_emails.add(email)
            print(f"    ✅ Method 2 - Line {line_index+1}: {clean_course_code}, {name_found}, {email}")
        else:
            print(f"    ⚠️  Method 2 - Line {line_index+1}: Incomplete match for {email} (course: {course_found}, name: {name_found})")
    
    # METHOD 3: Ultra-aggressive - try to reconstruct missing data
    print(f"    🚀 Method 3: Ultra-aggressive reconstruction (found {len(extracted_rows)} so far)...")
    
    # Find all remaining emails that we haven't processed
    remaining_emails = []
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
        
        email_matches = re.findall(email_pattern, line)
        for email in email_matches:
            if email not in processed_emails:
                remaining_emails.append((i, line, email))
    
    print(f"    📧 {len(remaining_emails)} emails still unprocessed")
    
    # For remaining emails, try to find ANY course code and ANY name
    for line_index, line, email in remaining_emails:
        course_found = None
        name_found = None
        
        # Look for ANY course code in the entire text
        for search_line in lines:
            course_match = re.search(course_code_pattern, search_line)
            if course_match:
                course_found = course_match.group()
                break  # Use the first course code we find
        
        # Look for ANY reasonable name (not email, not course code, reasonable length)
        for search_line in lines:
            search_line = search_line.strip()
            if (search_line and 
                '@' not in search_line and 
                not re.search(course_code_pattern, search_line) and
                len(search_line) >= 2 and 
                len(search_line) <= 50):
                
                clean_name = re.sub(r'[^\w\s-]', '', search_line)
                clean_name = re.sub(r'\s+', ' ', clean_name).strip()
                
                if clean_name and len(clean_name) >= 2:
                    name_found = clean_name
                    break  # Use the first reasonable name we find
        
        # Add the row even with minimal data
        if course_found and name_found:
            clean_course_code = re.sub(r'\s*-\s*', '-', course_found)
            
            extracted_rows.append({
                'course_code': clean_course_code,
                'first_name': name_found,
                'email': email
            })
            processed_emails.add(email)
            print(f"    ✅ Method 3 - Line {line_index+1}: {clean_course_code}, {name_found}, {email}")
        elif email:  # Even if we can't find course/name, add the email
            # Use a default course code if we found any
            default_course = "UNKNOWN"
            for row in extracted_rows:
                if row['course_code'] != "UNKNOWN":
                    default_course = row['course_code']
                    break
            
            extracted_rows.append({
                'course_code': default_course,
                'first_name': "UNKNOWN",
                'email': email
            })
            processed_emails.add(email)
            print(f"    ⚠️  Method 3 - Line {line_index+1}: {default_course}, UNKNOWN, {email}")
    
    print(f"    📊 Total extracted rows: {len(extracted_rows)}")
    print(f"    📧 Unique emails processed: {len(processed_emails)}")
    
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
        print(f"  Average rows per page: {len(all_rows) / total_pages:.1f}" if total_pages > 0 else "  Average rows per page: N/A")
        print(f"  Target: ~46 rows per page")
        
        # Check if we're close to the target
        if total_pages > 0:
            avg_per_page = len(all_rows) / total_pages
            if avg_per_page >= 40:
                print(f"  ✅ Good extraction rate! ({avg_per_page:.1f} rows/page)")
            elif avg_per_page >= 30:
                print(f"  ⚠️  Moderate extraction rate ({avg_per_page:.1f} rows/page) - some data may be missing")
            else:
                print(f"  ❌ Low extraction rate ({avg_per_page:.1f} rows/page) - significant data may be missing")
        
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
