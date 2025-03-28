from PyPDF2 import PdfReader, PdfWriter
from bs4 import BeautifulSoup
import re
import os

def extract_timesheet_data(html_file):
    """Extract student and lesson data from the HTML timesheet file."""
    try:
        with open(html_file, "r") as f:
            html = f.read()
    except FileNotFoundError:
        print(f"Error: HTML file '{html_file}' not found")
        return None, None, None
    except Exception as e:
        print(f"Error reading HTML file: {str(e)}")
        return None, None, None
    
    if not html:
        print("Error: HTML file is empty")
        return None, None, None
    
    soup = BeautifulSoup(html, 'html.parser')
    
    # Extract ABN
    abn_span = soup.find('span', id='abnLabel')
    if not abn_span:
        print("Error: ABN information not found in HTML")
        abn_number = ""
    else:
        abn_number = abn_span.text
    
    # Extract month and year
    month_year_element = soup.find('a', id='mainContentPlaceHolder_timeSheetsRepeater_selectTimesheetLinkButton_0')
    if not month_year_element:
        print("Error: Month/Year information not found in HTML")
        month, year = "", ""
    else:
        try:
            month_year_text = month_year_element.text
            month, year = month_year_text.split(" ")
        except ValueError:
            print(f"Error: Could not parse month and year from '{month_year_element.text}'")
            month, year = "", ""
    
    # Extract student information
    pattern = re.compile(r'^mainContentPlaceHolder_rptTimesheetStudents_timesheetDetails_')
    matching_elements = soup.find_all(id=pattern)
    
    if not matching_elements:
        print("Error: No student information found in HTML")
        return None, month, abn_number
    
    students_info = []
    for element in matching_elements:
        student_info = {}
        
        # Get the specific ID suffix
        id_suffix = element['id'].replace('mainContentPlaceHolder_rptTimesheetStudents_timesheetDetails_', '')
        
        # Get student name
        student_name_element = soup.find('span', id=f'mainContentPlaceHolder_rptTimesheetStudents_lblStudentName_{id_suffix}')
        if not student_name_element:
            print(f"Warning: Student name not found for element with ID suffix {id_suffix}")
            student_info['student_name'] = f"Unknown Student {id_suffix}"
        else:
            student_info['student_name'] = student_name_element.text
        
        student_info['lessons'] = []
        
        # Get lesson information
        lesson_table_element = soup.find(id=f"mainContentPlaceHolder_rptTimesheetStudents_timeSheetDetailsDataList_{id_suffix}")
        if not lesson_table_element:
            print(f"Warning: No lesson table found for student {student_info['student_name']}")
            student_info['total_hours'] = "0"
        else:
            lesson_rows = lesson_table_element.find_all('tr')
            if not lesson_rows:
                print(f"Warning: No lesson rows found for student {student_info['student_name']}")
                student_info['total_hours'] = "0"
            else:
                total_found = False
                
                for tr in lesson_rows:
                    # Extract total hours
                    if 'GroupFooter1' in tr.get('class', []):
                        hours_cell = tr.find('td', class_='hours')
                        if hours_cell:
                            hours_span = hours_cell.text.strip()
                            student_info['total_hours'] = hours_span
                            total_found = True
                    
                    # Extract lesson details
                    if 'ThinBorder' in tr.get('class', []):
                        date_cell = tr.find('td', class_='date')
                        hours_cell = tr.find('td', class_='hours')
                        
                        if not date_cell:
                            print(f"Warning: Date cell missing for a lesson of student {student_info['student_name']}")
                            continue
                            
                        if not hours_cell:
                            print(f"Warning: Hours cell missing for a lesson on {date_cell.text.strip()} for student {student_info['student_name']}")
                            continue
                        
                        date = date_cell.text.strip()
                        
                        hours_span = hours_cell.find('span')
                        if not hours_span:
                            print(f"Warning: Hours span missing for a lesson on {date} for student {student_info['student_name']}")
                            continue
                            
                        hours = hours_span.text.strip()
                        
                        # Skip entries with 0 hours
                        if hours and hours[0] != "0":
                            lesson_info = {
                                "date": date,
                                "length": hours
                            }
                            student_info['lessons'].append(lesson_info)
                
                # Set default total_hours if not found
                if not total_found:
                    try:
                        calculated_total = sum(float(lesson['length']) for lesson in student_info['lessons'])
                        student_info['total_hours'] = str(calculated_total)
                        print(f"Note: Calculated total hours for {student_info['student_name']}: {calculated_total}")
                    except (ValueError, KeyError) as e:
                        print(f"Error calculating total hours for {student_info['student_name']}: {str(e)}")
                        student_info['total_hours'] = "0"
        
        if not student_info['lessons']:
            print(f"Warning: No valid lessons found for student {student_info['student_name']}")
            
        students_info.append(student_info)
    
    return students_info, month, year, abn_number

def create_student_timesheet(template_pdf, output_pdf, field_updates):
    """Create a new PDF with updated fields for a specific student."""
    try:
        reader = PdfReader(template_pdf)
    except FileNotFoundError:
        print(f"Error: Template PDF '{template_pdf}' not found")
        return False
    except Exception as e:
        print(f"Error opening template PDF: {str(e)}")
        return False
    
    writer = PdfWriter()
    
    # Copy all pages from the original PDF to the writer
    for page in reader.pages:
        writer.add_page(page)
    
    # Get available fields for debugging
    fields = reader.get_fields()
    if not fields:
        print(f"Warning: No form fields found in the template PDF '{template_pdf}'")
    else:
        # Check if required fields exist
        missing_fields = []
        for field in ["Month", "Your ABN"]:
            if field not in fields:
                missing_fields.append(field)
        
        if missing_fields:
            print(f"Warning: Required fields missing in PDF template: {', '.join(missing_fields)}")
    
    # Check which fields from our updates actually exist in the PDF
    if fields:
        unknown_fields = [field for field in field_updates.keys() if field not in fields]
        if unknown_fields:
            print(f"Warning: The following fields don't exist in the PDF and will be ignored: {', '.join(unknown_fields)}")
            # Remove non-existent fields to prevent errors
            for field in unknown_fields:
                field_updates.pop(field)
    
    try:
        # Update all pages (usually fields are on specific pages, but this updates all to be safe)
        for page_num in range(len(writer.pages)):
            writer.update_page_form_field_values(
                writer.pages[page_num],
                field_updates
            )
        
        # Create directory if it doesn't exist
        output_dir = os.path.dirname(output_pdf)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Save the modified PDF
        with open(output_pdf, "wb") as output_file:
            writer.write(output_file)
        
        print(f"Created timesheet: {output_pdf}")
        return True
    except Exception as e:
        print(f"Error creating PDF '{output_pdf}': {str(e)}")
        return False

def main():
    # File paths
    html_file = "timesheet.html"
    template_pdf = "template.pdf"
    
    # Validate input files exist
    if not os.path.exists(html_file):
        print(f"Error: HTML timesheet file '{html_file}' not found")
        return
    
    if not os.path.exists(template_pdf):
        print(f"Error: PDF template file '{template_pdf}' not found")
        return
    
    # Extract data from HTML
    students_info, month, year, abn_number = extract_timesheet_data(html_file)

    output_dir = month + "-" + year + "-" + "timesheets"
    
    if not students_info:
        print("Error: No student information found")
        return
    
    if not month:
        print("Warning: Month information is missing")
    
    if not abn_number:
        print("Warning: ABN information is missing")
    
    print(f"Found {len(students_info)} students")
    print(f"Month: {month or 'NOT FOUND'}, ABN: {abn_number or 'NOT FOUND'}")

    try:
        os.makedirs(output_dir, exist_ok=True)
    except Exception as e:
        print(f"Error creating output directory: {str(e)}")
        return
    
    success_count = 0
    # For each student, create a new timesheet
    for student in students_info:
        student_name = student['student_name']
        
        # Skip students with no lessons
        if not student['lessons']:
            print(f"Skipping {student_name} - no lessons found")
            continue
        
        # Create a safe filename
        safe_name = ''.join(c if c.isalnum() or c in [' ', '_'] else '_' for c in student_name)
        safe_name = safe_name.replace(" ", "_")
        output_pdf = os.path.join(output_dir, f"{safe_name}_timesheet.pdf")
        
        # Create initial field updates with common information based on actual PDF fields
        field_updates = {
            "Month": month,
            "Your ABN": abn_number,
            "Your Name": "Sam Towney",  # Using the default value from the fields
            "STUDENT JOB NAME": student_name,
            "Page No": "1",  # Default page number
            "of": "1",       # Default total pages
        }
        
        # Calculate total hours
        total_hours = student['total_hours']
        field_updates["Total Job Time"] = total_hours
        
        # Add lesson information based on actual PDF field structure
        # PDF has fields for 10 lessons maximum
        max_lessons = min(len(student['lessons']), 10)
        
        for i, lesson in enumerate(student['lessons'], 1):
            if i > 10:  # Form only has fields for 10 lessons
                print(f"Warning: Student {student_name} has more than 10 lessons, only the first 10 will be included")
                break
                
            # Check for missing data
            if not lesson.get('date'):
                print(f"Warning: Missing date for lesson {i} of student {student_name}")
                lesson['date'] = "N/A"
                
            if not lesson.get('length'):
                print(f"Warning: Missing length for lesson {i} of student {student_name}")
                lesson['length'] = "0"
            
            # Format the date and time fields according to PDF structure
            date_parts = lesson['date'].split("/") if "/" in lesson['date'] else ["", "", ""]
            if len(date_parts) == 3:
                date_formatted = f"{date_parts[0].zfill(2)}  {date_parts[1].zfill(2)}  {date_parts[2][-2:]}"
            else:
                date_formatted = lesson['date']
            
            time_formatted = lesson['length']
            
            # Update fields using the exact field names from the PDF
            field_updates[f"dd  mm  yy{i}"] = date_formatted
            field_updates[f"hh  mm{i}_3"] = time_formatted    # Duration

        # Create the student-specific timesheet
        if create_student_timesheet(template_pdf, output_pdf, field_updates):
            success_count += 1
            print(f"Created timesheet for {student_name} with {len(student['lessons'])} lessons")
    
    print(f"\nSummary: Successfully created {success_count} of {len(students_info)} timesheets")

if __name__ == "__main__":
    main()