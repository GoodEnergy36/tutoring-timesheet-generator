# tutoring-timesheet-generator

I work for a tutoring company and have to fill out timesheets for each month for all my students both on the website and on pdf's. This requires duplicate processes which annoy me a little. Because, I imagine, it annoys other tutors within the company and I like making things, I decided to build a POC that takes the html of my filled timesheets online and populates copies of the companies pdf timesheet template with the contents of said html. This isn't meant to be used by other tutors, but for the company to have as a starting point if they are interested in adding a feature like this into their website.

## Documentation

### Overview

This Python script automates the process of generating individual PDF timesheets for multiple students by:
1. Extracting student and lesson data from an HTML timesheet file
2. Creating individual PDF timesheets for each student with their specific lesson information
3. Filling the PDF form fields with appropriate values

### Requirements

- Python 3.6 or higher
- Required Python packages:
  - PyPDF2 (for PDF manipulation)
  - BeautifulSoup4 (for HTML parsing)
  - re (for regular expressions)

Install with:
```
pip install PyPDF2 beautifulsoup4
```

### Files Needed

1. **HTML timesheet file**: `timesheet.html` - HTML file containing all student lesson information
2. **PDF template**: `template.pdf` - PDF form template for the timesheets

### How It Works

#### Data Extraction
The script parses the HTML timesheet to extract:
- ABN number
- Month/year
- Student information:
  - Student names
  - Lesson dates
  - Lesson durations
  - Total hours

#### PDF Generation
For each student, the script:
- Creates a copy of the PDF template
- Fills form fields with student-specific data
- Saves as a new PDF with the student's name

#### Field Mapping
The script maps extracted data to the following PDF fields:
- `Month` - Month from HTML data
- `Your ABN` - ABN number from HTML data
- `Your Name` - Set to "Sam Towney"
- `STUDENT JOB NAME` - Student's name
- `Total Job Time` - Total lesson hours
- `dd mm yy#` - Lesson dates (for up to 10 lessons)
- `hh mm#_3` - Lesson durations (for up to 10 lessons)

### Usage

1. Place the HTML timesheet file as `timesheet.html` in the same directory as the script
2. Place the PDF template as `template.pdf` in the same directory
3. Run the script:
   ```
   python timesheet_generator.py
   ```
4. The script will create a `timesheets` directory with individual PDF timesheets for each student

### Error Handling

The script includes robust error handling to:
- Detect missing input files
- Identify missing student or lesson data
- Handle malformed dates and durations
- Validate PDF form fields
- Skip students with no valid lessons
- Provide clear warnings and error messages

### Limitations

- The PDF template must have specific field names as listed in the Field Mapping section
- The script processes a maximum of 10 lessons per student (PDF template limitation)
- The HTML file structure must match the expected format from the original timesheet example

### Customization

To customize the script:
1. Change the input filenames in the `main()` function
2. Modify the field mapping in the field_updates dictionary to match different PDF forms
3. Adjust the date and time formatting for different requirements
