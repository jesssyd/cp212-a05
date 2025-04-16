# Student Marking Application - CP212 Assignment 5 

## Purpose

The Student Marking Application was developed to analyze and report on student grade data stored in a Microsoft Access database. The tool is integrated into an Excel workbook and enables users to interactively import data, view course enrollment, and generate comprehensive statistical reports.

## How It Works

In the Excel file `tidd9161_a05`, users can access the application through the custom ribbon tab titled **"Student Marking"**, or from a button on the **"Introduction"** worksheet. When launched, a user form appears with a multi-page navigation system.

### Page 1: File Upload  
Users are prompted to import a `.mdb` file (specifically `Registrar.mbd`) using a FileOpen dialog. This step is required to move forward in the application.

### Page 2: Application Options  
Once a valid file is uploaded, the user can select from three options:
- **List Courses**: Populates a list box with all available courses from the database.
- **Course Enrollment**: Requires the user to select a course. Once selected, the app generates a new worksheet named after the course and appends "Enrollment", listing enrolled students.
- **Generate Report**: Requires both a course and an assignment to be selected. Generates a statistical report summarizing student grades.

### Page 3: Report Output  
The report includes:
- Assignment-level grade statistics: min, max, mean, median, mode, and standard deviation
- Class average and standard deviation
- Grade frequency distribution
- A histogram of grade frequencies
- A formatted Microsoft Word document containing the same information

## Tools & Technologies
- **Excel VBA**
- **MS Access (.mdb)**
- **SQL**
- **MS Word (automated export)**

## Key Features

- **User-Friendly Interface**  
  Developed using MS User Forms in Excel, allowing users to navigate a multi-step process for uploading files, selecting actions, and generating outputs.

- **File Integration**  
  Users select an `.mdb` file via a file dialog, which is validated before proceeding.

- **Functional Options**  
  The application provides three core functionalities:
  - **List Courses**: Displays available courses from the database.
  - **Course Enrollment**: Allows users to select a course and view enrolled students in a newly generated worksheet.
  - **Generate Report**: Produces a detailed statistical report including:
    - Grade summaries (min, max, mean, median, mode, standard deviation)
    - Class average and standard deviation
    - Grade frequency tables
    - A histogram to visualize grade distribution
    - An automatically generated MS Word document containing the report data

## Output

The application creates structured Excel worksheets and a formatted Word document summarizing course and assignment performance data. These outputs support academic analysis and administrative reporting.
