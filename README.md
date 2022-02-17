# Watermarker

## Introduction

Watermarker is a program to help deter student cheating. This program accepts as input a list of students (stored in a CSV format) and an assignment pdf. The program then automatically sends copies of this assignment via email to all given students, with each assignment (pdf) watermarked with the student's full name and ID. Moreover, each assignment sent will be encrypted and protected from editing. 

For auditing purposes, the program records the details of student processed and stores copies of each file sent. 

## Installation

* Create a new directory. 
* Download the [watermarker.exe](watermarker.exe) file and copy it into this new directory.
* Download the starter [watermarker.json](watermarker.json) file and copy it into this new directory.
* Edit the watermaker.json file using any common text editor.

## Configuration and Usage

The watermarker configuration file [watermarker.json](watermarker.json) contains the necessary settings for the program to run:

  * program_path - This is the path to the folder that you downloaded the watermarker.exe and watermarker.json files
  * course_num - This is the number of the course. This is used to create the email subject for emails sent to students (and cc'd to prof's/TAs)
  * students_csv_fname - This is the filename for the list of students to send assignments to. See the sample [students__class01](https://github.com/prof-tcsmith/watermarker/tree/main/student_data/students_class01.csv) for format details. 
  * assignment_fname - This is the pdf filename for the assignment you wish to send students.
  * owner_pword - Each personalized assignment is encrypted. Students must use their pword to read the file. The owner_pword is used to open the file for editing. 
  * cc_list - List of cc's for the message sent to each student (i.e. "somebody@somewhere.com;someoneelse@nowhere.com",
  * email_body_fname - Text used to create the body of the emails sent.
  * watermarks_path - Subfolder name for where files will be stored. 
  * assignments_path - Subfolder name for where assignments are stored.
  * sent_items_path - Subfolder name for where sent items will be stored.
  * student_data_path - Subfolder name for where student information is stored (see sample [students_class_01.csv](https://github.com/prof-tcsmith/watermarker/tree/main/student_data/students_class01.csv)
  * sent_items_path - Subfolder where each of the generated pdf's will be stored.

> NOTE: If any of the paths and/or files that you entered in the JSON file do not exist, the program will crash. Be sure to double-check that each exists before attempting to run the program.

## Sample output

An example of a watermarked assignment:

![Sample Assignment](https://github.com/prof-tcsmith/watermarker/blob/main/assignments/assignment_XYZ123_1000111.png)

The full sample PDF for student "John Doe" is found [here](https://github.com/prof-tcsmith/watermarker/tree/main/assignments/assignment_XYZ123_1000111.pdf) (NOTE: The file is encrypted, so it will not display in GitHub. You will need to download the file to view it with a pdf reader. The password to open the file for reading is 1234)

## Requirements

* Windows 7, 8, 10 or 11
* MS Outlook (must be installed)

## Contributing

This project is under [active development](https://github.com/prof-tcsmith/watermarker/projects/1))

Contributions are welcome! Please feel free to submit a Pull Request or suggest additional features/fixed by opening an [issue](https://github.com/prof-tcsmith/watermarker/issues).

## Limitations

* Assignments must not be more than 1 page (functionality for multiple pages will soon be added)


## Notes

> NOTE1: This software is 'alpha' - it's a fully functional program but does not yet have well-documented code or a user interface.

> NOTE2: The content found in the Assignments and Watermarks folders are examples of the output produced by the program. The input file for the demo is hw2q01.pdf.



