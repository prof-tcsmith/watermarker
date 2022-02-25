# Watermarker

## Introduction

Watermarker is a program to help deter student cheating. This program accepts as input a list of students (stored in a CSV format), an assignment pdf, and an email body/message. The program then automatically sends copies of this assignment via email to all given students, with each assignment (pdf) watermarked with the student's full name and ID. Moreover, each assignment sent is encrypted and password protected from editing. The program records the students' details and stores copies of each file sent for auditing purposes. 

## Installation

Though this software is fully functional, the installation and configuration process is manual and does not have a graphical user interface. 

The installation process is as follows:

* Create a new directory. 
* Download the [watermarker.exe](watermarker.exe) file and copy it into this new directory.
* Download the starter [watermarker.json](watermarker.json) file and copy it into this new directory.
* Create an assignment PDF, a CSV containing student information, and an email body/template (samples of these are included in repo).
* Edit the watermaker.json file using any common text editor. Note that default directories/filenames for input files are set in this json configuration file.  
* Run the watermarker.exe file. 


## Configuration and Usage

The watermarker configuration file [watermarker.json](watermarker.json) contains the necessary settings for the program to run:

  * program_path - This is the path to the folder that you downloaded the watermarker.exe and watermarker.json files. The need to enter this will be removed in future versions.
  * course_num - This is the course number (i.e. MIS207, ITM251). This is used to create the email subject for emails sent to students (and cc'd to prof's/TAs)
  * students_csv_fname - This is the filename of the CSV file containing the list of students to send assignments to. See the sample [students__class01](https://github.com/prof-tcsmith/watermarker/tree/main/student_data/students_class01.csv) for format details. 
  * student_data_path - Subfolder name for where student information is stored (see sample [students_class_01.csv](https://github.com/prof-tcsmith/watermarker/tree/main/student_data/students_class01.csv) for format details).
  * assignment_fname - This is the filename of the assignment pdf you wish to send students.
  * assignments_path - Subfolder name for where assignments are stored.
  * owner_pword - Each personalized assignment is encrypted. Students must use their passwords to read the file. The owner_pword is used to open the file for editing and should be kept secret and used only by a TA or professor.
  * cc_list - List of cc's for the message sent to each student (i.e., "somebody@somewhere.com;someoneelse@nowhere.com"). Setting this will allow for confirmation and auditing of assignments sent.
  * email_body_fname - Text used to create the body of the emails sent to students.
  * watermarks_path - Subfolder name for where the student watermark files will be stored. This is a temporary storage area used during the creation of the watermarked files.
  * sent_items_path - Subfolder where each of the generated pdf's is stored.


The following settings specify the color, opacity, font size and rotation of the watermark:

  * red - Red color value (0-1)
  * green - Green color value (0-1)
  * blue - Blue color value (0-1)
  * opacity - Opacity value (0-1) 
  * font_size - Font size 
  * rotation - Rotation value (0-360)

> NOTE: If any of the paths and/or files that you entered in the JSON file do not exist, the program will crash. Be sure to double-check that each exists before attempting to run the program.

## Sample output

An example of a watermarked assignment:

The full sample PDF for student "John Doe" is found [here](https://github.com/prof-tcsmith/watermarker/blob/master/sample_assignment.pdf)

> NOTE: The pdf file above is encrypted, so it will not display in GitHub. You will need to download the file to view it with a pdf reader. The student password to open the file for reading is 1234. Students have limited rights and can only view the file. The admin password is 4321. Admins have full rights to the file. The admin password is set in the JSON config file. The unique student password is read from the student CSV file)


## Requirements

* Windows 7, 8, 10 or 11
* MS Outlook 2013 or later (the default email account set in outlook will be used to send the student assignments)

## Contributing

This project is under [active development]. Please feel free to [submit a pull request](https://github.com/prof-tcsmith/watermarker/pulls) to add features or fix bugs.


## Notes

> NOTE1: This software is 'alpha' - it's a fully functional program but does not yet have well-documented code or a user interface.

> NOTE2: The content found in the Assignments and Watermarks folders are sample files to assist in testing. You need to provide your own csv data and input pdf.



