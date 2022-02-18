import csv
import PyPDF2 # install pypdf2 module
import win32com.client # install pywin32 module
import time
from hashlib import md5
from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2.generic import NameObject, DictionaryObject, ArrayObject, \
    NumberObject, ByteStringObject
from PyPDF2.pdf import _alg33, _alg34, _alg35
from PyPDF2.utils import b_
from reportlab.pdfgen import canvas
import pathlib

################################################################
# parameters... set these to fit local install preferences

#path = pathlib.Path(__file__).parent.resolve()
#print("v"*20)
#print(path)
#print("^"*20)

import json
with open('watermarker.json') as config_file:
    configs = json.load(config_file)
course_num = configs['course_num']
students_csv_fname = configs['students_csv_fname']
assignment_fname = configs['assignment_fname']
owner_pword = configs['owner_pword']
cc_list = configs['cc_list']
email_body_fname = configs['email_body_fname']

program_path = configs['program_path']
watermarks_path = configs['watermarks_path']
assignments_path = configs['assignments_path']
sent_items_path = configs['sent_items_path']
student_data_path = configs['student_data_path']
sent_items_path = configs['sent_items_path']


################################################################
# define helper functions

def get_students(student_csv_fname):
    """
    Reads the csv file containing student information and returns a list of dictionaries
    """
    students = []
    with open(student_csv_fname, encoding='utf-8') as csvfile:
        student_reader = csv.DictReader(csvfile, skipinitialspace=True)
        for next_student in student_reader:
            students.append(next_student)
    return students

def get_num_pages(assignment_fname):
#    pdf = PdfFileReader(open('path/to/file.pdf', 'rb'))
#    pdf.getNumPages()

    input_file = open(assignment_fname, 'rb')
    input_pdf = PyPDF2.PdfFileReader(input_file)
    num_pages = input_pdf.getNumPages()
    input_file.close()
    return num_pages

def create_watermark(output_fname, student_id, student_name, pages=1):
    """
    Creates a watermark file based on student ID and student name.

    :param str output_fname: File name for the newly created watermark file
    :param str student_id: Student ID that will be added to watermark
    :param str student_name: Student name that will be added to watermark
    """
    # use report lab to create a watermark file
    # that contains username --

    c = canvas.Canvas(output_fname)
    for page in range(pages):
        c.setFont("Helvetica", 17)
        c.rotate(30)
        c.setFillColorRGB(0.05, 0.5, 0.1, 0.1)
        for x in range(-300, 1200, 17):
            c.drawString(0, x, f'{student_name}-{student_id} '*10)
        c.showPage()
    c.save()



def create_merge(assignment_fname, watermark_fname, merged_fname, num_pages):
    """
    Creates a merged file between the question file and the watermark file

    :param str question_fname: Full path to the original question file to be watermarked
    :param str watermark_fname: Full path to watermark file to use to watermark the question
    :param str merged_fname: Name of the new merged file that will be created
    """
    # open the original question pdf
    assignment_file = open(assignment_fname, 'rb')
    assignment_pdf = PyPDF2.PdfFileReader(assignment_file)

    # open the watermark pdf
    watermark_file = open(watermark_fname, 'rb')
    watermark_pdf = PyPDF2.PdfFileReader(watermark_file)

    # create a merged pdf
    output = PyPDF2.PdfFileWriter()

    for i in range(num_pages):
        pdf_page = assignment_pdf.getPage(i)
        watermark_page = watermark_pdf.getPage(i)
        pdf_page.mergePage(watermark_page)
        output.addPage(pdf_page)

    merged_file = open(merged_fname, 'wb')
    output.write(merged_file)

    # close all open files
    merged_file.close()
    watermark_file.close()
    assignment_file.close()


def send_email(recipient, cc, subject, body, attachment, mail):
    """
    Send email via Outlook client .

    :param str recipient: Email address of recipient
    :param str cc: List of email addresses to cc (using semi-colon to separate multiple addresses)
    :param str subject: Subject of the email
    :param str body: Text body of the email (will also be used to create HTML body)
    :param str attachment: Full path to file to attach to email
    :param obj mail: Win32 Com object that's initialized to Outlook
    """
    print(f"..........{attachment}")
    mail.To = recipient
    mail.Subject = subject
    mail.HTMLBody = f'<h3>{body}</h3>'
    mail.Body = body
    mail.CC = cc
    mail.Attachments.Add(attachment)
    mail.Send()

def encrypt(writer_obj: PdfFileWriter, user_pwd, owner_pwd=None, use_128bit=True):
    """
    Encrypt this PDF file with the PDF Standard encryption handler.

    :param str user_pwd: The "user password", which allows for opening
        and reading the PDF file with the restrictions provided.
    :param str owner_pwd: The "owner password", which allows for
        opening the PDF files without any restrictions.  By default,
        the owner password is the same as the user password.
    :param bool use_128bit: flag as to whether to use 128bit
        encryption.  When false, 40bit encryption will be used.  By default,
        this flag is on.
    """
    import time, random
    if owner_pwd == None:
        owner_pwd = user_pwd
    if use_128bit:
        V = 2
        rev = 3
        keylen = int(128 / 8)
    else:
        V = 1
        rev = 2
        keylen = int(40 / 8)
    # permit copy and printing only:
    # P = -44
    # prevent everything:
    P = -3904
    O = ByteStringObject(_alg33(owner_pwd, user_pwd, rev, keylen))
    ID_1 = ByteStringObject(md5(b_(repr(time.time()))).digest())
    ID_2 = ByteStringObject(md5(b_(repr(random.random()))).digest())
    writer_obj._ID = ArrayObject((ID_1, ID_2))
    if rev == 2:
        U, key = _alg34(user_pwd, O, P, ID_1)
    else:
        assert rev == 3
        U, key = _alg35(user_pwd, rev, keylen, O, P, ID_1, False)
    encrypt = DictionaryObject()
    encrypt[NameObject("/Filter")] = NameObject("/Standard")
    encrypt[NameObject("/V")] = NumberObject(V)
    if V == 2:
        encrypt[NameObject("/Length")] = NumberObject(keylen * 8)
    encrypt[NameObject("/R")] = NumberObject(rev)
    encrypt[NameObject("/O")] = ByteStringObject(O)
    encrypt[NameObject("/U")] = ByteStringObject(U)
    encrypt[NameObject("/P")] = NumberObject(P)
    writer_obj._encrypt = writer_obj._addObject(encrypt)
    writer_obj._encrypt_key = key

##################################################################
# Main program
if __name__ == '__main__':
    print(get_students(f"{student_data_path}\\{students_csv_fname}"))

    students = get_students(f"{student_data_path}\\{students_csv_fname}")

    num_pages = get_num_pages(f"{assignments_path}\\{assignment_fname}") # get the number of pages in the assignment pdf

    for student in students:
        print(f"Processing {student}")
        print("\n"+"#"*80)
        print(f"Processing student: {student['name']}")


        # check the send value, if 0, then this student is skipped
        if student['send'] == '0':
            print("This student is flagged to be skipped")
            continue


        # create watermark file, and merge with assignment file
        # NOTE: Assignment fname must exist

        print("Creating watermark file for", student['name'])
        create_watermark(output_fname=f"{watermarks_path}\\{student['name'].replace(' ', '_')}_{student['id']}.pdf",
                         student_id=student['id'],
                         student_name=student['name'], # user student's first name
                         pages = num_pages
        )
        create_merge(assignment_fname = f"{assignments_path}\\{assignment_fname}",
                     watermark_fname = f"{watermarks_path}\\{student['name'].replace(' ', '_'):s}_{student['id']:s}.pdf",
                     merged_fname = f"{sent_items_path}\\{assignment_fname.split('.')[0]}_{student['id']:s}.pdf",
                     num_pages = num_pages
        )

        # encrypt the merged pdf
        # TODO: add this to the current encrypt function -- clean this up,
        #  also add ability to have user password and admin password to be
        #  parameters.
        print("Encrypting assignment file for", student['name'])
        unmeta = PdfFileReader(f"{sent_items_path}\\{assignment_fname.split('.')[0]}_{student['id']}.pdf")
        writer = PdfFileWriter()
        writer.appendPagesFromReader(unmeta)
        encrypt(writer, user_pwd=student['pword'], owner_pwd=owner_pword)
        with open(f"{sent_items_path}\\{assignment_fname.split('.')[0]}_{student['id']}.pdf", 'wb') as fp:
            writer.write(fp)

        # Send email with merged document attached
        # COM objects expose methods and properties using the IDispatch interface
        print("Sending the generated assignment file to", student['email'])
        with open(f"{assignments_path}\\{email_body_fname}") as fin:
            email_body = fin.read()

        outlook = win32com.client.Dispatch('outlook.application') # here connect with outlook client (which must be installed)
        send_email(recipient=student['email'],
                   cc=cc_list,
                   subject=f"[{course_num} Assignment Bot] {assignment_fname.split('.')[0]} [studentID={student['id']}]",
                   body=f"{email_body}",
                   attachment=f"{program_path}\\{sent_items_path}\\{assignment_fname.split('.')[0]}_{student['id']}.pdf", # need full path here
                   mail=outlook.CreateItem(0)
             )

        # initial test indicated that without a pause, outlook may miss sending some emails.
        # I've not had issues with a 0.5 second delay
        # TODO: try to eliminate the need for delay in send loop
        print("Finished processing student ", student['name'])
        time.sleep(0.5)

input("\nAll students processed. Hit return to exit the program.")

