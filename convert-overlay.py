#!/usr/bin/python3
from pdfrw import PdfReader, PdfWriter
import pdfrw
import subprocess
import shlex
import os
# Annotation Key used by pdfrw
ANNOT_KEY = '/Annots'

# The values below are the order in wich these categories appear in
# page 2 of the DDAH forms
DDAH_CATEGORIES = {"CONTACT HOURS" : 3,
                   "MARKING HOURS": 4,
                   "PREP HOURS": 2,
                   "INVIGILATION HOURS": 5}

# The valeus below are the category names in the page 1 of the DDAH forms
CONTACT_KEY = "CONTACT HOURS"
MARKING_KEY = "MARKING HOURS"
PREP_KEY = "PREP HOURS"
INVIG_KEY = "INVIGILATION HOURS"

DDAH_CATEGORY_NAMES = {"CONTACT HOURS" : "Contact Time",
                       "MARKING HOURS": "Marking/Grading",
                       "PREP HOURS": "Preparation",
                       "INVIGILATION HOURS": "Other Duties"}

DEPARTMENT_KEY = "Department"
COURSE_CODE_KEY = "Course Code"
TUTORIAL_CAT_KEY = "Tutorial Category"
SUPERVISOR_KEY = "Supervising Professor"
SECTION_ENROLMENT_KEY = "Est. Enrolment / TA Section"
COURSE_ENROLMENT_KEY = "Expected Enrolment \(course\)"
COURSE_TITLE_KEY = "Course Title"

# These keys correspond exactly to the *field names in the PDF file*
# and should not be changed!
INFO_FIELDS = {"Department": "MCS",
               "Course Code": "CSC***",
               "Course Title": "Placeholder Course Title",
               "Tutorial Category": "Laboratory / Practical",
               SUPERVISOR_KEY: "Placeholder Instructor",
               "Est. Enrolment / TA Section": 0,
               "Expected Enrolment \(course\)": 0}

# TODO: Fill out the approver
APPROVER = ""

# TODO: Fill out the date
DATE = "July 27, 2020"


def parse_data(filename):
    """
    Parse data in filename. See sample_data.txt for an example.
    """
    ta = {
        "name": "",  # TA full name
        "total ": 0,  # total contract hours
        "detailed": [],  # materials for the first page
        "summary": {}  # materials for the second page
    }
    category = None
    total_hours = 0

    for line in open(filename):
        line = line.strip()
        if not line or line.startswith("====") or line.startswith("TA INFO"):
            # skip these lines. they are not meaningful
            continue

        key, info = line.split(":")
        key = key.strip()

        if key in DDAH_CATEGORIES:
            category = key

        if key.lower().startswith("Full Name".lower()):
            ta["name"] = info.strip()
        elif key.startswith("Total contract"):
            ta["total"] = float(info.strip().strip("_").strip("\t"))
        else:
            # convert info -> hours, filter out zero hours:
            info = info.strip("_").strip("\t")
            if not info:
                continue
            hours = float(info)
            if hours < 0.1:
                continue

            # detailed hours:
            ta["detailed"].append((key, category, hours), )
            # summary hours:
            assert category is not None
            if category not in ta["summary"]:
                ta["summary"][category] = 0
            ta["summary"][category] += hours
            # total hours:
            total_hours += hours

    # verify that the hours add up
    if total_hours != ta["total"]:
        raise ValueError(
            "Total contract hours is {0} but {1} hours are assigned".format(
                ta["total"], total_hours))
    # verify that there are at most 12 rows in ta["detailed"]
    if len(ta["detailed"]) > 12:
        raise ValueError(
            "DDAH form supports at most 12 rows of detailed activity, but there are {0}".format(
                len(ta["detailed"])))

    return ta

def generate_page1(ta_data):
    from reportlab.pdfgen import canvas
    c = canvas.Canvas("page1.pdf")
    line = 22
    c.drawString(110,660,INFO_FIELDS[DEPARTMENT_KEY])
    c.drawString(110,660-line,INFO_FIELDS[COURSE_CODE_KEY])
    c.drawString(110,660-2*line,INFO_FIELDS[COURSE_TITLE_KEY])
    c.drawString(110,660-3*line,INFO_FIELDS[TUTORIAL_CAT_KEY])

    #c.drawString(160,660-4*line,"y") # optional
    #c.drawString(210,660-4*line,"x") # mandatory

    c.drawString(430,660,INFO_FIELDS[SUPERVISOR_KEY])
    c.drawString(430,660-line,INFO_FIELDS[SECTION_ENROLMENT_KEY])
    c.drawString(430,660-2*line,INFO_FIELDS[COURSE_ENROLMENT_KEY])

    line = 32.5
    for i, (task, category, hour) in enumerate(ta_data["detailed"]):
        hour = "%.1f" % hour
        c.drawString(60,430 - i*line,task)
        c.drawString(385,430 - i*line, hour)
        c.drawString(470,430 - i*line,category)
    c.showPage()
    c.save()

def generate_page1(ta_data):
    from reportlab.pdfgen import canvas
    c = canvas.Canvas("page1.pdf")
    line = 22
    c.drawString(110,660,INFO_FIELDS[DEPARTMENT_KEY])
    c.drawString(110,660-line,INFO_FIELDS[COURSE_CODE_KEY])
    c.drawString(110,660-2*line,INFO_FIELDS[COURSE_TITLE_KEY])
    c.drawString(110,660-3*line,INFO_FIELDS[TUTORIAL_CAT_KEY])

    #c.drawString(160,660-4*line,"y") # optional
    #c.drawString(210,660-4*line,"x") # mandatory

    c.drawString(430,660,INFO_FIELDS[SUPERVISOR_KEY])
    c.drawString(430,660-line,INFO_FIELDS[SECTION_ENROLMENT_KEY])
    c.drawString(430,660-2*line,INFO_FIELDS[COURSE_ENROLMENT_KEY])

    line = 32.5
    for i, (task, category, hour) in enumerate(ta_data["detailed"]):
        hour = "%.1f" % hour
        c.drawString(60,430 - i*line,task)
        c.drawString(385,430 - i*line, hour)
        c.drawString(470,430 - i*line,category)

    # TODO: draw background here?
    c.drawString(385,430 - 12*line, str(ta_data["total"]))
    c.showPage()
    c.save()

def generate_page2(ta_data):
    from reportlab.pdfgen import canvas
    c = canvas.Canvas("page2.pdf")

    line = 29
    for i, key in enumerate(["_FIRSTCONTACT", "_ADDITIONAL", # are these not in here?
                             PREP_KEY, CONTACT_KEY, MARKING_KEY, INVIG_KEY]):
        if key in ta_data["summary"]:
            c.drawString(450,450-i*line, str(ta_data["summary"][key]))
            #c.drawString(520,450-i*line,"Revised")

    c.drawString(450,450-6*line, str(ta_data["total"]))


    c.drawString(30,235,INFO_FIELDS[SUPERVISOR_KEY])
    c.drawString(30,180,APPROVER)
    c.drawString(30,125,ta_data["name"])

    c.drawString(475,235,DATE)

    c.showPage()
    c.save()


def write_pdf_overlay(outfile, ta_data, TEMPLATE="DDAH.pdf"):
    from pdfrw import PdfReader, PdfWriter, PageMerge

    generate_page1(ta_data)
    generate_page2(ta_data)

    base_pdf = pdfrw.PdfReader(TEMPLATE)

    page1 = pdfrw.PdfReader("page1.pdf")
    merger = PageMerge(base_pdf.pages[0])
    merger.add(page1.pages[0]).render()

    page2 = pdfrw.PdfReader("page2.pdf")
    merger2 = PageMerge(base_pdf.pages[1])
    merger2.add(page2.pages[0]).render()

    writer = PdfWriter()
    writer.write(outfile, base_pdf)


def write_pdf(outfile, ta_data, TEMPLATE="DDAH.pdf"):
    # template_pdf = pdfrw.PdfReader(TEMPLATE)

    fdffile = open(outfile, 'w')

    # Page 1: Description of Duties
    # annotations = template_pdf.pages[0][ANNOT_KEY]

    '''for index, annotation in enumerate(annotations):
        key = annotation['/TU']
        if key:
            key = key[1:-2] # Remove "(" at beginning and ":)" and end of string
            if key in INFO_FIELDS:
                value = INFO_FIELDS[key]
                annotation.update(
                    pdfrw.PdfDict(V='{}'.format(value), AP='{}'.format(value))
                )'''

    fdffile.write('%FDF-1.2\n\n')
    fdffile.write('1 0 obj\n')
    fdffile.write('<<\n')
    fdffile.write('/FDF << /Fields 2 0 R>>\n')
    fdffile.write('>>\n')
    fdffile.write('endobj\n')
    fdffile.write('2 0 obj\n')
    fdffile.write(
        '[<< /T (form1[0].Page1[0].Table2[0].HeaderRow[0].TextField1[0]) /V (' +
        INFO_FIELDS[DEPARTMENT_KEY] + ') >>\n')
    fdffile.write(
        '<< /T (form1[0].Page1[0].Table2[0].HeaderRow[0].TextField5[0]) /V (' +
        INFO_FIELDS[SUPERVISOR_KEY] + ') >>\n')
    fdffile.write(
        '<< /T (form1[0].Page1[0].Table2[0].Row1[0].TextField1[0]) /V (' +
        INFO_FIELDS[COURSE_CODE_KEY] + ') >>\n')
    fdffile.write(
        '<< /T (form1[0].Page1[0].Table2[0].Row1[0].NumericField1[0]) /V (' + str(
            INFO_FIELDS[SECTION_ENROLMENT_KEY]) + ') >>\n')
    fdffile.write(
        '<< /T (form1[0].Page1[0].Table2[0].Row2[0].TextField1[0]) /V (' +
        INFO_FIELDS[COURSE_TITLE_KEY] + ') >>\n')
    fdffile.write(
        '<< /T (form1[0].Page1[0].Table2[0].Row2[0].NumericField3[0]) /V (' + str(
            INFO_FIELDS[COURSE_ENROLMENT_KEY]) + ') >>\n')
    fdffile.write(
        '<< /T (form1[0].Page1[0].Table2[0].#subformSet[0].Row2[1].TextField1[0]) /V (' +
        INFO_FIELDS[TUTORIAL_CAT_KEY] + ') >>\n')

    # Page 1: Allocations of Hours (Detailed)
    # annotations = template_pdf.pages[0][ANNOT_KEY]
    for i, (task, category, hour) in enumerate(ta_data["detailed"]):
        fdffile.write('<< /T (form1[0].Page1[0].Table3[0].Row' + str(
            1 + i) + '[0].Unit' + str(1 + i) + '[0]) /V (' + str(
            i + 1) + ') >>\n')
        fdffile.write('<< /T (form1[0].Page1[0].Table3[0].Row' + str(
            1 + i) + '[0].#field[1]) /V (' + task + ') >>\n')
        fdffile.write('<< /T (form1[0].Page1[0].Table3[0].Row' + str(
            1 + i) + '[0].Total[0]) /V (' + str(hour) + ') >>\n')
        fdffile.write('<< /T (form1[0].Page1[0].Table3[0].Row' + str(
            1 + i) + '[0].DropDownList1[0]) /V (' + str(
            DDAH_CATEGORY_NAMES[category]) + ') >>\n')
        '''# The "#" field
        annotations[12 + 7 * i].update(
            pdfrw.PdfDict(V='{}'.format((i + 1)), AP='{}'.format((i + 1)))
        )
        # The "Responsibility/Activity" field
        annotations[12 + 7 * i + 1].update(
            pdfrw.PdfDict(V='{}'.format(task), AP='{}'.format(task))
        )
        # The "Hour" field
        annotations[12 + 7 * i + 4].update(
            pdfrw.PdfDict(V='{}'.format(hour))
        )
        # The "Category" field"
        annotations[12 + 7 * i + 6].update(
            pdfrw.PdfDict(V='{}'.format(DDAH_CATEGORY_NAMES[category]))
        )'''
    # Page 1 Total:

    # fdffile.write('<< /T (form1[0].Page1[0].Zero[0]) /V (' + str(ta_data["total"]) + ') >>\n')
    fdffile.write(
        '<< /T (form1[0].Page1[0].Table3[0].Row6[0].Cell2[0]) /V (' + str(
            ta_data["total"]) + ') >>\n')

    '''annotations[96].update(
        pdfrw.PdfDict(V='{}'.format(ta_data["total"]))
    )'''

    # PAGE TWO: SUMMARY
    '''annotations = template_pdf.pages[1][ANNOT_KEY]
    for category, index in DDAH_CATEGORIES.items():
        if category in ta_data["summary"]:
            annotations[index*2].update(
                pdfrw.PdfDict(V='{}'.format(ta_data["summary"][category]))
            )'''
    if PREP_KEY in ta_data["summary"]:
        fdffile.write(
            '<< /T (form1[0].Page2[0].#subform[2].Table1[1].Row3[0].Total[0]) /V (' + str(
                ta_data["summary"][PREP_KEY]) + ') >>\n')

    if CONTACT_KEY in ta_data["summary"]:
        fdffile.write(
            '<< /T (form1[0].Page2[0].#subform[2].Table1[1].Row4[0].Total[0]) /V (' + str(
                ta_data["summary"][CONTACT_KEY]) + ') >>\n')

    if MARKING_KEY in ta_data["summary"]:
        fdffile.write(
            '<< /T (form1[0].Page2[0].#subform[2].Table1[1].Row5[0].Total[0]) /V (' + str(
                ta_data["summary"][MARKING_KEY]) + ') >>\n')

    if INVIG_KEY in ta_data["summary"]:
        fdffile.write(
            '<< /T (form1[0].Page2[0].#subform[2].Table1[1].Row6[0].Total[0]) /V (' + str(
                ta_data["summary"][INVIG_KEY]) + ') >>\n')

    '''annotations[12].update(
        pdfrw.PdfDict(V='{}'.format(ta_data["total"]))
    )'''

    fdffile.write(
        '<< /T (form1[0].Page2[0].#subform[2].Table1[1].Row6[0].Cell2[0]) /V (' + str(
            ta_data["total"]) + ') >>\n')

    # PAGE TWO: Information
    '''annotations[23].update( # Prepared by
        pdfrw.PdfDict(V='{}'.format(INFO_FIELDS[SUPERVISOR_KEY]))
    )
    annotations[24].update( # Approved By
        pdfrw.PdfDict(V='{}'.format(APPROVER))
    )
    annotations[25].update( # Accepted By
        pdfrw.PdfDict(V='{}'.format(ta_data["name"]))
    )
    for i in [26, 27, 28]: # Date fields
        annotations[i].update(
            pdfrw.PdfDict(V='{}'.format(DATE))
        )'''

    fdffile.write(
        '<< /T (form1[0].Page2[0].#subform[2].#field[21]) /V (' + INFO_FIELDS[
            SUPERVISOR_KEY] + ') >>\n')
    fdffile.write(
        '<< /T (form1[0].Page2[0].#subform[2].#field[22]) /V (' + APPROVER + ') >>\n')
    fdffile.write(
        '<< /T (form1[0].Page2[0].#subform[2].#field[23]) /V (' + ta_data[
            "name"] + ') >>\n')
    fdffile.write(
        '<< /T (form1[0].Page2[0].#subform[2].#field[24]) /V (' + DATE + ') >>\n')

    # template_pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))

    fdffile.write(']\n')
    fdffile.write('endobj\n')
    fdffile.write('trailer\n')
    fdffile.write('<<\n')
    fdffile.write('/Root 1 0 R\n\n')
    fdffile.write('>>\n')
    fdffile.write('%%EOF\n')
    fdffile.close()

    # PdfWriter().write(outfile, template_pdf)


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 8:
        print("Usage: {} "
              "CSC*** "
              "CourseTitleInQuotesOrWithoutSpaces "
              "\"Laboratory/Practical\"|\"SkillsDevelopment\" "
              "InstructorName "
              "EstEnrollmentPerTASection "
              "ExpectedEnrollment "
              "TemplateFile".format(sys.argv[0]))
        exit()

    INFO_FIELDS["Course Code"] = sys.argv[1]
    INFO_FIELDS["Course Title"] = sys.argv[2]
    INFO_FIELDS["Tutorial Category"] = sys.argv[3]
    INFO_FIELDS[SUPERVISOR_KEY] = sys.argv[4]
    INFO_FIELDS["Est. Enrolment / TA Section"] = sys.argv[5]
    INFO_FIELDS["Expected Enrolment \(course\)"] = sys.argv[6]

    #for data_file in os.listdir('.'):
    #    if data_file.endswith(".txt"):
    data_file = sys.argv[7]
    ta_data = parse_data(data_file)

    pdf_out_file=os.path.splitext(data_file)[0]+'.pdf'
    write_pdf_overlay(pdf_out_file, ta_data)

