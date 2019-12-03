from pdfrw import PdfReader, PdfWriter
import pdfrw

# Annotation Key used by pdfrw
ANNOT_KEY = '/Annots'

# The values below are the order in wich these categories appear in
# page 2 of the DDAH forms
DDAH_CATEGORIES = {"CONTACT HOURS" : 3,
                   "MARKING HOURS": 4,
                   "PREP HOURS": 2,
                   "INVIGILATION HOURS": 5}

# The valeus below are the category names in the page 1 of the DDAH forms
DDAH_CATEGORY_NAMES = {"CONTACT HOURS" : "Contact Time",
                       "MARKING HOURS": "Marking/Grading",
                       "PREP HOURS": "Preparation",
                       "INVIGILATION HOURS": "Other Duties"}


# This is the only key from INFO_FIELDS key that gets reused
SUPERVISOR_KEY = "Supervising Professor"

# These keys correspond exactly to the *field names in the PDF file*
# and should not be changed!
INFO_FIELDS = {"Department": "MCS",
               "Course Code": "CSC338",
               "Course Title": "Numerical Methods",
               "Tutorial Category": "Skills Development",
               SUPERVISOR_KEY: "Lisa Zhang",
               "Est. Enrolment / TA Section": 30,
               "Expected Enrolment \(course\)": 90}

# TODO: Fill out the approver
APPROVER = "Approver Name Goes Here"

# TODO: Fill out the date
DATE = "December 3, 2019"


def parse_data(filename):
    """
    Parse data in filename. See sample_data.txt for an example.
    """
    ta = {
        "name": "",     # TA full name
        "total ": 0,    # total contract hours
        "detailed": [], # materials for the first page
        "summary": {}   # materials for the second page
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

        if key.startswith("Full Name"):
            ta["name"] = info.strip()
        elif key.startswith("Total contract"):
            ta["total"] = float(info.strip().strip("_"))
        else:
            # convert info -> hours, filter out zero hours:
            info = info.strip("_")
            if not info:
                continue
            hours = float(info)
            if hours < 0.1:
                continue

            # detailed hours:
            ta["detailed"].append((key, category, hours),)
            # summary hours:
            assert category is not None
            if category not in ta["summary"]:
                ta["summary"][category] = 0
            ta["summary"][category] += hours
            # total hours:
            total_hours += hours

    # verify that the hours add up
    if total_hours != ta["total"]:
        raise ValueError("Total contract hours is {0} but {1} hours are assigned".format(ta["total"], total_hours))
    # verify that there are at most 12 rows in ta["detailed"]
    if len(ta["detailed"]) > 12:
        raise ValueError("DDAH form supports at most 12 rows of detailed activity, but there are {0}".format(len(ta["detailed"])))

    return ta

def write_pdf(outfile, ta_data, TEMPLATE="DDAH.pdf"):
    template_pdf = pdfrw.PdfReader(TEMPLATE)

    # Page 1: Description of Duties
    annotations = template_pdf.pages[0][ANNOT_KEY]
    for index, annotation in enumerate(annotations):
        key = annotation['/TU']
        if key:
            key = key[1:-2] # Remove "(" at beginning and ":)" and end of string
            if key in INFO_FIELDS:
                value = INFO_FIELDS[key]
                annotation.update(
                    pdfrw.PdfDict(V='{}'.format(value))
                )

    # Page 1: Allocations of Hours (Detailed)
    annotations = template_pdf.pages[0][ANNOT_KEY]
    for i, (task, category, hour) in enumerate(ta_data["detailed"]):
        # The "#" field
        annotations[12 + 7 * i].update(
            pdfrw.PdfDict(V='{}'.format((i + 1)))
        )
        # The "Responsibility/Activity" field
        annotations[12 + 7 * i + 1].update(
            pdfrw.PdfDict(V='{}'.format(task))
        )
        # The "Hour" field
        annotations[12 + 7 * i + 4].update(
            pdfrw.PdfDict(V='{}'.format(hour))
        )
        # The "Category" field"
        annotations[12 + 7 * i + 6].update(
            pdfrw.PdfDict(V='{}'.format(DDAH_CATEGORY_NAMES[category]))
        )
    # Page 1 Total:
    annotations[96].update(
        pdfrw.PdfDict(V='{}'.format(ta_data["total"]))
    )


    # PAGE TWO: SUMMARY
    annotations = template_pdf.pages[1][ANNOT_KEY]
    for category, index in DDAH_CATEGORIES.items():
        if category in ta_data["summary"]:
            annotations[index*2].update(
                pdfrw.PdfDict(V='{}'.format(ta_data["summary"][category]))
            )
    annotations[12].update(
        pdfrw.PdfDict(V='{}'.format(ta_data["total"]))
    )

    # PAGE TWO: Information
    annotations[23].update( # Prepared by
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
        )

    PdfWriter().write(outfile, template_pdf)


if __name__ == "__main__":
    import sys

    data_file = "sample_data.txt" # sys.argv[1]
    pdf_out_file = "out.pdf" # sys.argv[2]

    ta_data = parse_data(data_file)
    write_pdf(pdf_out_file, ta_data)
