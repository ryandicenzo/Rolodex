import vobject, xlsxwriter

"""
Program takes in a .vcf file (iCloud exports them nicely) and parses it
to give you a clean spreadsheet to which you can manage your contacts,
and keep track of information easily for networking.
"""


# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('rolodex.xlsx')
worksheet = workbook.add_worksheet()

# Header for first row.

row = 0
column = 0

content = ["Name", "Phone", "Connection", "Relationship",
           "Email", "Hometown", "Birthday", "Employer", "School", "Notes"]

# Add content to first line, completing header.
for item in content :

    # Write current item
    worksheet.write(0, column, item)

    # Increment column by one
    column += 1


# Get .vcf from user
foundFile = False

# Loop to get correct file name from user
while not foundFile:
    # Input for File
    filename = input('Enter file name of vcf file in program folder. (ex: \'contacts.vcf\'): ')

    # If file found, continue. Else, try again.
    try:
        vcf = open(filename, encoding="utf8")
        foundFile = True
    except:
        print('File not found. Are you sure it is in the program folder?')

# Read in a stream of the file object.
contact_stream = vobject.readComponents(vcf)

# Adjust row, want to start writing at the 1st row, not 0th.
row = 1

# For each parsed contact
for contact in contact_stream:
    # Handle Name Assignment
    try:
        name = contact.fn.value
    except:
        print('VCard must have a name')
        break

    # Handle Telephone Assignment
    try:
        tel = (contact.tel.value)
    except:
        tel = ''
        print('No phone number found for ' + contact.fn.value)

    # Handle Email Assignment
    try:
        email = (contact.email.value)
    except:
        email = ''
        print('No email number found for ' + contact.fn.value)

    # Write found elements to sheet
    worksheet.write(row, 0, name)
    worksheet.write(row, 1, tel)
    worksheet.write(row, 4, email)

    # Prepare for next contact
    row += 1

workbook.close()

print("\n\nGeneration Completed!")
