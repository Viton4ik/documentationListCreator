
import xlsxwriter
import os
import fnmatch
from datetime import datetime
from datetime import date

# start_path = input("Input a path: ")
start_path = "C:\Documentation"
current_path = os.getcwd()
print(start_path)
print(os.path)
print(f"folder '{start_path}' contains: {os.listdir(start_path)}\n")

#get our hostname (local network name)
import socket
host_name = socket.gethostname()

# create a xls -file
file_name = xlsxwriter.Workbook(f'{start_path}\context.xlsx')
file_name.set_properties({
    'title':    'Documentation List',
    # 'subject':  'With document properties',
    'author':   'Victor Vetoshkin',
    # 'manager':  'Dr. Heinz Doofenshmirtz',
    'company':  'created for Gastrans d.o.o Novi Sad',
    # 'category': 'Example spreadsheets',
    # 'keywords': 'Sample, Example, Properties',
    # 'created':  date.today(),
    'created':  datetime.utcnow(),
    'comments': 'https://t.me/Viton4ik',
    'hyperlink_base': 'https://t.me/Viton4ik',
})

# get a date
now = datetime.now()
current_date = now.date()
current_time = now.time()

# options for cells
def cell_color():
    # color line and format by cell's width
    # color #E2EFDA - olive
    if row >= 10 and row % 2 == 0:
        color_format = file_name.add_format({'font_size': 12, 'text_wrap': True})
    else:
        color_format = file_name.add_format({'font_size': 12, 'text_wrap': True, 'bg_color': '#E2EFDA'})
    return color_format

# read names and numbers for the Siemens docs
with open(f"C:\Documentation_ext\SIM_name.txt", 'r') as _file_name:
    with open(f"C:\Documentation_ext\SIM_number.txt", 'r') as _file_number:
        f_name = _file_name.readlines()
        f_number = _file_number.readlines()

# page shift
shift = 2
# row
row = 4
# colum
col = 1

worksheet = file_name.add_worksheet("DOCS")#file_name.add_worksheet("VENDOR DOCS")
for dir, folders, files in os.walk(start_path):
    dir_ = dir.split('\\')
    # create RFQ data
    RFQ_lst = dir_[shift + 1:shift:-1]
    RFQ = ''.join(RFQ_lst)
    # create System data
    fold_lst = dir_[shift + 2:shift + 1:-1]
    fold = ''.join(fold_lst)
    # create UNIT data
    endfold_lst = dir_[shift + 3:shift + 2:-1]
    endfold = ''.join(endfold_lst)

    # Set the height of Row
    #worksheet.set_row(0, 20)  # Set the height of Row 1 to 20.
    # worksheet.set_default_row(30)

    # Hide all rows without data.
    worksheet.set_default_row(hide_unused_rows=True)
    worksheet.write("A4", '_')  # to avoid this row hiding - with a filter

    # set a column's width
    worksheet.set_column("A:A", 3) # A
    worksheet.set_column("B:B", 30) # B
    worksheet.set_column("C:C", 15)  # C
    worksheet.set_column("D:D", 30) # D
    worksheet.set_column("E:E", 30) # E
    worksheet.set_column("F:F", 70) # F
    worksheet.set_column("G:G", 70) # G

    # text format
    text_format = file_name.add_format({'bold': True, 'align': 'center', 'font_size':16, 'bg_color':'#70AD47'})

    # # set a filter
    worksheet.autofilter(3, 1, 50000, 6)
    # # worksheet.filter_column(1, 'RFQ == ')

    # freeze the line
    worksheet.freeze_panes(4, 1)

    # worksheet.write_url('A4', "external:c:/test/testlink.xlsx",string="Link to other workbook")
    # worksheet.write_url('A4', "internal:c:/test/testlink.xlsx", string="Link to other workbook")

    cell_format2 = file_name.add_format({'font_size': 12, 'text_wrap': True, 'bg_color': '#E2EFDA','color': 'blue', 'underline': True}) # color - olive
    cell_format1 = file_name.add_format({'font_size': 12, 'text_wrap': True, 'color': 'blue', 'underline': True})
    cell_format3 = file_name.add_format({'font_size': 9, 'color': 'blue', 'underline': True})
    cell_format4 = file_name.add_format({'num_format': 'dd/mm/yy','font_size': 9,'align': 'left', 'color': 'blue'}) # num_format': 'dd/mm/yy hh:mm

    worksheet.insert_image('G1', 'C:\Documentation_ext\GST.png', {'x_offset': 350}) # image

    # write data in xls-file
    worksheet.write("B3", 'DOCUMENTATION', text_format)  # column's name
    worksheet.write("C3", 'TYPE', text_format)  # column's name
    worksheet.write("D3", 'SYSTEM', text_format)  # column's name
    worksheet.write("E3", 'UNIT', text_format)  # column's name
    worksheet.write("F3", 'FILES', text_format)  # column's name.
    worksheet.write("G3", 'DOCUMENT NAME', text_format)  # column's name

    url = "https://t.me/Viton4ik"
    worksheet.write_url("B2", f"external:{url}", string="Created by @Victor Vetoshkin", cell_format=cell_format3,
                            tip="text me")  # my contact
    dt = f"Created: {current_date}"
    worksheet.write("B1", dt, cell_format4)


    # # Add the VBA project binary.
    # import xlwings as xw
    #
    # vba_book = xw.Book("vba_python.xlsm")
    # vba_macro = vba_book.macro("SampleMacro")
    # vba_macro()

    # worksheet.set_vba_name('VENDOR DOCS')
    # file_name.add_vba_project('C:/Documentation_ext/vbaProject.bin')
    # # file_name.add_vba_project('./vbaProject.bin')

    for file in files: # create a list of files
        # cells options
        cell_format = cell_format1 if (row >= 10 and row % 2 == 0) else cell_format2

        # write data in xls-file
        try:
            worksheet.write(row, col, dir_[2], cell_color())  # B1
        except IndexError:
            pass
        worksheet.write(row, col + 1, RFQ, cell_color())  # C1
        worksheet.write(row, col + 2, fold, cell_color())  # D1
        link__ = f"\\\\{host_name}\\{dir[2:]}"
        worksheet.write_url(row, col + 3, f'external:{link__}', string=endfold, cell_format=cell_format,
                            tip="This is the link. Click here!")  # E1
        link_ = f"\\\\{host_name}\\{dir[2:]}\\{file}"
        worksheet.write_url(row, col + 4, f'external:{link_}', cell_format=cell_format, string=file,
                            tip="This is the link. Click here!")  # F1
        link = f"\\\\{host_name}\\{dir[2:]}\\{file}"

        # uncomment this line -> G1 will be = F1 if it's empty
        worksheet.write_url(row, col + 5, f'external:{link}', cell_format=cell_format, string=file,
                            tip="This is the link. Click here!")  # G1

        # create names relate to the numbers for Siemens docs in a xls-file
        for i in range(len(f_name)):
            f_number_ = f_number[i]
            f_name_ = f_name[i]
            name = f_name_.split('\n')
            number = f_number_.split('\n')
            compare_ = number[0] in file[:26]
            if compare_:
                # worksheet.write_url(row, col + 5, f'external:{dir}\{file}', cell_format=cell_format, string=name[0],
                #                     tip="This is the link. Click here!")  #
                # print(f"{number[0]} in {file[:26]}")
                link = f"\\\\{host_name}\\{dir[2:]}\\{file}"
                worksheet.write_url(row, col + 5, f'external:{link}', cell_format=cell_format, string=name[0],
                                    tip="This is the link. Click here!")  # G1

        row += 1
    row += 1

# hide empty rows: 4-9
for i in range(4,9):
    worksheet.set_row(i, None, None, {'hidden': 1})

# # Tips
# text = 'This column contains links to folders and files. Just click on it'
# worksheet.write_comment('E3', text, {'visible':True, 'x_scale': 1.25, 'y_scale': 0.4})#, 'x_offset': 30, 'y_offset': 30})
# worksheet.write_comment('F3', text ,{'visible':True, 'x_scale': 1.25, 'y_scale': 0.4})

## features for printing
header1 = "Vendor's documentation list"
footer1 = '@developed by Victor Vetoshkin'
worksheet.set_landscape()
worksheet.set_paper(9) #A4 paper
worksheet.set_header(header1)
worksheet.set_footer(footer1)

# file protecting option
# worksheet.protect(options={'autofilter': True})
# worksheet.unprotect_range('C3:C3')

# worksheet_log = file_name.add_worksheet("Log")
# worksheet_log.set_vba_name('Log')
# file_name.add_vba_project('C:/Documentation_ext/vbaProject1.bin')
#############
file_name.close()
#############

