import os
import win32com.client
import fitz
from cloudmesh.common.Shell import Shell

import win32com.client, win32com.client.makepy, winerror
from win32com.client.dynamic import ERRORS_BAD_CONTEXT


def PPTtoPDF(inputFileName, outputFileName, formatType = 32):
    powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
    powerpoint.Visible = 1

    if not outputFileName.endswith('.pdf'):
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
    deck.Close()
    powerpoint.Quit()

Shell.mkdir('./convertedPDFs')
print('Converting ppts to pdfs first...')

for root, dirs, files in os.walk(r'./originalppt'):

    for f in files:
        # print(f)
        actual = (os.path.join(root, f).replace('\\', '/'))
        actual = os.path.abspath(actual)
        # print(actual)
        if f.endswith(".pptx"):
            # this will not work if any other part of the original path
            # contains the words pptx or originalppt
            # so if we see those words occur twice or more, this fails
            try:
                PPTtoPDF(actual, actual.replace('.pptx', '.pdf').replace('originalppt', 'convertedPDFs'))
            except KeyboardInterrupt:
                exit()
            except Exception as e:
                pass

print('Combining files...')
result = fitz.open()
for root, dirs, files in os.walk(r'./convertedPDFs'):
    for file in files:
        if file.endswith('.pdf'):
            actual = (os.path.join(root, file).replace('\\', '/'))
            with fitz.open(actual) as mfile:
                result.insert_pdf(mfile)

result.save("milk_cereal_combine.pdf")


# Open PDF file, use Acrobat Exchange to save file as .docx file.


def PDF_to_Word(input_file, output_file):
    ERRORS_BAD_CONTEXT.append(winerror.E_NOTIMPL)
    src = os.path.abspath(input_file)

    # Lunch adobe
    win32com.client.makepy.GenerateFromTypeLibSpec('Acrobat')
    adobe = win32com.client.DispatchEx('AcroExch.App')
    avDoc = win32com.client.DispatchEx('AcroExch.AVDoc')
    # Open file
    avDoc.Open(src, src)
    pdDoc = avDoc.GetPDDoc()
    jObject = pdDoc.GetJSObject()
    # Save as word document
    jObject.SaveAs(output_file, "com.adobe.acrobat.docx")
    avDoc.Close(-1)

print('Converting pdf to giant docx...')
my_real_path = os.path.abspath('./milk_cereal_combine.pdf')
my_output = my_real_path.replace('.pdf', '.docx')
my_output = my_output.replace('milk_cereal', 'reesespuffs')
PDF_to_Word(my_real_path, my_output)

for root, dirs, files in os.walk(r'./convertedPDFs'):

    for f in files:
        if f.endswith('.pdf'):
            # print(f)
            actual = (os.path.join(root, f).replace('\\', '/'))
            actual = os.path.abspath(actual)
            Shell.rm(actual)

Shell.rmdir('./convertedPDFs')
Shell.rm(my_real_path)

print(f"SAVED TO {my_output}")
