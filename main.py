from PyPDF2 import PdfFileReader, PdfFileWriter
from clint.textui import colored, puts, indent, progress
from glob import glob
from os import path

def process_windows_pptx(inputFileName, outputFileName):
    # https://stackoverflow.com/a/51952043
    try:
        import win32com.client
        powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
        powerpoint.Visible = 1
        if outputFileName[-3:] != 'pdf':
            outputFileName = outputFileName + ".pdf"
        deck = powerpoint.Presentations.Open(path.abspath(inputFileName))
        deck.SaveAs(path.abspath(outputFileName), 32) # formatType = 32 for ppt to pdf
        deck.Close()
        powerpoint.Quit()
    except:
        puts(colored.red("We can't convert files! You might not have PowerPoint installed, or you might not be on Windows."))

template = 'template.pdf'
files = []

# Try and get a specific dropped file, otherwise just glob for anything likely in this directory
import sys
try:
    files = [sys.argv[1],]
except IndexError:
    puts(colored.red("No file dropped. Looking for other non-template files in directory..."))
    allfiles = glob("./[!~]*.pdf") + glob("./[!~]*.ppt?")
    files = [path.relpath(f) for f in allfiles if path.relpath(f) != template]

# List and identify files requiring conversion
puts(colored.green("Files to load (yellow will need converting, Windows Only):"))
needs_converting = []
with indent(4):
    # Filter out PPT files to separate list
    needs_converting = [f for f in files if "ppt" in path.splitext(f)[1] and not path.splitext(f)[0] + ".pdf" in files]
    # Filter out powerpoints already converted
    files = [f for f in files if not "ppt" in path.splitext(f)[1]]
    # Filter out files tagged with [C] (converted outputs) and their sources to separate lists
    already_converted = [f for f in files if path.splitext(f)[0].startswith("[C]")]
    already_converted_origin = [path.split(f)[1][4:] for f in already_converted]
    already_converted_origin_pptx = [path.splitext(f)[0][4:] + ".pptx" for f in already_converted]
    #  - Try and remove these
    files = [f for f in files if not (f in already_converted or f in already_converted_origin or f in already_converted_origin_pptx)]
    # Build results string
    results_list = [colored.yellow(f) for f in needs_converting]
    results_list.extend(files)
    # results = ", ".join()
    puts(str(results_list))

# Handle no files!
if not files:
    puts(colored.green("❓ - No files found! Try adding some powerpoint or PDF files to the folder. If there are already some there, you might have already converted them!"))

# Convert any files that need it, adding the converted files to our `files` list
if needs_converting:
    puts("Converting {} file{}...".format(len(needs_converting), "" if len(needs_converting) == 1 else "s"))
    with indent(4):
        for f in needs_converting:
            output_name = path.splitext(f)[0] + ".pdf"
            puts(f + ' → ' + output_name, newline=False)
            process_windows_pptx(f,output_name)
            files.append(output_name)
            puts(colored.green('✔'))
    puts(colored.green("🎉 Done converting required files!"))
else:
    puts(colored.green("🎉 No need to convert any files!"))

# Start overlaying PDF!
puts(colored.yellow("Building Templated Notes..."))
with indent(4):
    for f in files:
        puts(f, newline=False)
        # https://gist.github.com/vsajip/8166dc0935ee7807c5bd4daa22a20937
        slides = PdfFileReader(open(f, 'rb'), strict=False)
        output_pdf = PdfFileWriter()
        for i in progress.bar(range(slides.getNumPages())):
        # for i in progress.bar(range(1)):
            pdf_template = PdfFileReader(open(template, 'rb'), strict=False)
            template_page = pdf_template.getPage(0)
            slide_page = slides.getPage(i)
            # Scale page to fit in box
            new_height = 192.86 # magic number, (slideheight)/0.35277
            new_ratio = new_height/float(slide_page.mediaBox[3])
            new_width = new_ratio*float(slide_page.mediaBox[2])
            slide_page.scaleTo(new_width, new_height)
            # Figure out where box is
            page_width = 184.453
            # [x1,y1,width,height]
            template_page.mergeTranslatedPage(
                slide_page,
                float(template_page.mediaBox[2])-float(slide_page.mediaBox[2]),
                float(template_page.mediaBox[3])-float(slide_page.mediaBox[3])
            )
            output_pdf.addPage(template_page)
        output_pdf.write(open(
            "[C] {}".format(path.split(f)[1]),"wb"
        ))
        puts(colored.green('✔'))
