import pdfrw as pdf
import win32com.client
import xlrd
import os


def add_watermark(orientation='P',
                  names_list=r'//nycroot/data/equities/hfs/newbiz/3 Business Consulting/WatermarkTool/watermark_list.xlsm',
                  watermarks_path=r'//nycroot/data/equities/hfs/newbiz/3 Business Consulting/WatermarkTool/watermarks/',
                  save_path=r'//nycroot/data/equities/hfs/newbiz/3 Business Consulting/WatermarkTool/watermarked_files/',
                  files_path=r'//nycroot/data/equities/hfs/newbiz/3 Business Consulting/WatermarkTool/files_to_watermark/'
                  ):
    """
    :param orientation: defalts to potrait, but set to L for landscape
    :param names_list: path and file name for the excel file that generates the watermarks
    :param watermarks_path: path to a folder where watermark PDFs can be saved
    :param save_path: path to save to save the final files in
    :param files_path: path to folder with PDF files to be watermarked
    """
    macroname = 'watermarks'
    if orientation == "L":
        macroname = 'watermarksL'

    generate_watermarks(names_list, macroname)

    files_list = os.listdir(files_path)

    # Generate list of names to watermark with
    book = xlrd.open_workbook(names_list)
    sheet = book.sheet_by_index(0)
    names = []
    for i in range(1, sheet.nrows):
        names += [sheet.row_values(i)]

    # Generate list of files to watermark
    for file in files_list:
        output_file_name = file.split('.')[0]
        source_pdf = files_path + file
        for name in names:
            owner = name[1]
            name = name[0]
            writer = pdf.PdfWriter()
            source = pdf.PdfReader(source_pdf)
            watermark_pdf = watermarks_path + name + '.pdf'
            watermark = pdf.PageMerge().add(pdf.PdfReader(watermark_pdf).pages[0])[0]
            for page in source.pages:
                pdf.PageMerge(page).add(watermark).render()
                writer.addpage(page)
            file_name = save_path + owner +'/' + output_file_name + ' - ' + name + '.pdf'
            os.makedirs(os.path.dirname(file_name), exist_ok=True)
            writer.write(file_name)


def generate_watermarks(file_name, macro_name):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = 0
    wb = excel.Workbooks.Open(file_name)
    macro = wb.Name + '!' + macro_name
    excel.Run(macro)
    wb.Close(1)


if __name__ == "__main__":
    add_watermark(orientation='L')
