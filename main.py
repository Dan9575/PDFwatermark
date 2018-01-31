import pdfrw as pdf
import matplotlib.pyplot as plt
import os
import pandas as pd
import datetime

def add_watermark(
                  names_list=r'',
                  watermarks_path=r'',
                  save_path=r'',
                  files_path=r''
                  ):
    """
    :param names_list: path and file name for the excel file that generates the watermarks
    :param watermarks_path: path to a folder where watermark PDFs can be saved
    :param save_path: path to save to save the final files in
    :param files_path: path to folder with PDF files to be watermarked
    """

    now = datetime.datetime.now()
    date = now.strftime("%Y-%m-%d")

    # Generate list of names and sizes to watermark with
    files_list = os.listdir(files_path)
    watermarks_list = os.listdir(watermarks_path)
    names = pd.read_excel(names_list)

    for file in files_list:
        output_file_name = file.split('.')[0]
        source_pdf = files_path + file

        for name in names.iterrows():
            source = pdf.PdfReader(source_pdf)
            size = source.pages[0].MediaBox[2:4]
            h = round(float(size[0]) / 72, 2)
            w = round(float(size[1]) / 72, 2)
            watermark_file_name = name[1][0] + "_" + str(h) + "_" + str(w) + ".pdf"
            if watermark_file_name not in watermarks_list:
                make_water_mark(name[1][0], watermarks_path, watermark_file_name, h, w)

            owner = name[1][1]
            name = name[1][0]
            writer = pdf.PdfWriter()
            watermark_pdf = watermarks_path + name + "_" + str(h) + "_" + str(w) + ".pdf"
            watermark = pdf.PageMerge().add(pdf.PdfReader(watermark_pdf).pages[0])[0]

            for page in source.pages:
                pdf.PageMerge(page).add(watermark).render()
                writer.addpage(page)
            file_name = save_path + owner + '/' + output_file_name + ' - ' + date + '/' + output_file_name + ' - ' + name + '.pdf'
            os.makedirs(os.path.dirname(file_name), exist_ok=True)
            writer.write(file_name)


def make_water_mark(name, watermarks_path, watermark_file_name, h, w):
        save_file = watermarks_path + watermark_file_name
        watermark_text = 'Confidential - ' + name

        fig, ax = plt.subplots(figsize=(h, w))
        plt.ylabel('')
        plt.xlabel('')
        plt.xticks([])
        plt.yticks([])
        plt.text(x=0.5,
                 y=0.5,
                 s=watermark_text,
                 horizontalalignment='center',
                 verticalalignment='center',
                 rotation=45,
                 alpha=0.4,
                 size=30
                 )
        ax.set_xticklabels('')
        ax.set_yticklabels('')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_visible(False)
        ax.spines['bottom'].set_visible(False)
        plt.savefig(save_file, dpi=500, transparent=True)
        plt.close(fig)


if __name__ == "__main__":
    add_watermark()
