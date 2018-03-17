import packages.tkinter as tk
import packages.pdfrw as pdf
import packages.json as json
import matplotlib.pyplot as plt
import os
import sys


class App:
    def __init__(self, master):
        with open('- Tool\defaults.txt', 'r') as file:
            self.variables = json.loads(file.read())

        with open('- Tool\data.txt', 'r') as file:
            self.data = json.loads(file.read())

        self.owners = sorted([i for i in self.data.keys()])

        self.last_page = tk.IntVar()
        self.message = tk.StringVar()
        self.update_name = tk.StringVar()
        self.watermark_prefix_text = tk.StringVar()
        self.watermark_prefix_text.set("Current Watermark Prefix is: " + self.variables['watermark_prefix'])
        master.title("PDF Watermarking Tool")
        self.message.set("Watermark Tool!")

        tk.Frame(master).grid(row=0, pady=5)
        tk.Label(master, textvariable=self.message, font=('Arial', 12), width=30).grid(row=1, pady=10, columnspan=2)
        tk.Checkbutton(master, text='Watermark last page', variable=self.last_page, anchor='e', padx=5,
                       pady=10).grid(row=4, sticky='w')
        tk.Button(master, text='Run All', command=self.watermark_all, padx=5, width=10).grid(row=2, pady=5, padx=10, column=0, sticky='w')
        tk.Button(master, text='Settings', command=self.open_frame, padx=5, width=10).grid(row=3, pady=5, padx=10, column=0, sticky='w')

        self.run_some = tk.LabelFrame(master, text='Run Some')
        self.run_some.grid(row=5, pady=5, padx=10, sticky='w')

        # Owners list
        self.scrollbar = tk.Scrollbar(self.run_some, orient='vertical')
        self.ol = tk.Listbox(self.run_some, yscrollcommand=self.scrollbar.set, selectmode='extended', exportselection=0)
        self.scrollbar.config(command=self.ol.yview)
        self.ol.grid(row=0, column=0, pady=5, padx=0, sticky='w')

        self.scrollbar.grid(row=0, column=1, sticky='ns')

        # Files List
        self.fl = tk.Listbox(self.run_some, selectmode='extended', width=40,
                             exportselection=0)
        self.fl.grid(row=0, column=2, pady=5, padx=10, sticky='w')

        tk.Button(self.run_some, command=self.watermark_some, text='Run Some',  padx=5, width=10).grid(row=3, pady=5, padx=10, column=0,
                                                                                   sticky='w')
        tk.Button(self.run_some, command=self.build_owners_list_box, text='Update Files List',  padx=5, width=12).grid(row=3, pady=5, padx=10, column=2,
                                                                                   sticky='w')
        self.update_group = tk.LabelFrame(master, text='Update Owner Lists')
        self.update_group.grid(row=6, pady=5, padx=10, sticky='w')
        tk.Button(self.update_group, text='Update Watermarks', command=self.eidt_lists, padx=5).grid(row=0, column=1,
                                                                                                     pady=5, padx=20,
                                                                                                     sticky='w')
        tk.Button(self.update_group, text='Delete Owner', command=self.add_delete_button, padx=5).grid(row=0, column=2,
                                                                                                  pady=5, padx=20,
                                                                                                  sticky='w')

        self.delete_button = tk.Button(self.update_group, text='CONFIRM DELETE', command=self.delete_owner, state='disabled', padx=5)
        self.delete_button.grid(row=0, column=3, pady=5, padx=20, sticky='w')

        self.new_owner_name = tk.StringVar()
        self.new_owner = tk.LabelFrame(master, text='Add Owner')
        self.new_owner.grid(row=7, pady=5, padx=10, sticky='w')
        tk.Entry(self.new_owner, textvariable=self.new_owner_name).grid(row=0, column=0, pady=5, padx=5, sticky='w')
        tk.Button(self.new_owner, text='Add', command=self.add_owner, padx=5).grid(row=0, column=1, pady=5, padx=5,
                                                                                                     sticky='w')
        self.build_owners_list_box()

        self.check_version()

        master.bind('<Button-1>', self.disable_delete_button)

    def build_owners_list_box(self):
        self.owners = sorted([i for i in self.data.keys()])
        self.ol.delete(0, 'end')
        self.fl.delete(0, 'end')
        tk.OptionMenu(self.update_group, self.update_name, *[i for i in self.owners]).grid(row=0, column=0, pady=5,
                                                                                           padx=0)
        for i in self.owners:
            self.ol.insert('end', i)

        for i in os.listdir(self.variables['files_path']):
            self.fl.insert('end', i)

    def add_delete_button(self):
        self.delete_button.config(state='normal', background='red')

    def disable_delete_button(self, event):
        if event.widget is not self.delete_button:
            self.delete_button.config(state='disabled', background='SystemButtonFace')

    def delete_owner(self):
        del(self.data[self.update_name.get()])
        self.save_data()
        self.build_owners_list_box()
        self.delete_button.config(state='disabled', background='SystemButtonFace')

    def check_version(self):
        if sys.version_info[0] < 3:
            self.message.set("Must be using Python 3")


    def open_frame(self):
        DefaultsWindow(self)

    def eidt_lists(self):
        ListWindow(self)

    def add_owner(self):
        name = self.new_owner_name.get()
        if name not in self.data:
            self.data[name] = ['Internal Only']
            self.save_data()
            self.build_owners_list_box()

    def save_data(self):
        with open('- Tool\data.txt', 'w') as file:
            json.dump(self.data, file)

    def watermark_some(self):
        files_list = self.fl.curselection()
        owners = self.ol.curselection()
        files_list = [self.fl.get(i) for i in files_list]
        owners = [self.ol.get(i) for i in owners]
        self.run(files_list, owners)

    def watermark_all(self):
        files_path = self.variables['files_path']
        files_list = os.listdir(files_path)
        owners = self.data.keys()
        self.run(files_list, owners)

    def run(self, files_list, owners):
        last_page = self.last_page.get()
        watermarks_path = self.variables['watermarks_path']
        save_path = self.variables['save_path']
        files_path = self.variables['files_path']
        watermark_prefix = self.variables['watermark_prefix']

        # Generate list of names and sizes to watermark with
        watermarks_list = os.listdir(watermarks_path)
        wm = WatermarkMaker(watermarks_path, save_path)

        wm_count = 0
        for i in owners:
            wm_count += len(self.data[i])

        file_count = (len(files_list) * wm_count)
        progress = 0
        self.message.set('Running......' + str(progress) + ' of ' + str(file_count))
        master.update()

        for file in files_list:
            output_file_name = file.split('.')[0]
            source_pdf = files_path + file
            for owner in owners:
                for name in self.data[owner]:
                    source = pdf.PdfReader(source_pdf)
                    size = source.pages[0].MediaBox[2:4]
                    h = round(float(size[0]) / 72, 2)
                    w = round(float(size[1]) / 72, 2)
                    watermark_file_name = name + "_" + str(h) + "_" + str(w) + ".pdf"
                    if watermark_file_name not in watermarks_list:
                        wm.make_watermark(name, watermark_file_name, h, w, watermark_prefix)
                    wm.watermark(owner, name, h, w, source, last_page, output_file_name)
                    progress += 1
                    self.message.set('Running......' + str(progress) + ' of ' + str(file_count))
                    master.update()
        self.message.set("Done!")


class DefaultsWindow(tk.Toplevel):
    def __init__(self, master):
        self.main_window = master
        tk.Toplevel.__init__(self)
        self.title('Settings')

        tk.Label(self, text="Settings", pady=10, font=('Arial', 12), anchor='w').grid(row=0, pady=5, columnspan=2, sticky='w')

        self.defaults_dict = {"Watermarks Folder": 'watermarks_path',
                         'Save Watermarked Files': 'save_path',
                         'Files to Watermark': 'files_path',
                              'Watermark Prefix': 'watermark_prefix'}

        i = 1
        self.entry_dict = {}
        for k, v in self.defaults_dict.items():
            tk.Label(self, text=k, pady=5).grid(row=i, sticky='w')
            self.entry_dict[v] = tk.Entry(self, width=150)
            self.entry_dict[v].grid(row=i, column=2, pady=5)
            self.entry_dict[v].insert(0, self.main_window.variables[v])
            i += 1

        tk.Button(self, text='Save', command=self.save_paths).grid(row=11, pady=10, padx=10, sticky='w')

    def save_paths(self):
        for v in self.defaults_dict.values():
            self.main_window.variables[v] = self.entry_dict[v].get()

        with open('- Tool\defaults.txt', 'w') as file:
            json.dump(self.main_window.variables, file)
        self.destroy()


class ListWindow(tk.Toplevel):
    def __init__(self, master):
        self.main_window = master
        tk.Toplevel.__init__(self)
        self.title('Watermark Lists')
        self.update_name = self.main_window.update_name.get()

        tk.Label(self, text="Watermark List", pady=10, font=('Arial', 12)).grid(row=0, pady=5, columnspan=2)
        self.t = tk.Text(self)
        watermark_list = self.main_window.data[self.update_name]
        for i in sorted(watermark_list, reverse=True):
            self.t.insert(1.0, i + '\n')
        self.t.grid(row=2)

        tk.Button(self, text='Save', command=self.save).grid(row=11, pady=10, padx=10, sticky='w')

    def save(self):
        contents = self.t.get('1.0', 'end-1c')
        new_list = []
        for i in contents.splitlines():
            new_list.append(i.strip())
        self.main_window.data[self.update_name] = new_list
        self.main_window.save_data()

        self.destroy()


class WatermarkMaker:
    def __init__(self, watermarks_path, save_path):
        self.watermarks_path = watermarks_path
        self.save_path = save_path

    def make_watermark(self, name, watermark_file_name, h, w, watermark_prefix):
        save_file = self.watermarks_path + watermark_file_name
        watermark_text = watermark_prefix + name
        size = ((h + w) / 2) * 3
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
             size=size
             )
        ax.set_xticklabels('')
        ax.set_yticklabels('')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_visible(False)
        ax.spines['bottom'].set_visible(False)
        plt.savefig(save_file, dpi=500, transparent=True)
        plt.close(fig)

    def watermark(self, owner, name, h, w, source, last_page, output_file_name):
        writer = pdf.PdfWriter()
        watermark_pdf = self.watermarks_path + name + "_" + str(h) + "_" + str(w) + ".pdf"
        watermark = pdf.PageMerge().add(pdf.PdfReader(watermark_pdf).pages[0])[0]
        i = 1
        for page in source.pages:
            if len(source.pages) > i:
                pdf.PageMerge(page).add(watermark).render()
            else:
                if last_page == 1:
                    pdf.PageMerge(page).add(watermark).render()
            writer.addpage(page)
            i += 1
        file_name = self.save_path + owner + '/' + name + '/' + output_file_name + ' - ' + name + '.pdf'
        os.makedirs(os.path.dirname(file_name), exist_ok=True)
        writer.write(file_name)


if __name__ == "__main__":
    master = tk.Tk()
    app = App(master)
    master.mainloop()
