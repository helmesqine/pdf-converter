from customtkinter import *
from tkinter import filedialog
from threading import Timer
from PIL import Image
from tkinter import messagebox
import PDFTOOLS

root = CTk()
root.title('PDF CONVERTER')
root.iconbitmap('pdf_icon.ico')
root.geometry('800x600')
root.resizable(False, False)
set_appearance_mode("light")
set_default_color_theme('green')
tabview = CTkTabview(root, height=590)
tabview.pack(side=BOTTOM, anchor='center', fill=X)
IMG2PDF = tabview.add("IMG > PDF")
WORD2PDF = tabview.add("WORD > PDF")
PDF2IMG = tabview.add("PDF > IMG")
PDF2WORD = tabview.add("PDF > WORD")

CONVERTER = PDFTOOLS.TOOLS()

bar = CTkProgressBar(root)
bar_text = CTkLabel(root, text='Processing...')

files = []
file_type = None
active = str(tabview.get()).replace(' > ', '2')

output_folder = ''

move = 5
start = 0
end = 5


def brows_output():
    global output_folder
    output_folder = filedialog.askdirectory()
    if str(output_folder) != '':
        globals()[str(tabview.get()).replace(' > ', '2') + 'output_entry'].configure(state=NORMAL)
        globals()[str(tabview.get()).replace(' > ', '2') + 'output_entry'].delete(0, END)
        globals()[str(tabview.get()).replace(' > ', '2') + 'output_entry'].insert(0, output_folder)
        globals()[str(tabview.get()).replace(' > ', '2') + 'output_entry'].configure(state='readonly')


def display_function(file):
    file_name, file_extension = os.path.splitext(file)
    if file_extension in ['.docx', '.doc']:
        return CTkImage(Image.open('word.png'), size=(80, 90))
    elif file_extension in ['.png', '.jpg']:
        try:
            img = Image.open(file).resize((80, 90))
        except:
            img = Image.open('image.png')
        return CTkImage(img, size=(80, 90))
    elif file_extension in ['.pdf']:
        return CTkImage(Image.open('pdf.png'), size=(80, 90))


def global_import():
    global files
    global move
    global start
    global end
    if str(tabview.get()).replace(' > ', '2') == 'IMG2PDF':
        files = list(filedialog.askopenfilenames(filetypes=[("Images", ".jpg .png")]))
    elif str(tabview.get()).replace(' > ', '2') == 'WORD2PDF':
        files = list(filedialog.askopenfilenames(filetypes=[("WORD", ".doc .docx")]))
    else:
        files = list(filedialog.askopenfilenames(filetypes=[("PDF", ".pdf")]))
    for widgets in globals()[str(tabview.get()).replace(' > ', '2') + '_display_frame'].winfo_children():
        widgets.destroy()
    if len(files) > 0:
        if len(files) > 5:
            globals()[str(tabview.get()).replace(' > ', '2') + 'f_button'].configure(state=NORMAL)
            globals()[str(tabview.get()).replace(' > ', '2') + 'b_button'].configure(state=NORMAL)
        elif len(files) <= 5:
            globals()[str(tabview.get()).replace(' > ', '2') + 'f_button'].configure(state=DISABLED)
            globals()[str(tabview.get()).replace(' > ', '2') + 'b_button'].configure(state=DISABLED)
        """
        x = round(len(files) / 5)
        if len(files) - (x * 5) != 0:
            pages_number = x + 1
        else:
            pages_number = x
        """
        start = 0
        end = 5
        Timer(0, lambda: set_files(start, end)).start()
    else:
        globals()[str(tabview.get()).replace(' > ', '2') + '_entry'].configure(state=NORMAL)
        globals()[str(tabview.get()).replace(' > ', '2') + '_entry'].delete(0, END)
        globals()[str(tabview.get()).replace(' > ', '2') + '_entry'].configure(state=DISABLED)
        files = ''


def change_page(w):
    global move
    global start
    global end
    try:
        if w == 'f':
            start += move
            end += move
            if not files[start:end]:
                start -= move
                end -= move
            else:
                set_files(start, end)
        elif w == 'b':
            start -= move
            end -= move
            if not files[start:end]:
                start += move
                end += move
            else:
                set_files(start, end)
    except:
        pass


def set_files(start_, end_):
    global files
    global move
    global start
    global end
    i = 0
    for widgets in globals()[str(tabview.get()).replace(' > ', '2') + '_display_frame'].winfo_children():
        widgets.destroy()
    displayed_files = files[start_:end_]
    for file in displayed_files:
        image_to_display = display_function(file)
        globals()[str(tabview.get()).replace(' > ', '2') + f'frame{i}'] = CTkFrame(
            globals()[str(tabview.get()).replace(' > ', '2') + '_display_frame'], width=80)
        globals()[str(tabview.get()).replace(' > ', '2') + f'frame{i}'].pack(side=LEFT, fill=Y, padx=5, pady=3)
        globals()[str(tabview.get()).replace(' > ', '2') + f'lable{i}'] = CTkLabel(
            globals()[str(tabview.get()).replace(' > ', '2') + f'frame{i}'],
            text=str(tabview.get()).replace(' > ', '2') + f'lable{i}',
            width=80, image=image_to_display, font=('arial', 1))
        globals()[str(tabview.get()).replace(' > ', '2') + f'lable{i}'].pack(side=TOP, fill='both')

        globals()[str(tabview.get()).replace(' > ', '2') + f'text{i}'] = CTkLabel(
            globals()[str(tabview.get()).replace(' > ', '2') + f'frame{i}'],
            text=os.path.basename(file)[:8] + '...')
        globals()[str(tabview.get()).replace(' > ', '2') + f'get_text{i}'] = CTkLabel(
            globals()[str(tabview.get()).replace(' > ', '2') + f'frame{i}'],
            text=os.path.basename(file))
        globals()[str(tabview.get()).replace(' > ', '2') + f'get_path{i}'] = CTkLabel(
            globals()[str(tabview.get()).replace(' > ', '2') + f'frame{i}'],
            text=file)
        globals()[str(tabview.get()).replace(' > ', '2') + f'text{i}'].pack(side=BOTTOM)

        globals()[str(tabview.get()).replace(' > ', '2') + f'lable{i}'].bind("<Enter>", display_remove_sign)
        globals()[str(tabview.get()).replace(' > ', '2') + f'lable{i}'].bind("<Leave>", disappear_remove_sign)
        globals()[str(tabview.get()).replace(' > ', '2') + f'lable{i}'].bind("<Button-1>", remove_element)
        i += 1


def import_display_frames(list_display):
    for frame in list_display:
        globals()[str(frame) + '_display_frame'] = CTkFrame(globals()[str(frame)], height=120, width=450,
                                                            border_color='#2CC985', border_width=3)
        globals()[str(frame) + '_display_frame'].pack_propagate(False)
        globals()[str(frame) + '_display_frame'].place(relx=0.4, rely=0.3, anchor='center')

        globals()[str(frame) + '_display_filename'] = CTkLabel(globals()[str(frame)], width=450, height=20, text='',
                                                               text_color='#2CC985')
        globals()[str(frame) + '_display_filename'].pack_propagate(False)
        globals()[str(frame) + '_display_filename'].place(relx=0.4, rely=0.44, anchor='center')

        globals()[str(frame) + 'f_button'] = CTkButton(globals()[str(frame)], text='>',
                                                       command=lambda: change_page('f'), state=DISABLED)
        globals()[str(frame) + 'b_button'] = CTkButton(globals()[str(frame)], text='<',
                                                       command=lambda: change_page('b'), state=DISABLED)
        globals()[str(frame) + 'f_button'].place(relx=0.71, rely=0.25, anchor='center', width=25)
        globals()[str(frame) + 'b_button'].place(relx=0.71, rely=0.37, anchor='center', width=25)

        globals()[str(frame) + 'entry_text'] = CTkLabel(globals()[str(frame)], text='Files ')
        globals()[str(frame) + 'entry_text'].place(relx=0.14, rely=0.14, anchor='center')

        globals()[str(frame) + 'output_entry'] = CTkEntry(globals()[str(frame)], width=400, state='readonly')
        globals()[str(frame) + 'output_entry']._entry.configure(readonlybackground='white')
        globals()[str(frame) + 'output_entry'].place(relx=0.45, rely=0.55, anchor='center')
        globals()[str(frame) + 'output_text'] = CTkLabel(globals()[str(frame)], text='Output path ')
        globals()[str(frame) + 'output_text'].place(relx=0.15, rely=0.55, anchor='center')
        globals()[str(frame) + 'output_button'] = CTkButton(globals()[str(frame)], text='Browse', command=brows_output)
        globals()[str(frame) + 'output_button'].place(relx=0.8, rely=0.55, anchor='center')

        if str(frame) != 'PDF2IMG' and str(frame) != 'PDF2WORD':
            globals()[str(frame) + 'combo'] = CTkComboBox(globals()[str(frame)], values=['Single file', 'Multi files'],
                                                          state='readonly')
            globals()[str(frame) + 'combo']._entry.configure(readonlybackground='white')
            globals()[str(frame) + 'combo'].set('Single file')
            globals()[str(frame) + 'combo'].place(relx=0.3, rely=0.7, anchor='center')
            globals()[str(frame) + 'combo_text'] = CTkLabel(globals()[str(frame)], text='Output type')
            globals()[str(frame) + 'combo_text'].place(relx=0.15, rely=0.7, anchor='center')

        if str(frame) == 'IMG2PDF':
            globals()[str(frame) + 'combo_size'] = CTkComboBox(globals()[str(frame)],
                                                               values=['Same as image', 'A4', 'A5', 'A3'],
                                                               state='readonly')
            globals()[str(frame) + 'combo_size']._entry.configure(readonlybackground='white')
            globals()[str(frame) + 'combo_size'].set('Same as image')
            globals()[str(frame) + 'combo_size'].place(relx=0.7, rely=0.7, anchor='center')
            globals()[str(frame) + 'combo_size_text'] = CTkLabel(globals()[str(frame)], text='Output Size')
            globals()[str(frame) + 'combo_size_text'].place(relx=0.55, rely=0.7, anchor='center')

        globals()[str(frame) + '_btn'] = CTkButton(globals()[str(frame)], text='Import', command=global_import,
                                                   height=120)
        globals()[str(frame) + '_btn'].place(relx=0.82, rely=0.3, anchor='center')


def display_remove_sign(event):
    global old_image
    if "<class 'tkinter.Label'>" == str(type(event.widget)):
        old_image = globals()[str(event.widget.cget("text"))].cget('image')
        filepath = globals()[str(event.widget.cget("text"))].cget('text')
        img = Image.open('remove.png').resize((80, 100))
        globals()[str(event.widget.cget("text"))].configure(image=CTkImage(img, size=(80, 100)))
        globals()[str(event.widget.cget("text"))].image = CTkImage(img)
        main = str(event.widget.cget("text")).split('lable')[0]
        n_main = str(event.widget.cget("text")).split('lable')[1]
        globals()[str(main) + '_display_filename'].configure(text=globals()[f'{main}get_text{n_main}'].cget('text'))


def disappear_remove_sign(event):
    global old_image
    if "<class 'tkinter.Label'>" == str(type(event.widget)):
        globals()[str(event.widget.cget("text"))].configure(image=old_image)
        globals()[str(event.widget.cget("text"))].image = old_image
        main = str(event.widget.cget("text")).split('lable')[0]
        globals()[str(main) + '_display_filename'].configure(text='')


def remove_element(event):
    global files
    global move
    global start
    global end
    if "<class 'tkinter.Label'>" == str(type(event.widget)):
        files.remove(globals()[str(event.widget.cget("text")).replace('lable', 'get_path')].cget("text"))
        globals()[str(event.widget.cget("text")).replace('lable', 'frame')].pack_forget()
        main = str(event.widget.cget("text")).split('lable')[0]
        globals()[str(main) + '_display_filename'].configure(text='')
    for widgets in globals()[str(tabview.get()).replace(' > ', '2') + '_display_frame'].winfo_children():
        widgets.destroy()
    set_files(start, end)
    if len(files) <= 5:
        globals()[str(tabview.get()).replace(' > ', '2') + 'f_button'].configure(state=DISABLED)
        globals()[str(tabview.get()).replace(' > ', '2') + 'b_button'].configure(state=DISABLED)
    if len(files) > 5:
        globals()[str(tabview.get()).replace(' > ', '2') + 'f_button'].configure(state=NORMAL)
        globals()[str(tabview.get()).replace(' > ', '2') + 'b_button'].configure(state=NORMAL)
    elif len(files) <= 5:
        set_files(0, 5)
        globals()[str(tabview.get()).replace(' > ', '2') + 'f_button'].configure(state=DISABLED)
        globals()[str(tabview.get()).replace(' > ', '2') + 'b_button'].configure(state=DISABLED)


def set_theme(e):
    set_appearance_mode(theme_select.get())
    set_default_color_theme('green')

    if str(theme_select.get()) == 'Light':
        for frame in main_list:
            if str(frame) != 'PDF2IMG' and str(frame) != 'PDF2WORD':
                globals()[str(frame) + 'combo']._entry.configure(readonlybackground='white')
                if str(frame) != 'PDF2IMG' and str(frame) != 'PDF2WORD':
                    try:
                        globals()[str(frame) + 'combo_size']._entry.configure(readonlybackground='white')
                    except:pass
            try:
                globals()[str(frame) + 'output_entry']._entry.configure(readonlybackground='white')
            except:
                pass
        theme_select._entry.configure(readonlybackground='white')
    elif str(theme_select.get()) == 'Dark':
        for frame in main_list:
            if str(frame) != 'PDF2IMG':
                try:
                    globals()[str(frame) + 'combo']._entry.configure(readonlybackground='#343638')
                except:
                    pass
                if str(frame) != 'PDF2IMG' and str(frame) != 'PDF2WORD':
                    try:
                        globals()[str(frame) + 'combo_size']._entry.configure(readonlybackground='#343638')
                    except:
                        pass
            try:
                globals()[str(frame) + 'output_entry']._entry.configure(readonlybackground='#343638')
            except:
                pass
        theme_select._entry.configure(readonlybackground='#343638')


def img2pdf_():
    if globals()['IMG2PDFoutput_entry'].get() == '':
        messagebox.showerror('PDFCONVERTER', 'Please choose output directory!')
        return
    if not files:
        messagebox.showerror('PDFCONVERTER', 'Please choose files to convert!')
        return
    Timer(0, progress_start).start()
    CONVERTER.image_to_pdf(files, globals()['IMG2PDFoutput_entry'].get(), globals()['IMG2PDFcombo'].get(),
                           globals()['IMG2PDFcombo_size'].get())
    Timer(0, progress_end).start()


def word2pdf_():
    if globals()['WORD2PDFoutput_entry'].get() == '':
        messagebox.showerror('PDFCONVERTER', 'Please choose output directory!')
        return
    if not files:
        messagebox.showerror('PDFCONVERTER', 'Please choose files to convert!')
        return
    Timer(0, progress_start).start()
    CONVERTER.word_to_pdf(files, globals()['WORD2PDFoutput_entry'].get(), globals()['WORD2PDFcombo'].get())
    Timer(0, progress_end).start()


def pdf2img_():
    if globals()['PDF2IMGoutput_entry'].get() == '':
        messagebox.showerror('PDFCONVERTER', 'Please choose output directory!')
        return
    if not files:
        messagebox.showerror('PDFCONVERTER', 'Please choose files to convert!')
        return
    Timer(0, progress_start).start()
    CONVERTER.pdf_to_img(files, globals()['PDF2IMGoutput_entry'].get())
    Timer(0, progress_end).start()


def pdf2doc_():
    if globals()['PDF2WORDoutput_entry'].get() == '':
        messagebox.showerror('PDFCONVERTER', 'Please choose output directory!')
        return
    if not files:
        messagebox.showerror('PDFCONVERTER', 'Please choose files to convert!')
        return
    Timer(0, progress_start).start()
    CONVERTER.pdf_to_doc(files, globals()['PDF2WORDoutput_entry'].get())
    Timer(0, progress_end).start()


def progress_start():
    theme_select.place_forget()
    tabview.pack_forget()
    bar.start()
    bar_text.place(relx=0.5, rely=0.45, anchor='center')
    bar.place(relx=0.5, rely=0.5, anchor='center')


def progress_end():
    global files
    for widgets in globals()[str(tabview.get()).replace(' > ', '2') + '_display_frame'].winfo_children():
        try:
            widgets.destroy()
        except:
            pass
    files = ''
    bar.stop()
    bar_text.place_forget()
    bar.place_forget()
    theme_select.place(relx=0.9, rely=0.97, anchor='center', height=25)
    tabview.pack(side=BOTTOM, anchor='center', fill=X)


main_list = ['IMG2PDF', 'WORD2PDF', 'PDF2IMG', 'PDF2WORD']
import_display_frames(main_list)

theme_select = CTkComboBox(root, values=['Light', 'Dark'], command=set_theme, state='readonly')
theme_select._entry.configure(readonlybackground='white')
theme_select.set('Light')
theme_select.place(relx=0.9, rely=0.97, anchor='center', height=25)

IMG2PDF_convert_button = CTkButton(globals()['IMG2PDF'], text='Convert', command=lambda: Timer(0, img2pdf_).start())
IMG2PDF_convert_button.place(relx=0.5, rely=0.85, anchor='center')

WORD2PDF_convert_button = CTkButton(globals()['WORD2PDF'], text='Convert', command=lambda: Timer(0, word2pdf_).start())
WORD2PDF_convert_button.place(relx=0.5, rely=0.85, anchor='center')

PDF2IMG_convert_button = CTkButton(globals()['PDF2IMG'], text='Convert', command=lambda: Timer(0, pdf2img_).start())
PDF2IMG_convert_button.place(relx=0.45, rely=0.7, anchor='center')

PDF2WORD_convert_button = CTkButton(globals()['PDF2WORD'], text='Convert', command=lambda: Timer(0, pdf2doc_).start())
PDF2WORD_convert_button.place(relx=0.45, rely=0.7, anchor='center')

root.mainloop()
