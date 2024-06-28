from tkinter import *
from tkinter import filedialog
import concurrent.futures
from tkinter import messagebox


# COMPARE WORD FILE PROGRAMME

import difflib
import docx2txt

total_point = 100.0
main_file_path = None
compare_file_path = None
accuracy_holder = None
cases = []

comma_count = 0
dot_count = 0
space_count = 0
inverted_comma_count = 0
double_inverted_comma_count = 0
exclamation_count = 0
other_count = 0
spelling_capital_small = 0

from reportlab.platypus import SimpleDocTemplate, Image, Table, TableStyle, Spacer, Paragraph, PageBreak
from reportlab.lib.styles import ParagraphStyle, TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors


def create_pdf_reportlab():
    logo = "./img.png"

    total_count = comma_count + dot_count + space_count + inverted_comma_count + double_inverted_comma_count \
                    + exclamation_count + other_count + spelling_capital_small

    total_error = float(comma_count * 0.50) + float(dot_count * 0.50) + float(space_count * 0.75) \
                    + float(inverted_comma_count * 0.50) + float(double_inverted_comma_count * 0.50) \
                    + float(exclamation_count * 0.50) + float(other_count * 0.50) + float(spelling_capital_small * 0.75)

    data = [
        ["No.", "Description", "Total Mistakes", "Per Mistake Count", "Total"],
        ["1", "Comma", str(comma_count), "0.50", str(float(comma_count * 0.50))],
        ["2", "Dot", str(dot_count), "0.50",str(float(dot_count * 0.50))],
        ["3", "Space", str(space_count), "0.75",str(float(space_count * 0.75))],
        ["4", "Inverted Comma", str(inverted_comma_count), "0.50",str(float(inverted_comma_count * 0.50))],
        ["5", "Double Inverted Comma",str(double_inverted_comma_count), "0.50", str(float(double_inverted_comma_count * 0.50))],
        ["6", "Exclamation",str(exclamation_count), "0.50",str(float(exclamation_count * 0.50))],
        # ["7", "Other",str(other_count), "0.50",str(float(other_count * 0.50))],
        ["7", "Spelling Mistake",str(spelling_capital_small), "0.75", str(float(spelling_capital_small * 0.75))],
        ["", "Total", str(total_count), "2.5", str(total_error)]
    ]

    files_extentions = [('pdf', '*.pdf*')]
    file_location = filedialog.asksaveasfile(filetypes = files_extentions, defaultextension = files_extentions)

    if file_location == None or file_location == '':
        return


    pdf = SimpleDocTemplate(
        filename=file_location.name,
        pagesize=letter,
        topMargin=10

    )

    table = Table(data)

    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.black),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Courier-Bold'),
        ('FONTNAME', (0, 0), (-1, 0), 'Courier-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 14),
        ('FONTSIZE', (0, 4), (-1, 0), 14),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('TOPPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.black)
    ])

    para_style = ParagraphStyle(
        name='bold',
        fontSize=30,
        alignment=TA_CENTER
    )

    footer_style = ParagraphStyle(
        name='bold',
        fontSize=12,
        alignment=TA_CENTER
    )

    table.setStyle(style)

    final_pdf = []

    im = Image(logo)

    accuracy_text_pdf = Paragraph("Your Accuracy is "+str(total_point)+"%", para_style)


    space = Spacer(width=0, height=40)

    final_pdf.append(im)
    final_pdf.append(space)
    final_pdf.append(Paragraph("Typing Test Result", para_style))
    final_pdf.append(Spacer(width=0, height=60))
    final_pdf.append(table)
    final_pdf.append(Spacer(width=0, height=40))
    final_pdf.append(accuracy_text_pdf)
    final_pdf.append(Spacer(width=0, height=240))
    final_pdf.append(Paragraph("For further information contact \n https://github.com/codeterrayt - rohanprajapati7860@gmail.com", footer_style))



    pdf.build(final_pdf)
    messagebox.showinfo("showinfo", "Pdf Generated Successfully!")

# _ignore = [".",",","-","_"]
#
# def get_word(index):
#     main_file_sentence = cases[0][0]
#     back_count  = index
#     add_count = index
#
#     check_first = False
#     check_second = False
#
#     for check_space in range(0,100):
#         try:
#             if main_file_sentence[back_count] == " " or main_file_sentence[back_count] == "" or main_file_sentence[back_count] in _ignore:
#                 if main_file_sentence[back_count] == " ":
#                     back_count +=1
#                 else:
#                     check_first = True
#             else:
#                back_count -= 1
#         except:
#             pass
#         try:
#             if main_file_sentence[add_count] == " " or main_file_sentence[add_count] == "" or main_file_sentence[
#                 add_count] in _ignore:
#                 check_second = True
#             else:
#                 add_count +=1
#         except:
#             pass
#
#         if check_first and check_second:
#             break
#
#     if back_count < 0:
#         back_count = 0
#     return [back_count,add_count]
    # print(main_file_sentence[back_count:add_count])
    # print([back_count, add_count])



def depth_compare(s,i):
    global total_point
    global dot_count
    global comma_count
    global space_count
    global inverted_comma_count
    global double_inverted_comma_count
    global exclamation_count
    global other_count
    global spelling_capital_small

    single_inverted_comma = ["\u0027", "\u0060", "\u00B4", "\u2018", "\u2019"]
    double_invered_comma = ["\u02EE", "\u0022", "\u3003", "\u05F2", "\u1CD3", "\u02f6", "\u201D", "\u201F", "\u201C"]

    if s[0] == ' ':
        pass
    elif s[0] == '-':
        if s[-1] in single_inverted_comma or s[-1] in double_invered_comma or s[-1] == "." or s[-1] == "," or s[-1] == "'" or s[-1] == '"' or s[-1] == '“' or  s[-1] == '‘' or s[-1] == "!":
            if s[-1] == ".":
                dot_count +=1
            if s[-1] == ",":
                comma_count +=1
            if s[-1] in single_inverted_comma:
                inverted_comma_count +=1
            if s[-1] in double_invered_comma:
                double_inverted_comma_count +=1

            if s[-1] == "!":
                exclamation_count +=1
            total_point = total_point - 0.50
        elif s[-1] == " ":
            space_count +=1
            total_point = total_point - 0.75
        else:
            spelling_capital_small +=1
            total_point = total_point - 0.75
        # print(u'not Present "{}" from position {}'.format(s[-1], i))
    elif s[0] == '+':
        if s[-1] in single_inverted_comma or s[-1] in double_invered_comma or s[-1] == "." or s[-1] == "," or s[
            -1] == "'" or s[-1] == '"' or s[-1] == '“' or s[-1] == '‘' or s[-1] == "!":
            if s[-1] == ".":
                dot_count += 1
            if s[-1] == ",":
                comma_count += 1
            if s[-1] in single_inverted_comma:
                inverted_comma_count += 1
            if s[-1] in double_invered_comma:
                double_inverted_comma_count += 1
            if s[-1] == "!":
                exclamation_count +=1
            total_point = total_point - 0.50
        elif s[-1] == " ":
            space_count +=1
            total_point = total_point - 0.75
        else:
            spelling_capital_small +=1
            total_point = total_point - 0.75

        # spelling_capital_small +=1
        # total_point = total_point - 0.75
        # myWord = get_word(i)
        # print(cases[0][1][myWord[0]:myWord[1]]+" => "+cases[0][0][myWord[0]:myWord[1]] )
        # print(u'Extra  "{}" to position {}'.format(s[-1], i))

def compare():
    global cases

    main_doc = ''
    compare_doc = ''
    main_file_path_list = list(main_file_path)
    compare_file_path_list = list(compare_file_path)

    try:
        main_file_path_list.sort(key=lambda f: int(re.sub('\D', '', f)))
        compare_file_path_list.sort(key=lambda f: int(re.sub('\D', '', f)))
    except:
        pass

    for docx_file in main_file_path_list:
        main_doc += docx2txt.process(docx_file)

    for compare_docx_File in compare_file_path_list:
        compare_doc += docx2txt.process(compare_docx_File)

    cases = [(main_doc, compare_doc)]
    for a, b in cases:
        for i, s in enumerate(difflib.ndiff(a, b)):
            depth_compare(s,i)
    return total_point


def generate_pdf():
    if accuracy_holder != None:
        create_pdf_reportlab()
    else:
        messagebox.showinfo("showinfo", "Please First Compare The File")



# GUI

root = Tk()
root.title("Compare Docx")
root.minsize(500,500)
root.maxsize(500,500)
root.iconphoto(False,PhotoImage(file="./myicon.png"))

def reset_all():
    global total_point
    global main_file_path
    global compare_file_path
    global accuracy_holder
    global dot_count
    global comma_count
    global space_count
    global cases
    global inverted_comma_count
    global double_inverted_comma_count
    global exclamation_count
    global other_count
    global spelling_capital_small

    total_point = 100.0
    main_file_path = None
    compare_file_path = None
    accuracy_holder = None
    dot_count = 0
    comma_count = 0
    space_count = 0
    inverted_comma_count = 0
    double_inverted_comma_count = 0
    exclamation_count = 0
    other_count = 0
    spelling_capital_small = 0
    cases = []

    generate_pdf_btn.pack_forget()
    accuracy.config(text="")
    reset_btn.pack_forget()

    main_path_label.config(text="File Not Selected" , fg="red")
    compare_path_label.config(text="File Not Selected" , fg="red")

    choose_main_file_btn.config(text="Choose", bg="black" )
    choose_compare_file_btn.config(text="Choose", bg="black")

    start_comparision_btn.config(text="Start Comparison", bg="black")


def chooseMainFile():
    global main_file_path
    global accuracy_holder
    accuracy_holder = None
    main_file_path = filedialog.askopenfilenames(filetypes=[("Excel files", ".docx .doc")])
    if main_file_path != '':
        main_path_label.config(text="File Selected!" , fg="green")
        choose_main_file_btn.config(bg="blue" , text="Selected")

def chooseCompareFile():
    global compare_file_path
    global accuracy_holder
    accuracy_holder = None
    compare_file_path = filedialog.askopenfilenames(filetypes=[("Excel files", ".docx .doc")])
    if compare_file_path != '':
        compare_path_label.config(text="File Selected!" , fg="green")
        choose_compare_file_btn.config(bg="blue", text="Selected")

def start_compare():
    global accuracy
    global accuracy_holder
    if main_file_path == None or main_file_path == '' or compare_file_path == None or compare_file_path == '':
        if main_file_path == None or main_file_path == '':
            messagebox.showinfo("showinfo", "Please Choose Main File")
        if compare_file_path == None or compare_file_path == '':
            messagebox.showinfo("showinfo", "Please Choose Compare File")
        return

    elif len(main_file_path) != len(compare_file_path):
        messagebox.showinfo("showinfo", "Please Choose Equal Files")
        return

    with concurrent.futures.ThreadPoolExecutor() as executor:
        start_comparision_btn.config(text="Comparing...Please Wait" , bg="blue")
        root.update()
        future = executor.submit(compare)
        return_value = future.result()
        accuracy_holder = return_value
        if return_value < 0:
            accuracy_holder = 0
        accuracy.config(text = "Accuracy = "+str(accuracy_holder)+"%")
        start_comparision_btn.config(text="File Compared Successfully!" , bg="green")
        generate_pdf_btn.pack(pady=10)
        reset_btn.pack()


Label(text="Compare Docx" , font =("Arial",20,"bold")).pack(pady=15)

Label(text="Choose Main File" , font =("Arial",15,"italic")).pack(pady=10)
choose_main_file_btn = Button(text="Choose" , fg="white" , bg="black" , padx=20 ,borderwidth=0 ,command=chooseMainFile)
choose_main_file_btn.pack(pady=10)
main_path_label = Label(text="File Not Selected" , font =("Arial",10,"italic") , fg="red")
main_path_label.pack()

Label(text="Choose Compare File" , font =("Arial",15,"italic")).pack(pady=10)
choose_compare_file_btn = Button(text="Choose" , fg="white" , bg="black" , padx=20 ,borderwidth=0 ,command=chooseCompareFile)
choose_compare_file_btn.pack(pady=10)
compare_path_label = Label(text="File Not Selected" , font =("Arial",10,"italic") , fg="red")
compare_path_label.pack()

start_comparision_btn = Button(text="Start Comparison" , fg="white" , bg="black" , padx=20 ,borderwidth=0 ,command= lambda  : start_compare())
start_comparision_btn.pack(pady=10)

accuracy = Label(text="" , font =("Arial",25,"italic"))
accuracy.pack(pady=10)

generate_pdf_btn = Button(text="Generate PDF" , fg="white" , bg="green" ,borderwidth =0 , padx=60 , pady=10, command=generate_pdf)
reset_btn = Button(text="RESET" , fg="white" , bg="red" ,borderwidth =0 , padx=60 , command=reset_all)

root.mainloop()