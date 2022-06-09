from calendar import Calendar
from tkcalendar import Calendar, DateEntry
from tkinter import *
import numpy as np
import pytesseract
from PIL import ImageTk, Image
from tkinter import ttk, filedialog
import os
import shutil
import cv2
from time import sleep
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import asksaveasfile
import mysql.connector
from datetime import date
from datetime import datetime
import pandas as pd

filepath = None
FileSizeLimit = 5000000
Image_data = []
myData = []
countSubmit = 1
modifiedCount = 1
DataOutputPath = 'UseForms\DataOutput.csv'
selecteddate1 = "not selected"
selecteddate2 = "not selected"

def Data_Show():
    mydb = mysql.connector.connect(host="localhost", user="root", passwd="", database="form_data")
    mycursor = mydb.cursor()

    def select_date1():
        if __name__ == "__main__":
            def destroy_frame_calender():
                frame_calendar.destroy()

            def print_sel():
                global selecteddate1
                selecteddate1 = cal.selection_get()
                frame_calendar.destroy()

            frame_calendar = tk.Frame(root1)
            sapretor = ttk.Separator(frame_calendar)
            frame_calendar.configure(background="#b5ffff")
            frame_calendar.place(relx=0.0, rely=0.0, relheight=1, relwidth=1)

            heading = Label(frame_calendar)
            btncancel = Button(frame_calendar)
            btnok = Button(frame_calendar)
            cal = Calendar(frame_calendar, font="Arial 14", selectmode='day', year=2022, month=4, day=1)

            heading.configure(text="Form Details Reader",
                              font="-family {Arial} -size 20", justify='center')
            btncancel.configure(text="Cancel", activeforeground="#000000", activebackground="#ececec",
                                disabledforeground="#a3a3a3",
                                font="-family {Arial} -size 12", highlightbackground="#d9d9d9", highlightcolor="black",
                                overrelief="sunken", pady="0", command=destroy_frame_calender)
            btnok.configure(text="Ok", activeforeground="#000000", activebackground="#ececec",  disabledforeground="#a3a3a3", font="-family {Arial} -size 12",
                            highlightbackground="#d9d9d9", highlightcolor="black", overrelief="sunken", pady="0",
                            command=print_sel)

            heading.place(relx=0.0, rely=0.059, height=48, width=655)
            btncancel.place(relx=0.29, rely=0.866, height=33, width=106)
            btnok.place(relx=0.549, rely=0.866, height=33, width=106)
            sapretor.place(relx=-0.015, rely=0.207, relwidth=1.037)
            cal.place(relx=0.076, rely=0.276, relheight=0.541, relwidth=0.846)

    def select_date2():
        if __name__ == "__main__":
            def destroy_frame_calender():
                frame_calendar.destroy()

            def print_sel():
                global selecteddate2
                selecteddate2 = cal.selection_get()
                frame_calendar.destroy()

            frame_calendar = tk.Frame(root1)
            sapretor = ttk.Separator(frame_calendar)
            frame_calendar.configure(background="#b5ffff")
            frame_calendar.place(relx=0.0, rely=0.0, relheight=1, relwidth=1)

            heading = Label(frame_calendar)
            btncancel = Button(frame_calendar)
            btnok = Button(frame_calendar)
            cal = Calendar(frame_calendar, font="Arial 14", selectmode='day', year=2022, month=4,
                           day=29)

            heading.configure(text="Form Details Reader",
                              font="-family {Arial} -size 20", justify='center')
            btncancel.configure(text="Cancel", activeforeground="#000000", activebackground="#ececec",
                               disabledforeground="#a3a3a3",
                                font="-family {Arial} -size 12", highlightbackground="#d9d9d9", highlightcolor="black",
                                overrelief="sunken", pady="0", command=destroy_frame_calender)
            btnok.configure(text="Ok", activeforeground="#000000", activebackground="#ececec", disabledforeground="#a3a3a3", font="-family {Arial} -size 12",
                            highlightbackground="#d9d9d9", highlightcolor="black", overrelief="sunken", pady="0",
                            command=print_sel)

            heading.place(relx=0.0, rely=0.059, height=48, width=655)
            btncancel.place(relx=0.29, rely=0.866, height=33, width=106)
            btnok.place(relx=0.549, rely=0.866, height=33, width=106)
            sapretor.place(relx=-0.015, rely=0.207, relwidth=1.037)
            cal.place(relx=0.076, rely=0.276, relheight=0.541, relwidth=0.846)

    def show_entry_data():
        # print(selecteddate1)
        # print(selecteddate2)
        if (selecteddate1 == "not selected"):
            statusbar['text'] = "*Select Starting Date"
            messagebox.showinfo("Empty Field", "Select Start Date")
        elif (selecteddate2 == "not selected"):
            statusbar['text'] = "*Select Ending Date"
            messagebox.showinfo("Empty field", "Select End Date")
        else:
            query = "SELECT Name,ID_No,Mobile,Branch,AdmissionYear,Email,AddressLine1,AddressLine2,Declaration,EntDate FROM form_entry WHERE EntDate >= %s and EntDate <= %s"
            mycursor.execute(query, (selecteddate1, selecteddate2,))
            myresult = mycursor.fetchall()
            for row in myresult:
                Scrolledtreeview1.insert("", tk.END, values=row)
            statusbar['text'] = 'Data From "' + str(selecteddate1) + '" to "' + str(selecteddate2) + '"'

    def delete():
        if (selecteddate1 == "not selected"):
            statusbar['text'] = "*Select Starting Date"
            messagebox.showinfo("Empty Field", "Select Start Date")
        elif (selecteddate2 == "not selected"):
            statusbar['text'] = "*Select Ending Date"
            messagebox.showinfo("Empty field", "Select End Date")
        else:
            itemfocus = Scrolledtreeview1.focus()
            itemlist = Scrolledtreeview1.item(itemfocus)
            itemvalue = itemlist['values']
            print(itemvalue)
            try:
                ID_No = itemvalue[1]
                mycursor.execute("DELETE FROM form_entry where ID_No=%s;", (ID_No,))
                mydb.commit()
                Scrolledtreeview1.delete(itemfocus)
                statusbar['text'] = ("Deleted ID Number : " + str(ID_No))
            except IndexError:
                statusbar['text'] = "*select record to delete"
                # messagebox.showinfo("Empty Selection","Select any Record")

    def download_entry_data():
        if (selecteddate1 == "not selected"):
            statusbar['text'] = "*Select Starting Date"
            messagebox.showinfo("Empty Field", "Select Stat Date")
        elif (selecteddate2 == "not selected"):
            statusbar['text'] = "*Select Ending Date"
            messagebox.showinfo("Empty field", "Select End Date")
        else:
            query = "SELECT Name,ID_No,Mobile,Branch,AdmissionYear,Email,	AddressLine1,AddressLine2,Declaration FROM form_entry WHERE EntDate >= %s and EntDate <= %s"
            mycursor.execute(query, (selecteddate1, selecteddate2,))
            myresult = mycursor.fetchall()

            time = date.today()
            s1 = selecteddate2
            if (s1 > time):
                s1 = time
            df = pd.DataFrame(myresult,
                              columns=['Name', 'ID', 'Mobile No', 'Branch', 'Admission Year', 'Email', 'Address Line 1',
                                       'Address Line 2', 'Declaration'])

            files = [('ExcelSheet', '*.xlsx'), ('ExcelSheet 2013', '*.xls')]
            initialfilename = "Entry (" + str(selecteddate1) + " to " + str(s1) + ").xlsx"
            initialpath = "c://"
            file1 = asksaveasfile(filetypes=files, defaultextension=".xlsx", initialfile=initialfilename,
                                  initialdir=initialpath, title="FRA")
            if file1:
                datatoexcel = pd.ExcelWriter(file1.name, engine='xlsxwriter')
                df.to_excel(datatoexcel, index=False, sheet_name="Sheet1")
                worksheet = datatoexcel.sheets['Sheet1']
                worksheet.set_column('A:A', 20)
                worksheet.set_column('B:B', 25)
                worksheet.set_column('C:C', 25)
                worksheet.set_column('F:F', 20)
                datatoexcel.save()
                statusbar['text'] = "Download success  from " + str(selecteddate1) + " to " + str(s1)

    def destroy_users_attendance():
        frame_show_entry.destroy()

    if __name__ == "__main__":
        frame_show_entry = tk.Frame(root1)
        frame_show_entry.configure(background="#b5ffff")
        frame_show_entry.place(relx=0.0, rely=0.0, relheight=1, relwidth=1)
        sapretor = ttk.Separator(frame_show_entry)

        heading = Label(frame_show_entry)
        sapretor = ttk.Separator(frame_show_entry)
        lblentry = Label(frame_show_entry)
        btnstart = Button(frame_show_entry)
        btnend = Button(frame_show_entry)
        btnback = Button(frame_show_entry)
        btndelete = Button(frame_show_entry)
        btndownload = Button(frame_show_entry)
        btndisplay = Button(frame_show_entry)
        statusbar = Label(frame_show_entry)
        Scrolledtreeview1 = ttk.Treeview(frame_show_entry)
        scroll_y = Scrollbar(frame_show_entry, orient=VERTICAL)

        heading.configure(text="Form Details Reader",
                          font="-family {Arial} -size 20", justify='center')
        lblentry.configure(text="Filled Form", background="#b5ffff", foreground="#000000",
                                font="-family {Arial} -size 15", justify='center')
        btnstart.configure(text="Start Date", activeforeground="#000000", activebackground="#ececec", disabledforeground="#a3a3a3",
                           font="-family {Arial} -size 15", highlightbackground="#d9d9d9", highlightcolor="black",
                           overrelief="sunken", pady="0", command=select_date1)
        btnend.configure(text="End Date", activeforeground="#000000", activebackground="#ececec", disabledforeground="#a3a3a3", font="-family {Arial} -size 15",
                         highlightbackground="#d9d9d9", highlightcolor="black", overrelief="sunken", pady="0",
                         command=select_date2)
        btnback.configure(text="Back", activeforeground="#000000", activebackground="#ececec", disabledforeground="#a3a3a3", font="-family {Arial} -size 15",
                          highlightbackground="#d9d9d9", highlightcolor="black", overrelief="sunken", pady="0",
                          command=destroy_users_attendance)
        btndelete.configure(text="Delete", activeforeground="#000000", activebackground="#ececec", disabledforeground="#a3a3a3", font="-family {Arial} -size 15",
                            highlightbackground="#d9d9d9", highlightcolor="black", overrelief="sunken", pady="0",
                            command=delete)
        btndownload.configure(text="Download", activeforeground="#000000", activebackground="#ececec",
                              disabledforeground="#a3a3a3",
                              font="-family {Arial} -size 15", highlightbackground="#d9d9d9", highlightcolor="black",
                              overrelief="sunken", pady="0", command=download_entry_data)
        btndisplay.configure(text="Display", activeforeground="#000000", activebackground="#ececec",
                            disabledforeground="#a3a3a3",
                             font="-family {Arial} -size 15", highlightbackground="#d9d9d9", highlightcolor="black",
                             overrelief="sunken", pady="0", command=show_entry_data)
        statusbar.configure(text="OCR Based Text Extraction",
                            font="-family {Arial} -size 12", justify='center')
        Scrolledtreeview1.configure(columns=("1", "2", "3", "4", "5", "6", "7", "8", "9", "10"),
                                    yscrollcommand=scroll_y.set)
        scroll_y.configure(command=Scrolledtreeview1.yview)

        heading.place(relx=0.0, rely=0.059, height=48, width=655)
        lblentry.place(relx=0.32, rely=0.217, height=30, width=240)
        btnstart.place(relx=0.168, rely=0.295, height=30, width=150)
        btnend.place(relx=0.61, rely=0.295, height=30, width=150)
        btnback.place(relx=0.091, rely=0.807, height=30, width=130)
        btndelete.place(relx=0.518, rely=0.807, height=30, width=130)
        btndownload.place(relx=0.305, rely=0.807, height=30, width=130)
        btndisplay.place(relx=0.732, rely=0.807, height=30, width=130)
        statusbar.place(relx=0.0, rely=0.925, height=26, width=662)
        sapretor.place(relx=-0.015, rely=0.207, relwidth=1.037)
        Scrolledtreeview1.place(relx=0.091, rely=0.394, relheight=0.388, relwidth=0.837)
        scroll_y.place(relx=0.899, rely=0.399, relwidth=0.031, relheight=0.38)
        Scrolledtreeview1.heading("1", text="Name")
        Scrolledtreeview1.heading("2", text="ID")
        Scrolledtreeview1.heading("3", text="Mobile No")
        Scrolledtreeview1.heading("4", text="Branch")
        Scrolledtreeview1.heading("5", text="Admission Year")
        Scrolledtreeview1.heading("6", text="Email")
        Scrolledtreeview1.heading("7", text="Address Line 1")
        Scrolledtreeview1.heading("8", text="Address Line 2")
        Scrolledtreeview1.heading("9", text="Declaration")
        Scrolledtreeview1.heading("10", text="Date")
        Scrolledtreeview1['show'] = 'headings'
        Scrolledtreeview1.column("1", width=100)
        Scrolledtreeview1.column("2", width=100)
        Scrolledtreeview1.column("3", width=100)
        Scrolledtreeview1.column("4", width=100)
        Scrolledtreeview1.column("5", width=100)
        Scrolledtreeview1.column("6", width=100)
        Scrolledtreeview1.column("7", width=100)
        Scrolledtreeview1.column("8", width=100)
        Scrolledtreeview1.column("9", width=100)
        Scrolledtreeview1.column("10", width=100)


def destroy_frame():
    root1.destroy()


if __name__ == "__main__":
    root1 = Tk()
    frame = tk.Frame(root1)
    sapretor = ttk.Separator(frame)
    root1.geometry("656x508+450+150")
    root1.minsize(148, 1)
    root1.maxsize(1924, 1030)
    root1.resizable(False, False)
    root1.title("Form Details Reader")
    root1.protocol('WM_DELETE_WINDOW', exit)
    frame.configure(background="#b5ffff")
    frame.place(relx=0.0, rely=0.0, relheight=1, relwidth=1)


    def Second_page():
        def Open_image_window():
            global myData
            vname = StringVar()
            vid = StringVar()
            vmobile = StringVar()
            vbranch = StringVar()
            vyear = StringVar()
            vemail = StringVar()
            vaddress1 = StringVar()
            vaddress2 = StringVar()
            vcheck = IntVar()
            temp = 0
            if (temp == 0):
                vname.set(myData[0])
                vid.set(myData[1])
                vmobile.set(myData[2])
                vbranch.set(myData[3])
                vyear.set(myData[4])
                vemail.set(myData[5])
                vaddress1.set(myData[6])
                vaddress2.set(myData[7])
                vcheck.set(myData[8])
                temp = 1

            if __name__ == "__main__":
                Image_frame = tk.Frame(root1)
                sapretor = ttk.Separator(Image_frame)
                Image_frame.configure(background="#b5ffff")
                Image_frame.place(relx=0.0, rely=0.0, relheight=1, relwidth=1)

                def destroy_Image_frame():
                    Image_frame.destroy()
                    myData.clear()

                def btn_submit():
                    print('add')
                    mydb = mysql.connector.connect(host="localhost", user="root", passwd="", database="form_data")
                    mycursor = mydb.cursor()
                    global countSubmit, modifiedCount
                    with open(DataOutputPath, 'a+') as f:
                        for data in myData:
                            f.write((str(data) + ','))
                        f.write('\n')
                    sql = "INSERT INTO form_entry (Name,ID_No,Mobile,	Branch,AdmissionYear,Email,	AddressLine1,	AddressLine2,Declaration,EntDate) VALUES (%s, %s, %s,%s,%s,%s,%s,%s,%s,%s)"
                    now = datetime.now()
                    date = now.strftime('%Y-%m-%d')
                    val = (
                        myData[0], myData[1], myData[2], myData[3], myData[4], myData[5], myData[6], myData[7],
                        myData[8],
                        now)
                    mycursor.execute(sql, val)
                    mydb.commit()

                    statusbar['text'] = str(countSubmit) + " Entry Submitted Successfully"
                    countSubmit = countSubmit + 1

                # Image_label_list.append(Image1)

                def btn_Modify():
                    print('submit')
                    myData[0] = vname.get()
                    myData[1] = vid.get()
                    myData[2] = vmobile.get()
                    myData[3] = vbranch.get()
                    myData[4] = vyear.get()
                    myData[5] = vemail.get()
                    myData[6] = vaddress1.get()
                    myData[7] = vaddress2.get()
                    myData[8] = vcheck.get()
                    for x in myData:
                        print(x)
                    vname.set(myData[0])
                    vid.set(myData[1])
                    vmobile.set(myData[2])
                    vbranch.set(myData[3])
                    vyear.set(myData[4])
                    vemail.set(myData[5])
                    vaddress1.set(myData[6])
                    vaddress2.set(myData[7])
                    vcheck.set(myData[8])
                    statusbar['text'] = "Entry Modified Successfully"

                # create label
                # heading = Label(Image_frame)
                # lblname = Label(Image_frame)
                image1 = Label(Image_frame)
                LName = Label(Image_frame)
                ID = Label(Image_frame)
                Mobil = Label(Image_frame)
                Branch = Label(Image_frame)
                Admission_year = Label(Image_frame)
                Email = Label(Image_frame)
                Address1 = Label(Image_frame)

                name = Entry(Image_frame)
                id = Entry(Image_frame)
                mobile = Entry(Image_frame)
                branch = Entry(Image_frame)
                admssion_year = Entry(Image_frame)
                email = Entry(Image_frame)
                address1 = Entry(Image_frame)
                address2 = Entry(Image_frame)

                check = Checkbutton(Image_frame)
                submit = Button(Image_frame)
                Modify = Button(Image_frame)
                Back = Button(Image_frame)
                statusbar = Label(Image_frame)

                # heading.configure(text="OCR Based Text Extraction", background="#00ffff", foreground="#000000", font="-family {Arial} -size 20",justify='center')
                # lblname.configure(text="Processed Image", background="#b5ffff", font="-family {Arial} -size 15 -weight bold")
                LName.configure(text="Name", background="#b5ffff", font="-family {Arial} -size 10")
                ID.configure(text="ID No", background="#b5ffff", font="-family {Arial} -size 10")
                Mobil.configure(text="Moblie No", background="#b5ffff", font="-family {Arial} -size 10 ")
                Branch.configure(text="Branch", background="#b5ffff", font="-family {Arial} -size 10 ")
                Admission_year.configure(text="Admission Year", background="#b5ffff", font="-family {Arial} -size 10")
                Email.configure(text="Email", background="#b5ffff", font="-family {Arial} -size 10")
                Address1.configure(text="Address", background="#b5ffff", font="-family {Arial} -size 10")

                image1.configure(image=Image_data[0], background="#b5ffff", font="-family {Arial} -size 10")

                name.configure(font="-family {Arial} -size 10", textvariable=vname)
                id.configure(font="-family {Arial} -size 10", textvariable=vid)
                mobile.configure(font="-family {Arial} -size 10", textvariable=vmobile)
                branch.configure(font="-family {Arial} -size 10", textvariable=vbranch)
                admssion_year.configure(font="-family {Arial} -size 10", textvariable=vyear)
                email.configure(font="-family {Arial} -size 10", textvariable=vemail)
                address1.configure(font="-family {Arial} -size 10", textvariable=vaddress1)
                address2.configure(font="-family {Arial} -size 10", textvariable=vaddress2)


                check.configure(text="I hereby declare that all the \n information given above is true.",
                                variable=vcheck, onvalue=1, offvalue=0)
                submit.configure(text="Submit", activeforeground="#000000", activebackground="#ececec",
                                 font="-family {Arial} -size 15", pady="0", command=btn_submit)
                Modify.configure(text="Modify", activeforeground="#000000", activebackground="#ececec",
                                 font="-family {Arial} -size 15", pady="0", command=btn_Modify)
                Back.configure(text="Back", activeforeground="#000000", activebackground="#ececec",
                               font="-family {Arial} -size 15", pady="0", command=destroy_Image_frame)
                statusbar.configure(text="Processed Image", font="-family {Arial} -size 12", justify='center')

                # heading.place(relx=0.0, rely=0.059, height=48, width=655)
                # lblname.place(relx=0.137, rely=0.394, height=30, width=120)
                LName.place(relx=0.5, rely=0.02, height=12, width=32)
                ID.place(relx=0.5, rely=0.07, height=12, width=35)
                Mobil.place(relx=0.5, rely=0.12, height=12, width=63)
                Branch.place(relx=0.5, rely=0.17, height=12, width=42)
                Admission_year.place(relx=0.5, rely=0.22, height=12, width=92)
                Email.place(relx=0.5, rely=0.27, height=12, width=35)
                Address1.place(relx=0.5, rely=0.32, height=12, width=49)

                image1.place(relx=0.02, rely=0.02, height=450, width=300)
                name.place(relx=0.65, rely=0.02, height=13, relwidth=0.3)
                id.place(relx=0.65, rely=0.07, height=13, relwidth=0.3)
                mobile.place(relx=0.65, rely=0.12, height=13, relwidth=0.3)
                branch.place(relx=0.65, rely=0.17, height=13, relwidth=0.3)
                admssion_year.place(relx=0.65, rely=0.22, height=13, relwidth=0.3)
                email.place(relx=0.65, rely=0.27, height=13, relwidth=0.3)
                address1.place(relx=0.65, rely=0.32, height=13, relwidth=0.3)
                address2.place(relx=0.65, rely=0.37, height=13, relwidth=0.3)

                check.place(relx=0.58, rely=0.42, height=30, relwidth=0.45)
                submit.place(relx=0.79, rely=0.8, height=30, width=100)
                Modify.place(relx=0.79, rely=0.7, height=30, width=100)
                Back.place(relx=0.55, rely=0.8, height=30, width=100)
                # sapretor.place(relx=-0.015, rely=0.207, relwidth=1.037)
                statusbar.place(relx=0.0, rely=0.925, height=26, width=662)

        if __name__ == "__main__":
            frame2 = tk.Frame(root1)
            sapretor = ttk.Separator(frame2)
            frame2.configure(background="#b5ffff")
            frame2.place(relx=0.0, rely=0.0, relheight=1, relwidth=1)

            # create label
            heading = Label(frame2)
            loading = Label(frame2)
            statusbar = Label(frame2)

            heading.configure(text="OCR Based Text Extraction", font="-family {Arial} -size 20", justify='center')
            loading.configure(text="Loading....Please wait..", font="-family {Arial} -size 15 -weight bold")
            statusbar.configure(text="OCR Based Text Extraction", font="-family {Arial} -size 12", justify='center')

            progress = ttk.Progressbar(frame2, orient=HORIZONTAL, length="430", value=1, mode='determinate')
            heading.place(relx=0.0, rely=0.059, height=48, width=655)
            loading.place(relx=0.274, rely=0.394, height=30, width=300)
            progress.place(relx=0.168, rely=0.551, relwidth=0.654, relheight=0.0, height=22)
            sapretor.place(relx=-0.015, rely=0.207, relwidth=1.037)
            statusbar.place(relx=0.0, rely=0.925, height=26, width=662)

        global filepath
        if filepath:
            per = 25
            pixelThreshold = 50
            roi = [[(349, 598), (1054, 655), 'text', 'Name'],
                   [(198, 792), (532, 848), 'text', 'ID'],
                   [(604, 794), (1053, 849), 'text', 'Mobile'],
                   [(198, 983), (758, 1040), 'text', 'Branch'],
                   [(836, 985), (1036, 1041), 'text', 'Admission Year'],
                   [(430, 1114), (1072, 1169), 'text', 'Email'],
                   [(429, 1218), (1073, 1274), 'text', 'Address Line1'],
                   [(428, 1305), (1072, 1361), 'text', 'Address Line2'],
                   [(230, 1370), (273, 1412), 'box', 'Declaration']]

            pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

            imgQ = cv2.imread('Query0.jpg')
            h, w, c = imgQ.shape
            imgQ = cv2.resize(imgQ, (w, h))

            orb = cv2.ORB_create(1200)
            kp1, des1 = orb.detectAndCompute(imgQ, None)

            img = cv2.imread(str(filepath))
            # cv2.imshow(y, img)
            kp2, des2 = orb.detectAndCompute(img, None)
            bf = cv2.BFMatcher(cv2.NORM_HAMMING)
            matches = bf.match(des2, des1)
            matches = tuple(sorted(matches, key=lambda x: x.distance))
            good = matches[:int(len(matches) * (per / 100))]
            imgMatch = cv2.drawMatches(img, kp2, imgQ, kp1, good[:10], None, flags=2)
            # cv2.imshow(y, imgMatch)

            srcPoints = np.float32([kp2[m.queryIdx].pt for m in good]).reshape(-1, 1, 2)
            dstPoints = np.float32([kp1[m.trainIdx].pt for m in good]).reshape(-1, 1, 2)

            M, _ = cv2.findHomography(srcPoints, dstPoints, cv2.RANSAC, 5.0)
            imgScan = cv2.warpPerspective(img, M, (w, h))
            # cv2.imshow(y, imgScan)

            imgShow = imgScan.copy()
            imgMask = np.zeros_like(imgShow)

            global myData
            print('############## Extraction Data from Form  ############')
            for x, r in enumerate(roi):

                progress["value"] = x * 10
                progress.update()

                # str1=StringVar()
                cv2.rectangle(imgMask, (r[0][0], r[0][1]), (r[1][0], r[1][1]), (0, 255, 0), cv2.FILLED)
                imgShow = cv2.addWeighted(imgShow, 0.99, imgMask, 0.1, 0)

                imgCrop = imgScan[r[0][1]:r[1][1], r[0][0]:r[1][0]]
                # print(x)
                if r[2] == 'text':
                    print('{} :{}'.format(r[3], pytesseract.image_to_string(imgCrop)))
                    str1 = pytesseract.image_to_string(imgCrop)
                    str1 = str1.replace(",", "")
                    myData.append(str1.rstrip())

                if r[2] == 'box':
                    imgGray = cv2.cvtColor(imgCrop, cv2.COLOR_BGR2GRAY)
                    imgThresh = cv2.threshold(imgGray, 170, 255, cv2.THRESH_BINARY_INV)[1]
                    totalPixels = cv2.countNonZero(imgThresh)
                    # print(totalPixels)
                    if totalPixels > pixelThreshold:
                        totalPixels = 1
                    else:
                        totalPixels = 0
                    print(f'{r[3]} : {totalPixels}')
                    myData.append(totalPixels)
                cv2.putText(imgShow, str(myData[x]) + " ", (r[0][0], r[0][1]), cv2.FONT_HERSHEY_PLAIN, 2, (0, 0, 255),
                            2)

            imgShow = cv2.resize(imgShow, (w // 3, h // 3))

            im = cv2.cvtColor(imgShow, cv2.COLOR_BGR2RGB)
            im = Image.fromarray(im)
            resized_image = im.resize((420, 590))
            newImage = ImageTk.PhotoImage(resized_image)
            Image_data.clear()
            Image_data.append(newImage)
            frame2.destroy()
            Open_image_window()

            cv2.waitKey(0)


    def Capture_Image():
        global filepath

        key = cv2.waitKey(1)
        webcam = cv2.VideoCapture("http://192.168.43.1:4747/video")
        # webcam = cv2.VideoCapture(0)
        img_counter = 0
        sleep(2)

        while True:

            try:
                check, frame = webcam.read()
                print(check)  # prints true as long as the webcam is running
                print(frame)  # prints matrix values of each framecd
                cv2.imshow("Capturing", frame)
                key = cv2.waitKey(1)
                if key == ord('s'):
                    cv2.imwrite(os.path.join(
                        "H:\\College\\Sem-6\\3IT31\\formTesting\\UseForms\\" + "opencv_frame_{}.png".format(
                            img_counter)), img=frame)
                    webcam.release()
                    print("Image saved!")
                    filepath = "H:\\College\\Sem-6\\3IT31\\formTesting\\UseForms\\" + "opencv_frame_{}.png".format(
                        img_counter)
                    img_counter = 1
                    cv2.destroyAllWindows()
                    break

                elif key == ord('q'):
                    webcam.release()
                    cv2.destroyAllWindows()
                    break

            except(KeyboardInterrupt):
                print("Turning off camera.")
                webcam.release()
                print("Camera off.")
                print("Program ended.")
                cv2.destroyAllWindows()
                break
        if(img_counter==1):
                Second_page()
        # Filepath_label = Label(win, text="The File is located at : " + str(filepath), font=('Aerial 11')).pack()
        # # win.withdraw()
        # First_Next_button.pack(anchor=CENTER, padx=100, ipadx=2, ipady=2)


    def Browse_Image():
        """ Open form  image and read input field form  image """
        global filepath
        global filesize
        file = filedialog.askopenfile(mode='r', filetypes=[('Images', '*.png'), ('Images', '*.jpg'), ('Images', '*.jpeg')])
        if file:
            filepath = os.path.abspath(file.name)
            filesize = os.path.getsize(filepath)
            # print('Size of file is', filesize)
            # print(type(filesize))

            if filesize > FileSizeLimit:
                # print(type(filesize))
                statusbar['text'] = "Your File size is " + str(round(filesize / 1000000, 3)) + " MB"
                messagebox.showinfo("Invalid Size", "Please Select File With Less Than 5MB")
            # win.withdraw()
            else:
                Second_page()


    def Download_excel():
        files = [('ExcelSheet', '*.csv'), ('ExcelSheet 2013', '*.xls')]
        file1 = asksaveasfile(filetypes=files, defaultextension=".xlsx", initialfile='DataOutput',
                              initialdir='Downloads/', title="OCR")
        shutil.copy(DataOutputPath, file1.name)
        messagebox.showinfo("Download", "Download Successful")



    heading = Label(frame)
    btnbrw = Button(frame)
    btncap = Button(frame)
    btnexl = Button(frame)
    btnexit = Button(frame)
    statusbar = Label(frame)

    heading.configure(text="Form Details Reader", font="-family {Arial} -size 20", justify='center')
    btnbrw.configure(text="Browse Image", activeforeground="#000000", activebackground="#ececec",
                      font="-family {Arial} -size 15", pady="0", command=Browse_Image)
    btncap.configure(text="Capture Image", activeforeground="#000000", activebackground="#ececec",
                      font="-family {Arial} -size 15", pady="0", command=Capture_Image)
    btnexl.configure(text="Database", activeforeground="#000000", activebackground="#ececec",
                      font="-family {Arial} -size 15", pady="0", command=Data_Show)
    btnexit.configure(text="Exit", activeforeground="#000000", activebackground="#ececec",
                      font="-family {Arial} -size 15", pady="0", command=destroy_frame)
    statusbar.configure(text="Form Details Reader", font="-family {Arial} -size 12", justify='center')

    heading.place(relx=0.0, rely=0.059, height=48, width=655)
    btnbrw.place(relx=0.29, rely=0.276, height=50, width=300)
    btncap.place(relx=0.29, rely=0.433, height=50, width=300)
    btnexl.place(relx=0.29, rely=0.591, height=50, width=300)
    btnexit.place(relx=0.29, rely=0.748, height=50, width=300)
    statusbar.place(relx=0.0, rely=0.925, height=26, width=662)
    sapretor.place(relx=-0.015, rely=0.207, relwidth=1.037)
    root1.mainloop()