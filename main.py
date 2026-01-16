import random
import tkinter as tk
from tkinter import ttk, scrolledtext
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
from PIL import Image, ImageTk
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from datetime import datetime
import os
from docx2pdf import convert
import sys

def getDepartment():
    selectedDepartment = []
    if departmentValue[0].get() == 1:
        selectedDepartment.append("BBA")
    if departmentValue[1].get() == 1:
        selectedDepartment.append("BCA")
    if departmentValue[2].get() == 1:
        selectedDepartment.append("BCOM")
    if departmentValue[3].get() == 1:
        selectedDepartment.append("MBA")
    if departmentValue[4].get() == 1:
        selectedDepartment.append("MCA")
    if departmentValue[5].get() == 1:
        selectedDepartment.append("MCOM")
    return ",".join(selectedDepartment)

def getEventCategory():
    if eventCategory.get() == "Others":
        return eventCategoryText.get()
    else:
        return eventCategory.get()

def validateDate(event):
    date1 = datetime.strptime(startDateEntry.get(), "%d/%m/%Y")
    date2 = datetime.strptime(endDateEntry.get(),  "%d/%m/%Y")

    # Compare the dates and show the result
    if date1 > date2:
        messagebox.showinfo(title="Error:" ,message=f"Date {startDateEntry.get()} comes after {endDateEntry.get()}")
        endDateEntry.set_date(date1)

def getDate():
    if startDateEntry.get() != endDateEntry.get():
        return f"{startDateEntry.get()} - {endDateEntry.get()}"
    else:
        return startDateEntry.get()
def generatePDF():
    data = {
        "Title": eventTitle.get(),
        "Department": getDepartment(),
        "Event_Category": getEventCategory(),
        "Organising_Committee":organisingCommittee.get(),
        "Date": getDate(),
        "Time": f"{hourTime.get()} : {minuteTime.get()}  {amPmTime.get()}",
        "Venue": venue.get(),
        "Resource_person": f"{resourcePersonTitle.get()}  {resourcePersonName.get()}",
        "Target_Audience": targetAudience.get(),
        "No_of_Participants": noOfParticipants.get(),
        "Report_Submitted_By": f"{reportSubmittedByTitle.get()}  {reportSubmittedByName.get()}",
        "Convener_Of_Program": f"{convenerOfTheProgramTitle.get()}  {convenerOfTheProgramName.get()}",
        "Description": description.get("1.0", tk.END).strip(),
        "Program_Outcome": programOutcome.get("1.0", tk.END).strip()
    }
    global doc_file

    try:
        # exe_dir = os.path.dirname(sys.executable)
        # template_path = os.path.join(exe_dir, "2023 IQAC formats for Documentation.docx")
        template_path = "Template.docx"
        doc = DocxTemplate(template_path)

        # Add images to the data dictionary
        for i, photo_path in enumerate(imageFilePaths, start=1):
            if photo_path:
                data[f"image{i}"] = InlineImage(doc, photo_path, height=Mm(57))

        # Render the document
        doc.render(data)
        doc_file = f"temp{random.randint(1,1000)}.docx"
        doc.save(doc_file)

        # Ask user for file name and location
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if file_path:
            # Convert docx to pdf
            convert(doc_file, file_path)
            messagebox.showinfo("Submission Successful",
                                f"Event details submitted successfully and PDF generated: {file_path}")
        else:
            messagebox.showinfo("Submission Successful", "Event details submitted successfully!")

    except Exception as e:
        messagebox.showerror("Submission Failed", f"An error occurred: {e}")
        return  # Return to avoid accessing doc_file in finally block if an error occurs

    finally:
        # Delete the temporary docx
        if doc_file and os.path.exists(doc_file):
            os.remove(doc_file)

def resetForm():
        eventTitle.delete(0, tk.END)
        for i  in range(0,6):
            department[i].deselect()
        organisingCommittee.delete(0, tk.END)
        eventCategory.set('')
        eventCategoryText.delete(0, tk.END)
        startDateEntry.delete(0, tk.END)
        endDateEntry.delete(0, tk.END)
        hourTime.set('')
        minuteTime.set('')
        amPmTime.set('')
        venue.delete(0, tk.END)
        resourcePersonTitle.set('')
        resourcePersonName.delete(0, tk.END)
        targetAudience.delete(0, tk.END)
        noOfParticipants.delete(0, tk.END)
        reportSubmittedByTitle.set('')
        reportSubmittedByName.delete(0, tk.END)
        convenerOfTheProgramTitle.set('')
        convenerOfTheProgramName.delete(0, tk.END)
        description.delete("1.0", tk.END)
        programOutcome.delete("1.0", tk.END)

        for i,image in enumerate(imageLabels):
            if image:
                image.grid_forget()
            imageLabels[i]= ""
        imageFilePaths=["","","",""]


window = tk.Tk()
window.config(bg="white")

setString = "SESHADRIPURAM EDUCATIONAL TRUST"
sctString = "Seshadripuram College Tumakuru"
TUString = "Permanently Affiliated to Tumkur University"
addressString = '"Vikasa Bharathi”, #3, Kalpatharu Badavane , Gangasandra Main Road, Tumakuru – 572105'
isoString = "ISO 9001:2015 Certified Institution | Recognised by Government of Karnataka | College Code -1086"
sdclogoImage = tk.PhotoImage(file="sdclogo.png")

headingFontFamily = "Times New Roman"
headingFontSize = 11
headingFontStyle = ""
headingFrame = tk.Frame(window, bg="white")
tk.Label(headingFrame, text=setString, font=(headingFontFamily, headingFontSize, headingFontStyle), bg="white").grid(row=1, column=1)
tk.Label(headingFrame, text=sctString, font=(headingFontFamily, headingFontSize + 15, headingFontStyle), bg="white").grid(row=2, column=1)
tk.Label(headingFrame, text=TUString, font=(headingFontFamily, headingFontSize, headingFontStyle), bg="white").grid(row=3, column=1)
tk.Label(headingFrame, text=addressString, font=(headingFontFamily, headingFontSize, headingFontStyle), bg="white").grid(row=4, column=1)
tk.Label(headingFrame, text=isoString, font=(headingFontFamily, headingFontSize, headingFontStyle), bg="white").grid(row=5, column=1)
tk.Label(headingFrame, image=sdclogoImage, bg="white").grid(rowspan=6, row=0, column=2)
headingFrame.pack()




formEventFont = ("poppins", 16)

canvas_frame = tk.Frame(window)
canvas_frame.pack(fill=tk.BOTH, expand=True)

canvas = tk.Canvas(canvas_frame, bg="white")
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = tk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview, width=20)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

canvas.configure(yscrollcommand=scrollbar.set)

formFrame = tk.Frame(canvas, bg="white")

canvas.create_window((0, 150), window=formFrame, anchor="nw")

def on_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

formFrame.bind("<Configure>", on_configure)

def on_mouse_wheel(event):
    if event.delta:  # Windows
        canvas.yview_scroll(-1 * int(event.delta / 120), "units")
    else:  # macOS and Linux
        if event.num == 4:
            canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            canvas.yview_scroll(1, "units")

# Bind mouse wheel scroll to the canvas
canvas.bind_all("<MouseWheel>", on_mouse_wheel)  # Windows


# Form fields go here
tk.Label(formFrame, text="Event Title:", font=formEventFont, bg="white").grid(row=0, column=0, padx=20, pady=20, sticky="w")
eventTitle = tk.Entry(formFrame, font=formEventFont, width=40, relief="groove")
eventTitle.grid(row=0, column=1, padx=20, pady=20, columnspan=6, sticky="w")

tk.Label(formFrame, text="Department:", font=formEventFont, bg="white").grid(row=1, column=0, padx=20, pady=20, sticky="w")
departmentValue = [tk.IntVar(window) for _ in range(6)]
department = []
checkButtonFont = ("poppins", 13)
departmentsName = ["BBA", "BCA", "BCOM", "MBA", "MCA", "MCOM"]
for i, dep in enumerate(departmentsName):
    department.append(tk.Checkbutton(formFrame, variable=departmentValue[i], text=dep, font=checkButtonFont, bg="white"))
    department[-1].grid(row=1, column=1 + i)


tk.Label(formFrame, text="Event Category:", font=formEventFont, bg="white").grid(row=2, column=0, padx=20, pady=20, sticky="w")
eventCategoryOptions = ["FDP", "SDP", "Seminar", "Workshop", "Others"]
eventCategory = ttk.Combobox(formFrame, values=eventCategoryOptions, state="readonly", font=checkButtonFont, width=10)
eventCategory.grid(row=2, column=1, padx=10, pady=10, columnspan=4, sticky="w")

eventCategoryText = tk.Entry(formFrame, font=formEventFont, width=20, relief="groove")
eventCategoryText.grid(row=2, column=2, padx=20, pady=20, sticky="w", columnspan=3)
eventCategoryText.config(state=tk.DISABLED)

def isOthers(event, eventCategory):
    if eventCategory.get() != "Others":
        eventCategoryText.config(state=tk.DISABLED)
    else:
        eventCategoryText.config(state=tk.NORMAL)

eventCategory.bind("<<ComboboxSelected>>", lambda event: isOthers(event, eventCategory))

tk.Label(formFrame, text="Organising Committee:", font=formEventFont, bg="white").grid(row=3, column=0, padx=20, pady=20, sticky="w")
organisingCommittee = tk.Entry(formFrame, font=formEventFont, width=40, relief="groove")
organisingCommittee.grid(row=3, column=1, padx=20, pady=20, columnspan=6, sticky="w")

tk.Label(formFrame, text="Start Date:", font=formEventFont, bg="white").grid(row=4, column=0, padx=20, pady=20, sticky="w")
startDateEntry = DateEntry(formFrame, width=12, borderwidth=2, font=checkButtonFont, date_pattern='dd/mm/yyyy', state="readonly")
startDateEntry.grid(row=4, column=1, padx=20, pady=20, columnspan=6, sticky="w")
startDateEntry.delete(0, tk.END)

tk.Label(formFrame, text="End Date:", font=formEventFont, bg="white").grid(row=5, column=0, padx=20, pady=20, sticky="w")
endDateEntry = DateEntry(formFrame, width=12, borderwidth=2, font=checkButtonFont, date_pattern='dd/mm/yyyy', state="readonly")
endDateEntry.bind("<FocusOut>", validateDate)
endDateEntry.grid(row=5, column=1, padx=20, pady=20, columnspan=6, sticky="w")
endDateEntry.delete(0, tk.END)

tk.Label(formFrame, text="Time:", font=formEventFont, bg="white").grid(row=6, column=0, padx=20, pady=20, sticky="w")
timeFrame = tk.Frame(formFrame)
hourList = [str(i).zfill(2) for i in range(1, 13)]
minuteList = [str(i).zfill(2) for i in range(0, 60)]
hourTime = ttk.Combobox(timeFrame, values=hourList, state="readonly", font=checkButtonFont, width=4)
minuteTime = ttk.Combobox(timeFrame, values=minuteList, state="readonly", font=checkButtonFont, width=4)
amPmTime = ttk.Combobox(timeFrame, values=["AM", "PM"], state="readonly", font=checkButtonFont, width=4)

hourTime.grid(row=0, column=0, padx=2, pady=2)
tk.Label(timeFrame, text=":", font=formEventFont, bg="white").grid(row=0, column=1, padx=2, pady=2)
minuteTime.grid(row=0, column=3, padx=2, pady=2)
amPmTime.grid(row=0, column=4, padx=1, pady=1)

timeFrame.grid(row=6, column=1, padx=20, pady=20, columnspan=6, sticky="w")

tk.Label(formFrame, text="Venue:", font=formEventFont, bg="white").grid(row=7, column=0, padx=20, pady=20, sticky="w")
venue = tk.Entry(formFrame, font=checkButtonFont, width=40, relief="groove")
venue.grid(row=7, column=1, padx=20, pady=20, columnspan=6, sticky="w")

nameTitles = ["Dr.", "Mr.", "Mrs.", "Miss.", "Prof."]
tk.Label(formFrame, text="Resource Person:", font=formEventFont, bg="white").grid(row=8, column=0, padx=20, pady=20, sticky="w")
resourcePersonTitle = ttk.Combobox(formFrame, values=nameTitles, state="readonly", font=checkButtonFont, width=7)
resourcePersonTitle.grid(row=8, column=1, padx=10, pady=10, sticky="w")
resourcePersonName = tk.Entry(formFrame, font=checkButtonFont, width=40, relief="groove")
resourcePersonName.grid(row=8, column=2, padx=20, pady=20, sticky="w", columnspan=4)

tk.Label(formFrame, text="Target Audience:", font=formEventFont, bg="white").grid(row=9, column=0, padx=20, pady=20, sticky="w")
targetAudience = tk.Entry(formFrame, font=checkButtonFont, width=40, relief="groove")
targetAudience.grid(row=9, column=1, padx=20, pady=20, columnspan=6, sticky="w")

tk.Label(formFrame, text="Ex: BCA 1 C,  BCOM 1,  MCA,  Lecturers,  Others", font=formEventFont, bg="white").grid(row=10, column=1, columnspan=6, sticky="w")

tk.Label(formFrame, text="No of Participants:", font=formEventFont, bg="white").grid(row=11, column=0, padx=20, pady=20, sticky="w")
noOfParticipants = tk.Spinbox(formFrame, from_=1, to=99999999, font=checkButtonFont, width=10)
noOfParticipants.grid(row=11, column=1, padx=20, pady=20, columnspan=6, sticky="w")
noOfParticipants.delete(0, tk.END)

tk.Label(formFrame, text="Report Submitted By:", font=formEventFont, bg="white").grid(row=12, column=0, padx=20, pady=20, sticky="w")
reportSubmittedByTitle = ttk.Combobox(formFrame, values=nameTitles, state="readonly", font=checkButtonFont, width=7)
reportSubmittedByTitle.grid(row=12, column=1, padx=10, pady=10,sticky="w")
reportSubmittedByName = tk.Entry(formFrame, font=checkButtonFont, width=40, relief="groove")
reportSubmittedByName.grid(row=12, column=2, padx=20, pady=20, sticky="w", columnspan=4)

tk.Label(formFrame, text="Convener Of The Program:", font=formEventFont, bg="white").grid(row=13, column=0, padx=20, pady=20, sticky="w")
convenerOfTheProgramTitle = ttk.Combobox(formFrame, values=nameTitles, state="readonly", font=checkButtonFont, width=7)
convenerOfTheProgramTitle.grid(row=13, column=1, padx=10, pady=10,sticky="w")
convenerOfTheProgramName = tk.Entry(formFrame, font=checkButtonFont, width=40, relief="groove")
convenerOfTheProgramName.grid(row=13, column=2, padx=20, pady=20, sticky="w", columnspan=4)

tk.Label(formFrame, text="Images:", font=formEventFont, bg="white").grid(row=14, column=0, padx=20, pady=20, sticky="w")

imageFilePaths = ["","","",""]
imageLabels = ["","","",""]
def selectImagePath(n, row, column):
    imageFilePaths[n] = filedialog.askopenfilename()
    image_path = imageFilePaths[n]
    image = Image.open(image_path)
    resized_image = image.resize((300, 200), Image.Resampling.LANCZOS)
    image1 = ImageTk.PhotoImage(resized_image)
    imageLabel = tk.Label(formFrame, image=image1, bg="white")
    imageLabel.image = image1
    imageLabel.grid(row=row, column=column, padx=20, pady=20, columnspan=3)
    imageLabels[n] = imageLabel


image1Button = tk.Button(formFrame, text="upload Image 1", command= lambda: selectImagePath(0,15,1))
image1Button.grid(row=14, column=1, padx=20, pady=20, columnspan=3)

image2Button = tk.Button(formFrame, text="upload Image 2", command= lambda: selectImagePath(1, 15, 5))
image2Button.grid(row=14, column=5, padx=20, pady=20, columnspan=3)

image3Button = tk.Button(formFrame, text="upload Image 3", command= lambda: selectImagePath(2,17,1))
image3Button.grid(row=16, column=1, padx=20, pady=20, columnspan=3)

image4Button = tk.Button(formFrame, text="upload Image 4", command= lambda: selectImagePath(3,17,5))
image4Button.grid(row=16, column=5, padx=20, pady=20, columnspan=3)

tk.Label(formFrame, text="Description:", font=formEventFont, bg="white").grid(row=18, column=0, padx=20, pady=20, sticky="w")
description = scrolledtext.ScrolledText(formFrame, height=5, width=60, font=checkButtonFont)
description.grid(row=18, column=1, padx=20, pady=20, columnspan=6, sticky="w")


tk.Label(formFrame, text="Program Outcome:", font=formEventFont, bg="white").grid(row=19, column=0, padx=20, pady=20, sticky="w")
programOutcome = scrolledtext.ScrolledText(formFrame, height=5, width=60, font=checkButtonFont)
programOutcome.grid(row=19, column=1, padx=20, pady=20, columnspan=6, sticky="w")

submitButton = tk.Button(formFrame, text="Submit",width=10, height=1, font=formEventFont, command= lambda: generatePDF())
submitButton.grid(row=20, column=0, padx=20, pady=20, columnspan=3)

resetButton = tk.Button(formFrame, text="Clear",width=10, height=1, font=formEventFont,  command= lambda: resetForm())
resetButton.grid(row=20, column=2, padx=20, pady=20, columnspan=3)

window.mainloop()



