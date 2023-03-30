import docx2txt
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import os

# Function for opening the
# file explorer window
def browseFiles():
    global job_des_path
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    filename = filedialog.askopenfilename(initialdir=desktop,
                                          title="Select a File",
                                          filetypes=(("Docx files",
                                                      "*.docx*"),
                                                     ("All files",
                                                      "*.*")))

    job_des_path = filename
    job_des_label.configure(text="File Selected: " + filename)

def submit():
    global resume_path, job_des_path, window

    if resume_path != "" and job_des_path != "":
        job_description = docx2txt.process(job_des_path)
        resume = docx2txt.process(resume_path)
        content = [job_description, resume]
        cv = CountVectorizer()
        count_matrix = cv.fit_transform(content)
        mat = cosine_similarity(count_matrix)
        print('Resume Matches by: ' + str(mat[1][0] * 100) + '%')
        messagebox.showinfo("Result",'Your Resume matches ' +str(int(mat[1][0] * 100)) + '%'+' with Job Description')

def browseResume():
    global resume_path
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    filename = filedialog.askopenfilename(initialdir=desktop,
                                          title="Select a File",
                                          filetypes=(("Docx files",
                                                      "*.docx*"),
                                                     ("All files",
                                                      "*.*")))

    # Change label contents
    resume_path=filename
    resume_label.configure(text="File Selected: " + filename)

# Create the root window
window = Tk()

# Set window title
window.title('Scrutiny of Resume')

# Set window size
window.geometry("720x480")

window.config(background="#F2B33D")
frame = Frame(window, bg='#F2B33D')
# Create a File Explorer label
t = Label(window,text="Scrutiny of Resume",
                            fg="white",bg="black",font=('Times', 24))

t.pack(fill=X);

job_des_label = Label(frame, text="No file chosen",fg="black",bg='#F2B33D')

job_des_title = Label(frame, text="Job Description",fg="black",bg='#F2B33D',font=('Helvetica', 10, 'bold'))


resume_label = Label(frame, text="No file chosen",fg="black",bg='#F2B33D')
resume_title = Label(frame, text="Resume",fg="black",bg='#F2B33D',font=('Helvetica', 10, 'bold'))

button_explore = Button(frame,
                        text="Choose File",
                        command=browseFiles,width=15)

button_exit = Button(frame,
                     text="Exit",
                     command=exit)


button_resume = Button(frame, text="Choose File", command=browseResume,width=15)
submit = Button(window, text="Analyze",font=('Helvetica', 10, 'bold'), command=submit,height=2,width=10,bg='silver')
# Grid method is chosen for placing
# the widgets at respective positions
# in a table like structure by
# specifying rows and columns
job_des_title.grid(column=1,row=1,sticky='w')
button_explore.grid(column=1, row=2,padx=20,pady=5)
job_des_label.grid(column=2,row=2,sticky='w')
resume_title.grid(column=1,row=3,sticky='w')
button_resume.grid(column=1, row=4)
resume_label.grid(column=2,row=4,sticky='w')
frame.pack(expand=True);
submit.pack(fill=X);

window.mainloop()
job_des_path = ""
resume_path = ""


