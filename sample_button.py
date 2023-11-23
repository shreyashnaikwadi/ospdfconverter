import tkinter as tk

root = tk.Tk()
root.geometry('800x600')
root.resizable(width=False, height=False)

colour1 = '#020f12'
colour2 = '#05d7ff'
colour3 = '#65e7ff'
colour4 = 'BLACK'

main_frame = tk.Frame(root, bg=colour1, pady=40)
main_frame.pack(fill=tk.BOTH, expand=True)
main_frame.columnconfigure(0, weight=1)
main_frame.rowconfigure(0, weight=1)
main_frame.rowconfigure(1, weight=1)

button1 = tk.Button(
    main_frame,
    background=colour2,
    foreground=colour4,
    activebackground=colour3,
    activeforeground=colour4,
    highlightthickness=2,
    highlightbackground=colour2,
    highlightcolor='WHITE',
    width=20,
    height=2,
    border=0,
    cursor='hand1',
    text='Convert PDF to Word',
    font=('Arial', 16, 'bold')
)

button1.grid(column=0, row=0)

button2 = tk.Button(
    main_frame,
    background=colour1, 
    foreground=colour2,
    activebackground=colour3,
    activeforeground=colour4,
    highlightthickness=2,
    highlightbackground=colour2,
    highlightcolor='WHITE',
    width=20,
    height=2,
    border=0,
    cursor='hand1',
    text='Convert Word to PDF',
    font=('Arial', 16, 'bold')
)

button2.grid(column=0, row=1)


root.mainloop()
