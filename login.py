from customtkinter import *
from PIL import Image
from CTkMessagebox import CTkMessagebox
import sys
import os 

def validate_login():
    if email_entry.get() == "admin@gmail.com" and password_entry.get() == "1234":
        CTkMessagebox(title="Login Successful", message="Login Successful!, Redirecting you....")
        app.after(4000, redirect_to_main)
    else:
        CTkMessagebox(title="Error", message="Login Failed!!!", icon="cancel")

def redirect_to_main():
    app.destroy()  # Close the current window
    import main
    main.main()    # Run the main module
    sys.exit()

app = CTk()
app.title("OS PDF CONVERTER")
app.geometry("600x480")
app.resizable(0, 0)

import ctypes

myappid = 'mycompany.myproduct.subproduct.version'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

app.iconbitmap(default='assets/logo.ico')

script_dir = os.path.dirname(os.path.realpath(__file__))

side_image_path = os.path.join(script_dir, 'assets', 'sideimg.png')
email_icon_image_path = os.path.join(script_dir, 'assets', 'emailicon.png')
password_icon_image_path = os.path.join(script_dir, 'assets', 'passwordicon.png')


side_img_data = Image.open(side_image_path)
email_icon_data = Image.open(email_icon_image_path)     
password_icon_data = Image.open(password_icon_image_path)

side_img = CTkImage(dark_image=side_img_data, light_image=side_img_data, size=(300, 480))
email_icon = CTkImage(dark_image=email_icon_data, light_image=email_icon_data, size=(20, 20))
password_icon = CTkImage(dark_image=password_icon_data, light_image=password_icon_data, size=(17, 17))

CTkLabel(master=app, text="", image=side_img).pack(expand=True, side="left")

frame = CTkFrame(master=app, width=300, height=480, fg_color="#ffffff")
frame.pack_propagate(0)
frame.pack(expand=True, side="right")

CTkLabel(master=frame, text="Welcome Back!", text_color="#601E88", anchor="w", justify="left", font=("Arial Bold", 24)).pack(anchor="w", pady=(50, 5), padx=(25, 0))

CTkLabel(master=frame, text="Sign in to your account", text_color="#7E7E7E", anchor="w", justify="left", font=("Arial Bold", 12)).pack(anchor="w", padx=(25, 0))

# Login

username = StringVar()
password = StringVar()

CTkLabel(master=frame, text="  Email:", text_color="#601E88", anchor="w", justify="left", font=("Arial Bold", 14), image=email_icon, compound="left").pack(anchor="w", pady=(38, 0), padx=(25, 0))

email_entry = CTkEntry(master=frame, width=225, fg_color="#EEEEEE", border_color="#601E88", border_width=1, text_color="#000000")

email_entry.pack(anchor="w", padx=(25, 0))

CTkLabel(master=frame, text="  Password:", text_color="#601E88", anchor="w", justify="left", font=("Arial Bold", 14), image=password_icon, compound="left").pack(anchor="w", pady=(21, 0), padx=(25, 0))


# Password

password_entry = CTkEntry(master=frame, width=225, fg_color="#EEEEEE", border_color="#601E88", border_width=1, text_color="#000000", show="*")
password_entry.pack(anchor="w", padx=(25, 0))

# Login Button

CTkButton(master=frame, text="Login", fg_color="#601E88", hover_color="#E44982", font=("Arial Bold", 12),text_color="#ffffff", width=225, command=validate_login).pack(anchor="w", pady=(40, 0), padx=(25, 0))

app.mainloop()

import main