#Importing the necessary libraries
import tkinter as tk
from PIL import Image, ImageTk
import webbrowser
import customtkinter as ctk
from itertools import count, cycle
import openpyxl
import xlwt

#Importing the necessary libraries
import tkinter as tk
from PIL import Image, ImageTk
import webbrowser
import customtkinter as ctk
from itertools import count, cycle
import openpyxl
import xlwt

#Creating lists for the values of parameters to be displayed in Excel
power_results = list()
voltage_results = list()
current_results = list()
eon_results =list()
eoff_results = list()
efs_results = list()
out_curr_res = list()

#Class to display the gif image at the beginning
class ImageLabel(tk.Label):
    """
    A Label that displays images, and plays them if they are gifs
    :im: A PIL Image instance or a string filename

    """
    def load(self, im):
        if isinstance(im, str):
            im = Image.open(im)
        frames = []
        try:
            for i in count(1):
                frames.append(ImageTk.PhotoImage(im.copy()))
                im.seek(i)
        except EOFError:
            pass
        self.frames = cycle(frames)
        try:
            self.delay = im.info['duration']
        except:
            self.delay = 100
        if len(frames) == 1:
            self.config(image=next(self.frames))
        else:
            self.next_frame()
    def unload(self):
        self.config(image=None)
        self.frames = None
    def next_frame(self):
        if self.frames:
            self.config(image=next(self.frames))
            self.after(self.delay, self.next_frame)

#demo of the interface:
root1 = tk.Tk()
lbl = ImageLabel(root1)
lbl.pack()
lbl.load('welcome3.gif')
root1.after(5000,lambda:root1.destroy()) #First argument controlling the # of milliseconds to keep the gif played
root1.mainloop()

#To calculate the power when "Enter" key is pressed and triggered
def on_enter_key(event):
    button_calculate.invoke()

#Function to calculate the power of all the components
def calculate_power():
    component = var_component.get()
    try:
        voltage = float(entry_voltage.get())
        current = float(entry_current.get())
        global user_name
        user_name = str(entry_user_name.get())
        power = voltage * current
        power_c = (1/2)*current*(voltage)**2
        if component == "Resistor":
            voltage_results.append(voltage)
            current_results.append(current)
            label_result.configure(text=f"Power of Resistor: {power:.2f} Watts")
            power_results.append(f"{power:.2f}")
        elif component == "BJT":
            voltage_results.append(voltage)
            current_results.append(current)
            label_result.configure(text=f"BJT Power Calculation: {power:.2f} Watts")
            power_results.append(f"{power:.2f}")
        elif component == "MOSFET":
            eon = float(entry_eon.get())
            eoff = float(entry_eoff.get())
            efs = float(entry_fs.get())
            voltage_results.append(voltage)
            current_results.append(current)
            eon_results.append(eon)
            eoff_results.append(eoff)
            efs_results.append(efs)
            conduct_loss = float(entry_voltage.get()) * float(entry_current.get())
            switch_loss = (float(entry_eon.get())*0.001 + float(entry_eoff.get())*0.001) * float(entry_fs.get())
            power_mosfet = conduct_loss + switch_loss
            label_result.configure(text=f"MOSFET Power Calculation: {power_mosfet:.2f} Watts")
            power_results.append(f"{power_mosfet:.2f}")
        elif component == "IGBT":
            eon = float(entry_eon.get())
            eoff = float(entry_eoff.get())
            efs = float(entry_fs.get())
            voltage_results.append(voltage)
            current_results.append(current)
            eon_results.append(eon)
            eoff_results.append(eoff)
            efs_results.append(efs)
            conduct_loss = float(entry_voltage.get()) * float(entry_current.get())
            switch_loss = (float(entry_eon.get())*0.001 + float(entry_eoff.get())*0.001) * float(entry_fs.get())
            power_igbt = conduct_loss + switch_loss
            label_result.configure(text=f"IGBT Power Calculation: {power_igbt:.2f} Watts")
            power_results.append(f"{power_igbt:.2f}")
        elif component == "LDO":
            out_curr = float(entry_output_curr_ldo.get())
            voltage_results.append(voltage)
            current_results.append(current)
            out_curr_res.append(out_curr)
            power_ldo = (voltage - current)*out_curr
            label_result.configure(text=f"LDO Power Calculation: {power_ldo:.2f} Watts")
            power_results.append(f"{power_ldo:.2f}")
        elif component == "Capacitor":
            label_result.configure(text=f"Capacitor Energy Calculation: {power_c:.2f} Joule")
            voltage_results.append(voltage)
            current_results.append(current)
            power_results.append(f"{power_c:.2f} Joule")
        elif component == "Varistor":
            power_var = (voltage**2)/current
            label_result.configure(text=f"Varistor Power Calculation: {power_var:.2f} Watts")
            power_results.append(f"{power_var:.2f}")
            voltage_results.append(voltage)
            current_results.append(current)
        elif component == "Diode":
            label_result.configure(text=f"Power of Diode: {power:.2f} Watts")
            voltage_results.append(voltage)
            current_results.append(current)
            power_results.append(f"{power:.2f}")
    except ValueError:
        label_result.configure(text="Invalid input. Please enter numbers only.")

#The function to bind the link to the powerpoint presentations
def open_link(event):
    comp = var_component2.get()
    if comp == "More on Resistor":
        url = "https://arcelik-my.sharepoint.com/:p:/p/26041265/EQCvirgR_PNAgttpcK-2zcsB5nkguBrMkIpL4ebdJyH2gQ?wdOrigin=TEAMS-WEB.p2p.bim&wdExp=TEAMS-CONTROL&wdhostclicktime=1690974618401&web=1"
    elif comp == "More on MOSFET":
        url = "https://arcelik-my.sharepoint.com/:p:/p/26041265/EWjsr-W9uOxFpmh1Uj0ZCT8BsLrYE45VKumnm1c-3ZgTbg?wdOrigin=TEAMS-WEB.p2p.bim&wdExp=TEAMS-CONTROL&wdhostclicktime=1690974661651&web=1"
    elif comp == "More on IGBT":
        url = "https://arcelik-my.sharepoint.com/:p:/p/26041265/EYxn9Xddx99Cn5HetwRs0wABsmqYpFfdc2kcRBf8tpepQA?wdOrigin=TEAMS-WEB.p2p.bim&wdExp=TEAMS-CONTROL&wdhostclicktime=1690974515140&web=1"
    webbrowser.open(url)

#The function to act accordingly when the radiobutton is changed to a diff component
def component_change():
    selected_component = var_component.get()
    if selected_component == "Resistor":
        label_voltage.configure(text="V_R (Volts)")
        label_current.configure(text="I_R (Amperes)")
        #Delete the result and the entries potentially displayed on the window
        label_result.configure(text='')
        entry_voltage.delete(0, tk.END)
        entry_current.delete(0, tk.END)
        entry_eon.delete(0, tk.END)
        entry_eoff.delete(0, tk.END)
        entry_fs.delete(0, tk.END)
        entry_eon.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        entry_eoff.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0")
        entry_fs.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0")
        entry_output_curr_ldo.delete(0, tk.END)
        entry_output_curr_ldo.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0")
        label_output_curr_ldo.configure(text="")
        label_eon.configure(text="")
        label_eoff.configure(text="")
        label_fs.configure(text="")

        #Change the position of the button and result accordingly
        button_calculate.place(x=540, y=330)
        button_calculate.focus()
        root.bind('<Return>', on_enter_key)
        label_result.place(x=360, y=380)

        #Add an image of Component and schematic explaining the functioning
        logo_path = "Resistor_back.jpg"
        logo_path2 = "res_sym.jpg"
    elif selected_component == "BJT":
        #The same for the other components as well
        label_voltage.configure(text="V_CE (Volts)")
        label_current.configure(text="I_C (Amperes)")
        label_result.configure(text='')
        entry_voltage.delete(0, tk.END)
        entry_current.delete(0, tk.END)
        entry_eon.delete(0, tk.END)
        entry_eoff.delete(0, tk.END)
        entry_fs.delete(0, tk.END)
        entry_eon.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        entry_eoff.configure(width=0, height=0, corner_radius=0,border_width = 0, fg_color="#739bd0") 
        entry_fs.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        entry_output_curr_ldo.delete(0, tk.END)
        entry_output_curr_ldo.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        label_output_curr_ldo.configure(text="")
        label_eon.configure(text="")
        label_eoff.configure(text="")
        label_fs.configure(text="")

        button_calculate.place(x=540, y=330)
        button_calculate.focus()
        root.bind('<Return>', on_enter_key)
        label_result.place(x=360, y=380)

        logo_path = "BJT_back.jpg"
        logo_path2 = "bjt_sym.png"
    elif selected_component == "MOSFET":
        label_voltage.configure(text="V_DS (Volts)")
        label_current.configure(text="I_D (Amperes)")
        label_result.configure(text='')
        entry_voltage.delete(0, tk.END)
        entry_current.delete(0, tk.END)
        entry_eon.delete(0, tk.END)
        entry_eoff.delete(0, tk.END)
        entry_fs.delete(0, tk.END)
        entry_output_curr_ldo.delete(0, tk.END)
        entry_output_curr_ldo.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        label_output_curr_ldo.configure(text="")
        label_eon.configure(text="E_on(mJ):", fg="black", bg="#739bd0", font=("Arial Black", 13))
        label_eoff.configure(text="E_off(mJ):", fg="black", bg="#739bd0", font=("Arial Black", 13))
        label_fs.configure(text="E_fs(mJ):", fg="black", bg="#739bd0", font=("Arial Black", 13))
        entry_eon.configure(width=180, height=30, corner_radius=10, fg_color="white") 
        entry_eoff.configure(width=180, height=30, corner_radius=10, fg_color="white") 
        entry_fs.configure(width=180, height=30, corner_radius=10, fg_color="white") 
        label_eon.place(x=360, y=330)
        entry_eon.place(x=540, y=330)
        label_eoff.place(x=360, y=370)
        entry_eoff.place(x=540, y=370)
        label_fs.place(x=360, y=410)
        entry_fs.place(x=540, y=410)

        button_calculate.place(x=540, y=450)
        button_calculate.focus()
        root.bind('<Return>', on_enter_key)
        label_result.place(x=360, y=490) 
        logo_path = "Mosfet_back.jpg"  # Replace with the actual path to your logo image
        logo_path2 = "MOSFET_sym.png"
    elif selected_component == "IGBT":
        #Add an image of Component and schematic explaining the functioning
        label_voltage.configure(text="V_CE(sat) (Volts)")
        label_current.configure(text="I_C (Amperes)")
        label_result.configure(text='')
        entry_voltage.delete(0, tk.END)
        entry_current.delete(0, tk.END)
        entry_eon.delete(0, tk.END)
        entry_eoff.delete(0, tk.END)
        entry_fs.delete(0, tk.END)
        entry_output_curr_ldo.delete(0, tk.END)
        entry_output_curr_ldo.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        label_output_curr_ldo.configure(text="")
        label_eon.configure(text="E_on(mJ):", fg="black", bg="#739bd0", font=("Arial Black", 13))
        label_eoff.configure(text="E_off(mJ):", fg="black", bg="#739bd0", font=("Arial Black", 13))
        label_fs.configure(text="E_fs(mJ):", fg="black", bg="#739bd0", font=("Arial Black", 13))
        entry_eon.configure(width=180, height=30, corner_radius=10, fg_color="white") 
        entry_eoff.configure(width=180, height=30, corner_radius=10, fg_color="white") 
        entry_fs.configure(width=180, height=30, corner_radius=10, fg_color="white") 
        label_eon.place(x=360, y=330)
        entry_eon.place(x=540, y=330)
        label_eoff.place(x=360, y=370)
        entry_eoff.place(x=540, y=370)
        label_fs.place(x=360, y=410)
        entry_fs.place(x=540, y=410)

        button_calculate.place(x=540, y=450)
        button_calculate.focus()
        root.bind('<Return>', on_enter_key)
        label_result.place(x=360, y=490)

        logo_path = "Igbt.jpg"  # Replace with the actual path to your logo image
        logo_path2 = "IGBT_sym.png"
    elif selected_component == "LDO":
        label_voltage.configure(text="V_in (Volts)")
        label_current.configure(text="V_out (Volts):")
        label_result.configure(text='')
        entry_voltage.delete(0, tk.END)
        entry_current.delete(0, tk.END)
        entry_eon.delete(0, tk.END)
        entry_eoff.delete(0, tk.END)
        entry_fs.delete(0, tk.END)
        entry_eon.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        entry_eoff.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        entry_fs.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        label_eon.configure(text="")
        label_eoff.configure(text="")
        label_fs.configure(text="")
        label_output_curr_ldo.configure(text="I_out (Amperes):", fg="black", bg="#739bd0", font=("Arial Black", 13))
        entry_output_curr_ldo.configure(width=180, height=30, corner_radius=10, fg_color="white")
        label_output_curr_ldo.place(x=360, y=330)
        entry_output_curr_ldo.place(x=540, y=330)

        button_calculate.place(x=540, y=370)
        button_calculate.focus()
        root.bind('<Return>', on_enter_key)
        label_result.place(x=360, y=430)

        logo_path = "Ldo_back.jpg"
        logo_path2 = "LDO_sym.png"
    elif selected_component == "Capacitor":
        label_voltage.configure(text="V_C (Volts)")
        label_current.configure(text="C (Farads)")
        label_result.configure(text='')
        entry_voltage.delete(0, tk.END)
        entry_current.delete(0, tk.END)
        entry_eon.delete(0, tk.END)
        entry_eoff.delete(0, tk.END)
        entry_fs.delete(0, tk.END)

        entry_eon.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        entry_eoff.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        entry_fs.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 

        entry_output_curr_ldo.delete(0, tk.END)
        entry_output_curr_ldo.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        label_output_curr_ldo.configure(text="")

        label_eon.configure(text="")
        label_eoff.configure(text="")
        label_fs.configure(text="")

        label_current.place(x=360, y=290)
        entry_current.place(x=540, y=290)
        
        button_calculate.place(x=540, y=330)
        button_calculate.focus()
        root.bind('<Return>', on_enter_key)
        label_result.place(x=360, y=380)

        logo_path = "Capacitor_back.jpg"
        logo_path2 = "capacitor_sym.png"
    elif selected_component == "Varistor":
        label_voltage.configure(text="Vₒₚ,ₘₐₓ (Volts)")
        label_current.configure(text="Rₒₚ (Ohms)")
        label_result.configure(text='')
        entry_voltage.delete(0, tk.END)
        entry_current.delete(0, tk.END)
        entry_eon.delete(0, tk.END)
        entry_eoff.delete(0, tk.END)
        entry_fs.delete(0, tk.END)
        entry_eon.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        entry_eoff.configure(width=0, height=0, corner_radius=0,border_width = 0, fg_color="#739bd0") 
        entry_fs.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        entry_output_curr_ldo.delete(0, tk.END)
        entry_output_curr_ldo.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        label_output_curr_ldo.configure(text="")
        label_eon.configure(text="")
        label_eoff.configure(text="")
        label_fs.configure(text="")
        button_calculate.place(x=540, y=330)
        button_calculate.focus()
        root.bind('<Return>', on_enter_key)
        label_result.place(x=360, y=380)
        logo_path = "Varistor_back.jpg"  
        logo_path2 = "varistor_sym.png"
    elif selected_component == "Diode": 
        label_voltage.configure(text="V_Forward (V)")
        label_current.configure(text="I_Forward (A)")
        label_result.configure(text='')
        entry_voltage.delete(0, tk.END)
        entry_current.delete(0, tk.END)
        entry_eon.delete(0, tk.END)
        entry_eoff.delete(0, tk.END)
        entry_fs.delete(0, tk.END)
        entry_eon.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0") 
        entry_eoff.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0")
        entry_fs.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0")

        entry_output_curr_ldo.delete(0, tk.END)
        entry_output_curr_ldo.configure(width=0, height=0, corner_radius=0, border_width = 0, fg_color="#739bd0")
        label_output_curr_ldo.configure(text="")
        label_eon.configure(text="")
        label_eoff.configure(text="")
        label_fs.configure(text="")
        button_calculate.place(x=540, y=330)
        button_calculate.focus()
        root.bind('<Return>', on_enter_key)
        label_result.place(x=360, y=380)
        logo_path = "diode.jpg"
        logo_path2 = "diode_sym.jpg"

    #Display the appropriate component
    logo_image4 = Image.open(logo_path)
    logo_image4 = logo_image4.resize((120, 115))  # Adjust the size of the logo as needed
    logo_photo4 = ImageTk.PhotoImage(logo_image4)
    logo_component.configure(image=logo_photo4)
    logo_component.image = logo_photo4

    # and its schematic 
    logo_symb = Image.open(logo_path2)
    logo_symb = logo_symb.resize((160, 125))  # Adjust the size of the logo as needed
    logo_symbol = ImageTk.PhotoImage(logo_symb)
    logo_symbol2.configure(image=logo_symbol)
    logo_symbol2.image = logo_symbol

# Create the main window
root = tk.Tk()
root.title("Electronic Component Power Calculator")

# Styling
root.configure(bg="#739bd0")
root.geometry("1000x600")

ctk.set_appearance_mode("Light")

# Add the company logo
logo_path = "arcelik_logo_final.jpg"
logo_image = Image.open(logo_path)
logo_image = logo_image.resize((360, 240))  # Adjust the size of the logo as needed
logo_photo = ImageTk.PhotoImage(logo_image)
root.iconphoto(False, logo_photo)
logo_label = tk.Label(root, image=logo_photo, highlightthickness=0, borderwidth=0, relief="flat")
logo_label.place(x=-40, y=-55)


#Add the department logo:
logo_path = "mt-logo2.png"  # Replace with the actual path to your logo image
logo_image3 = Image.open(logo_path)
logo_image3 = logo_image3.resize((90, 80))  # Adjust the size of the logo as needed
logo_photo3 = ImageTk.PhotoImage(logo_image3)
logo_label3 = tk.Label(root, image=logo_photo3,  highlightthickness=0, borderwidth=0, relief="flat")
logo_label3.place(x=870, y=20)

#Add the name of the department below the logo
label_logo = tk.Label(root, text="TEST&ONAY DONANIM", fg="black", bg="#739bd0", font=("Arial", 9))
label_logo.place(x=845, y=105)

#Add the aesthetic photo:
logo_path = "wave.jpg"  # Replace with the actual path to your logo image
logo_image7 = Image.open(logo_path)
logo_image7 = logo_image7.resize((1100, 408))  # Adjust the size of the logo as needed
logo_photo7 = ImageTk.PhotoImage(logo_image7)
logo_label7 = tk.Label(root, image=logo_photo7,  highlightthickness=0, borderwidth=0, relief="flat")
logo_label7.place(x=0, y=244)

#Add the resistor photo: (The default that will change on the button change according to the component)
logo_path = "Resistor_back.jpg"  # Replace with the actual path to your logo image
logo_image5 = Image.open(logo_path)
logo_image5 = logo_image5.resize((120, 115))  # Adjust the size of the logo as needed
logo_photo5 = ImageTk.PhotoImage(logo_image5)
logo_component = tk.Label(root, image=logo_photo5,  highlightthickness=0, borderwidth=0, relief="flat")
logo_component.place(x=840, y=220)

#Add the symbol photo:
logo_path = "res_sym.jpg"  # Replace with the actual path to your logo image
logo_symb = Image.open(logo_path)
logo_symb = logo_symb.resize((120, 115))  # Adjust the size of the logo as needed
logo_symbol = ImageTk.PhotoImage(logo_symb)
logo_symbol2 = tk.Label(root, image=logo_symbol,  highlightthickness=0, borderwidth=0, relief="flat")
logo_symbol2.place(x=840, y=330)


# Component selection buttons
var_component = tk.StringVar()
var_component.set("Resistor")  # Default selection
button_resistor = tk.Radiobutton(root, text="Resistor", variable=var_component, value="Resistor", bg="#739bd0", fg="#000000", font=("Arial", 11, 'bold'), command=component_change)
button_bjt = tk.Radiobutton(root, text="BJT", variable=var_component, value="BJT", bg="#739bd0", fg="#000000", font=("Arial", 11, 'bold'), command=component_change)
button_mosfet = tk.Radiobutton(root, text="MOSFET", variable=var_component, value="MOSFET", bg="#739bd0", fg="#000000", font=("Arial", 11, 'bold'), command=component_change)
button_igbt = tk.Radiobutton(root, text="IGBT", variable=var_component, value="IGBT", bg="#739bd0", fg="#000000", font=("Arial", 11, 'bold'), command=component_change)
button_ldo = tk.Radiobutton(root, text="LDO", variable=var_component, value="LDO", bg="#739bd0", fg="#000000", font=("Arial", 11, 'bold'), command=component_change)
button_capacitor = tk.Radiobutton(root, text="Capacitor", variable=var_component, value="Capacitor", bg="#739bd0", fg="#000000", font=("Arial", 11, 'bold'), command=component_change)
button_varistor = tk.Radiobutton(root, text="Varistor", variable=var_component, value="Varistor", bg="#739bd0", fg="#000000", font=("Arial", 11, 'bold'), command=component_change, relief="flat")
button_diode = tk.Radiobutton(root, text="Diode", variable=var_component, value="Diode", bg="#739bd0", fg="#000000", font=("Arial", 11, 'bold'), command=component_change, relief="flat")

# Input widgets
label_voltage = tk.Label(root, text="V_R (Volts):", fg="black", bg="#739bd0", font=("Arial Black", 13))
entry_voltage = ctk.CTkEntry(master = root, width=180, height=30, corner_radius=10, fg_color="white")

label_current = tk.Label(root, text="I_R (Amperes):", fg="black", bg="#739bd0", font=("Arial Black", 13))
entry_current = ctk.CTkEntry(master = root, width=180, height=30, corner_radius=10, fg_color="white")

label_eon = tk.Label(root, text="E_on(mJ):", fg="black", bg="#739bd0", font=("Arial Black", 13))
entry_eon = ctk.CTkEntry(master = root, width=180, height=30, corner_radius=10, fg_color="white")

label_eoff = tk.Label(root, text="E_off (mJ):", fg="black", bg="#739bd0", font=("Arial Black", 13))
entry_eoff = ctk.CTkEntry(master = root, width=180, height=30, corner_radius=10, fg_color="white")

label_fs = tk.Label(root, text="F_s (Hz):", fg="black", bg="#739bd0", font=("Arial Black", 13))
entry_fs = ctk.CTkEntry(master = root, width=180, height=30, corner_radius=10, fg_color="white")

label_output_curr_ldo = tk.Label(root, text="I_out (Amperes):", fg="black", bg="#739bd0", font=("Arial Black", 13))
entry_output_curr_ldo = ctk.CTkEntry(master = root, width=180, height=30, corner_radius=10, fg_color="white")

label_title = tk.Label(root, text="Welcome to Arçelik Component Power Calculator!",fg="#660033", bg="#739bd0", font=("Forte", 16, "bold"))

# Prompt the user to enter their name
entry_user_name = ctk.CTkEntry(master = root, width=220, height=30, corner_radius=10, fg_color="white")
entry_user_name.place(x=540, y=190)
user = tk.Label(root, text="Please enter your name",fg="black", bg="#739bd0", font=("Arial Black", 13))
user.place(x=290, y=190)

#Button creating with Custom Tkinter
button_calculate = ctk.CTkButton(master=root,
                                 text="Calculate Power",
                                 command=calculate_power,
                                 width=140,
                                 height=30,
                                 border_width=0,
                                 corner_radius=8,
                                 fg_color = "Green")

label_result = tk.Label(root, text="", font=("Arial", 14, 'bold'), fg="#cc0000", bg="#739bd0")

# Creating the dropdown menu for presentations about the components
var_component2 = tk.StringVar()
var_component2.set("More on Resistor")  # Default selection

# Create a Label widget for the hyperlink
link_option = ctk.CTkOptionMenu(master = root, variable=var_component2, values=["More on Resistor", "More on BJT", "More on MOSFET", "More on IGBT", "More on LDO", "More on Capacitor", "More on Varistor"], fg_color="#ff6666", dropdown_fg_color = "#ff6666", dropdown_hover_color = "white", text_color = "black")
link_option.place(x=35, y=255)

#The text that will appear as the link
link_label = tk.Label(root, text="Info on Component", fg="#003300", bg="#739bd0", font=("Arial", 10, 'bold'), cursor="hand2")
link_label.place(x=35, y=295)

# Bind the mouse click event to the open_link function
link_label.bind("<Button-1>", open_link) 

#Author Rights (ahahahaha)
label_author = tk.Label(root, text="Prepared by 2023 summer intern Nihat Ahmadli", fg="black", bg="#3798dc", font=("Arial", 11, "bold"))

# Place widgets on the window
button_resistor.place(x=219, y=250)
button_bjt.place(x = 219, y= 280)
button_mosfet.place(x=219, y=310)
button_igbt.place(x=219, y=340)
button_ldo.place(x=219, y=370)
button_capacitor.place(x=219, y=400)
button_varistor.place(x=219, y=430)
button_diode.place(x=219, y=460)
label_title.place(x=270, y=150)
label_voltage.place(x=360, y=250)
entry_voltage.place(x=540, y=250)
label_current.place(x=360, y=290)
entry_current.place(x=540, y= 290)
button_calculate.place(x=540, y=330)
button_calculate.focus() 

# Bind the Enter key press event to the button click function
root.bind('<Return>', on_enter_key)

#Display the result
label_result.place(x=360, y=380)
label_author.place(x=660, y=575)

# Run the application
root.mainloop()


'''

    Importing the results to Excel

'''
# Create a new workbook to add the result of the power calculation to an Excel sheet
workbook = xlwt.Workbook()

# Select the active worksheet
worksheet = workbook.add_sheet('Sheet 1')

style = xlwt.easyxf('pattern: pattern solid, fore_colour light_yellow;'
                    'font: name Arial, bold on, color green, height 260;' # Font size = 280 (14pt)
                    'border: left thin, right thin, top thin, bottom thin;') 
style_label = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;'
                    'font: name Arial, bold on, color green, height 260;' # Font size = 280 (14pt)
                    'border: left thin, right thin, top thin, bottom thin;') 

style_label2 = xlwt.easyxf('pattern: pattern solid, fore_colour lime;'
                    'font: name Arial, bold on, color green, height 260;' # Font size = 280 (14pt)
                    'border: left thin, right thin, top thin, bottom thin;') 

style_label3 = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;'
                    'font: name Arial, bold on, color green, height 260;' # Font size = 280 (14pt)
                    'border: left thin, right thin, top thin, bottom thin;') 
style_label4 = xlwt.easyxf('pattern: pattern solid, fore_colour magenta_ega;'
                    'font: name Arial, bold on, color green, height 260;' # Font size = 280 (14pt)
                    'border: left thin, right thin, top thin, bottom thin;') 
style_label5 = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;'
                    'font: name Arial, bold on, color green, height 260;' # Font size = 280 (14pt)
                    'border: left thin, right thin, top thin, bottom thin;') 
style_label6 = xlwt.easyxf('pattern: pattern solid, fore_colour turquoise;'
                    'font: name Arial, bold on, color green, height 260;' # Font size = 280 (14pt)
                    'border: left thin, right thin, top thin, bottom thin;') 
style_label7 = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;'
                    'font: name Arial, bold on, color green, height 260;' # Font size = 280 (14pt)
                    'border: left thin, right thin, top thin, bottom thin;') 
style_label8 = xlwt.easyxf('pattern: pattern solid, fore_colour silver_ega;'
                    'font: name Arial, bold on, color green, height 260;' # Font size = 280 (14pt)
                    'border: left thin, right thin, top thin, bottom thin;') 
style_label9 = xlwt.easyxf('alignment: horizontal left;')
                    
style_label_user = xlwt.easyxf(
                    'font: name Arial, bold on, color black, height 260;' # Font size = 280 (14pt)
                    'border: left thin, right thin, top thin, bottom thin;') 

column_width = max([len("Power (Watts)")]) * 400  # 256 units = 1 character width

worksheet.col(0).width = column_width
worksheet.col(2).width = column_width
worksheet.col(4).width = column_width
worksheet.col(6).width = column_width
worksheet.col(8).width = column_width
worksheet.col(10).width = column_width
worksheet.col(12).width = column_width
worksheet.col(14).width = column_width
column_width = max([len("Resistor")]) * 400  # 256 units = 1 character width
worksheet.col(1).width = column_width
column_width = max([len("BJT")]) * 950  # 256 units = 1 character width
worksheet.col(3).width = column_width
column_width = max([len("MOSFET")]) * 550  # 256 units = 1 character width
worksheet.col(5).width = column_width
column_width = max([len("IGBT")]) * 850  # 256 units = 1 character width
worksheet.col(7).width = column_width
column_width = max([len("LDO")]) * 850  # 256 units = 1 character width
worksheet.col(9).width = column_width
column_width = max([len("Capacitor")]) * 400  # 256 units = 1 character width
worksheet.col(11).width = column_width
column_width = max([len("Varistor")]) * 500  # 256 units = 1 character width
worksheet.col(13).width = column_width
column_width = max([len("Diode")]) * 850  # 256 units = 1 character width
worksheet.col(15).width = column_width

try:
    worksheet.write_merge(0, 0, 0, 3, f"Hesaplayan kisi: {user_name}", style_label_user)
except:
    pass
worksheet.write(2, 0, "Komponent:", style_label)
worksheet.write(3, 0, "Güç (Watts):", style_label)
worksheet.write(4, 0, "V_R (Volts)", style_label)
worksheet.write(5, 0, "I_R (Amperes)", style_label)
worksheet.write(2, 1, "Resistor", style)
worksheet.write(2, 2, "Komponent:", style_label2)
worksheet.write(3, 2, "Güç (Watts):", style_label2)
worksheet.write(4, 2, "V_CE (Volts)", style_label2)
worksheet.write(5, 2, "I_C (Amperes)", style_label2)
worksheet.write(2, 3, "BJT", style)
worksheet.write(2, 4, "Komponent:", style_label)
worksheet.write(3, 4, "Güç (Watts):", style_label)
worksheet.write(4, 4, "V_DS (Volts)", style_label)
worksheet.write(5, 4, "I_D (Amperes)", style_label)
worksheet.write(6, 4, "E_on (mJ)", style_label)
worksheet.write(7, 4, "E_off (mJ)", style_label)
worksheet.write(8, 4, "E_fs (mJ):", style_label)
worksheet.write(2, 5, "MOSFET", style)
worksheet.write(2, 6, "Komponent:", style_label4)
worksheet.write(3, 6, "Güç (Watts):", style_label4)
worksheet.write(4, 6, "V_CE (Volts)", style_label4)
worksheet.write(5, 6, "I_C (Amperes)", style_label4)
worksheet.write(6, 6, "E_on (mJ)", style_label4)
worksheet.write(7, 6, "E_off (mJ)", style_label4)
worksheet.write(8, 6, "E_fs (mJ):", style_label4)
worksheet.write(2, 7, "IGBT", style)
worksheet.write(2, 8, "Komponent:", style_label)
worksheet.write(3, 8, "Güç (Watts):", style_label)
worksheet.write(4, 8, "V_in (Volts)", style_label)
worksheet.write(5, 8, "V_out (Volts)", style_label)
worksheet.write(6, 8, "I_out (Amperes)", style_label)
worksheet.write(2, 9, "LDO", style)
worksheet.write(2, 12, "Komponent:", style_label)
worksheet.write(3, 12, "Güç (Watts):", style_label)
worksheet.write(4, 12, "Vₒₚ,ₘₐₓ (Volts)", style_label)
worksheet.write(5, 12, "Rₒₚ (Ohms)", style_label)
worksheet.write(2, 13, "Varistor", style)
worksheet.write(2, 14, "Komponent:", style_label8)
worksheet.write(3, 14, "Güç (Watts):", style_label8)
worksheet.write(4, 14, "V_Forward (V):", style_label8)
worksheet.write(5, 14, "I_Forward (A):", style_label8)
worksheet.write(2, 15, "Diode", style)
worksheet.write(2, 10, "Komponent:", style_label6)
worksheet.write(3, 10, "Güç (Watts):", style_label6)
worksheet.write(4, 10, "V_C (Volts)", style_label6)
worksheet.write(5, 10, "C (Farads)", style_label6)
worksheet.write(2, 11, "Capacitor", style)

style = xlwt.easyxf('font: name Arial, height 240;')
if len(power_results) != 0:
    count = 1
    count2 = 1
    count3 = 1
    count4 = 5
    count5 = 5
    count6 = 5
    count7 = 9
    for pow in power_results:
        worksheet.write(3, count, pow)
        count += 2
    for vol in voltage_results:
        worksheet.write(4, count2, vol, style_label9)
        count2 +=2
    for curr in current_results:
        worksheet.write(5, count3, curr, style_label9)
        count3 +=2
    for eons in eon_results:
        worksheet.write(6, count4, eons, style_label9)
        count4 +=2
    for eoffs in eoff_results:
        worksheet.write(7, count5, eoffs, style_label9)
        count5 +=2
    for efss in efs_results:
        worksheet.write(8, count6, efss, style_label9)
        count6 +=2
    for out_cur in out_curr_res:
        worksheet.write(6, count7, out_cur, style_label9)


# Save the Excel file containing the results
workbook.save('output_power.xls')

