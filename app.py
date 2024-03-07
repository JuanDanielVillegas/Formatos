from tkinter import *
import customtkinter 
from cells import *
from main import select_emp, start

def show_apps(window):

    params = {
        "onvalue": 1,
        "text_color": "#37393b",
        "border_width": 2,
        "corner_radius": 3
    }

    suiefi_cb = customtkinter.CTkCheckBox(window, text="SUIEFI", **params)
    suiefi_cb.grid(row=9, column=0,  padx=(20, 0), pady=(10, 0), sticky="nsew")       

    dyp_cb = customtkinter.CTkCheckBox(window, text="DyP",**params)
    dyp_cb.grid(row=9, column=1,  padx=(0), pady=(10, 0), sticky="nsew")

    cuweb_cb = customtkinter.CTkCheckBox(window, text="CU WEB",**params)
    cuweb_cb.grid(row=9, column=2,  padx=(0,70), pady=(10, 0), sticky="nsew") 

    cognos_cb = customtkinter.CTkCheckBox(window, text="COGNOS WEB",**params)
    cognos_cb.grid(row=9, column=3,  padx=(0), pady=(10, 0), sticky="nsew") 

    crp_cb = customtkinter.CTkCheckBox(window, text="CRP", **params)
    crp_cb.grid(row=10, column=3,  padx=(0), pady=(20, 0), sticky="nsew")    

    sima_plan_cb = customtkinter.CTkCheckBox(window, text="SIMA PLAN.", **params)
    sima_plan_cb.grid(row=10, column=0,  padx=(20, 0), pady=(20, 0), sticky="nsew") 

    sima_seg_cb = customtkinter.CTkCheckBox(window, text="SIMA SEG.", **params)
    sima_seg_cb.grid(row=10, column=1,  padx=(0, 0), pady=(20, 0), sticky="nsew") 

    satcloud_cb = customtkinter.CTkCheckBox(window, text="SAT CLOUD", **params)
    satcloud_cb.grid(row=10, column=2,  padx=(0), pady=(20, 0), sticky="nsew") 

    vucem_cb = customtkinter.CTkCheckBox(window, text="VUCEM", **params)
    vucem_cb.grid(row=11, column=0,  padx=(20,0), pady=(20, 0), sticky="nsew") 

    return [suiefi_cb, dyp_cb, cuweb_cb,  cognos_cb, crp_cb, sima_plan_cb, satcloud_cb, vucem_cb, sima_seg_cb]
    

def call_main(window, entry):
    #f5f8fa
    textbox_bg = "#f5f8fa"

    
    radiobutton_params = {
        "border_width_unchecked": 3,
        "border_width_checked":7
    }

    checkbox_params = {
        "onvalue": 1,
        "text_color": "#37393b",
        "border_width": 2,
        "corner_radius": 3,
        "fg_color":"#279650",
        "hover_color":"#279650"
    }


    text = entry.get()
    nombre = text.replace(' ','%')
  
    emp = select_emp(nombre)

    
    label = customtkinter.CTkLabel(window, font=('Arial', 14), text_color="#0078d7",text="Número Empleado", fg_color="transparent")
    label.grid(row=1, column=0,  padx=(20, 0), pady=(0, 0), sticky="nw")

    label2 = customtkinter.CTkLabel(window, font=('Arial', 14), text_color="#0078d7",text="Nombre", fg_color="transparent")
    label2.grid(row=1, column=1,  padx=(0, 0), pady=(0, 0), sticky="nw")


    # lbl_background = customtkinter.CTkFrame(master=window, height=40, bg_color="#fff")
    # lbl_background.configure(width=900)
    # lbl_background.grid(row=2, column=0, columnspan=6, padx=20, pady=20, sticky="nw")

    #1d1e1e
    textbox = customtkinter.CTkTextbox(window, fg_color=textbox_bg, corner_radius=1, text_color="#37393b")
    textbox.configure(height=40, width=300)
    textbox.grid(row=2, column=0, padx=(20, 0), pady=(0), sticky="nsew")


    textbox.delete(1.0, END)  
    textbox.insert(END,  emp['A21'])


    textbox2 = customtkinter.CTkTextbox(window, fg_color=textbox_bg, corner_radius=1, text_color="#37393b")
    textbox2.configure(height=40, width=520)
    textbox2.grid(row=2, column=1, columnspan=4, padx=(0, 20), pady=0, sticky="nsew")


    textbox2.delete(1.0, END)  
    textbox2.insert(END, emp['A18'])


    label3 = customtkinter.CTkLabel(window, text="Movimiento:", fg_color="transparent")
    label3.grid(row=4, column=3, columnspan=1, padx=(0, 0), pady=(20, 0), sticky="nsew")       
    combobox = customtkinter.CTkComboBox(window, values=["", "Alta", "Baja", "Cambio", "Reactivación"], corner_radius=5, border_width=1,button_color="#3b8ed0" )
    combobox.grid(row=4, column=4, columnspan=1, padx=(20, 20), pady=(20, 0), sticky="nsew")


    label4 = customtkinter.CTkLabel(window, text="Acceso SUIEFI:", fg_color="transparent")
    label4.grid(row=5, column=3, columnspan=1, padx=(0, 0), pady=(20, 0), sticky="nsew")       
    combobox2 = customtkinter.CTkComboBox(window, values=["", "Alta", "Baja", "Cambio"], corner_radius=5, border_width=1, button_color="#3b8ed0")
    combobox2.grid(row=5, column=4, columnspan=1, padx=(20, 20), pady=(20, 0), sticky="nsew")


    checkbox = customtkinter.CTkCheckBox(window, text="Envío de Cédula", **checkbox_params)
    checkbox.configure(width=2, height=2)
    checkbox.grid(row=4, column=0, padx=(20, 0), pady=(20, 0), sticky="nsew")       


    radio_var = IntVar(value=0)

    radiobutton_1 = customtkinter.CTkRadioButton(window, text="S=Usuario", variable= radio_var, value=1, **radiobutton_params)
    radiobutton_1.grid(row=5, column=0,  padx=(20, 0), pady=(20, 0), sticky="nw")       
    radiobutton_2 = customtkinter.CTkRadioButton(window, text="T=Transferencia", variable= radio_var, value=2, **radiobutton_params)
    radiobutton_2.grid(row=5, column=1, padx=(0, 0), pady=(20, 0), sticky="nw")       


    sr_cb = customtkinter.CTkCheckBox(window, text="Seguimiento Revisiones", **checkbox_params)
    sr_cb.grid(row=4, column=1, padx=(0, 0), pady=(20, 0), sticky="nw")   

    apps_label = customtkinter.CTkLabel(window, height=25, font=('Arial', 16),text_color="#0078d7", text="Aplicaciones", corner_radius=5)
    apps_label.grid(row=7, column=0, columnspan=5, padx=20, pady=10, sticky="nsew")

    checkboxes = show_apps(window)


    fill_btn_params = {
        "text":"Llenar Formatos", "corner_radius":5, "fg_color":"#279650", "hover_color":"#1e5934"
    }
    
    
    if emp['A34'] == 'ip_not_found' :
        
        button3 = customtkinter.CTkButton(window, command=lambda:show_ip_window(window, emp, combobox.get().strip(), combobox2.get().strip(), checkbox.get(), radio_var.get(), sr_cb.get(), checkboxes), text="Ingrese IP", corner_radius=5, fg_color="#f23d3d", hover_color="#d63636")
        button3.configure(height=40)
        button3.grid(row=12, column=0, columnspan=5, padx=(0, 0), pady=(35, 20))

    else:
        #button2 = customtkinter.CTkButton(window, command=lambda:start(emp, combobox.get().strip(), combobox2.get().strip(), checkbox.get(), radio_var.get(), sr_cb.get(), checkboxes), **fill_btn_params)
        button2 = customtkinter.CTkButton(window, command=lambda:start(emp, combobox.get().strip(), combobox2.get().strip(), checkbox.get(), radio_var.get(), sr_cb.get(), checkboxes), **fill_btn_params)
        button2.configure(height=40)
        button2.grid(row=12, column=0, columnspan=5, padx=(0, 0), pady=(35, 20))


def set_field(emp, ip_entry):
    
    cells['A34'] = ip_entry
   
    return emp


def show_ip_window(window, emp, combobox, combobox2, checkbox, radio_var, sr_cb, checkboxes):
    ip_window = customtkinter.CTkToplevel()
    ip_window.title('IP')
    ip_window.geometry('450x90')
    ip_window.grab_set()

    ip_window.grid_columnconfigure(0, weight=1)

    ip_entry = customtkinter.CTkEntry(ip_window, placeholder_text="IP", corner_radius=5, font=('Arial', 12), border_width=1)
    ip_entry.configure( height=40)
    ip_entry.grid(row=0, column=0, columnspan=4, padx=(20, 0), pady=(20, 20), sticky="nsew")

    ip_button = customtkinter.CTkButton(ip_window, command=lambda: ( start(set_field(emp, ip_entry.get()), combobox, combobox2, checkbox, radio_var, sr_cb, checkboxes), ip_window.withdraw(),window.withdraw(), init()), text="Llenar Formatos", corner_radius=5, fg_color="#279650", hover_color="#1e5934")
    ip_button.configure(height=40)
    ip_button.grid(row=0, column=4, columnspan=1, padx=(20, 20), pady=(20, 20), sticky="nw")



def init():
    customtkinter.set_appearance_mode('ligth')
    window = customtkinter.CTk()
    window.title('Llenado de Formatos')
    window.geometry('820x520')

    window.grid_columnconfigure(0, weight=1)

    entry = customtkinter.CTkEntry(window, placeholder_text="Nombre del Empleado", corner_radius=5, font=('Arial', 14), border_width=1)
    entry.configure( height=40)
    entry.grid(row=0, column=0, columnspan=4, padx=(20, 0), pady=(20, 20), sticky="nsew")

    button = customtkinter.CTkButton(window, text="Buscar",  command=lambda:call_main(window, entry), corner_radius=5, )
    button.configure(height=40)
    button.grid(row=0, column=4, columnspan=1, padx=(20, 20), pady=(20, 20), sticky="nw")

    window.mainloop()


init()


