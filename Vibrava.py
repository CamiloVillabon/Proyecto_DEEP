from cgitb import text
from distutils.cmd import Command
from email import utils
from email.mime import image
from importlib.metadata import PathDistribution
from sqlite3 import Cursor
from tkinter import BOTH, BOTTOM, CENTER, END, FLAT, LEFT, RIGHT, SUNKEN, Y, LabelFrame, OptionMenu, StringVar, Text, Tk, Label, Button, Frame, Toplevel,  messagebox, filedialog, ttk, Scrollbar, VERTICAL, HORIZONTAL, PhotoImage
from turtle import width
import pandas as pd
from pathlib import Path
import os
from multiprocessing.util import info
import time
import webbrowser
from re import L
from unicodedata import category
from bs4 import BeautifulSoup
import pandas as pd
from typing import final
import numpy as np
import math


ventana = Tk()
ventana.config(bg='black')
#width= ventana.winfo_screenwidth()
ventana.iconbitmap('Vibrava.ico')
ventana.geometry('600x400')
#ventana.attributes('-fullscreen', True)
ventana.state('zoomed')
#ventana.attributes('-alpha',0.8)
ventana.minsize(width=600, height=400)
ventana.title('Operating Limit Summary Report')




ventana.columnconfigure(0, weight = 1)
ventana.rowconfigure(0, weight= 1)
ventana.columnconfigure(0, weight = 25)
ventana.rowconfigure(1, weight= 25)
ventana.columnconfigure(0, weight = 1)
ventana.rowconfigure(2, weight= 1)


frame0 = Frame(ventana, bg='black')
frame0.grid(column=0,row=0,sticky='nsew')

welcome_label=Label(frame0,text="Welcome to Vibrava", fg="white", bg='black', font=("Broadway",30,"bold"))
welcome_label.grid(row=0, column=1, rowspan=1, padx=250, columnspan=1, sticky= 'w')

# adding image (remember image should be PNG and not JPG)
img_1 = PhotoImage(file = 'Halliburton_logo_svg.png')
img_2 = PhotoImage(file = 'Vibrava_1.png')
#img_Bit_Bounce = PhotoImage(file = r"C:\Users\camil\OneDrive\Escritorio\DOCUMENTOS_HALLIBURTON\PROYECTO_DEEP\Bit_Bounce.jpg")
img1 = img_1.subsample(8, 8)
img2 = img_2.subsample(4, 4)
#img_Bit_Bounce = img_1.subsample(8, 8)

icon_label = Label (frame0)
icon_label.grid(row=0, column=0)

# setting image with the help of label
Label(frame0, image = img1, bg='black').grid(row = 0, column = 3, sticky='ne', columnspan = 2, rowspan = 2, padx = 25, pady = 25)
Label(frame0, image = img2, bg='black').grid(row = 0, column = 0, sticky='nw', padx=30, pady=10)

frame1 = Frame(ventana, bg='black')
frame1.grid(column=0,row=1,sticky='nsew')

about_us_label = Label (frame1, fg= 'white', bg='Black', font= ('Arial',10,'bold'))
about_us_label.grid(row=0, column=2, sticky='nsew')

frame2 = Frame(ventana, bg='black')
frame2.grid(column=0,row=2,sticky='nsew')

frame0.columnconfigure(0, weight = 1)
frame0.rowconfigure(0, weight= 1)
frame0.columnconfigure(1, weight = 1)
frame0.rowconfigure(0, weight= 1)
frame0.columnconfigure(2, weight = 1)
frame0.rowconfigure(0, weight= 1)

frame1.columnconfigure(0, weight = 1000)
frame1.rowconfigure(0, weight= 1)
frame1.columnconfigure(1, weight = 1)
frame1.rowconfigure(0, weight= 1)
frame1.columnconfigure(2, weight = 1)
frame1.rowconfigure(0, weight= 1)

frame2.columnconfigure(0, weight = 1)
frame2.rowconfigure(0, weight= 1)

frame2.columnconfigure(1, weight = 1)
frame2.rowconfigure(0, weight= 1)

frame2.columnconfigure(2, weight = 1)
frame2.rowconfigure(0, weight= 1)

frame2.columnconfigure(3, weight = 1)
frame2.rowconfigure(0, weight= 1)

frame2.columnconfigure(4, weight = 3)
frame2.rowconfigure(0, weight= 1)

menu= StringVar()
menu.set("Select Vibration Tool")



#Create a dropdown Menu
drop= OptionMenu(frame1, menu, "iCruise" , "DDSr-HCIM","DDSr-DGR")
drop.grid(column=2, row=0, sticky='n', pady=50, padx= 30, ipadx=70)
drop.config(bg='Spring green')
drop['menu'].config(bg='Spring green')

drop_label = ttk.Label(frame1, text="Vibration Mechanisms", background = 'black', foreground='white',font= ('Helvetica', 12))
drop_label.grid(column=2, row=0, sticky='n', pady=90)

new = 1
url = "https://app.powerbi.com/groups/me/reports/1acb9560-e50a-4a51-a9df-2b7cb7a8e43f/ReportSection"

#FILE MANAGER
#*********************************************************************************************************************************

def folder_opener(user_folder):

    # user_folder = input("Por favor agregue la ruta de su carpeta de trabajo: ")
    #user_folder = "C:/Users/camil/OneDrive/Escritorio/DOCUMENTOS_HALLIBURTON/PROYECTO_DEEP/Documentos_backup"
    
    os.chdir(user_folder)
    dir_list = os.listdir()
    print('dir_list', dir_list)
    final = allow_format(dir_list)
    print('final', final)
    return final

def allow_format(directory_list):

    # Checks if the file format is one of the allowed ones and avoids duplicated files
    # with different formats. Also ignores folders
    allowed_formats = ['xls', 'html', 'txt', 'xlsx']
    final_list = []
    unique_file_list = []

    for file in directory_list:
        dot_index = None
        
        for i in range(len(file)):
            if file[i] == ".":
                dot_index = i

        if dot_index is None:
            continue

        file_name = file[:dot_index]
        file_format = file[dot_index + 1:]
        
        if file_name not in unique_file_list and file_format in allowed_formats:
            unique_file_list.append(file_name)
            final_list.append(file)

    return final_list

#*******************************************************************************************************************************
#FUNCIONES DE VIBRACIONES
#*********************************************************************************************************************************
'''DATA EXTRACTION'''

def report_types():
    # Type of reports generated by the inSite platform.
    
    return ['Detailed Vibration Report', 'Operating Limit Summary Report']

def report_categ_ind(soup_file):
    # Returns an index based on the file type of category.

    title_tag = str(soup_file.h1)
    category_list = report_types()
    for cat in category_list:
        if cat in title_tag:
            return category_list.index(cat)

def info_parser(soup_file, table_ind: int):
    # Returns a list (or dictionary?) with the parsed information of the well or the analysis report.
    # through the use of the BeautifulSoup.find_all() method it extracts all the tags inside the html

    table_tag = soup_file.find_all("table")[table_ind]
    row_tag = table_tag.find_all("tr")

    final_lst = []
    for tr in row_tag:
        cell_lst = []

        for td in tr:
            if td.name:
                cell_lst.append(td.string)

        final_lst.append(str(x) for x in cell_lst)

    return dict(final_lst)

def vib_val_extr(parent_div):

    final_lst = []

    # Extract values from thead tags.
    thead_list = []
    for td in parent_div.thead.tr:
        if td.name:
            thead_list.append(str(td.string))

    final_lst.append(thead_list)

    # Extract values from tbody tags.
    for tr in parent_div.tbody:
        if tr.name:
            td_list = []

            for td in tr:
                if td.name:
                    td_list.append(str(td.string))

            final_lst.append(td_list)

    return final_lst

def vol_summary_extr(soup_file, file_category_ind):
    # Returns a datastructure with the parsed values from the vibrating
    # operating limit summary. It takes as parameter the soup from each xls
    # file and the table index 3.
    # Also, it's composed of two functions...

    value_lst = []
    table_ind = file_category_ind + 2
    table_tag = soup_file.find_all("table")[table_ind]

    div_tag = table_tag.find_all("div")
    for div in div_tag:
        if div.name:
            value_lst.append([str(div.string), vib_val_extr(div.parent)])

    return value_lst

'''CREATING DATAFRAMES'''

def row_merger(df_lst):

    # Note: Not used anymore due to changes in the design of the final dataframe.
    # Help to merge rows

    new_df_lst = []
    peak_df_merged = {}

    for df in df_lst:
        string_name = str(df.loc[0, "Measure Type"]).split()
        if "Peak" in string_name:
            key_p = string_name[1]
            if key_p not in peak_df_merged:
                peak_df_merged[key_p] = [df]

            else:
                peak_df_merged[key_p].append(df)

        else:
            new_df_lst.append(df)

    if len(peak_df_merged) > 0:
            for v in peak_df_merged.values():
                if len(v) > 1:
                    v[0] = v[0].combine_first(v[1]) 

                new_df_lst.append(v[0])

    return new_df_lst

def peak_name_fix(df):
    
    # Adds the type of measure done in Peak Bins to measure type, (Mins) or (Events).
 
    bit_run_units = ['(Mins)', '(Events)', '(count)']
    for word in bit_run_units:

        if word in df.columns[2]:
            
            if word == '(count)' or word == '(Events)':
                word = '(Events)'
            filt = df['Measure Type'].str.contains('Peak', na=False)
            df.loc[filt, 'Measure Type'] = df['Measure Type'] + ' ' + word

def vib_val_df(vib_val_list):
    # Function that returns a dataframe from the list of lists from vol_summary_extr
    # merged with other data like job number, run, sensor and even tool size for average bins.

    copy_lst = vib_val_list.copy()
    # df = pd.DataFrame([extrainfo_dict])
    # df.drop(columns = "Well Name", inplace = True)
    new_list = []

    for v_lst in copy_lst:
        header = v_lst[1].pop(0)

        v_df = pd.DataFrame(v_lst[1], columns = header)
        v_df.insert(0, "Measure Type", v_lst[0])
        v_df.iloc[:, 2] = [float(x) for x in v_df.iloc[:, 2]]
        peak_name_fix(v_df)
        new_list.append(v_df)

    return new_list

def info_merger_df(well_dict, vols_dict):
    # It takes both well information and the vibration operating limits summary
    # information data structures and returns a merged dataframe.

    well_copy = well_dict.copy()
    well_copy.update(vols_dict)
    
    # If needed here the dataframe can be cleaned for a future update...
    df = pd.DataFrame([well_copy])
    return df

def final_merger_df(well_dict, vols_dict, vib_val_dict_df):
    # Returns the raw final dataframe for this file.
    fin_df = pd.DataFrame()
    well_copy = well_dict.copy()
    well_copy.update(vols_dict)
    key_list = list(well_copy.keys())
    value_list = list(well_copy.values())

    for df in vib_val_dict_df:
        temp_df = df.copy()
        for i in range(len(key_list) - 1, -1, -1):
            temp_df.insert(0, key_list[i], value_list[i])

        fin_df = pd.concat([fin_df, temp_df]).reset_index(drop = True)
    
    return fin_df


'''DATA CLEANING'''

def df_modifier(df, file_category, file_name):
    
    # This function integers different functions related to modify the final
    # dataframe. it's only purpose is give more order to the script.
    tool_name = na_vib_tool_finder(df, file_name)
    del_av_bin_neg(df, file_category)
    column_eraser(df, file_category)
    column_replacer(df, 'M/LWD Tool Size', '6 Â¾" and smaller', '6 ¾" and smaller')
    column_replacer(df, 'Vibration Tool', 'N/A', tool_name)
    df.replace('None', np.nan, inplace=True)
    
def column_eraser(df, file_category):
    # Deletes columns that are considered useless for either reports.
    col_stnd = ['Rig Name','Activity Code', 'Report Generation Date and Time']
    col_stnd.extend(col_to_del(file_category))
    df.drop(columns=col_stnd, inplace=True)

def col_to_del(file_category):

    valid = report_types()
    if file_category not in valid:
        raise ValueError(f'Resultados: status debe ser una opción de {valid.join()}')
    
    if file_category == valid[0]:
        # Columns considered not useful for a Detailed Vibration Report.
        return ['Depth Range selected', 'Date/Time Range selected']

        # Columns considered not useful for a Operating Limit Summary Report.
    return ['GP RPM Filter Min Value', 'GP RPM Filter Max Value']
    
def del_av_bin_neg(df, file_category):

    # Reports contain two tables with "Delta Average Bins" however one of them has
    # negative Bands (G) so this function add a "(-)" to differenciate between each
    # table.
    valid = report_types()
    if file_category == valid[1]:
        filt = df['Band (G)'].str.contains('-', na=False)
        df.loc[filt, 'Measure Type'] = 'Delta Average Bins (-)'

def column_replacer(df, colmn, old_content, new_content):

    filter = (df[colmn] == old_content)
    df.loc[filter, colmn] = new_content

def na_vib_tool_finder(df, file_name):
    # Finds and returns the tool name from the file name in case it isn't available in the report.
    substring = ['PCM', 'PCDC', 'BaseStar'] # New tool's name can be added here.
    file_name = file_name.upper()

    for word in substring:
        if word == 'BASESTAR':
            return 'BaseStar'
        if word in file_name:
            return f'SVSS-{word}'
        

    return 'N/A'



    ######################################################################################################################################
    #end_index = file_name.find('-')
    #string_lst = file_name[:end_index].split()
    #start_index = 0
    #for i in range(len(string_lst)):
        #if string_lst[i] == 'VLA':
            #start_index = i

    #return '-'.join(string_lst[start_index + 1:])

'''ACCUMULATIVE FILTERS'''

def df_adapter(basic_df):

    # The function merges the different Bands and Bit Run columns into a single one for each case.
    df = basic_df.copy()
    df['Band'] = df['Band (G)'].combine_first(df['Band (%)'])
    df['Bit Run'] = df['Bit Run (Mins)'].combine_first(df['Bit Run (count)'])
    df.drop(['Band (G)', 'Band (%)', 'Bit Run (Mins)', 'Bit Run (count)'], axis=1, inplace=True)

    if 'Op Limit (Events)' in df.columns and 'Op Limit (Mins)' in df.columns:
        df['Op Limit'] = df['Op Limit (Mins)'].combine_first(df['Op Limit (Events)'])
        df.drop(['Op Limit (Mins)', 'Op Limit (Events)'], axis=1, inplace=True)
        df['Op Limit'] = df['Op Limit'].astype('float32')

    elif 'Op Limit (Mins)' in df.columns:
        df['Op Limit'] = df['Op Limit (Mins)']
        df.drop(['Op Limit (Mins)'], axis=1, inplace=True)
        df['Op Limit'] = df['Op Limit'].astype('float32')

    return df

def sum_data_filter(df):
    # Filters columns for a specific set of values and returns a new column with sums.
    
    if 'Op Limit' in df.columns:
        new_df = df.groupby(['Job Number','Vibration Tool',  'M/LWD Tool Size', 'Measure Type', 'Band']).agg(Bit_Run_Tot = ('Bit Run', 'sum'), Op_Lim = ('Op Limit', 'max')).round(2)  
        return new_df
    
    new_df = df.groupby(['Job Number', 'Vibration Tool', 'M/LWD Tool Size', 'Measure Type', 'Band'])['Bit Run'].apply(sum).rename('Bit_Run_Tot').round(2)
    return new_df

def surpass_op_lim(df):
    # Checks out if the accumulative Bit Run is higher than the operating limits values and returns a list of lists.
    df = df.reset_index()

    if 'Op_Lim' in df.columns:
        filt = (df['Bit_Run_Tot'] > df['Op_Lim'])
        #CREA DATAFRAME 
        #df_list = df[filt]
        #CREA LISTA DE LISTAS
        df_list = df[filt].values.tolist()
        print('Listado de limites', filt)

        return df_list

def available_tools(df):
    # Returns a list of vibration tools found in the dataframe.
    return list(df['Vibration Tool'].unique())

'''DATA EXPORTING'''

def export_xls(df, file_name, output_path):
    # Exports dataframes to excel.
    df_name = f'{file_name}.xlsx'
    with pd.ExcelWriter(f'{output_path}/output/{df_name}') as writer:
        df.to_excel(writer, index=False, header=True)
#**********************************************************************************************************************************

def openweb():
    webbrowser.open(url,new=new)

def update_progress_label():
    return f"Current Progress: {pb['value']}%"

def progress():
	GB = 100
	download = 0
	speed = 1
	while(download<GB):
		time.sleep(0.05)
		pb['value']+=(speed/GB)*100
		download+=speed
		value_label['text'] = update_progress_label()
	else:
		messagebox.showinfo(message='The progress completed!')

"""def progress():
    if pb['value'] < 100:
        pb['value'] += 20
        value_label['text'] = update_progress_label()
    else:
        messagebox.showinfo(message='The progress completed!')"""

def stop():
    pb.stop()
    value_label['text'] = update_progress_label()

pb = ttk.Progressbar(
    frame1,
    orient='horizontal',
    mode='determinate',
    length=300,
	style="red.Horizontal.TProgressbar"

)
pb.grid(column=2, row=0, sticky='n', pady=0, padx= 30)

value_label = ttk.Label(frame1, text=update_progress_label(), background = 'black', foreground='white',font= ('Helvetica', 10))
value_label.grid(column=2, row=0, sticky='n', pady=20)

#start_button = ttk.Button(frame1, text='Progress',command=progress)
#start_button.grid(column=2, row=3, padx=10, pady=10, sticky='w')

#stop_button = ttk.Button(frame1,text='Stop',command=stop)
#stop_button.grid(column=2, row=3, padx=10, pady=10, sticky='e')

#function calculate
def my_fun():
	my_dir= filedialog.askdirectory(title= 'Select your current directory')
	report_catgs = report_types()
	reports_dic = {v: pd.DataFrame() for v in report_catgs}
	files = folder_opener(user_folder=my_dir)
	cwd = os.getcwd()

	folder_path = f'{cwd}'.replace('\\', '/')

	for file in files:
		fixed_name = f'{folder_path}/{file}'
		try:
			with open(fixed_name) as html_file:
				soup = BeautifulSoup(html_file, 'html.parser')
		except:
			print(f"El archivo {file} no puede ser leído.")
			continue

		file_categ_ind = report_categ_ind(soup)
		file_categ = report_catgs[file_categ_ind]
		well_info = info_parser(soup, 0)
		vols_info = info_parser(soup, 1)
		values_dct = vol_summary_extr(soup, file_categ_ind)
		values_df_dict = vib_val_df(values_dct)
		raw_df = final_merger_df(well_info, vols_info, values_df_dict)

		df_modifier(raw_df, file_categ, file)
    
		reports_dic[file_categ] = pd.concat([reports_dic[file_categ],raw_df], ignore_index=True)
		print(f'¡El archivo {file} se ha analizado exitosamente!')

	pd.set_option("display.max_columns", 40)
	pd.set_option("display.max_rows", 200)

	os.makedirs(f'{folder_path}/output', exist_ok=True)

	for df_key, df_value in reports_dic.items():
		if len(df_value.columns) == 0:
			continue

        # Generates a modified dataframe and calculates values for it.
		temp_df = df_adapter(df_value)
        # Use it to show on the graphic user interface
		sum_gui_df = sum_data_filter(temp_df)
		print(sum_gui_df)

		print('df_key', df_key)
		df_name =  f'{df_key}_final_accumulated.xlsx'
		sum_gui_df.to_excel(f'{folder_path}/output/{df_name}', index=True, encoding='utf-8')
        # Final list with rows which overcame operating limits.
		surpassed_limits_list = surpass_op_lim(sum_gui_df)
        # Exports to excel files from either raw and calculated sum dataframes. 
		export_xls(df_value, f'{df_key}_final', folder_path)
		export_xls(sum_gui_df.reset_index(), f'{df_key}_accumulated', folder_path)
        
        # Tools available
		tools = available_tools(temp_df)
		print(tools)  
		

	return tools

def abrir_archivo():

	archivo = filedialog.askopenfilename(initialdir ='/', title='Selecione archivo', filetype=(('xls files', '*.xls*'),('All files', '*.*')))
	indica['text'] = archivo


def datos_excel():

	datos_obtenidos = indica['text']
	try:
		archivoexcel = r'{}'.format(datos_obtenidos)
		
		df = pd.read_excel(archivoexcel)

	except ValueError:
		messagebox.showerror('Informacion', 'Formato incorrecto')
		return None

	except FileNotFoundError:
		messagebox.showerror('Informacion', 'El archivo esta \n malogrado')
		return None

	Limpiar()
	df.fillna('', inplace=True)
	tabla['column'] = list(df.columns)
	tabla['show'] = "headings"  #encabezado

	for columna in tabla['column']:
		tabla.heading(columna, text= columna)
	

	df_fila = df.to_numpy().tolist()
	for fila in df_fila:
		tabla.insert('', 'end', values =fila)


def Limpiar():
	tabla.delete(*tabla.get_children())

def Bit_Bounce():
	toplvl = Toplevel(bg='black') #created Toplevel widger
	toplvl.resizable(False,False)
	toplvl.title("Bit Bounce")
	toplvl.grid_columnconfigure(0,weight=1)
	toplvl.grid_rowconfigure(0,weight=1)
	#toplvl.geometry("1000x500")
	photo = PhotoImage(file = 'Bit_Bounce.gif')
	lbl = Label(toplvl ,image = photo)
	#lbl2 = Label(toplvl, width=50, text='BIT BOUNCE', justify=LEFT, bg='black', fg='white', font=('Helvetica', 14, 'bold'))
	lbl.image = photo #keeping a reference in this line


	text2 = Text(toplvl, height=30, width=50, bg= 'black', fg='white', autoseparators=True)
	scroll = Scrollbar(toplvl, orient='vertical', command=text2.yview)
	text2.configure(yscrollcommand=scroll.set)
	text2.tag_configure('bold_italics', font=('Helvetica', 12, 'bold', 'italic'))
	text2.tag_configure('big', font=('Helvetica', 20, 'bold'), foreground='red', justify=CENTER)
	text2.tag_configure('color', foreground='white', font=('Helvetica', 14), justify=LEFT, relief= SUNKEN)
	#text2.tag_bind('follow','<1>',lambda e, t=text2: t.insert(END, "Not now, maybe later!"))
	text2.insert(END,'\nBIT BOUNCE\n', 'big')
	Description = """\nDescription: Axial or longitudinal motion of the drillstring resulting in large WOB fluctuations causing the bit to repeatedly lift-off and impact the formation.\n"""
	text2.insert(END, Description, 'color')
	Typical_Environment= """\nTypical Environment: Vertical wells, roller cone bits in hard rock, undergauge hole, ledges, and stringers.\n"""
	text2.insert(END, Typical_Environment,'color')
	Consequences = """\nConsequences: The impact loading will damage the drill bit cutting structure, bearings, and seals. The drillstring can sustain damage from the axial shocks and lateral shocks induced by the string flexing. Hoisting equipment may be damaged in shallow wells.\n"""
	text2.insert(END, Consequences, 'color')
	Recommended_Real_Time_Actions = """\nRecommended Real-Time Actions: Increase WOB and/or decrease RPM. If vibration persists, stop the rotation, and then restart drilling under a lower WOB and/or lower RPM.\n"""
	text2.insert(END, Recommended_Real_Time_Actions, 'color')
	Other_Solutions_Postrun = """\nOther Solutions (Post Run): Consider using a less aggressive roller cone bit and/or a shock sub.\n"""
	text2.insert(END, Other_Solutions_Postrun, 'color')

	lbl.grid(row=0, column=0)
	#lbl2.grid(row=0, column=1)
	text2.grid(row=0, column=1, sticky= 'ew')
	scroll.grid(row=0,column=2, sticky= 'ns')
	text2['yscrollcommand'] = scroll.set


	toplvl.mainloop()

def Stick_Slip():
	toplvl = Toplevel(bg='black') #created Toplevel widger
	toplvl.resizable(False,False)
	toplvl.title("Stick-Slip – A Torsional Motion")
	toplvl.grid_columnconfigure(0,weight=1)
	toplvl.grid_rowconfigure(0,weight=1)
	#toplvl.geometry("1000x500")
	photo = PhotoImage(file = 'Stick_Slip.gif')
	lbl = Label(toplvl ,image = photo)
	#lbl2 = Label(toplvl, width=50, text='BIT BOUNCE', justify=LEFT, bg='black', fg='white', font=('Helvetica', 14, 'bold'))
	lbl.image = photo #keeping a reference in this line


	text2 = Text(toplvl, height=30, width=50, bg= 'black', fg='white', autoseparators=True)
	scroll = Scrollbar(toplvl, orient='vertical', command=text2.yview)
	text2.configure(yscrollcommand=scroll.set)
	text2.tag_configure('bold_italics', font=('Helvetica', 12, 'bold', 'italic'))
	text2.tag_configure('big', font=('Helvetica', 20, 'bold'), foreground='red', justify=CENTER)
	text2.tag_configure('color', foreground='white', font=('Helvetica', 14), justify=LEFT, relief= SUNKEN)
	#text2.tag_bind('follow','<1>',lambda e, t=text2: t.insert(END, "Not now, maybe later!"))
	text2.insert(END,'\nSTICK SLIP\n', 'big')
	Description = """\nDescription: Non-uniform bit rotation in which the bit stops rotating momentarily at regular intervals, which causes the string to torque up periodically and then spin free. This mechanism sets up the primary torsional vibrations in the string.\n"""
	text2.insert(END, Description, 'color')
	Typical_Environment= """\nTypical Environment: High angle and deep wells, hard formations or salt, use of aggressive PDC bits with high WOB.\n"""
	text2.insert(END, Typical_Environment,'color')
	Consequences = """\nConsequences: Surface torque fluctuation > 15% of average. Stick-slip can cause PDC bit damage, lower ROP, connection overtorque, back-off, and drillstring twist-offs. Interference with mud pulse telemetry, wear on stabilizer and bit gauge.\n"""
	text2.insert(END, Consequences, 'color')
	Recommended_Real_Time_Actions = """\nRecommended Real-Time Actions: Increase RPM and/or decrease WOB. If stick-slip persists, stop the rotary and restart drilling under a higher RPM and/or lower WOB.\n"""
	text2.insert(END, Recommended_Real_Time_Actions, 'color')
	Other_Solutions_Postrun = """\nOther Solutions (Post Run): Consider using less aggressive PDC bit or a torque feedback system (i.e., a “soft torque”). Reduce stabilizer rotational drag (change blade design or number of blades, non-rotating stabilizer or roller reamer). Smooth well profile.\n"""
	text2.insert(END, Other_Solutions_Postrun, 'color')

	lbl.grid(row=0, column=0)
	#lbl2.grid(row=0, column=1)
	text2.grid(row=0, column=1, sticky= 'ew')
	scroll.grid(row=0,column=2, sticky= 'ns')
	text2['yscrollcommand'] = scroll.set

	toplvl.mainloop()

def Bit_Whirl():
	toplvl = Toplevel(bg='black') #created Toplevel widger
	toplvl.resizable(False,False)
	toplvl.title("Bit Whirl – A Lateral Motion")
	toplvl.grid_columnconfigure(0,weight=1)
	toplvl.grid_rowconfigure(0,weight=1)
	#toplvl.geometry("1000x500")
	photo = PhotoImage(file = 'Bit_Whirl.gif')
	lbl = Label(toplvl ,image = photo)
	#lbl2 = Label(toplvl, width=50, text='BIT BOUNCE', justify=LEFT, bg='black', fg='white', font=('Helvetica', 14, 'bold'))
	lbl.image = photo #keeping a reference in this line


	text2 = Text(toplvl, height=30, width=50, bg= 'black', fg='white', autoseparators=True)
	scroll = Scrollbar(toplvl, orient='vertical', command=text2.yview)
	text2.configure(yscrollcommand=scroll.set)
	text2.tag_configure('bold_italics', font=('Helvetica', 12, 'bold', 'italic'))
	text2.tag_configure('big', font=('Helvetica', 20, 'bold'), foreground='red', justify=CENTER)
	text2.tag_configure('color', foreground='white', font=('Helvetica', 14), justify=LEFT, relief= SUNKEN)
	#text2.tag_bind('follow','<1>',lambda e, t=text2: t.insert(END, "Not now, maybe later!"))
	text2.insert(END,'\nBIT WHIRL\n', 'big')
	Description = """\nDescription: Eccentric rotation of the bit about a point other than its geometric center caused by bit/wellbore gearing (analogous to a planetary gear). The mechanism induces high-frequency lateral vibration of the bit and drillstring.\n"""
	text2.insert(END, Description, 'color')
	Typical_Environment= """\nTypical Environment: High angle and deep wells, hard formations or salt, use of aggressive PDC bits with high WOB.\n"""
	text2.insert(END, Typical_Environment,'color')
	Consequences = """\nConsequences: Bit cutter impact damage, overgauge hole, BHA connection failures, and MWD component failures.\n"""
	text2.insert(END, Consequences, 'color')
	Recommended_Real_Time_Actions = """\nRecommended Real-Time Actions: Reduce RPM. If vibration persists, stop the rotary and restart drilling under a lower RPM. \n"""
	text2.insert(END, Recommended_Real_Time_Actions, 'color')
	Other_Solutions_Postrun = """\nOther Solutions (Post Run): Consider changing the bit (flatter profile, anti-whirl PDC bit or a roller cone bit), using stabilized BHA with full gauge near-bit stabilizer.\n"""
	text2.insert(END, Other_Solutions_Postrun, 'color')

	lbl.grid(row=0, column=0)
	#lbl2.grid(row=0, column=1)
	text2.grid(row=0, column=1, sticky= 'ew')
	scroll.grid(row=0,column=2, sticky= 'ns')
	text2['yscrollcommand'] = scroll.set

	toplvl.mainloop()

def BHA_Whirl():
	toplvl = Toplevel(bg='black') #created Toplevel widger
	toplvl.resizable(False,False)
	toplvl.title("BHA Whirl (Forward and Backward) – A Lateral Motion")
	toplvl.grid_columnconfigure(0,weight=1)
	toplvl.grid_rowconfigure(0,weight=1)
	#toplvl.geometry("1000x500")
	photo = PhotoImage(file = 'BHA_Whirl.gif')
	lbl = Label(toplvl ,image = photo)
	#lbl2 = Label(toplvl, width=50, text='BIT BOUNCE', justify=LEFT, bg='black', fg='white', font=('Helvetica', 14, 'bold'))
	lbl.image = photo #keeping a reference in this line


	text2 = Text(toplvl, height=30, width=50, bg= 'black', fg='white', autoseparators=True)
	scroll = Scrollbar(toplvl, orient='vertical', command=text2.yview)
	text2.configure(yscrollcommand=scroll.set)
	text2.tag_configure('bold_italics', font=('Helvetica', 12, 'bold', 'italic'))
	text2.tag_configure('big', font=('Helvetica', 20, 'bold'), foreground='red', justify=CENTER)
	text2.tag_configure('color', foreground='white', font=('Helvetica', 14), justify=LEFT, relief= SUNKEN)
	#text2.tag_bind('follow','<1>',lambda e, t=text2: t.insert(END, "Not now, maybe later!"))
	text2.insert(END,'\nBHA WHIRL\n', 'big')
	Description = """\nDescription: Similar to bit whirl, the BHA gears around the borehole and results in severe lateral shocks between the BHA and the wellbore. BHA whirl has been proven as the major cause of many drillstring and MWD component failures. BHA whirl can also occur while rotating/ reaming off-bottom. Whirl can occur in a forward or backward motion.\n"""
	text2.insert(END, Description, 'color')
	Typical_Environment= """\nTypical Environment: High angle and deep wells, hard formations or salt, use of aggressive PDC bits with high WOB.\n"""
	text2.insert(END, Typical_Environment,'color')
	Consequences = """\nConsequences: Bit cutter impact damage, overgauge hole, BHA connection failures, and MWD component failures.\n"""
	text2.insert(END, Consequences, 'color')
	Recommended_Real_Time_Actions = """\nRecommended Real-Time Actions: Reduce RPM. If vibration persists, stop the rotary and restart drilling under a lower RPM. \n"""
	text2.insert(END, Recommended_Real_Time_Actions, 'color')
	Other_Solutions_Postrun = """\nOther Solutions (Post Run): Consider changing the bit (flatter profile, anti-whirl PDC bit or a roller cone bit), using stabilized BHA with full gauge near-bit stabilizer.\n"""
	text2.insert(END, Other_Solutions_Postrun, 'color')

	lbl.grid(row=0, column=0)
	#lbl2.grid(row=0, column=1)
	text2.grid(row=0, column=1, sticky= 'ew')
	scroll.grid(row=0,column=2, sticky= 'ns')
	text2['yscrollcommand'] = scroll.set

	toplvl.mainloop()

def Lateral_Shocks():
	toplvl = Toplevel(bg='black') #created Toplevel widger
	toplvl.resizable(False,False)
	toplvl.title("Lateral Shocks – A Lateral Motion")
	toplvl.grid_columnconfigure(0,weight=1)
	toplvl.grid_rowconfigure(0,weight=1)
	#toplvl.geometry("1000x500")
	photo = PhotoImage(file = 'Lateral_Shocks.gif')
	lbl = Label(toplvl ,image = photo)
	#lbl2 = Label(toplvl, width=50, text='BIT BOUNCE', justify=LEFT, bg='black', fg='white', font=('Helvetica', 14, 'bold'))
	lbl.image = photo #keeping a reference in this line


	text2 = Text(toplvl, height=30, width=50, bg= 'black', fg='white', autoseparators=True)
	scroll = Scrollbar(toplvl, orient='vertical', command=text2.yview)
	text2.configure(yscrollcommand=scroll.set)
	text2.tag_configure('bold_italics', font=('Helvetica', 12, 'bold', 'italic'))
	text2.tag_configure('big', font=('Helvetica', 20, 'bold'), foreground='red', justify=CENTER)
	text2.tag_configure('color', foreground='white', font=('Helvetica', 14), justify=LEFT, relief= SUNKEN)
	#text2.tag_bind('follow','<1>',lambda e, t=text2: t.insert(END, "Not now, maybe later!"))
	text2.insert(END,'\nLATERAL SHOCKS\n', 'big')
	Description = """\nDescription: The BHA moves sideways or sometimes whirls forward and backwards randomly (chaos). Unlike backward whirl, this chaotic motion often results in medium/high peak lateral accelerations but low average accelerations of the DDS data. Lateral shocks have also been linked to many MWD and downhole tool connection failures. Lateral shocks of the BHA can be induced from either bit whirl or from rotating an unbalanced drillstring.\n"""
	text2.insert(END, Description, 'color')
	Typical_Environment= """\nTypical Environment: Hard rock and unbalanced or long unstabilized drillstring. Lateral shocks can be induced from bit whirl or lateral movements caused when the drillstring moves sideways during bit bounce.\n"""
	text2.insert(END, Typical_Environment,'color')
	Consequences = """\nConsequences: MWD component failures (motor, MWD tool, etc.) localized tool joint and/or stabilizer wear, washouts or twist-offs due to connection fatigue cracks, increased average torque.\n"""
	text2.insert(END, Consequences, 'color')
	Recommended_Real_Time_Actions = """\nRecommended Real-Time Actions: Reduce RPM to reduce the drillstring energy. If vibration persists, stop rotating and restart drilling with a lower RPM. \n"""
	text2.insert(END, Recommended_Real_Time_Actions, 'color')
	Other_Solutions_Postrun = """\nOther Solutions (Post Run): Use largest practical drill collar size and/or a packed hole assembly with full gauge stabilizers. Reduce eccentricity of the drillstring. In very hard formations, avoid using an aggressive PDC bit.\n"""
	text2.insert(END, Other_Solutions_Postrun, 'color')

	lbl.grid(row=0, column=0)
	#lbl2.grid(row=0, column=1)
	text2.grid(row=0, column=1, sticky= 'ew')
	scroll.grid(row=0,column=2, sticky= 'ns')
	text2['yscrollcommand'] = scroll.set

	toplvl.mainloop()

def Torsional_Resonance():
	toplvl = Toplevel(bg='black') #created Toplevel widger
	toplvl.resizable(False,False)
	toplvl.title("Torsional Resonance – A Torsional Motion")
	toplvl.grid_columnconfigure(0,weight=1)
	toplvl.grid_rowconfigure(0,weight=1)
	#toplvl.geometry("1000x500")
	photo = PhotoImage(file = 'Torsional_Resonance.gif')
	lbl = Label(toplvl ,image = photo)
	#lbl2 = Label(toplvl, width=50, text='BIT BOUNCE', justify=LEFT, bg='black', fg='white', font=('Helvetica', 14, 'bold'))
	lbl.image = photo #keeping a reference in this line


	text2 = Text(toplvl, height=30, width=50, bg= 'black', fg='white', autoseparators=True)
	scroll = Scrollbar(toplvl, orient='vertical', command=text2.yview)
	text2.configure(yscrollcommand=scroll.set)
	text2.tag_configure('bold_italics', font=('Helvetica', 12, 'bold', 'italic'))
	text2.tag_configure('big', font=('Helvetica', 20, 'bold'), foreground='red', justify=CENTER)
	text2.tag_configure('color', foreground='white', font=('Helvetica', 14), justify=LEFT, relief= SUNKEN)
	#text2.tag_bind('follow','<1>',lambda e, t=text2: t.insert(END, "Not now, maybe later!"))
	text2.insert(END,'\nTORSIONAL RESONANCE\n', 'big')
	Description = """\nDescription: This is specifically drill collar torsional resonance as a natural torsional frequency of the drill collars that is being excited. It is thought to be caused by impacts of individual cutters or by localized excessive side forces in the BHA generating a juddering motion. The change in RPM is very small and the frequency of the vibration is high.\n"""
	text2.insert(END, Description, 'color')
	Typical_Environment= """\nTypical Environment: This specific type of vibration occurs predominantly in very hard rocks when drilled with a PDC bit.\n"""
	text2.insert(END, Typical_Environment,'color')
	Consequences = """\nConsequences: It is most damaging at higher rotational speeds where higher amplitude resonance at harmonics of the collars natural frequency can occur. Impact damage can occur to downhole equipment.\n"""
	text2.insert(END, Consequences, 'color')
	Recommended_Real_Time_Actions = """\nRecommended Real-Time Actions: Change RPM to move away from the excitation frequency, typically 10%. If vibration persists, stop rotating and restart drilling with a different RPM. \n"""
	text2.insert(END, Recommended_Real_Time_Actions, 'color')
	#Other_Solutions_Postrun = """\nOther Solutions (Post Run): Use largest practical drill collar size and/or a packed hole assembly with full gauge stabilizers. Reduce eccentricity of the drillstring. In very hard formations, avoid using an aggressive PDC bit.\n"""
	#text2.insert(END, Other_Solutions_Postrun, 'color')

	lbl.grid(row=0, column=0)
	#lbl2.grid(row=0, column=1)
	text2.grid(row=0, column=1, sticky= 'ew')
	scroll.grid(row=0,column=2, sticky= 'ns')
	text2['yscrollcommand'] = scroll.set

	toplvl.mainloop()

def Parametric_Resonance():
	toplvl = Toplevel(bg='black') #created Toplevel widger
	toplvl.resizable(False,False)
	toplvl.title("Parametric Resonance – An Axial/Lateral Motion")
	toplvl.grid_columnconfigure(0,weight=1)
	toplvl.grid_rowconfigure(0,weight=1)
	#toplvl.geometry("1000x500")
	photo = PhotoImage(file = 'Parametric_Resonance.gif')
	lbl = Label(toplvl ,image = photo)
	#lbl2 = Label(toplvl, width=50, text='BIT BOUNCE', justify=LEFT, bg='black', fg='white', font=('Helvetica', 14, 'bold'))
	lbl.image = photo #keeping a reference in this line


	text2 = Text(toplvl, height=30, width=50, bg= 'black', fg='white', autoseparators=True)
	scroll = Scrollbar(toplvl, orient='vertical', command=text2.yview)
	text2.configure(yscrollcommand=scroll.set)
	text2.tag_configure('bold_italics', font=('Helvetica', 12, 'bold', 'italic'))
	text2.tag_configure('big', font=('Helvetica', 20, 'bold'), foreground='red', justify=CENTER)
	text2.tag_configure('color', foreground='white', font=('Helvetica', 14), justify=LEFT, relief= SUNKEN)
	#text2.tag_bind('follow','<1>',lambda e, t=text2: t.insert(END, "Not now, maybe later!"))
	text2.insert(END,'\nPARAMETRIC RESONANCE\n', 'big')
	Description = """\nDescription: Severe lateral vibration induced as a result of axial excitations caused by bit/formation interaction. The dynamic component of axial load is primarily caused by bit/formation interaction, which results in fluctuations of weight on bit. Axial fluctuations at a specific frequency will cause lateral deflection of the drillstring through the small lateral displacements that are already occurring (i.e., the small bends that already exist will be magnified due to the wave traveling through them).\n"""
	text2.insert(END, Description, 'color')
	Typical_Environment= """\nTypical Environment: Interbedded formations, under gauge hole.\n"""
	text2.insert(END, Typical_Environment,'color')
	Consequences = """\nConsequences: Severe lateral vibration can induce accelerated failure in the drillstring. It can also create the opportunity for borehole enlargement, which may lead to poor directional control and also lead on to whirl and other mechanisms of vibration.\n"""
	text2.insert(END, Consequences, 'color')
	Recommended_Real_Time_Actions = """\nRecommended Real-Time Actions: Increase WOB and decrease RPM by 10%. If vibration persists, stop rotating and restart drilling with modified parameters, RPM first.\n"""
	text2.insert(END, Recommended_Real_Time_Actions, 'color')
	Other_Solutions_Postrun = """\nOther Solutions (Post Run): Modify bit design or use shock sub to dampen axial motion.\n"""
	text2.insert(END, Other_Solutions_Postrun, 'color')

	lbl.grid(row=0, column=0)
	#lbl2.grid(row=0, column=1)
	text2.grid(row=0, column=1, sticky= 'ew')
	scroll.grid(row=0,column=2, sticky= 'ns')
	text2['yscrollcommand'] = scroll.set

	toplvl.mainloop()

def Bit_Chatter():
	toplvl = Toplevel(bg='black') #created Toplevel widger
	toplvl.resizable(False,False)
	toplvl.title("Bit Chatter – A Lateral/Torsional Motion")
	toplvl.grid_columnconfigure(0,weight=1)
	toplvl.grid_rowconfigure(0,weight=1)
	#toplvl.geometry("1000x500")
	photo = PhotoImage(file = 'Bit_Chatter.gif')
	lbl = Label(toplvl ,image = photo)
	#lbl2 = Label(toplvl, width=50, text='BIT BOUNCE', justify=LEFT, bg='black', fg='white', font=('Helvetica', 14, 'bold'))
	lbl.image = photo #keeping a reference in this line


	text2 = Text(toplvl, height=30, width=50, bg= 'black', fg='white', autoseparators=True)
	scroll = Scrollbar(toplvl, orient='vertical', command=text2.yview)
	text2.configure(yscrollcommand=scroll.set)
	text2.tag_configure('bold_italics', font=('Helvetica', 12, 'bold', 'italic'))
	text2.tag_configure('big', font=('Helvetica', 20, 'bold'), foreground='red', justify=CENTER)
	text2.tag_configure('color', foreground='white', font=('Helvetica', 14), justify=LEFT, relief= SUNKEN)
	#text2.tag_bind('follow','<1>',lambda e, t=text2: t.insert(END, "Not now, maybe later!"))
	text2.insert(END,'\nBIT CHATTER\n', 'big')
	Description = """\nDescription: This is the high-frequency resonance of the bit and BHA. The excitation is caused by slightly eccentric bit rotation where there is cutter interference with the bottom hole cutting pattern. The cutters ride up on to the ridge between previously cut grooves and then drop back into the grove.\n"""
	text2.insert(END, Description, 'color')
	Typical_Environment= """\nTypical Environment: PDC bits drilling in high compressive strength rocks will create this vibration where each cutter is impacted on the formation.\n"""
	text2.insert(END, Typical_Environment,'color')
	Consequences = """\nConsequences: Bit cutter impact damage. High-frequency vibration can cause failure of electronic equipment due to vibration of electronic components and solder joints. This bit dysfunction can lead to bit whirl.\n"""
	text2.insert(END, Consequences, 'color')
	Recommended_Real_Time_Actions = """\nRecommended Real-Time Actions: Adjust RPM up or down to a region away from the RPM being used. Adjust the WOB if necessary to remove the condition. If vibration persists, stop rotating and restart drilling with modified parameters. It may be necessary to break the bit in to reestablish a cutting pattern.\n"""
	text2.insert(END, Recommended_Real_Time_Actions, 'color')
	Other_Solutions_Postrun = """\nOther Solutions (Post Run): Modified bit design or bit selection.\n"""
	text2.insert(END, Other_Solutions_Postrun, 'color')

	lbl.grid(row=0, column=0)
	#lbl2.grid(row=0, column=1)
	text2.grid(row=0, column=1, sticky= 'ew')
	scroll.grid(row=0,column=2, sticky= 'ns')
	text2['yscrollcommand'] = scroll.set

	toplvl.mainloop()

def Modal_Coupling():
	toplvl = Toplevel(bg='black') #created Toplevel widger
	toplvl.resizable(False,False)
	toplvl.title("Vibration Modal Coupling – Involves All Three Motions")
	toplvl.grid_columnconfigure(0,weight=1)
	toplvl.grid_rowconfigure(0,weight=1)
	#toplvl.geometry("1000x500")
	photo = PhotoImage(file = 'Modal_Coupling.gif')
	lbl = Label(toplvl ,image = photo)
	#lbl2 = Label(toplvl, width=50, text='BIT BOUNCE', justify=LEFT, bg='black', fg='white', font=('Helvetica', 14, 'bold'))
	lbl.image = photo #keeping a reference in this line


	text2 = Text(toplvl, height=30, width=50, bg= 'black', fg='white', autoseparators=True)
	scroll = Scrollbar(toplvl, orient='vertical', command=text2.yview)
	text2.configure(yscrollcommand=scroll.set)
	text2.tag_configure('bold_italics', font=('Helvetica', 12, 'bold', 'italic'))
	text2.tag_configure('big', font=('Helvetica', 20, 'bold'), foreground='red', justify=CENTER)
	text2.tag_configure('color', foreground='white', font=('Helvetica', 14), justify=LEFT, relief= SUNKEN)
	#text2.tag_bind('follow','<1>',lambda e, t=text2: t.insert(END, "Not now, maybe later!"))
	text2.insert(END,'\nMODAL COUPLING\n', 'big')
	Description = """\nDescription: A coupling motion among axial, torsional, and lateral vibrations. The coupling motion creates axial and torque oscillations and high lateral shocks of the BHA. The motion is similar to chaos so the DDS’s average data will not be very high.\n"""
	text2.insert(END, Description, 'color')
	Typical_Environment= """\nTypical Environment: Vertical or near vertical wells, pendulum or unstabilized BHA, and hard rock.\n"""
	text2.insert(END, Typical_Environment,'color')
	Consequences = """\nConsequences: MWD component failures, bit cutter impact damage, collar and stabilizer wear, wash-outs and twist-offs due to connection fatigue cracks.\n"""
	text2.insert(END, Consequences, 'color')
	Recommended_Real_Time_Actions = """\nRecommended Real-Time Actions: Stop rotating and pick up off bottom. Resume drilling with modified WOB and RPM. Attempt a lower RPM first.\n"""
	text2.insert(END, Recommended_Real_Time_Actions, 'color')
	Other_Solutions_Postrun = """\nOther Solutions (Post Run): Consider changing bit style and/or modifying BHA (packed hole assembly). Reduce stabilizer drag (blade design, non-rotating). Use a torque feedback system. Consider using the downhole mud motor.\n"""
	text2.insert(END, Other_Solutions_Postrun, 'color')

	lbl.grid(row=0, column=0)
	#lbl2.grid(row=0, column=1)
	text2.grid(row=0, column=1, sticky= 'ew')
	scroll.grid(row=0,column=2, sticky= 'ns')
	text2['yscrollcommand'] = scroll.set

	toplvl.mainloop()


tabla = ttk.Treeview(frame1 , height=30)

tabla.grid(column=0, row=0, sticky='SNEW')

ladox = Scrollbar(frame1, orient = HORIZONTAL, command= tabla.xview)
ladox.grid(column=0, row = 1, sticky='ew') 

ladoy = Scrollbar(frame1, orient = VERTICAL, command = tabla.yview)
ladoy.grid(column = 1, row = 0, sticky='ns')

tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)

estilo = ttk.Style(frame1)
estilo.theme_use('clam') #  ('clam', 'alt', 'default', 'classic')
#estilo.configure(".",font= ('Arial', 14), foreground='red2')
#estilo.configure("Treeview", font= ('Helvetica', 12), foreground='black',  background='white')
#estilo.map('Treeview',background=[('selected', 'green2')], foreground=[('selected','black')] )
estilo.configure("Treeview.Heading",font= ('Helvetica', 14), foreground='red2',  background='black')
estilo.configure("Treeview", font= ('Helvetica', 12), foreground='white',  background='black', fieldbackground='black')
estilo.map('Treeview',background=[('selected', 'sea green')], foreground=[('selected','white')] )
estilo.configure("red.Horizontal.TProgressbar", foreground='red', background='Spring green')
estilo.configure("drop", foreground='black', background='Spring green')

boton0 = Button(frame2, text= 'Calculate', bg='Spring green', command=lambda:[my_fun(),progress() ])
#boton0 = Button(frame2, text= 'Calculate', bg='dark turquoise', command=my_fun)
boton0.grid(column = 0, row = 0, sticky='nsew', padx=10, pady=10)

boton1 = Button(frame2, text= 'Final Report', bg='Spring green', command= abrir_archivo)
boton1.grid(column = 1, row = 0, sticky='nsew', padx=10, pady=10)

boton2 = Button(frame2, text= 'Show', bg='Spring green', command= datos_excel)
boton2.grid(column = 2, row = 0, sticky='nsew', padx=10, pady=10)

boton3 = Button(frame2, text= 'Reset', bg='Spring green', command=lambda:[stop(),Limpiar() ])
#boton3 = Button(frame2, text= 'Reset', bg='dark turquoise', command= Limpiar)
boton3.grid(column = 3, row = 0, sticky='nsew', padx=10, pady=10)

boton4 = Button(frame1, text= 'Power BI Visualization', bg='Spring green', command= openweb())
boton4.grid(column = 2, row = 0, padx=20, ipadx=70, pady=0, sticky= 's')

indica = Label(frame2, fg= 'white', bg='black', text= 'File location', font= ('Arial',10,'bold') )
indica.grid(column=4, row = 0)

Bit_bounce = Button(frame1, text= 'Bit Bounce', bg='Spring green', command= Bit_Bounce)
Bit_bounce.grid(column = 2, row = 0, padx=20, ipadx=70, pady=120, sticky= 'wen')

Stick_slip = Button(frame1, text= 'Stick Slip', bg='Red', command= Stick_Slip)
Stick_slip.grid(column = 2, row = 0, padx=20, ipadx=70, pady=160, sticky= 'wen')

Bit_whirl = Button(frame1, text= 'Bit Whirl', bg='Spring green', command= Bit_Whirl)
Bit_whirl.grid(column = 2, row = 0, padx=20, ipadx=70, pady=200, sticky= 'wen')

BHA_whirl = Button(frame1, text= 'BHA Whirl', bg='Spring green', command= BHA_Whirl)
BHA_whirl.grid(column = 2, row = 0, padx=20, ipadx=70, pady=240, sticky= 'wen')

Lateral_shocks = Button(frame1, text= 'Lateral Shocks', bg='Spring green', command= Lateral_Shocks)
Lateral_shocks.grid(column = 2, row = 0, padx=20, ipadx=70, pady=280, sticky= 'wen')

Torsional_resonance = Button(frame1, text= 'Torsional Resonance', bg='Orange', command= Torsional_Resonance)
Torsional_resonance.grid(column = 2, row = 0, padx=20, ipadx=70, pady=300, sticky= 'wes')

Parametric_resonance = Button(frame1, text= 'Parametric Resonance', bg='Spring green', command= Parametric_Resonance)
Parametric_resonance.grid(column = 2, row = 0, padx=20, ipadx=70, pady=260, sticky= 'wes')

Bit_chatter = Button(frame1, text= 'Bit Chatter', bg='Spring green', command= Bit_Chatter)
Bit_chatter.grid(column = 2, row = 0, padx=20, ipadx=70, pady=220, sticky= 'wes')

x=3
y=0

if y<x:
	Vibration_Modal_Coupling = Button(frame1, text= 'Vibration Modal Coupling', bg='Spring green', command= Modal_Coupling)
	Vibration_Modal_Coupling.grid(column = 2, row = 0, padx=20, ipadx=70, pady=180, sticky= 'wes')
else:
	Vibration_Modal_Coupling = Button(frame1, text= 'Vibration Modal Coupling', bg='red', command= Modal_Coupling)
	Vibration_Modal_Coupling.grid(column = 2, row = 0, padx=20, ipadx=70, pady=180, sticky= 'wes')


ventana.mainloop()