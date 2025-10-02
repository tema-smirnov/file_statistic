#Общая статистика
import pandas as pd
import matplotlib.pyplot as plt
import datetime
import tkinter
from tkinter import ttk
from tkinter import *
from tkinter import filedialog
from matplotlib.figure import Figure 
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
from PIL import ImageGrab
import pyglet, os

#--------------------Создаем окно--------------------
window = tkinter.Tk()
window.geometry('1280x720')
window.title("Статистика")
window.configure(bg='white')
#--------------------
pyglet.font.add_file('RostelecomBasis-Bold.otf')
pyglet.font.add_file('RostelecomBasis-Regular.otf')
#--------------------Создаем и добавляем лейбл--------------------

#--------------------
Tk().withdraw()  # Скрыть главное окно
fileex = filedialog.askopenfilename()
fileex_nam = os.path.splitext(os.path.basename(fileex))[0]
#--------------------



#--------------------Получаем датафрейм таблицы--------------------
df = pd.read_excel(fileex)
#--------------------print(df.head(len(df)))--------------------
columns1 = []
names = []


#--------------------Получаем список из заголовков столбцов(месяца)-------------------
columns = list()
columns = df.columns.ravel()[1:]    #получили спиок столбоцов в виде строк
columns1 = list()
for i in range(len(columns)):
    col = datetime.datetime.strptime(columns[i], '%m.%Y').strftime("%m.%Y")
    columns1.append(col)
#--------------------print(columns1)--------------------

#--------------------получаем список содержащий строки с результатами каждого сотрудника--------------------
res = df.values.tolist()    
#--------------------print(res)--------------------

#--------------------получаем список сотрудников--------------------
names = list()
for i in range(len(res)):
    names.append(res[i][0])
#--------------------print(names)--------------------

#список цветов (50)
colors = [
    "red", "blue", "green", "yellow", "orange", "purple", "pink", "brown", "cyan", "magenta",
    "lime", "maroon", "navy", "olive", "teal", "violet", "gold", "silver", "coral", "turquoise",
    "indigo", "salmon", "khaki", "lavender", "plum", "orchid", "crimson", "darkgreen", "darkblue", "darkred",
    "lightblue", "lightgreen", "lightpink", "lightyellow", "lightgray", "darkgray", "aqua", "beige", "chocolate", "firebrick",
    "forestgreen", "fuchsia", "gainsboro", "hotpink", "ivory", "linen", "moccasin", "peru", "rosybrown", "sienna"
]

#--------------------Функция для получения общего графика-------------------- 
def general_schedule():
    for widget in frame.winfo_children():
       widget.destroy() 
    # Создаем фигуру matplotlib
    fig = Figure(figsize=(14, 9), dpi=100)
    ax = fig.add_subplot(111)
    for i in range(len(df)):
        person = list()
        person = res[i]    #поучаем результаты отдельного сотрудника
        dates = columns1
        person_res = person[1:]
        ax.set_title("Статистика удержаний всех сторудников")
        ax.plot(dates, person_res, label = person[0], color = colors[i])
        ax.grid()
        ax.set_ylabel('Процент удержаний')
        ax.set_xlabel('Месяц')
    ax.set_xlim(left = 0)
    ax.set_ylim(bottom = 0)
    ax.legend(loc='lower center',
               mode='expand',
               borderaxespad=0,
               ncol=6)

    # Встраиваем фигуру в существующий виджет frame
    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(fill=tkinter.BOTH, expand=True)
    

#-------------------- --------------------

#--------------------Функция для получения статистики всех сотрудников на нескольких изображениях--------------------
# for i in range(len(df)):
#     person = list()
#     person = res[i]    #поучаем результаты отдельного сотрудника
#     dates = columns1
#     person_res = person[1:]
#     plt.title(f'Статистика удержаний {names[i]}')
#     plt.ylabel('Процент удержаний')
#     plt.xlabel('Месяц')
#     plt.plot(dates, person_res, label = names[i])
#     plt.legend(loc="best")
#     plt.grid()
#     plt.xlim(left=0)
#     plt.ylim(bottom=0, top = max(person_res)+ 5)
#     #plt.show()
#-------------------- --------------------

#--------------------Функция для получения статистики всех сотрудников на нескольких изображениях на одном--------------------
def schedules_of_all_employees():
    for widget in frame.winfo_children():
       widget.destroy() 
    n_plots = len(df)
    ncols = 4
    nrows = (n_plots // ncols) + (1 if n_plots % ncols else 0)

    fig, ax = plt.subplots(nrows=nrows, ncols=ncols, figsize=(14, 9), dpi = 100)
    ax = ax.flatten()  # упрощаем обход подграфиков в одномерном виде

    for i in range(len(ax)):
        if i < n_plots:
            person = res[i]           # данные сотрудника
            dates = columns1          # даты (ось X)
            person_res = person[1:]   # значения (ось Y)
        
            ax[i].plot(dates, person_res, label=names[i], color = colors [i])
            ax[i].set_title(f'Статистика удержаний {names[i]}', fontdict={"fontsize" : 6})
            ax[i].set_ylabel('Процент удержаний', fontdict={"fontsize" : 6})
            # ax[i].set_xlabel('Месяц', fontdict={"fontsize" : 5})
            ax[i].grid()
            ax[i].set_xlim(left=0)
            ax[i].set_ylim(bottom=0, top=max(person_res) + 5)
            ax[i].set_yticks(range(0, 40, 10))
            ax[i].tick_params(axis='y', which='major', labelsize=6)
            ax[i].tick_params(axis='x', which='major', labelsize=6, rotation=30)
        else:
            ax[i].axis('off')  # скрываем пустые подграфики

    plt.tight_layout()
    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(fill=tkinter.BOTH, expand=True)
# plt.show()

#--------------------Удаление графика--------------------
# def destr_fr():
#     for widget in frame.winfo_children():
#        widget.destroy()  
  
#--------------------вписал функцию во все кнопки с графиками --------------------
def show():
    for widget in frame.winfo_children():
       widget.destroy() 
    lbl.config(text=cb.get())
    selection = cb.get()
    # Создаем фигуру matplotlib
    fig = Figure(figsize=(14, 9), dpi=100)
    ax = fig.add_subplot(111)
    person = list()
    r_i =names.index(selection)
    person = res[r_i]    #поучаем результаты отдельного сотрудника
    dates = columns1
    person_res = person[1:]
    person_res_l = list()
    for u in range (len(person_res)):
        p = str(person_res[u])
        if p.isdigit():
            person_res_l.append(p)
    person_res_list = list(map(int, person_res_l))
    pers_perc_last_month = person_res_list[-1]
    pers_perc_befor_last_month = person_res_list[-2]
    if pers_perc_befor_last_month > 0:
        icncrease = pers_perc_last_month/pers_perc_befor_last_month
        icn_per = ((pers_perc_last_month - pers_perc_befor_last_month)/pers_perc_befor_last_month) * 100
        if icn_per > 0:
            text1 = f'Процент удержаний за этот месяц возрос на {icn_per:.2f}%!'
        if icn_per == 0:
            text1 = f'Процент удержаний за этот месяц не изменился'
        if icn_per < 0:
            text1 = f'Процент удержаний за этот месяц снизился на {icn_per:.2f}%!'
        if icncrease > 1:
            text = 'Молодец! Ты в этом месяце превзошёл свой прошлый результат! Не сбавляй темп!'
        elif icncrease == 1:
            text= 'Хороший результат, но ты можешь лучше. У тебя всё получится!'
        elif 0.99 >= icncrease >= 0.8:
            text= 'Ты сбавил темп, стоит приложить усилия и поработать над улучшением результата!'
        else:
            text = 'Пора взяться за ум, процент удержаний сильно упал, ты можешь лучше!\nЕсли нужна помощь, то обратись к наставнику или бизнес-тренеру, тебе обязательно помогу!'
    else:
        text, text1 = '', ''
    ax.set_title(f"Статистика {person[0]}")
    ax.plot(dates, person_res, label = person[0], color = colors[r_i])
    ax.grid()
    ax.set_ylabel('Процент удержаний')
    ax.set_xlabel('Месяц')
    ax.set_xlim(left = 0)
    ax.set_ylim(bottom = 0)
    ax.set_yticks(range(0, 50, 5))
    ax.legend(title= f'Процент удержаний за последний месяц {person_res_list[-1]}%\n{text1}\n{text}',
              loc='lower center',
               mode='expand',
               borderaxespad=0
               )

    # Встраиваем фигуру в существующий виджет frame
    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(fill=tkinter.BOTH, expand=True)
#-------------------- --------------------

#-------------------- --------------------
def save_frame_as_image(frame):
    filename = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png"), ("All files", "*.*")])
    if not filename:
        return
    # Обновляем окно, чтобы получить актуальные размеры и координаты
    frame.update_idletasks()
    # Получаем абсолютные координаты фрейма на экране

    # Делаем скриншот нужной области экрана
    img = ImageGrab.grab(bbox=(270, 128, 1670, 1028))
    img.save(filename)


def select_new_file():
    global fileex, df, columns, columns1, res, names
    Tk().withdraw()  # Скрыть главное окно
    fileex = filedialog.askopenfilename()
    fileex_nam = os.path.splitext(os.path.basename(fileex))[0]
    #--------------------Получаем датафрейм таблицы--------------------
    df = pd.read_excel(fileex)
    #--------------------print(df.head(len(df)))--------------------



    #--------------------Получаем новый список из заголовков столбцов(месяца)-------------------
    columns = list()
    columns = df.columns.ravel()[1:]    #получили спиок столбоцов в виде строк
    columns1 = list()
    for i in range(len(columns)):
        col = datetime.datetime.strptime(columns[i], '%m.%Y').strftime("%m.%Y")
        columns1.append(col)
   
    #--------------------получаем новый список содержащий строки с результатами каждого сотрудника--------------------
    res = df.values.tolist()    
   

    #--------------------получаем новый список сотрудников--------------------
    names = list()
    for i in range(len(res)):
        names.append(res[i][0])

    #--------------------получаем список содержащий строки с результатами каждого сотрудника--------------------
    res = df.values.tolist()    
    
    cb['values'] = names
    cb.set(names[0] if names else "")
    label.config(text=f"Статистика удержаний {fileex_nam}")
    




# filename_entry = Entry()
# filename_entry.place(x = 1450, y = 90)
# кнопки
Button(window, 
       text="Общий график",
       bg='white', 
       font = ('RostelecomBasis-Regular', 20),
       relief='ridge',
       justify = 'center', 
       command=general_schedule
       ).place(x = 0, y = 100)
Button(window, 
       text="Графики всех сотрудников", 
       bg='white', 
       font = ('RostelecomBasis-Regular', 20),
       relief='ridge',
       justify = 'center',
       wraplength=200,  
       command=schedules_of_all_employees
       ).place(x = 0, y = 170)
Button(window, 
       text="Выбрать сотрудника", 
       bg='white', 
       font = ('RostelecomBasis-Regular', 20),
       relief='ridge',
       justify = 'center', 
       wraplength=200, 
       command=show
       ).place(x = 0, y = 275)
Button(window, 
       text="Сохранить изображение", 
       bg='white', 
       font = ('RostelecomBasis-Regular', 20),
       relief='ridge',
       justify = 'center',
       wraplength=200,  
       command=lambda: save_frame_as_image(frame)
       ).place(x = 1726, y = 100)
# Button(window, 
#        text="Очистить", 
#        bg='white', 
#        font = ('RostelecomBasis-Regular', 20),
#        relief='ridge',
#        justify = 'center',  
#        command=destr_fr
#        ).place(x = 1776, y = 205) данная кнопка больше не требуется
Button(window, 
       text="Выбрать другую группу", 
       bg='white', 
       font = ('RostelecomBasis-Regular', 20),
       relief='ridge',
       justify = 'center',
       wraplength=200,  
       command=select_new_file
       ).place(x = 1713, y = 205)


#--------------------создаём комбобокс из всех сотрудников--------------------
names_var = StringVar(value=names[0]) 
cb = ttk.Combobox(window, 
                  textvariable=names_var, 
                  values=names,
                  font = ('RostelecomBasis-Regular', 11),
                  justify = 'center', 
                  state="readonly",
                  
                  )
cb.place(x = 0, y = 420)

#--------------------

#--------------------создаём лейбл для отображения выбранного сотрудника--------------------   
lbl = Label(window, 
            text=" "
            ) 
# lbl.place(x = 705, y = 50)
#--------------------
# Загружаем изображение
canvas = Canvas(height=96, width=200, borderwidth=0, highlightthickness=0)
logo_img = PhotoImage(file="логотип.png")
canvas.create_image(100, 50, image=logo_img)
canvas.configure(bg='white')
canvas.place(x = 0, y = 0)
#--------------------Создаём фрейм для отображения графиков--------------------
frame = tkinter.Frame(window, 
                      width=1400, 
                      height=900, 
                      bg='white'
                      )
frame.place(x = 270, y = 100)

label = tkinter.Label(window, 
                      text=f"Статистика удержаний {fileex_nam}", 
                      bg = 'white', 
                      font = ('RostelecomBasis-Bold', 30)
                      )
label.pack()

window.mainloop()