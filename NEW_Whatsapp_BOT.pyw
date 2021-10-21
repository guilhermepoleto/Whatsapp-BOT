from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import sys
import os
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import plyer.platforms.win.notification
from plyer import notification
import pandas as pd
import chromedriver_autoinstaller

def buscar(contato):
    campo = driver.find_element_by_xpath('//div[contains(@class,"copyable-text selectable-text")]')
    time.sleep(3)
    campo.click()
    campo.send_keys(contato)
    campo.send_keys(Keys.ENTER)

def enviar_msg(mensagem):
    campo_msg = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[1]/div/div[2]')
    campo_msg.click()
    time.sleep(3)
    campo_msg.send_keys(mensagem)
    campo_msg.send_keys(Keys.ENTER)

def enviar_foto(mensagem):
    attach = driver.find_element_by_xpath('//div[@title = "Anexar"]')
    attach.click()
    imgbox = driver.find_element_by_xpath('//input[@accept = "image/*,video/mp4,video/3gpp,video/quicktime"]')
    imgbox.send_keys(filepath)
    time.sleep(2)
    msg = driver.find_element_by_xpath('//div[contains(@class,"copyable-text selectable-text")]')
    msg.send_keys(mensagem)
    send = driver.find_element_by_xpath('//span[@data-icon="send"]')
    send.click()

def openFile():
    global filepath
    filepath = filedialog.askopenfilename(title="Open file",
                                          filetypes= (("JPG Files","*.jpg"),
                                          ("ALL Files","*.*")))
    nameFile = os.path.basename(filepath)
    l2.config(text=nameFile,font=("Arial", 10, "bold"))

def on_closing():
    global filepath
    filepath = ""
    win.destroy()

def on_OK():
    global getMsg
    getMsg = msgEntry.get("1.0",'end-1c')
    win.destroy()

def openWindowFile():
    global win
    win = tk.Tk()
    win.title("Abrir arquivo")
    win.geometry('470x300')
    global l2
    global msgEntry
    l2 = tk.Label(win,text="")
    l3 = tk.Label(win,text="")
    l4 = tk.Label(win,text="")
    msgEntry = tk.Text(win,width=50,height=8,borderwidth=2, relief="groove",font=("Helvetica", 11))
    button = tk.Button(win, text="Open Image",command=openFile)
    buttonOK = tk.Button(win, text="OK",command=on_OK)
    buttonSair = tk.Button(win, text="Cancelar",command=on_closing)
    l = tk.Label(text="")
    l.pack()
    button.pack()
    l4.pack()
    l2.pack()
    msgEntry.pack()
    l3.pack()
    buttonOK.pack(side='right', ipadx=50, padx=30)
    buttonSair.pack(side='left', ipadx=50, padx=30)
    win.protocol("WM_DELETE_WINDOW", on_closing)
    win.mainloop()

def getContato(cell):
    global contatos
    contatos.append(cell)

try:
    root = tk.Tk()
    root.withdraw()
    contatos = []

    EXCEL_path = filedialog.askopenfilename(title="Open Excel file",
                                            filetypes= (("EXCEL Files","*.xlsx"),
                                            ("ALL Files","*.*")))
    if EXCEL_path == "":
        messagebox.showerror("Erro", "Programa finalizado: Insira um arquivo")
        sys.exit()
    root.destroy()



    df = pd.read_excel(EXCEL_path, converters = {
            'Contatos': getContato
        })

    getMsg = ""
    filepath = ""
    openWindowFile()

    if filepath == "":
        if getMsg == "":
            sys.exit()

    driver = webdriver.Chrome()
    driver.get('https://web.whatsapp.com/')

    while len(driver.find_elements_by_id("side")) < 1:  
        time.sleep(1)
    notification.notify("Notificação", "QR Code Escaneado")

    if filepath == "":
        for cnt in contatos:
            buscar(cnt)
            time.sleep(3)
            enviar_msg(getMsg)
            time.sleep(3)
    else:
        for cnt in contatos:
            buscar(cnt)
            time.sleep(3)
            enviar_foto(getMsg)
            time.sleep(3)


except:
    driver.quit()
    sys.exit()
