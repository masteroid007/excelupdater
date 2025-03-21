import schedule
import time
from openpyxl import load_workbook
import os
from win32com.client import Dispatch
import tkinter as tk
from tkinter import ttk, scrolledtext
from tkinter import filedialog as fd
from tkinter import END
from tkinter import messagebox

programma_in_esecuzione = False #stato del programma all'avvio

def aggiornamento_stato_led(color, text):

    led_canvas.itemconfig(led, fill=color)
    
    status_text = f"Stato: {text}"
    status_label.config(text=status_text)

def file_browse():
    file = fd.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
    entry_file.delete(0, tk.END)
    entry_file.insert(0, file) #tasto per selezionare il file tramite explorer

def aggiorna_excel():
    try:
        file_path = entry_file.get()
        
        if not os.path.exists(file_path):
            messagebox.showerror("Errore", "Il file è stato rimosso durante l'esecuzione!")
            interrompi_programma()
            return

        excel = Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(os.path.abspath(file_path))
        wb.Save()
        log_message(f"File aggiornato con successo alle {time.strftime('%H:%M:%S')}")
        
    except Exception as e:
        log_message(f"Errore durante l'aggiornamento: {str(e)}")
        messagebox.showerror("Errore critico", f"Si è verificato un errore:\n{str(e)}")
        interrompi_programma()

def avvia_programma():
    global programma_in_esecuzione
    try:
        file_path = entry_file.get()
        
        if not file_path:
            messagebox.showerror("Errore", "Inserisci un percorso file valido!")
            return
            
        if not os.path.exists(file_path):
            messagebox.showerror("Errore", "Il file specificato non esiste!")
            log_message("Tentativo di avvio con file inesistente")
            return
            
        if not file_path.lower().endswith(('.xlsx', '.xls')):
            messagebox.showerror("Errore", "Il file deve essere un documento Excel!")
            log_message("Tentativo di avvio con file non Excel")
            return

        tempo_aggiornamento = float(entry_tempo.get())
        schedule.every(tempo_aggiornamento).seconds.do(aggiorna_excel)
        programma_in_esecuzione = True
        log_message(f"Programma avviato - Aggiornamento ogni {tempo_aggiornamento}s")
        aggiornamento_stato_led('green', 'In esecuzione')
        
    except ValueError:
        messagebox.showerror("Errore", "Inserisci un numero valido per l'intervallo!")
        log_message("Valore temporale non valido inserito")

def interrompi_programma():
    global programma_in_esecuzione
    if programma_in_esecuzione:
        job = schedule.get_jobs()
        job.clear()
        schedule.clear()  
        log_message("Programma interrotto.")
        aggiornamento_stato_led('red', 'Fermo')
        programma_in_esecuzione = False
    else:
        log_message("Nessun programma in esecuzione.") 

def log_message(message):
    console_log.insert(tk.END, message + "\n")
    console_log.see(tk.END)   # costante per il log dei mess nella console

#sezione grafica del programma

root = tk.Tk()
root.title("Excel Updater")
root.iconbitmap("logo.ico")

frame_input = ttk.Frame(root)
frame_input.pack(padx=10, pady=10)

label_file = ttk.Label(frame_input, text="Percorso del file Excel:")
label_file.grid(row=0, column=0, sticky=tk.W)
entry_file = ttk.Entry(frame_input, width=50)
entry_file.grid(row=0, column=1, padx=5, pady=5)

ttk.Button(frame_input, text="Sfoglia File", command=file_browse).grid(row=0, column=2, padx=5)


label_tempo = ttk.Label(frame_input, text="Tempo di aggiornamento (secondi):")
label_tempo.grid(row=1, column=0, sticky=tk.W)
entry_tempo = ttk.Entry(frame_input, width=10)
entry_tempo.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)

button_avvia = ttk.Button(frame_input, text="Avvia", command=avvia_programma)
button_avvia.grid(row=2, column=0, pady=10)

button_interrompi = ttk.Button(frame_input, text="Interrompi", command=interrompi_programma)
button_interrompi.grid(row=2, column=1, pady=10)

frame_log = ttk.Frame(root)
frame_log.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

label_log = ttk.Label(frame_log, text="Log:")
label_log.pack(anchor=tk.W)

console_log = scrolledtext.ScrolledText(frame_log, width=60, height=15, state='normal')
console_log.pack(fill=tk.BOTH, expand=True)

frame_status = ttk.Frame(root)
frame_status.pack(pady=5)


led_canvas = tk.Canvas(frame_status, width=30, height=30, bd=0, highlightthickness=0)
led_canvas.grid(row=0, column=2, padx=5)


led = led_canvas.create_oval(5, 5, 25, 25, fill='red', outline='')


status_label = ttk.Label(frame_status, text="Stato: Fermo")
status_label.grid(row=0, column=3, padx=5)

def mainloop():
    while True:
        try:
            root.update()
            schedule.run_pending()
            time.sleep(0.1)
        except KeyboardInterrupt:
            log_message("Programma interrotto con successo")
            break

if __name__ == "__main__":
    mainloop()
