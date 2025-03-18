import schedule
import time
from openpyxl import load_workbook
import os
from win32com.client import Dispatch
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

programma_in_esecuzione = False

def aggiorna_excel():
    try:
        file_path = entry_file.get()
        
        excel = Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(os.path.abspath(file_path))

        wb.Save()
        log_message(f"File aperto aggiornato con successo alle {time.strftime('%H:%M:%S')}")
        
    except Exception as e:
        log_message(f"Errore win32com: {str(e)}")

def avvia_programma():
    global programma_in_esecuzione
    try:
        tempo_aggiornamento = int(entry_tempo.get())
        schedule.every(tempo_aggiornamento).seconds.do(aggiorna_excel)
        programma_in_esecuzione = True
        log_message(f"Programma avviato con aggiornamento ogni {tempo_aggiornamento} secondi.")
    except ValueError:
        log_message("Errore: inserisci un numero valido per il tempo di aggiornamento.")

def interrompi_programma():
    global programma_in_esecuzione
    if programma_in_esecuzione:
        schedule.clear()  
        programma_in_esecuzione = False
        log_message("Programma interrotto.")
    else:
        log_message("Nessun programma in esecuzione.")

def log_message(message):
    console_log.insert(tk.END, message + "\n")
    console_log.see(tk.END)  

root = tk.Tk()
root.title("Excel Updater")

frame_input = ttk.Frame(root)
frame_input.pack(padx=10, pady=10)

label_file = ttk.Label(frame_input, text="Percorso del file Excel:")
label_file.grid(row=0, column=0, sticky=tk.W)
entry_file = ttk.Entry(frame_input, width=50)
entry_file.grid(row=0, column=1, padx=5, pady=5)

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

def mainloop():
    while True:
        try:
            root.update()
            schedule.run_pending()
            time.sleep(0.1)
        except KeyboardInterrupt:
            log_message("Programma interrotto.")
            break

if __name__ == "__main__":
    mainloop()