import tkinter as tk
from tkinter import messagebox
import fitz  # PyMuPDF
import pandas as pd
import os
import threading
from tkinter import scrolledtext

def load_search_keywords_from_file(filename):
    if not os.path.exists(filename):
        return []
    with open(filename, 'r', encoding="utf-8") as file:
        keywords = [line.strip() for line in file if line.strip()]
    return keywords 

def read_folder_and_subfolders(folder_path):
    pile = []
    for dirpath, dirnames, filenames in os.walk(folder_path):
        for filename in filenames:
            pile.append(os.path.join(dirpath, filename))
    return pile

def count_multiple_words_in_pdf(pdf_path, words_to_search, log_widget):
    doc = fitz.open(pdf_path)
    words_count = {word: [] for word in words_to_search}
    total_occurrences = {word: 0 for word in words_to_search}

    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        for word in words_to_search:
            text_instances = page.search_for(word)
            word_count = len(text_instances)
            words_count[word].append((page_num + 1, word_count))
            total_occurrences[word] += word_count

    doc.close()
    for word in words_to_search:
        for page, count in words_count[word]:
            log_widget.insert(tk.END, f"Kata '{word}' ditemukan {count} kali di Halaman {page}. \n")
    
    return total_occurrences

def core(log_widget):
    log_widget.config(state=tk.NORMAL)
    key = 'key.ini'
    log_widget.insert(tk.END, "Load Key\n") 
    if not os.path.exists(key):
        return []         
    with open(key, 'r', encoding="utf-8") as file:
        keywords = [line.strip() for line in file if line.strip()]
        pile = []

        polder = 'source'
        log_widget.insert(tk.END, "Load File\n") 
        for dirpath, dirnames, filenames in os.walk(polder):
            for filename in filenames:
                pile.append(os.path.join(dirpath, filename))

        country = pile[0].split('\\')[1]
        log_widget.insert(tk.END, f"Load File {country} \n") 
        list_dict = []
        for piles in pile :
            book = piles.split('\\')[2]
            log_widget.insert(tk.END, f"Load File {book} \n") 
            book_data = {'Book': [book]}
            result = count_multiple_words_in_pdf(piles, keywords, log_widget)
            data = {key.capitalize(): [value] for key, value in result.items()}
            list_dict.append({**book_data, **data})
        
        combined_data = {}
        for data in list_dict:
            for key, value in data.items():
                if key in combined_data:
                    if key == 'Book':
                        combined_data[key].extend(value)
                    else:
                        combined_data[key].append(value[0])
                else:
                    combined_data[key] = value

        log_widget.insert(tk.END, f"Generate excel {country} \n")             
        data = {key: value for key, value in combined_data.items()}
        df = pd.DataFrame(data)
        file_path = f"{country}.xlsx" 
        df.to_excel(file_path, index=False, engine='openpyxl') 
        log_widget.insert(tk.END, f"Done \n") 
        log_widget.see(tk.END)
        log_widget.config(state=tk.DISABLED)

def on_button_click():
    log_widget.pack(padx=10, pady=10) 
    res = load_search_keywords_from_file('key.ini')
    pile = read_folder_and_subfolders('source')
    if not res:
        messagebox.showinfo("Warning", "Key not found")      
    else:
        if not pile:
            messagebox.showinfo("Warning", "File not found")    
        else:
            button.pack_forget() 
            threading.Thread(target=core, args=(log_widget,), daemon=True).start()   

window = tk.Tk()
window.title("Tools Analysis")

screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()

window_width = 500
window_height = 300

center_x = int(screen_width / 2 - window_width / 2)
center_y = int(screen_height / 2 - window_height / 2)

window.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
window.resizable(False, False)

label = tk.Label(window, text="Tools Data Content Analysis", font=("Arial", 14))
label.pack(pady=20)

button = tk.Button(window, text="Start", font=("Arial", 12), command=on_button_click)
button.pack(pady=10)

# Pindahkan guide labels ke sini
guide = [
    "Silahkan isi key di file key", 
    "Silahkan isi file disertai nama folder yang akan di generate file excel di dalam folder source"
]
for idx, item in enumerate(guide, start=1):
    label1 = tk.Label(window, text=f"{idx}. {item}", font=("Arial", 9), anchor="w", justify="left")
    label1.pack(anchor="w", padx=50, pady=1)

# Definisikan ScrolledText setelah labels
log_widget = scrolledtext.ScrolledText(window, width=50, height=10, font=("Arial", 10))
log_widget.config(state=tk.DISABLED)


window.mainloop()
