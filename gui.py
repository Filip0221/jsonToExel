import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
# Import własnej logiki
import jsons_to_excel

class JsonToExcelGUI(ttk.Frame):
    def __init__(self, root):
        super().__init__(root, padding=12)
        self.root = root
        self.root.title("JSON → Excel (GUI)")
        self.pack(fill="both", expand=True)

        # Zmienne kontrolne Tkinter przechowujące ścieżki
        self.input_dir = tk.StringVar()
        self.output_file = tk.StringVar()

        # --- Tworzenie i rozmieszczanie elementów GUI ---
        
        # Folder wejściowy
        ttk.Label(self, text="Katalog z JSON-ami:").grid(row=0, column=0, sticky="w")
        ttk.Entry(self, textvariable=self.input_dir, width=50).grid(row=1, column=0, sticky="w", padx=(0,8))
        ttk.Button(self, text="Wybierz folder...", command=self.browse_folder).grid(row=1, column=1, sticky="w")

        # Plik wynikowy
        ttk.Label(self, text="Plik wynikowy (.xlsx):").grid(row=2, column=0, sticky="w", pady=(10,0))
        ttk.Entry(self, textvariable=self.output_file, width=50).grid(row=3, column=0, sticky="w", padx=(0,8))
        ttk.Button(self, text="Zapisz jako...", command=self.save_as).grid(row=3, column=1, sticky="w")

        # Start i status
        self.btn_start = ttk.Button(self, text="Start", command=self.on_start)
        self.btn_start.grid(row=4, column=0, sticky="w", pady=(12,0))

        self.status_label = ttk.Label(self, text="", foreground="blue")
        self.status_label.grid(row=5, column=0, columnspan=2, sticky="w", pady=(8,0))

    # Obsługa okna dialogowego wyboru folderu
    def browse_folder(self):
        folder = filedialog.askdirectory(title="Wybierz folder z plikami JSON")
        if folder:
            self.input_dir.set(folder)

    # Obsługa okna dialogowego zapisu pliku
    def save_as(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Zapisz jako (plik wynikowy)"
        )
        if path:
            self.output_file.set(path)

    # Walidacja danych wejściowych
    def validate(self):
        inp = self.input_dir.get().strip()
        out = self.output_file.get().strip()
        
        if not inp or not Path(inp).is_dir():
            messagebox.showwarning("Błąd folderu", "Wybierz istniejący folder z plikami JSON.")
            return False
        if not out or not out.lower().endswith(".xlsx"):
            messagebox.showwarning("Brak pliku", "Wybierz nazwę pliku wynikowego (.xlsx).")
            return False
            
        return True

    # Metoda uruchamiana po naciśnięciu "Start"
    def on_start(self):
        if not self.validate():
            return
        
        # Wyłączenie przycisku i zmiana statusu
        self.btn_start.config(state="disabled")
        self.status_label.config(text="Przetwarzanie...")

        # Uruchomienie konwersji w osobnym wątku, aby nie blokować GUI
        thread = threading.Thread(target=self._worker, daemon=True)
        thread.start()

    # Logika konwersji uruchamiana w tle
    def _worker(self):
        try:
            input_path = Path(self.input_dir.get())
            output_path = Path(self.output_file.get())

            # Wywołanie głównej funkcji konwertującej
            jsons_to_excel.jsonToExel(input_path, output_path)

        except Exception as e:
            # Powrót do wątku głównego (GUI) w przypadku błędu
            self.root.after(0, self._on_error, e)
        else:
            # Powrót do wątku głównego (GUI) w przypadku sukcesu
            self.root.after(0, self._on_success, output_path)
        finally:
            # Ponowne włączenie przycisku "Start" w wątku głównym
            self.root.after(0, lambda: self.btn_start.config(state="normal"))

    # Metody do obsługi wyników w wątku GUI (uruchamiane przez self.root.after(0, ...))
    def _on_error(self, exc):
        self.status_label.config(text="Błąd", foreground="red")
        messagebox.showerror("Błąd", f"Wystąpił błąd:\n{exc}")

    def _on_success(self, output_path):
        self.status_label.config(text=f"Gotowe: {output_path}", foreground="green")
        messagebox.showinfo("Sukces", f"Zapisano plik:\n{output_path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = JsonToExcelGUI(root)
    root.mainloop()