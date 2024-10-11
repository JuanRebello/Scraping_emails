import customtkinter as ctk
from tkinter import filedialog
from appnovo import iniciar_selenium  # Assegure-se de que iniciar_selenium está importado corretamente
import threading

class MyApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configurações básicas da janela
        self.title("Extração de emails")
        self.geometry("400x600")

        self.setup_layout()
        self.pause_duration = 0
        self.is_running = False
        self.automation_thread = None
        self.filepath = None  # Inicializa a variável filepath

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_layout(self):
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(padx=20, pady=20, fill="both", expand=True)

        self.header_label = ctk.CTkLabel(self.main_frame, text="Bem-vindo!")
        self.header_label.pack(pady=10, padx=10)

        self.file_label = ctk.CTkLabel(self.main_frame, text="Nenhum arquivo escolhido.")
        self.file_label.pack(pady=10, padx=10)

        self.choose_file_button = ctk.CTkButton(self.main_frame, text="Escolher arquivo", command=self.choose_file)
        self.choose_file_button.pack(pady=10, padx=10) 

        self.url_quantity_label = ctk.CTkLabel(self.main_frame, text="Quantidade de URLs a varrer:")
        self.url_quantity_label.pack(pady=10, padx=10)

        self.url_quantity_options = [str(i) for i in range(10, 101, 10)]
        self.url_quantity_menu = ctk.CTkOptionMenu(self.main_frame, values=self.url_quantity_options)
        self.url_quantity_menu.set('Escolha o lote')  # Define o texto padrão
        self.url_quantity_menu.pack(pady=10, padx=10)

        self.pause_label = ctk.CTkLabel(self.main_frame, text="Pausa entre os lotes:")
        self.pause_label.pack(pady=10, padx=10)

        self.pause_options = ["10s", "25s", "1min", "2min", "5min", "10min", "25min"]
        self.pause_menu = ctk.CTkOptionMenu(self.main_frame, values=self.pause_options, command=self.set_pause_duration)
        self.pause_menu.set('Escolha a pausa')
        self.pause_menu.pack(pady=10, padx=10)

        self.pause_message_label = ctk.CTkLabel(self.main_frame, text="")
        self.pause_message_label.pack(pady=10, padx=10)

        self.submit_button = ctk.CTkButton(self.main_frame, text="Iniciar varredura", command=self.start_scanning)
        self.submit_button.pack(pady=10, padx=10)

        self.feedback_label = ctk.CTkLabel(self.main_frame, text="", fg_color="lightgray")
        self.feedback_label.pack(pady=10, padx=10)

    def choose_file(self):
        self.filepath = filedialog.askopenfilename(title="Escolher arquivo", filetypes=[("Arquivos Excel", "*.xlsx")])
        if self.filepath:
            self.output_path = self.filepath.replace(".xlsx", " Atualizado.xlsx")  # Define o caminho do arquivo de saída
            self.file_label.configure(text=f"Arquivo escolhido: {self.filepath}")

    def set_pause_duration(self, value):
        if "s" in value:
            self.pause_duration = int(value[:-1])
        elif "min" in value:
            self.pause_duration = int(value[:-3]) * 60

    def start_scanning(self):
        if not self.is_running and self.filepath:  # Verifica se o arquivo foi escolhido
            self.is_running = True
            self.feedback_label.configure(text="Automação em execução")
            self.pause_message_label.configure(text="")  # Limpa a mensagem de pausa
            self.automation_thread = threading.Thread(target=self.run_automation)
            self.automation_thread.start()
        elif not self.filepath:
            self.feedback_label.configure(text="Por favor, escolha um arquivo antes de iniciar.")

    def run_automation(self):
        try:
            url_quantity = int(self.url_quantity_menu.get())
            pause_selection = f"{self.pause_duration}s" if self.pause_duration < 60 else f"{self.pause_duration // 60}min"
            iniciar_selenium(self.filepath, url_quantity, pause_selection, self.update_pause_message)
        finally:
            self.is_running = False
            self.feedback_label.configure(text="Automação finalizada")

    def update_pause_message(self, message):
        self.pause_message_label.configure(text=message)  # Atualiza a mensagem de pausa

    def show_pause_message(self, duration):
        self.pause_message_label.configure(text=f"Execução em pausa por {duration} segundos")

    def clear_pause_message(self):
        """ Limpa a mensagem de pausa. """
        self.pause_message_label.configure(text="")

    def on_closing(self):
        if self.is_running:
            self.feedback_label.configure(text="A automação está em execução. Aguarde...")
        else:
            self.destroy()

if __name__ == "__main__":
    app = MyApp()
    app.mainloop()
