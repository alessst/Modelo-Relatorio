import openpyxl
from openpyxl.drawing.image import Image
import tkinter as tk
from tkinter import filedialog, simpledialog
from shutil import copyfile



class InterfacePlanilha:


    def __init__(self, master):
        self.master = master
        self.master.title("Interface para Planilha")
        
        self.master.geometry("800x600")

        # Copiar o modelo para um novo arquivo
        copyfile('modelo.xlsx', 'planilha_atual.xlsx')
          # Aplicar o tema escuro
  

        # Abrir a planilha
        self.workbook = openpyxl.load_workbook('planilha_atual.xlsx')
        self.sheet = self.workbook.active

        # Armazenar imagens adicionadas
        self.imagens = []

        # Adicionar widgets
        self.label_dados1 = tk.Label(master, text="Local / LUC:")
        self.label_dados1.pack()
        self.entry_local_luc = tk.Entry(master)
        self.entry_local_luc.insert(tk.END, 'Novo Valor 1')
        self.entry_local_luc.pack()

        self.label_dados2 = tk.Label(master, text="Horário:")
        self.label_dados2.pack()
        self.entry_horario = tk.Entry(master)
        self.entry_horario.insert(tk.END, 'Novo Valor 2')
        self.entry_horario.pack()

        self.label_dados3 = tk.Label(master, text="Número da probe/telemetria:")
        self.label_dados3.pack()
        self.entry_numero_probe = tk.Entry(master)
        self.entry_numero_probe.insert(tk.END, 'Novo Valor 3')
        self.entry_numero_probe.pack()

        self.label_dados4 = tk.Label(master, text="Local (barramento) de instalação:")
        self.label_dados4.pack()
        self.entry_local_barramento = tk.Entry(master)
        self.entry_local_barramento.insert(tk.END, 'Novo Valor 4')
        self.entry_local_barramento.pack()

        self.label_dados5 = tk.Label(master, text="Identificação Ponto:")
        self.label_dados5.pack()
        self.entry_identificacao_ponto = tk.Entry(master)
        self.entry_identificacao_ponto.insert(tk.END, 'Novo Valor 5')
        self.entry_identificacao_ponto.pack()

        self.label_dados6 = tk.Label(master, text="Endereço:")
        self.label_dados6.pack()
        self.entry_endereco = tk.Entry(master)
        self.entry_endereco.insert(tk.END, 'Novo Valor 6')
        self.entry_endereco.pack()

        self.label_dados7 = tk.Label(master, text="A:")
        self.label_dados7.pack()
        self.entry_a = tk.Entry(master)
        self.entry_a.insert(tk.END, 'Novo Valor 7')
        self.entry_a.pack()

        self.label_dados8 = tk.Label(master, text="V:")
        self.label_dados8.pack()
        self.entry_v = tk.Entry(master)
        self.entry_v.insert(tk.END, 'Novo Valor 8')
        self.entry_v.pack()

        self.label_dados9 = tk.Label(master, text="kWh:")
        self.label_dados9.pack()
        self.entry_kwh = tk.Entry(master)
        self.entry_kwh.insert(tk.END, 'Novo Valor 9')
        self.entry_kwh.pack()

        self.label_dados10 = tk.Label(master, text="NOC:")
        self.label_dados10.pack()
        self.entry_noc = tk.Entry(master)
        self.entry_noc.insert(tk.END, 'Novo Valor 10')
        self.entry_noc.pack()

        self.label_dados11 = tk.Label(master, text="Instador:")
        self.label_dados11.pack()
        self.entry_instador = tk.Entry(master)
        self.entry_instador.insert(tk.END, 'Novo Valor 11')
        self.entry_instador.pack()

        self.label_imagem = tk.Label(master, text="Imagens:")
        self.label_imagem.pack()

        self.button_adicionar_imagem = tk.Button(master, text="Adicionar Imagem", command=self.adicionar_imagem)
        self.button_adicionar_imagem.pack()

        self.button_adicionar_multiplas_imagens = tk.Button(master, text="Adicionar Múltiplas Imagens", command=self.adicionar_multiplas_imagens)
        self.button_adicionar_multiplas_imagens.pack()

        self.button_salvar_planilha = tk.Button(master, text="Salvar Planilha", command=self.salvar_planilha)
        self.button_salvar_planilha.pack()
    
    def adicionar_imagem_prompt(self):
        cell = simpledialog.askstring("Célula da Imagem", "Informe a célula para adicionar a imagem (ex: B21):", parent=self.master)
        if cell:
            self.adicionar_imagem(cell)
            
    def adicionar_imagem(self, cell='B21', width=200, height=200):
        file_path = filedialog.askopenfilename(title="Selecione uma imagem", filetypes=[("Imagens", "*.png;*.jpg;*.jpeg;*.gif")])
        if file_path:
            img = Image(file_path)
            if width and height:
                img.width = width
                img.height = height
            self.sheet.add_image(img, cell)
            self.imagens.append(img)

    def adicionar_multiplas_imagens(self):
        num_imagens = simpledialog.askinteger("Número de Imagens", "Quantas imagens deseja adicionar?", parent=self.master)
        if num_imagens is not None:
            for _ in range(num_imagens):
                self.adicionar_imagem_prompt()

    def salvar_planilha(self):
        # Obter valores das entradas
        local_luc = self.entry_local_luc.get()
        horario = self.entry_horario.get()
        numero_probe = self.entry_numero_probe.get()
        local_barramento = self.entry_local_barramento.get()
        identificacao_ponto = self.entry_identificacao_ponto.get()
        endereco = self.entry_endereco.get()
        a = self.entry_a.get()
        v = self.entry_v.get()
        kwh = self.entry_kwh.get()
        noc = self.entry_noc.get()
        instador = self.entry_instador.get()

        # Atualizar valores nas células
        self.sheet['B1'] = local_luc
        self.sheet['E6'] = horario
        self.sheet['H13'] = numero_probe
        self.sheet['C15'] = local_barramento
        self.sheet['I3'] = identificacao_ponto
        self.sheet['D10'] = endereco
        self.sheet['B11'] = a
        self.sheet['F11'] = v
        self.sheet['H11'] = kwh
        self.sheet['C12'] = noc
        self.sheet['G12'] = instador

        # Salvar a planilha
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Planilhas Excel", "*.xlsx")])
        if file_path:
            self.workbook.save(file_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = InterfacePlanilha(root)
    root.mainloop()
