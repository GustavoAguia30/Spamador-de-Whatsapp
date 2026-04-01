import pywhatkit as kit
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from datetime import datetime, timedelta
import pandas as pd
import time
import sys
import os


def caminho_recurso(rel_path):
    try:
        base_path = sys._MEIPASS 
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, rel_path)


def carregar_excel():
    caminho = filedialog.askopenfilename(
        title="Selecionar arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )

    if caminho:
        try:
            df = pd.read_excel(caminho)
            numeros_excel = df.iloc[:, 0].dropna().astype(str).tolist()

            entrada_numeros.delete("1.0", tk.END)

            for numero in numeros_excel:
                entrada_numeros.insert(tk.END, numero + "\n")

            label_status.config(text=f"{len(numeros_excel)} números carregados")

            messagebox.showinfo("Sucesso", "Números carregados do Excel!")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler Excel:\n{str(e)}")


def enviar_mensagens():
    try:
        numeros = entrada_numeros.get("1.0", tk.END).strip().split("\n")
        mensagem = entrada_mensagem.get("1.0", tk.END).strip()

        if not numeros or not mensagem:
            messagebox.showwarning("Aviso", "Preencha todos os campos!")
            return

        hora = int(entrada_hora.get())
        minuto = int(entrada_minuto.get())

        tempo_envio = datetime.now().replace(hour=hora, minute=minuto, second=0)

        if tempo_envio < datetime.now():
            tempo_envio += timedelta(days=1)

        tempo_envio += timedelta(minutes=2)

        total = len([n for n in numeros if n.strip()])
        if total == 0:
            messagebox.showwarning("Aviso", "Nenhum número válido!")
            return

        progresso["maximum"] = total
        progresso["value"] = 0

        label_status.config(text="Iniciando envio...")
        label_progresso.config(text="0% (0/0)")

        enviados = 0

        for i, numero in enumerate(numeros):
            numero = numero.strip()

            if numero:
                if not numero.startswith("+"):
                    numero = "+55" + numero

                envio = tempo_envio + timedelta(minutes=i * 2)

                label_status.config(text=f"Enviando para: {numero}")

                kit.sendwhatmsg(
                    numero,
                    mensagem,
                    envio.hour,
                    envio.minute
                )

                enviados += 1
                progresso["value"] = enviados

                porcentagem = int((enviados / total) * 100)
                label_progresso.config(
                    text=f"{porcentagem}% ({enviados}/{total})"
                )

                janela.update_idletasks()

        label_status.config(text="Envio concluído ✅")
        messagebox.showinfo("Sucesso", "Mensagens agendadas!")

    except ValueError:
        messagebox.showerror("Erro", "Hora e minuto devem ser números!")
    except Exception as e:
        messagebox.showerror("Erro", str(e))


# JANELA PRINCIPAL
janela = tk.Tk()
janela.title("Disparador WhatsApp")
janela.geometry("960x880")
janela.configure(bg="#eef2f3")


style = ttk.Style()
style.theme_use("clam")

style.configure("TLabel", font=("Segoe UI", 10))
style.configure("TButton", font=("Segoe UI", 10, "bold"))


# Header
header = tk.Frame(janela, bg="#25D366", height=60)
header.pack(fill="x")

titulo = tk.Label(
    header,
    text=" Disparador de WhatsApp PRO",
    bg="#25D366",
    fg="white",
    font=("Segoe UI", 16, "bold")
)
titulo.pack(pady=10)


# Container
container = ttk.Frame(janela, padding=20)
container.pack(fill="both", expand=True)


# Números
frame_numeros = ttk.LabelFrame(container, text="📱 Números", padding=10)
frame_numeros.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

entrada_numeros = tk.Text(frame_numeros, height=6, width=35)
entrada_numeros.pack(side="left", fill="both", expand=True)

scroll_numeros = ttk.Scrollbar(frame_numeros, command=entrada_numeros.yview)
scroll_numeros.pack(side="right", fill="y")
entrada_numeros.config(yscrollcommand=scroll_numeros.set)

botao_excel = tk.Button(
    frame_numeros,
    text="📂 Importar Excel",
    command=carregar_excel,
    bg="#094f2a",
    fg="white",
    bd=0
)
botao_excel.pack(pady=5)


# Mensagem
frame_msg = ttk.LabelFrame(container, text="💬 Mensagem", padding=10)
frame_msg.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

entrada_mensagem = tk.Text(frame_msg, height=6, width=35)
entrada_mensagem.pack(side="left", fill="both", expand=True)

scroll_msg = ttk.Scrollbar(frame_msg, command=entrada_mensagem.yview)
scroll_msg.pack(side="right", fill="y")
entrada_mensagem.config(yscrollcommand=scroll_msg.set)


# Agendamento
frame_agendamento = ttk.LabelFrame(container, text="⏰ Agendamento", padding=15)
frame_agendamento.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")

ttk.Label(frame_agendamento, text="Hora:").grid(row=0, column=0, padx=5)
entrada_hora = ttk.Entry(frame_agendamento, width=8)
entrada_hora.grid(row=0, column=1, padx=5)

ttk.Label(frame_agendamento, text="Minuto:").grid(row=0, column=2, padx=5)
entrada_minuto = ttk.Entry(frame_agendamento, width=8)
entrada_minuto.grid(row=0, column=3, padx=5)


# Status
label_status = tk.Label(
    janela,
    text="Aguardando ação...",
    bg="#eef2f3",
    fg="#333"
)
label_status.pack()


# Barra de progresso
progresso = ttk.Progressbar(
    janela,
    orient="horizontal",
    length=500,
    mode="determinate"
)
progresso.pack(pady=5)


# Texto da barra
label_progresso = tk.Label(
    janela,
    text="0% (0/0)",
    bg="#eef2f3",
    font=("Segoe UI", 10, "bold")
)
label_progresso.pack()


# Botão
botao = tk.Button(
    janela,
    text="Enviar Mensagens",
    command=enviar_mensagens,
    bg="#25D366",
    fg="white",
    font=("Segoe UI", 12, "bold"),
    padx=15,
    pady=10,
    bd=0
)
botao.pack(pady=15)


container.columnconfigure(0, weight=1)
container.columnconfigure(1, weight=1)

janela.mainloop()