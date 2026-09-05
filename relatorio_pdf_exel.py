import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pypdf

# Mapeamento dos meses em maiúsculas
MESES_PT = {
    1: 'JANEIRO',
    2: 'FEVEREIRO',
    3: 'MARÇO',
    4: 'ABRIL',
    5: 'MAIO',
    6: 'JUNHO',
    7: 'JULHO',
    8: 'AGOSTO',
    9: 'SETEMBRO',
    10: 'OUTUBRO',
    11: 'NOVEMBRO',
    12: 'DEZEMBRO',
}


def converter_mes_vencimento(data_str):
  """Converte datas como '10/10/2026' para 'OUTUBRO' ou retorna 'SEM DATA'."""
  if (
      not data_str
      or "sem data" in str(data_str).lower()
      or "//" in str(data_str)
  ):
    return "SEM DATA"

  m = re.search(r"\d{2}/(\d{2})/\d{4}", str(data_str).strip())
  if m:
    mes_num = int(m.group(1))
    return MESES_PT.get(mes_num, str(data_str).strip())
  return str(data_str).strip()


def processar_pdf_siga(caminho_pdf):
  reader = pypdf.PdfReader(caminho_pdf)
  texto_completo = ""
  for page in reader.pages:
    texto_completo += page.extract_text() or ""

  linhas = texto_completo.split("\n")
  start = False
  table_lines = []

  for l in linhas:
    if "Listados:" in l:
      break
    if (
        "------------------------------------------------------------------------------------------------------------------------------------"
        in l
    ):
      if not start:
        start = True
        continue
    if start:
      table_lines.append(l)

  registros = []
  for l_str in table_lines:
    l_clean = l_str.strip()
    if not l_clean or "CLIENTE" in l_clean or "---" in l_clean:
      continue

    m = re.match(r"^([A-Z0-9_\s]{7,9})\s+(\d{8})\s+(.*)$", l_clean)
    if m:
      cli = m.group(1).strip()
      cep = m.group(2).strip()
      resto = m.group(3).strip()

      # Captura técnico, ignora comercial, captura quantidade e marca
      m_tech = re.search(
          r"\s+([A-Z])\s+(\d)\s+(\d{1,2})\s+([A-Z0-9/_-]+)\s*(.*)$", resto
      )
      if m_tech:
        end = resto[: m_tech.start()].strip()
        tec = m_tech.group(1)
        # m_tech.group(2) corresponde ao comercial e foi removido
        qtd = int(m_tech.group(3))
        marca = m_tech.group(4)
        datas_str = m_tech.group(5)

        datas = re.findall(r"\d{2}/\d{2}/\d{4}", datas_str)
        if len(datas) >= 3:
          vencto_raw = datas[2]
        elif len(datas) == 1:
          vencto_raw = datas[0]
        else:
          vencto_raw = "Sem Data"

        registros.append({
            "CLIENTE": cli,
            "CEP": f"{cep[:5]}-{cep[5:]}",
            "ENDEREÇO": end,
            "ST_TEC": tec,
            "QTDE": qtd,
            "MARCA": marca,
            "MÊS EXEC.": converter_mes_vencimento(vencto_raw),
        })

  return pd.DataFrame(registros)


class AppConversor:

  def __init__(self, root):
    self.root = root
    self.root.title("Conversor SIGA RIA — CREL Elevadores")
    self.root.geometry("520x330")
    self.root.resizable(False, False)
    self.root.configure(bg="#F4F6F9")

    # Cabeçalho
    lbl_titulo = tk.Label(
        root,
        text="Conversor de PDF SIGA para Excel",
        font=("Segoe UI", 15, "bold"),
        fg="#1B365D",
        bg="#F4F6F9",
    )
    lbl_titulo.pack(pady=(20, 5))

    lbl_sub = tk.Label(
        root,
        text="Extrai dados formatados sem setor comercial e com mês por extenso",
        font=("Segoe UI", 9),
        fg="#666666",
        bg="#F4F6F9",
    )
    lbl_sub.pack(pady=(0, 20))

    # Seleção de arquivo
    frame_pdf = tk.Frame(root, bg="#F4F6F9")
    frame_pdf.pack(fill="x", padx=30, pady=5)

    tk.Label(
        frame_pdf,
        text="Arquivo PDF:",
        font=("Segoe UI", 9, "bold"),
        bg="#F4F6F9",
        fg="#333",
    ).pack(anchor="w")
    self.entry_pdf = tk.Entry(
        frame_pdf, font=("Segoe UI", 10), bg="#FFFFFF", relief="solid", bd=1
    )
    self.entry_pdf.pack(side="left", fill="x", expand=True, ipady=4, pady=3)

    btn_busca = tk.Button(
        frame_pdf,
        text="Selecionar...",
        command=self.selecionar_pdf,
        font=("Segoe UI", 9),
        bg="#E2E8F0",
        relief="groove",
    )
    btn_busca.pack(side="right", padx=(8, 0), ipady=2)

    # Status
    self.lbl_status = tk.Label(
        root,
        text="Aguardando seleção do arquivo PDF...",
        font=("Segoe UI", 9, "italic"),
        fg="#888888",
        bg="#F4F6F9",
    )
    self.lbl_status.pack(pady=(15, 15))

    # Botão de Ação
    self.btn_converter = tk.Button(
        root,
        text="🚀 Converter e Salvar em Excel (.xlsx)",
        command=self.converter,
        font=("Segoe UI", 11, "bold"),
        bg="#1B365D",
        fg="#FFFFFF",
        activebackground="#0f233d",
        activeforeground="#FFFFFF",
        relief="flat",
        cursor="hand2",
    )
    self.btn_converter.pack(fill="x", padx=30, ipady=8)

  def selecionar_pdf(self):
    caminho = filedialog.askopenfilename(
        title="Selecione o Relatório RIA em PDF",
        filetypes=[("Arquivos PDF", "*.pdf")],
    )
    if caminho:
      self.entry_pdf.delete(0, tk.END)
      self.entry_pdf.insert(0, caminho)
      self.lbl_status.config(
          text=f"Pronto para converter: {os.path.basename(caminho)}",
          fg="#1B365D",
      )

  def converter(self):
    caminho_pdf = self.entry_pdf.get().strip()

    if not caminho_pdf or not os.path.exists(caminho_pdf):
      messagebox.showwarning(
          "Atenção", "Por favor, selecione um arquivo PDF válido."
      )
      return

    try:
      self.lbl_status.config(text="Processando relatório...", fg="#D97706")
      self.root.update()

      df = processar_pdf_siga(caminho_pdf)

      if df.empty:
        messagebox.showerror(
            "Erro", "Não foi possível extrair registros válidos deste PDF."
        )
        self.lbl_status.config(text="Falha na extração.", fg="#DC2626")
        return

      nome_padrao = (
          os.path.splitext(os.path.basename(caminho_pdf))[0] + ".xlsx"
      )
      pasta_padrao = os.path.dirname(caminho_pdf)

      caminho_salvar = filedialog.asksaveasfilename(
          title="Salvar Planilha Excel Como",
          initialdir=pasta_padrao,
          initialfile=nome_padrao,
          defaultextension=".xlsx",
          filetypes=[("Planilha Excel", "*.xlsx")],
      )

      if caminho_salvar:
        df.to_excel(caminho_salvar, index=False)
        self.lbl_status.config(
            text=f"Concluído! {len(df)} registros salvos.", fg="#16A34A"
        )
        messagebox.showinfo(
            "Sucesso",
            f"Planilha Excel gerada com sucesso!\n\n"
            f"• Registros processados: {len(df)}\n"
            f"• Colunas: CLIENTE, CEP, ENDEREÇO, ST_TEC, QTDE, MARCA, MÊS EXEC.\n\n"
            f"Salvo em:\n{caminho_salvar}",
        )
      else:
        self.lbl_status.config(text="Operação cancelada.", fg="#888888")

    except Exception as e:
      messagebox.showerror("Erro ao Processar", f"Ocorreu um erro:\n{str(e)}")
      self.lbl_status.config(text="Erro durante o processamento.", fg="#DC2626")


if __name__ == "__main__":
  janela = tk.Tk()
  app = AppConversor(janela)
  janela.mainloop()