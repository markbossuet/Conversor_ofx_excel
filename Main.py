import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import os
import chardet
from xml.etree import ElementTree
import tempfile
import time

def convert_money_string_to_float(value_str):
    try:
        cleaned = re.sub(r'[^\d,.-]', '', value_str.strip())
        cleaned = cleaned.replace(',', '.')

        if cleaned.count('.') > 1:
            parts = cleaned.split('.')
            cleaned = ''.join(parts[:-1]) + '.' + parts[-1]

        return float(cleaned)
    except (ValueError, AttributeError):
        print(f"Erro ao converter valor monetário: '{value_str}'")
        return 0.0

def clean_ofx_file(ofx_file):
    try:
        with open(ofx_file, 'r', encoding='latin-1') as f:
            raw_content = f.readlines()

        start_index = next((i for i, line in enumerate(raw_content) if line.strip().startswith("<OFX>")), None)

        if start_index is None:
            raise ValueError("O conteúdo XML não foi encontrado no arquivo OFX.")

        cleaned_content = "".join(raw_content[start_index:])

        temp_file = tempfile.NamedTemporaryFile(mode='w', suffix='.ofx', encoding='latin-1', delete=False)
        try:
            temp_file.write(cleaned_content)
            return temp_file.name
        finally:
            temp_file.close()
    except Exception as e:
        raise Exception(f"Erro ao limpar arquivo OFX: {e}")

def verify_excel_saved(output_excel):

    for _ in range(5):
        if os.path.exists(output_excel):
            try:
                with open(output_excel, 'rb'):
                    return True
            except IOError:
                pass
        time.sleep(1)
    return False

def ofx_to_excel_format(ofx_file, output_excel):
    temp_file_path = None
    try:
        with open(ofx_file, 'rb') as f:
            raw_data = f.read()
            encoding = chardet.detect(raw_data)['encoding']

        temp_file_path = clean_ofx_file(ofx_file)

        with open(temp_file_path, 'r', encoding=encoding) as f:
            tree = ElementTree.parse(f)
            root = tree.getroot()

        transactions = []
        for transaction in root.findall(".//STMTTRN"):
            try:
                date = transaction.find("DTPOSTED").text[:8]
                memo = transaction.find("MEMO").text
                amount_str = transaction.find("TRNAMT").text
                amount = convert_money_string_to_float(amount_str)

                valor_positivo = amount if amount > 0 else None
                valor_negativo = abs(amount) if amount < 0 else None

                transactions.append({
                    'Data': pd.to_datetime(date, format='%Y%m%d').strftime('%d/%m/%Y'),
                    'Descrição': memo,
                    'Valor Positivo': valor_positivo,
                    'Valor Negativo': valor_negativo
                })
            except Exception as e:
                print(f"Erro ao processar transação: {e}")

        df = pd.DataFrame(transactions)

        if df.empty:
            raise ValueError("Nenhuma transação foi encontrada no arquivo OFX.")

        df['Valor Positivo'] = df['Valor Positivo'].map(lambda x: f'{x:.2f}' if pd.notnull(x) else '')
        df['Valor Negativo'] = df['Valor Negativo'].map(lambda x: f'{x:.2f}' if pd.notnull(x) else '')

        print(f"Salvando o arquivo Excel em: {output_excel}")

        for attempt in range(3):
            try:
                df.to_excel(output_excel, index=False, sheet_name='Transações')
                if verify_excel_saved(output_excel):
                    print("Arquivo Excel salvo com sucesso!")
                    return
            except Exception as e:
                print(f"Tentativa {attempt + 1} falhou: {e}")
                time.sleep(1)

        raise Exception("Não foi possível salvar o arquivo Excel após várias tentativas.")

    except Exception as e:
        print(f"Erro ao processar o arquivo OFX: {e}")
    finally:
        if temp_file_path and os.path.exists(temp_file_path):
            os.remove(temp_file_path)

def abrir_arquivo_extrato():
    global extrato_sequencias
    filepath = filedialog.askopenfilename(title="Selecione o arquivo do Extrato", filetypes=[("Excel/OFX files", "*.xlsx;*.ofx")])
    if filepath:
        try:
            if filepath.lower().endswith(".ofx"):
                output_excel = os.path.splitext(filepath)[0] + ".xlsx"
                ofx_to_excel_format(filepath, output_excel)
                filepath = output_excel

            df = pd.read_excel(filepath)
            extrato_sequencias = extrair_sequencias(df, "extrato")
            exibir_sequencias(extrato_sequencias, "extrato")
            if razao_sequencias:  
                comparar_valores()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o arquivo do extrato: {e}")

def abrir_arquivo_razao():
    global razao_sequencias
    filepath = filedialog.askopenfilename(title="Selecione o arquivo da Razão", filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        try:
            df = pd.read_excel(filepath)
            razao_sequencias = extrair_sequencias(df, "razao")
            exibir_sequencias(razao_sequencias, "razao")
            if extrato_sequencias:  # Se o arquivo extrato já foi carregado, comparar automaticamente
                comparar_valores()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir o arquivo da razão: {e}")

def selecionar_ofx():
    filepath = filedialog.askopenfilename(title="Selecione o arquivo OFX", filetypes=[("Arquivos OFX", "*.ofx")])
    if filepath:
        try:
            output_excel = os.path.splitext(filepath)[0] + ".xlsx"
            ofx_to_excel_format(filepath, output_excel)
            messagebox.showinfo("Sucesso", f"Arquivo convertido e salvo em: {output_excel}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o arquivo OFX: {e}")

def extrair_sequencias(df, origem):
    sequencias = []
    for coluna in df.columns:
        for valor in df[coluna]:
            if isinstance(valor, str):  # Verifica se o valor é uma string
                match = re.search(r':\s*(\d+)', valor)
                if match:
                    sequencia = match.group(1)
                    if len(sequencia) > 4:  # Ignora sequências de até 4 números
                        sequencias.append((sequencia, origem))
    return sequencias

def exibir_sequencias(sequencias, tipo):
    if tipo == "extrato":
        lista_extrato.delete(0, tk.END)
        for seq, _ in sequencias:
            lista_extrato.insert(tk.END, seq)
    elif tipo == "razao":
        lista_razao.delete(0, tk.END)
        for seq, _ in sequencias:
            lista_razao.insert(tk.END, seq)

def comparar_valores():
    extrato_numeros = {seq for seq, _ in extrato_sequencias}
    razao_numeros = {seq for seq, _ in razao_sequencias}

    iguais = extrato_numeros & razao_numeros
    diferentes_extrato = [(seq, origem) for seq, origem in extrato_sequencias if seq not in razao_numeros]
    diferentes_razao = [(seq, origem) for seq, origem in razao_sequencias if seq not in extrato_numeros]
    diferentes = diferentes_extrato + diferentes_razao

    if iguais:
        exibir_valores(lista_iguais, [(seq, "iguais") for seq in iguais], "iguais")
    else:
        lista_iguais.delete(0, tk.END)  # Limpar lista caso não haja valores iguais

    if diferentes:
        exibir_valores(lista_diferentes, diferentes, "diferentes")
    else:
        lista_diferentes.delete(0, tk.END)
        lista_diferentes.insert(tk.END, "Não tem")

    exibir_contagem(len(iguais), len(diferentes))

def exibir_valores(lista, valores, tipo):
    lista.delete(0, tk.END)
    for valor, origem in valores:
        if tipo == "diferentes":
            lista.insert(tk.END, f"{valor} ({'Extrato' if origem == 'extrato' else 'Razão'})")
        else:
            lista.insert(tk.END, valor)

def exibir_contagem(iguais_count, diferentes_count):
    label_contagem.config(text=f"Iguais: {iguais_count}\nDiferentes: {diferentes_count}")

extrato_sequencias = []
razao_sequencias = []

root = tk.Tk()
root.title("Ferramenta de Conversão e Comparação")
root.geometry("600x500")

titulo = tk.Label(root, text="Ferramenta OFX e Comparação", font=("Helvetica", 16, "bold"))
titulo.pack(pady=10)

frame_botoes = tk.Frame(root)
frame_botoes.pack(pady=10)

botao_ofx_to_excel = tk.Button(frame_botoes, text="Converter OFX para Excel", command=selecionar_ofx, width=25, height=2)
botao_ofx_to_excel.grid(row=0, column=0, padx=5, pady=5)

botao_selecionar_extrato = tk.Button(frame_botoes, text="Selecionar Extrato", command=abrir_arquivo_extrato, width=25, height=2)
botao_selecionar_extrato.grid(row=0, column=1, padx=5, pady=5)

botao_selecionar_razao = tk.Button(frame_botoes, text="Selecionar Razão", command=abrir_arquivo_razao, width=25, height=2)
botao_selecionar_razao.grid(row=0, column=2, padx=5, pady=5)

frame_comparacao = tk.Frame(root)
frame_comparacao.pack(pady=10)

lista_extrato = tk.Listbox(frame_comparacao, selectmode=tk.SINGLE, width=30, height=10)
lista_extrato.grid(row=0, column=0, padx=10)

lista_razao = tk.Listbox(frame_comparacao, selectmode=tk.SINGLE, width=30, height=10)
lista_razao.grid(row=0, column=1, padx=10)

frame_resultados = tk.Frame(root)
frame_resultados.pack(pady=10)

label_iguais = tk.Label(frame_resultados, text="Iguais:")
label_iguais.grid(row=0, column=0, sticky="w")

lista_iguais = tk.Listbox(frame_resultados, width=40, height=5)
lista_iguais.grid(row=1, column=0, padx=10, pady=5)

label_diferentes = tk.Label(frame_resultados, text="Diferentes:")
label_diferentes.grid(row=0, column=1, sticky="w")

lista_diferentes = tk.Listbox(frame_resultados, width=40, height=5)
lista_diferentes.grid(row=1, column=1, padx=10, pady=5)

label_contagem = tk.Label(root, text="Iguais: 0\nDiferentes: 0", font=("Helvetica", 12))
label_contagem.pack(pady=10)

root.mainloop()
