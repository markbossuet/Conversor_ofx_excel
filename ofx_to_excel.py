import ofxparse
import pandas as pd
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilenames, askdirectory
import sys
import chardet
from xml.etree import ElementTree
import re
import os
import time
import tempfile

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

def main():
    root = Tk()
    root.withdraw()

    try:
        print("Selecione os arquivos OFX...")
        ofx_files = askopenfilenames(
            title="Selecione os arquivos OFX",
            filetypes=[("Arquivos OFX", "*.ofx"), ("Todos os Arquivos", "*.*")]
        )

        if not ofx_files:
            print("Nenhum arquivo selecionado. Saindo...")
            return

        print(f"{len(ofx_files)} arquivo(s) OFX selecionado(s).")
        print("Selecione o diretório onde salvar os arquivos Excel...")
        output_dir = askdirectory(title="Selecione o diretório para salvar os arquivos Excel")

        if not output_dir:
            print("Nenhum diretório selecionado. Saindo...")
            return

        for ofx_file in ofx_files:
            try:
                file_name = os.path.basename(ofx_file)
                excel_name = os.path.splitext(file_name)[0] + ".xlsx"
                output_excel = os.path.join(output_dir, excel_name)

                print(f"Processando: {file_name}")
                ofx_to_excel_format(ofx_file, output_excel)
            except Exception as e:
                print(f"Erro ao processar {file_name}: {e}")

        print("Processamento concluído para todos os arquivos!")

    finally:
        root.destroy()

if __name__ == "__main__":
    main()
