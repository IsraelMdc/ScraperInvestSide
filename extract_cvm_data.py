"""
Script para extração automatizada de demonstrativos contábeis da CVM (2019-2023)
e preenchimento das planilhas 'Planilha1' e 'Python' no arquivo Artigo.xlsx.
"""

import os
import io
import csv
import zipfile
import shutil
import requests
import openpyxl

CACHE_DIR = "cache_cvm"
EXCEL_PATH = "Artigo.xlsx"
BACKUP_PATH = "Artigo_backup.xlsx"

# Mapeamento oficial dos 30 tickers para seus respectivos códigos CVM
TICKERS_CVM = {
    'VALE3': ['004170'],
    'CSNA3': ['004030'],
    'USIM5': ['014320'],
    'GGBR4': ['003980'],
    'GOAU4': ['008656'],
    'PETR4': ['009512'],
    'PRIO3': ['022187'],
    'RECV3': ['025780'],
    'UGPA3': ['018465'],
    'VBBR3': ['024295'],
    'SUZB3': ['013986'],
    'KLBN11': ['012653'],
    'JBSS3': ['020575'],
    'BRFS3': ['016292'],
    'MRFG3': ['020788'],
    'ELET3': ['002437'],
    'CMIG4': ['002453'],
    'CPLE6': ['014311'],
    'EGIE3': ['017329'],
    'TRPL4': ['018376'],
    'WEGE3': ['005410'],
    'RADL3': ['005258'],
    'LREN3': ['008133'],
    'MGLU3': ['022470'],
    'NTCO3': ['024783', '019550'],
    'VIVO3': ['017671'],
    'TIMS3': ['024929', '017639'],
    'RENT3': ['019739'],
    'CCRO3': ['018821'],
    'MDIA3': ['020338'],
}

def ensure_cache():
    """Garante que os 5 arquivos ZIP da CVM estejam baixados no diretório de cache."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    base_url = "https://dados.cvm.gov.br/dados/CIA_ABERTA/DOC/DFP/DADOS/dfp_cia_aberta_{}.zip"
    for year in range(2019, 2024):
        p = os.path.join(CACHE_DIR, f"dfp_{year}.zip")
        if not os.path.exists(p):
            print(f"Baixando dados CVM para o ano {year}...")
            r = requests.get(base_url.format(year), timeout=60)
            r.raise_for_status()
            with open(p, "wb") as f:
                f.write(r.content)
            print(f"Arquivo CVM {year} salvo em {p}.")
        else:
            print(f"Arquivo CVM {year} já em cache.")

def load_cvm_database():
    """Carrega e indexa em memória todos os dados contábeis relevantes de 2019 a 2023."""
    print("Indexando demonstrativos contábeis da CVM...")
    cvm_db = {}
    for year in range(2019, 2024):
        p = os.path.join(CACHE_DIR, f"dfp_{year}.zip")
        z = zipfile.ZipFile(p)
        year_db = {}
        for prefix in ['BPA', 'BPP', 'DRE']:
            for suffix in ['con', 'ind']:
                fname = f"dfp_cia_aberta_{prefix}_{suffix}_{year}.csv"
                if fname not in z.namelist():
                    continue
                reader = csv.reader(io.StringIO(z.read(fname).decode('latin1')), delimiter=';')
                header = next(reader)
                idx_cvm = header.index('CD_CVM')
                idx_ordem = header.index('ORDEM_EXERC')
                idx_conta = header.index('CD_CONTA')
                idx_vl = header.index('VL_CONTA')
                idx_ds = header.index('DS_CONTA')

                for row in reader:
                    cvm = row[idx_cvm]
                    ordem = row[idx_ordem]
                    key = (prefix, suffix, cvm, ordem)
                    if key not in year_db:
                        year_db[key] = {}
                    conta = row[idx_conta]
                    vl = float(row[idx_vl])
                    ds = row[idx_ds]
                    year_db[key][conta] = (vl, ds)
        cvm_db[year] = year_db
    print("Indexação concluída com sucesso!")
    return cvm_db

def extract_metrics(ticker, year, cvm_db):
    """Extrai os 5 dados financeiros para o ticker e ano especificados."""
    cvm_list = TICKERS_CVM.get(ticker, [])
    y_data = cvm_db[year]

    # Para PetroRecôncavo em 2019 (IPO recente), busca coluna comparativa em 2020
    if ticker == 'RECV3' and year == 2019:
        y_data = cvm_db[2020]
        ordem_target = 'PENÚLTIMO'
    else:
        ordem_target = 'ÚLTIMO'

    for cvm in cvm_list:
        # Prioridade para demonstrativos consolidados (con), com fallback para individuais (ind)
        for suffix in ['con', 'ind']:
            key_bpa = ('BPA', suffix, cvm, ordem_target)
            key_bpp = ('BPP', suffix, cvm, ordem_target)
            key_dre = ('DRE', suffix, cvm, ordem_target)

            if key_bpa in y_data and key_bpp in y_data:
                bpa = y_data[key_bpa]
                bpp = y_data[key_bpp]
                dre = y_data.get(key_dre, {})

                # 1. Ativo Total (Conta 1)
                at = bpa.get('1', (None,))[0]
                if at is None:
                    for cd, (v, ds) in bpa.items():
                        if ds.strip().lower() == 'ativo total':
                            at = v
                            break

                # 2. Empréstimos e Financiamentos Circulante (Conta 2.01.04)
                ec = bpp.get('2.01.04', (None,))[0]
                if ec is None:
                    for cd, (v, ds) in bpp.items():
                        if cd.startswith('2.01') and 'empréstimos' in ds.lower():
                            ec = v
                            break
                if ec is None:
                    ec = 0.0

                # 3. Empréstimos e Financiamentos Não Circulante (Conta 2.02.01)
                enc = bpp.get('2.02.01', (None,))[0]
                if enc is None:
                    for cd, (v, ds) in bpp.items():
                        if cd.startswith('2.02') and 'empréstimos' in ds.lower():
                            enc = v
                            break
                if enc is None:
                    enc = 0.0

                # 4. Patrimônio Líquido (Conta 2.03)
                pl = bpp.get('2.03', (None,))[0]
                if pl is None:
                    for cd, (v, ds) in bpp.items():
                        if any(k in ds.lower() for k in ['patrimônio líquido consolidado', 'patrimônio líquido']):
                            pl = v
                            break

                # 5. Despesa Financeira (Conta 3.06.02)
                df = dre.get('3.06.02', (None,))[0]
                if df is None:
                    for cd, (v, ds) in dre.items():
                        if cd.startswith('3.06') and 'despesa' in ds.lower():
                            df = v
                            break
                if df is None:
                    df = 0.0

                # 6. Lucro Líquido (Conta 3.11 ou 3.09)
                ll = dre.get('3.11', (None,))[0]
                if ll is None or ll == 0.0:
                    ll = dre.get('3.09', (None,))[0]
                if ll is None:
                    for cd, (v, ds) in dre.items():
                        if 'lucro/prejuízo consolidado' in ds.lower() or 'lucro líquido' in ds.lower():
                            ll = v
                            break
                if ll is None:
                    ll = 0.0

                return {
                    'ativo_total': at,
                    'emp_circ': ec,
                    'emp_ncirc': enc,
                    'pl': pl,
                    'desp_fin': df,
                    'lucro_liq': ll,
                    'cvm': cvm,
                    'tipo': suffix
                }
    raise ValueError(f"Não foi possível encontrar dados para {ticker} ({year})")

def update_excel(cvm_db):
    """Atualiza o arquivo Artigo.xlsx com os dados extraídos e as fórmulas calculadas."""
    if not os.path.exists(BACKUP_PATH):
        print(f"Criando cópia de backup em {BACKUP_PATH}...")
        shutil.copyfile(EXCEL_PATH, BACKUP_PATH)

    print(f"Carregando {EXCEL_PATH} com openpyxl...")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws_p1 = wb['Planilha1']
    ws_py = wb['Python']

    current_ticker = None
    updated_rows = 0

    print("Preenchendo Planilha1...")
    for row_idx in range(3, 153):
        ticker_val = ws_p1.cell(row=row_idx, column=3).value
        if ticker_val and str(ticker_val).strip():
            current_ticker = str(ticker_val).strip()

        ano_val = ws_p1.cell(row=row_idx, column=4).value
        if ano_val is None:
            continue
        ano = int(ano_val)

        m = extract_metrics(current_ticker, ano, cvm_db)

        # Formatação dos valores:
        # Col F: Fórmula =Circulante+NãoCirculante
        emp_c_int = int(round(m['emp_circ']))
        emp_nc_int = int(round(m['emp_ncirc']))
        ws_p1.cell(row=row_idx, column=6).value = f"={emp_c_int}+{emp_nc_int}"

        # Col G: Despesa Financeira
        df_val = int(round(m['desp_fin']))
        ws_p1.cell(row=row_idx, column=7).value = df_val

        # Col H: Ativo Total
        at_val = int(round(m['ativo_total']))
        ws_p1.cell(row=row_idx, column=8).value = at_val

        # Col I: Patrimônio Líquido
        pl_val = int(round(m['pl']))
        ws_p1.cell(row=row_idx, column=9).value = pl_val

        # Col J: Lucro Líquido
        ll_val = int(round(m['lucro_liq']))
        ws_p1.cell(row=row_idx, column=10).value = ll_val

        # Fórmulas das Variáveis Calculadas (Colunas M, N, O, P):
        # M: Custo da Dívida = ABS(Despesa_Financeira) / Divida_Bruta
        ws_p1.cell(row=row_idx, column=13).value = f"=ABS(G{row_idx})/F{row_idx}"

        # N: ROA = Lucro_Liquido / Ativo_Total
        ws_p1.cell(row=row_idx, column=14).value = f"=J{row_idx}/H{row_idx}"

        # O: Alavancagem = Divida_Bruta / Patrimonio_Liquido
        ws_p1.cell(row=row_idx, column=15).value = f"=F{row_idx}/I{row_idx}"

        # P: Tam_Ativo = LN(Ativo_Total)
        ws_p1.cell(row=row_idx, column=16).value = f"=LN(H{row_idx})"

        updated_rows += 1

    print(f"Planilha1: {updated_rows} linhas preenchidas/atualizadas com sucesso.")

    print("Preenchendo aba Python (painel econométrico)...")
    # Aba Python possui 1 linha de cabeçalho e linhas 2 a 151
    for py_row in range(2, 152):
        p1_row = py_row + 1  # Linha correspondente em Planilha1 (3 a 152)

        # Col C: Setor_Intensivo -> =Planilha1!E{p1_row}
        ws_py.cell(row=py_row, column=3).value = f"=Planilha1!E{p1_row}"

        # Col D: Custo_Dívida -> =Planilha1!M{p1_row}
        ws_py.cell(row=py_row, column=4).value = f"=Planilha1!M{p1_row}"

        # Col E: Disc_Carbono -> =Planilha1!L{p1_row}
        ws_py.cell(row=py_row, column=5).value = f"=Planilha1!L{p1_row}"

        # Col F: Gov_ISE -> =Planilha1!K{p1_row}
        ws_py.cell(row=py_row, column=6).value = f"=Planilha1!K{p1_row}"

        # Col G: Tam_Ativo -> =Planilha1!P{p1_row}
        ws_py.cell(row=py_row, column=7).value = f"=Planilha1!P{p1_row}"

        # Col H: Alavancagem -> =Planilha1!O{p1_row}
        ws_py.cell(row=py_row, column=8).value = f"=Planilha1!O{p1_row}"

        # Col I: ROA -> =Planilha1!N{p1_row}
        ws_py.cell(row=py_row, column=9).value = f"=Planilha1!N{p1_row}"

    print("Salvando alterações em Artigo.xlsx...")
    wb.save(EXCEL_PATH)
    print("Artigo.xlsx salvo com sucesso!")

def main():
    print("=== INICIANDO EXTRAÇÃO DE DADOS FINANCEIROS CVM ===")
    ensure_cache()
    cvm_db = load_cvm_database()
    update_excel(cvm_db)
    print("=== PROCESSO CONCLUÍDO COM SUCESSO! ===")

if __name__ == '__main__':
    main()

