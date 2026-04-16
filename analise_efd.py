import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os

# ==========================================
# CONFIGURAÇÃO VISUAL DO TEMA
# ==========================================
ctk.set_appearance_mode("Dark")  
ctk.set_default_color_theme("blue")  

class EFDAnalyzerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Suíte Master EFD-Contribuições - V3")
        self.geometry("650x500")
        self.resizable(False, False)

        self.caminho_txt_in = ""
        self.caminho_excel_out = ""
        self.caminho_excel_in = ""
        self.caminho_txt_out = ""

        self.create_widgets()

    def create_widgets(self):
        self.lbl_titulo = ctk.CTkLabel(self, text="Suíte de Manipulação EFD-Contribuições", font=ctk.CTkFont(size=20, weight="bold"))
        self.lbl_titulo.pack(pady=(15, 5))

        self.tabview = ctk.CTkTabview(self, width=600, height=250)
        self.tabview.pack(padx=20, pady=5)

        self.aba_auditor = self.tabview.add("🔍 Auditor (TXT -> Excel)")
        self.aba_conversor = self.tabview.add("⚙️ Conversor (Excel -> TXT)")

        # ------------------------------------------
        # ABA 1: AUDITOR
        # ------------------------------------------
        self.btn_selecionar_txt = ctk.CTkButton(self.aba_auditor, text="1. Selecionar SPED (.txt)", command=self.selecionar_txt_auditor)
        self.btn_selecionar_txt.grid(row=0, column=0, padx=20, pady=20)

        self.lbl_txt_path = ctk.CTkLabel(self.aba_auditor, text="Nenhum arquivo...", text_color="gray", width=350, anchor="w")
        self.lbl_txt_path.grid(row=0, column=1, padx=10)

        self.btn_salvar_excel = ctk.CTkButton(self.aba_auditor, text="2. Salvar Relatório (.xlsx)", command=self.selecionar_excel_auditor)
        self.btn_salvar_excel.grid(row=1, column=0, padx=20, pady=20)

        self.lbl_excel_path = ctk.CTkLabel(self.aba_auditor, text="Nenhum destino...", text_color="gray", width=350, anchor="w")
        self.lbl_excel_path.grid(row=1, column=1, padx=10)

        self.btn_analisar = ctk.CTkButton(self.aba_auditor, text="🚀 Iniciar Auditoria", command=self.iniciar_analise_thread, state="disabled", fg_color="green", hover_color="darkgreen", width=200)
        self.btn_analisar.grid(row=2, column=0, columnspan=2, pady=10)

        # ------------------------------------------
        # ABA 2: CONVERSOR
        # ------------------------------------------
        self.lbl_aviso_conversor = ctk.CTkLabel(self.aba_conversor, text="Transforma planilhas em blocos SPED (Restaura ordem exata)", text_color="orange", font=ctk.CTkFont(size=12, slant="italic"))
        self.lbl_aviso_conversor.grid(row=0, column=0, columnspan=2, pady=(5, 10))

        self.btn_selecionar_excel_in = ctk.CTkButton(self.aba_conversor, text="1. Selecionar Excel (.xlsx)", command=self.selecionar_excel_conversor)
        self.btn_selecionar_excel_in.grid(row=1, column=0, padx=20, pady=15)

        self.lbl_excel_in_path = ctk.CTkLabel(self.aba_conversor, text="Nenhum arquivo...", text_color="gray", width=350, anchor="w")
        self.lbl_excel_in_path.grid(row=1, column=1, padx=10)

        self.btn_salvar_txt = ctk.CTkButton(self.aba_conversor, text="2. Salvar SPED (.txt)", command=self.selecionar_txt_conversor)
        self.btn_salvar_txt.grid(row=2, column=0, padx=20, pady=15)

        self.lbl_txt_out_path = ctk.CTkLabel(self.aba_conversor, text="Nenhum destino...", text_color="gray", width=350, anchor="w")
        self.lbl_txt_out_path.grid(row=2, column=1, padx=10)

        self.btn_converter = ctk.CTkButton(self.aba_conversor, text="⚙️ Iniciar Conversão", command=self.iniciar_conversao_thread, state="disabled", fg_color="#b55e00", hover_color="#8f4a00", width=200)
        self.btn_converter.grid(row=3, column=0, columnspan=2, pady=10)

        # ------------------------------------------
        # TERMINAL DE LOG
        # ------------------------------------------
        self.log_box = ctk.CTkTextbox(self, width=600, height=120, state="disabled")
        self.log_box.pack(pady=10, padx=20)
        self.log("Sistema iniciado. Novo sistema de travas de colunas ativado!")

    def log(self, mensagem):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", mensagem + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    # ==========================================
    # FUNÇÕES DE INTERFACE
    # ==========================================
    def selecionar_txt_auditor(self):
        caminho = filedialog.askopenfilename(title="Selecione o arquivo SPED", filetypes=[("TXT", "*.txt")])
        if caminho:
            self.caminho_txt_in = caminho
            self.lbl_txt_path.configure(text=os.path.basename(caminho), text_color="white")
            if self.caminho_txt_in and self.caminho_excel_out: self.btn_analisar.configure(state="normal")

    def selecionar_excel_auditor(self):
        caminho = filedialog.asksaveasfilename(title="Onde salvar o Excel?", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if caminho:
            self.caminho_excel_out = caminho
            self.lbl_excel_path.configure(text=os.path.basename(caminho), text_color="white")
            if self.caminho_txt_in and self.caminho_excel_out: self.btn_analisar.configure(state="normal")

    def iniciar_analise_thread(self):
        self.btn_analisar.configure(state="disabled")
        threading.Thread(target=self.processar_efd).start()

    def selecionar_excel_conversor(self):
        caminho = filedialog.askopenfilename(title="Selecione o Excel", filetypes=[("Excel", "*.xlsx")])
        if caminho:
            self.caminho_excel_in = caminho
            self.lbl_excel_in_path.configure(text=os.path.basename(caminho), text_color="white")
            if self.caminho_excel_in and self.caminho_txt_out: self.btn_converter.configure(state="normal")

    def selecionar_txt_conversor(self):
        caminho = filedialog.asksaveasfilename(title="Onde salvar o SPED?", defaultextension=".txt", filetypes=[("TXT", "*.txt")])
        if caminho:
            self.caminho_txt_out = caminho
            self.lbl_txt_out_path.configure(text=os.path.basename(caminho), text_color="white")
            if self.caminho_excel_in and self.caminho_txt_out: self.btn_converter.configure(state="normal")

    def iniciar_conversao_thread(self):
        self.btn_converter.configure(state="disabled")
        threading.Thread(target=self.processar_excel_para_txt).start()

    # ==========================================
    # 🔍 MODO AUDITOR (Cria Travas e Preserva Ordem)
    # ==========================================
    def processar_efd(self):
        self.log("\n[AUDITOR] Lendo arquivo TXT e rastreando hierarquias...")
        dados_por_registro = {}
        erros_encontrados = []
        
        operacao_atual, situacao_atual, documento_atual = None, None, '-'
        csts_entrada = ['50', '51', '52', '53', '54', '55', '56', '60', '61', '62', '63', '64', '65', '66', '70', '71', '72', '73', '74', '75', '98', '99']
        csts_credito = [str(i) for i in range(50, 67)]

        try:
            with open(self.caminho_txt_in, 'r', encoding='latin-1') as file:
                linhas = file.readlines()

            for num_linha, linha in enumerate(linhas, start=1):
                linha_str = linha.strip('\n').strip('\r')
                if not linha_str.startswith('|'): continue
                
                # Preserva EXATAMENTE todos os campos e colunas vazias
                campos = linha_str.split('|')
                if len(campos) < 3: continue 
                
                registro = campos[1] 
                
                # ADICIONA AS TRAVAS DE SEGURANÇA NA LINHA
                linha_excel = [num_linha] + campos + ['[FIM]']
                
                if registro not in dados_por_registro: dados_por_registro[registro] = []
                dados_por_registro[registro].append(linha_excel)

                # ================== VALIDAÇÕES ==================
                if len(campos) > 8 and registro in ['C100', 'D100']:
                    operacao_atual, situacao_atual, documento_atual = campos[2], campos[6], campos[8]
                
                if registro == 'C170' and situacao_atual in ['00', '01']:
                    cst_pis = campos[26] if len(campos) > 26 else ''
                    cst_cofins = campos[32] if len(campos) > 32 else ''
                    conta_contabil = campos[38] if len(campos) > 38 else ''

                    if cst_pis in csts_credito or cst_cofins in csts_credito:
                        if not conta_contabil.strip():
                            erros_encontrados.append({'Linha TXT': num_linha, 'Documento': documento_atual, 'Registro': registro, 'Tipo de Erro': 'Conta Ausente', 'Detalhe': f'CST {cst_pis}/{cst_cofins} exige conta no 0500.'})
                    if operacao_atual == '0':
                        if cst_pis and cst_pis not in csts_entrada:
                            erros_encontrados.append({'Linha TXT': num_linha, 'Documento': documento_atual, 'Registro': registro, 'Tipo de Erro': 'CST PIS Inválido', 'Detalhe': cst_pis})
                        if cst_cofins and cst_cofins not in csts_entrada:
                            erros_encontrados.append({'Linha TXT': num_linha, 'Documento': documento_atual, 'Registro': registro, 'Tipo de Erro': 'CST COF Inválido', 'Detalhe': cst_cofins})

                if registro == 'F120':
                    conta_contabil_f120 = campos[16] if len(campos) > 16 else ''
                    if not conta_contabil_f120.strip():
                        erros_encontrados.append({'Linha TXT': num_linha, 'Documento': 'Depreciação', 'Registro': 'F120', 'Tipo de Erro': 'Conta Ausente no F120', 'Detalhe': 'Campo 16 vazio.'})

            self.log("[AUDITOR] Montando abas seguras no Excel...")
            with pd.ExcelWriter(self.caminho_excel_out, engine='xlsxwriter') as writer:
                if erros_encontrados:
                    df_erros = pd.DataFrame(erros_encontrados)
                    df_erros.groupby('Tipo de Erro').size().reset_index(name='Qtd').to_excel(writer, sheet_name='📊 RESUMO ERROS', index=False)
                    df_erros.to_excel(writer, sheet_name='🚨 RELATÓRIO DETALHADO', index=False)
                
                for reg in sorted(dados_por_registro.keys()):
                    linhas_reg = dados_por_registro[reg]
                    max_len = max(len(l) for l in linhas_reg)
                    
                    # Normaliza o tamanho preenchendo vazios antes da trava [FIM]
                    for l in linhas_reg:
                        while len(l) < max_len:
                            l.insert(-1, '')
                            
                    df = pd.DataFrame(linhas_reg)
                    
                    # Cria Cabeçalhos Estruturais
                    headers = ["ID_LINHA", "INICIO", "REGISTRO"]
                    for i in range(max_len - 4):
                        headers.append(f"CAMPO_{i+1:02d}")
                    headers.append("FIM_LINHA")
                    
                    df.columns = headers
                    df.to_excel(writer, sheet_name=f'Reg {reg}', index=False)
                    
                    # Ajuste de largura das colunas
                    worksheet = writer.sheets[f'Reg {reg}']
                    worksheet.set_column('A:A', 10) 
                    worksheet.set_column('B:B', 5)  
                    worksheet.set_column('C:C', 12) 
                    worksheet.set_column('D:Z', 18) 

            self.log("✅ [AUDITOR] Excel protegido contra perda de dados gerado!")
            messagebox.showinfo("Sucesso", "Auditoria concluída!\nO Excel gerado agora preserva as colunas vazias!")
        except Exception as e:
            self.log(f"❌ ERRO: {e}")
        finally:
            self.btn_analisar.configure(state="normal")

    # ==========================================
    # ⚙️ MODO CONVERSOR (Restabelece a Estrutura Original)
    # ==========================================
    def processar_excel_para_txt(self):
        self.log("\n[CONVERSOR] Extraindo blocos do Excel e desfazendo travas...")
        try:
            excel_data = pd.read_excel(self.caminho_excel_in, sheet_name=None, dtype=str)
            todas_as_linhas = []
            
            for sheet_name, df in excel_data.items():
                if sheet_name in ['📊 RESUMO ERROS', '🚨 RELATÓRIO DETALHADO']: continue
                
                if 'ID_LINHA' not in df.columns or 'FIM_LINHA' not in df.columns:
                    raise Exception(f"A aba '{sheet_name}' é de uma versão antiga ou inválida.\nUse o Auditor V3 para gerar o Excel primeiro.")
                
                df = df.fillna('')
                self.log(f"  > Mapeando aba: {sheet_name}...")
                
                col_inicio = df.columns.get_loc("INICIO")
                col_fim = df.columns.get_loc("FIM_LINHA")
                
                for index, row in df.iterrows():
                    # 1. Pega a linha original (permite decimais para inserção manual)
                    try:
                        id_linha = float(str(row['ID_LINHA']).replace(',', '.'))
                    except ValueError:
                        id_linha = 999999999 
                    
                    # 2. Pega EXATAMENTE as colunas entre INICIO e FIM_LINHA
                    campos_sped = row.values[col_inicio:col_fim]
                    
                    # 3. Reconstrói a linha perfeita
                    linha_sped = "|".join(campos_sped) + "\n"
                    todas_as_linhas.append((id_linha, linha_sped))

            self.log("[CONVERSOR] Aplicando ordenação Pai/Filho pelo ID...")
            todas_as_linhas.sort(key=lambda x: x[0])

            with open(self.caminho_txt_out, 'w', encoding='latin-1') as txt_file:
                for _, linha in todas_as_linhas:
                    txt_file.write(linha)

            self.log("✅ [CONVERSOR] Arquivo SPED perfeitamente costurado!")
            messagebox.showinfo("Sucesso", "TXT convertido com sucesso!\nO Validador vai ler a hierarquia perfeitamente.")

        except Exception as e:
            self.log(f"❌ ERRO NA CONVERSÃO: {e}")
            messagebox.showerror("Erro", f"Não foi possível converter:\n{e}")
        finally:
            self.btn_converter.configure(state="normal")

if __name__ == "__main__":
    app = EFDAnalyzerApp()
    app.mainloop()