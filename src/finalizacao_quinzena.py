from Sistema_Entrada_Dados import GestaoTaxasFixas

import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from openpyxl import load_workbook, Workbook
from datetime import datetime
from dateutil.relativedelta import relativedelta
import calendar
import os
import openpyxl

from config.utils import (
    validar_data,
    validar_data_quinzena,
    formatar_moeda,
    ARQUIVO_CLIENTES,
    ARQUIVO_CONTROLE,
    PASTA_CLIENTES,
    BASE_PATH
)

class ControleLancamentosTaxaADM:
    def __init__(self):
        self.planilha_controle = ARQUIVO_CONTROLE
        self.verificar_planilha_controle()

    def verificar_planilha_controle(self):
        """Cria a planilha de controle se não existir"""
        if not os.path.exists(self.planilha_controle):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = 'Controle'
            
            headers = ['Cliente', 'Data Referência', 'Data Lançamento', 'Valor', 'Status']
            for col, header in enumerate(headers, 1):
                ws.cell(row=1, column=col, value=header)
            
            wb.save(self.planilha_controle)

    def verificar_lancamento_existente(self, ws_contratos, cliente, data_ref):
        """Verifica se já existe lançamento na aba de controle"""
        data_str = data_ref.strftime("%d/%m/%Y")
        for row in ws_contratos.iter_rows(min_row=3, values_only=True):
            if (row[22] and  # Tem referência na coluna PARCELAS
                row[25] == data_str):  # Mesma data
                return True
        return False

    def registrar_lancamento(self, cliente, data_ref, valor):
        """Registra um novo lançamento de taxa"""
        try:
            print(f"\nRegistrando lançamento no controle...")
            print(f"Cliente: {cliente}")
            print(f"Data: {data_ref}")
            print(f"Valor: R$ {valor:.2f}")
            
            # Se a data for string, converter para datetime
            if isinstance(data_ref, str):
                data_ref = datetime.strptime(data_ref, '%d/%m/%Y')
            
            # Verificar se já existe
            existe, detalhes = self.verificar_lancamento_existente(cliente, data_ref)
            if existe:
                print("AVISO: Lançamento já existe no controle!")
                return False
                    
            
            wb = load_workbook(self.planilha_controle)
            ws = wb['Controle']
            
            # Adicionar novo registro
            proxima_linha = ws.max_row + 1
            
            
            ws.cell(row=proxima_linha, column=1, value=cliente)  # Cliente
            ws.cell(row=proxima_linha, column=2, value=data_ref)  # Data Referência
            data_atual = datetime.now()
            ws.cell(row=proxima_linha, column=3, value=data_atual)  # Data Lançamento
            ws.cell(row=proxima_linha, column=4, value=float(valor))  # Valor
            ws.cell(row=proxima_linha, column=5, value='LANÇADO')  # Status
            
            # Formatar células
            ws.cell(row=proxima_linha, column=2).number_format = 'DD/MM/YYYY'
            ws.cell(row=proxima_linha, column=3).number_format = 'DD/MM/YYYY'
            ws.cell(row=proxima_linha, column=4).number_format = '#,##0.00'
            
            
            wb.save(self.planilha_controle)
            print("Registro concluído com sucesso!")
            return True
                
        except Exception as e:
            print(f"Erro ao registrar lançamento no controle: {str(e)}")
            if 'wb' in locals():
                wb.close()
            raise Exception(f"Erro ao registrar lançamento no controle: {str(e)}")


class FinalizacaoQuinzena:
    def __init__(self, parent=None):
        self.parent = parent  # Guarda referência ao parent
        self.root = tk.Toplevel(parent) if parent else tk.Tk()
        self.root.title("Finalização de Quinzena")
        self.root.geometry("1000x600")
        self.gestao_taxas = GestaoTaxasFixas(self)
        
        # Variável de controle para logs
        self.calculando_final = tk.BooleanVar(value=False)
        
        # Inicializar controle de taxa ADM
        self.controle = ControleLancamentosTaxaADM()
        
        # Inicializar atributos que serão usados na interface
        self.data_ref_entry = None
        self.tree_clientes = None
        
        # Configurar a interface
        self.setup_gui()
        
        # Configurar o comportamento ao fechar a janela
        self.root.protocol("WM_DELETE_WINDOW", self.voltar_menu)

    def setup_gui(self):
        # Frame principal com padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill='both', expand=True)

        # Frame para seleção de data
        frame_data = ttk.LabelFrame(main_frame, text="Data de Referência")
        frame_data.pack(fill='x', pady=5)

        ttk.Label(frame_data, text="Data:").pack(side='left', padx=5)
        self.data_ref_entry = DateEntry(
            frame_data,
            format='dd/mm/yyyy',
            locale='pt_BR'
        )
        self.data_ref_entry.pack(side='left', padx=5)
        
        # Sempre adicionar o botão de carregar clientes
        ttk.Button(frame_data, 
                  text="Carregar Clientes",
                  command=self.carregar_clientes).pack(side='left', padx=5)

        # Frame para lista de clientes
        frame_lista = ttk.LabelFrame(main_frame, text="Clientes")
        frame_lista.pack(fill='both', expand=True, pady=5)

        # Treeview com todas as colunas necessárias
        self.tree_clientes = ttk.Treeview(frame_lista, 
                                        columns=('Cliente', 'Taxa ADM', 'Valor Taxa', 'Status'),
                                        show='headings')
        
        self.tree_clientes.heading('Cliente', text='Cliente')
        self.tree_clientes.heading('Taxa ADM', text='Taxa ADM (%)')
        self.tree_clientes.heading('Valor Taxa', text='Valor Taxa')
        self.tree_clientes.heading('Status', text='Status')
        
        self.tree_clientes.column('Cliente', width=300)
        self.tree_clientes.column('Taxa ADM', width=100)
        self.tree_clientes.column('Valor Taxa', width=150)
        self.tree_clientes.column('Status', width=150)
        
        self.tree_clientes.pack(fill='both', expand=True)

        # Frame para botões
        frame_botoes = ttk.Frame(main_frame)
        frame_botoes.pack(fill='x', pady=5)

        ttk.Button(frame_botoes, 
                  text="Processar Selecionados",
                  command=self.processar_clientes).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes,
                  text="Voltar ao Menu Principal",
                  command=self.voltar_menu).pack(side='right', padx=5)


    def validar_data(self, data_str):
        """Valida o formato da data"""
        return validar_data(data_str)

    def carregar_clientes(self):
        """Carrega os clientes com suas informações e status"""
        data_ref = self.data_ref_entry.get()
        if not self.validar_data(data_ref):
            messagebox.showerror("Erro", "Data inválida!")
            return

        try:
            # Limpar lista atual
            for item in self.tree_clientes.get_children():
                self.tree_clientes.delete(item)

            wb = load_workbook(ARQUIVO_CLIENTES)
            ws = wb['Clientes']
            clientes_processados = set()
            
            for nome_cliente in clientes_processados:
                arquivo_cliente = PASTA_CLIENTES / f"{nome_cliente}.xlsx"
                wb_cliente = load_workbook(arquivo_cliente)
                ws_contratos = wb_cliente['Contratos_ADM']
                
                # Buscar contratos ativos
                for row in ws_contratos.iter_rows(min_row=3, values_only=True):
                    if row[3] == 'ATIVO':  # Status do contrato
                        if row[9] == 'Fixo':  # Tipo Fixo
                            valor = f"R$ {float(row[10].replace(',', '.')):,.2f}"
                            tipo = "FIXO"
                        else:  # Percentual
                            valor = f"{float(row[10])}%"
                            tipo = "PERCENTUAL"
                            
                        # Verificar se já tem lançamento
                        tem_lancamento = self.verificar_lancamento_existente(
                            ws_contratos, nome_cliente, data_ref
                        )
                        
                        status = "JÁ LANÇADO" if tem_lancamento else "PENDENTE"
                        
                        self.tree_clientes.insert('', 'end', values=(
                            nome_cliente,
                            tipo,
                            valor,
                            status
                        ))
                
                wb_cliente.close()
            
            wb.close()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar clientes: {str(e)}")


    
        

    def calcular_taxa_adm(self, cliente, data_ref):
        """Calcula o valor da taxa de administração para o cliente"""
        try:
            print("\n=== Iniciando cálculo de taxa ADM ===")
            print(f"Cliente: {cliente}")
            print(f"Data Ref: {data_ref}")
            
            # Carregar arquivo do cliente
            wb = load_workbook(PASTA_CLIENTES / f"{cliente}.xlsx")
            ws_dados = wb["Dados"]
            
            # Carregar taxa de administração da planilha de clientes
            wb_clientes = load_workbook(ARQUIVO_CLIENTES)  # Correto: Abre o arquivo base de clientes
            ws_clientes = wb_clientes['Clientes']
            
            # Encontrar a taxa de administração total para o cliente
            taxa_adm_total = 0
            for row in ws_clientes.iter_rows(min_row=2, values_only=True):
                if row[0] == cliente and row[6]:  # Nome do cliente e tem percentual
                    percentual = float(row[6]) if row[6] else 0
                    taxa_adm_total += percentual / 100  # Converter percentual para decimal
            
            print(f"Taxa ADM total encontrada: {taxa_adm_total:.2%}")
            
            if taxa_adm_total == 0:
                wb.close()
                wb_clientes.close()
                return 0
                
            # Converter data de referência
            data_ref = datetime.strptime(data_ref, '%d/%m/%Y') if isinstance(data_ref, str) else data_ref
            print(f"\nData de referência convertida: {data_ref.strftime('%d/%m/%Y')}")
            
            # Calcular valor base
            valor_base = 0
            lancamentos_encontrados = 0
            
            for row in ws_dados.iter_rows(min_row=2, values_only=True):
                data_lancamento = row[0]
                if isinstance(data_lancamento, datetime):
                    data_str_planilha = data_lancamento.strftime('%d/%m/%Y')
                    data_str_ref = data_ref.strftime('%d/%m/%Y')
                    
                    if data_str_planilha == data_str_ref:
                        tipo_desp = row[1]
                        if isinstance(tipo_desp, (int, float)) and 1 <= tipo_desp <= 6:
                            valor = row[7]  # Coluna H com o valor
                            if valor:
                                valor_base += float(valor)
                                lancamentos_encontrados += 1

            print(f"Valor base calculado: R$ {valor_base:.2f}")
            
            valor_taxa = valor_base * taxa_adm_total
            print(f"\nValor final da taxa: R$ {valor_taxa:.2f}")
            
            wb.close()
            wb_clientes.close()
            return valor_taxa
                
        except Exception as e:
            print(f"\nERRO: {str(e)}")
            raise Exception(f"Erro ao calcular taxa: {str(e)}")

    def lancar_taxa_adm(self, cliente, data_ref, valor):
        """Lança a taxa de administração na planilha do cliente"""
        try:
            print(f"Lançando taxa ADM para {cliente}")
            # Obter dados do fornecedor
            cnpj_cpf, nome_fornecedor = self.obter_dados_fornecedor(cliente)
            if not cnpj_cpf or not nome_fornecedor:
                raise Exception("Dados do fornecedor não encontrados")


            wb = load_workbook(PASTA_CLIENTES / f"{cliente}.xlsx")
            ws = wb["Dados"]
            
            proxima_linha = ws.max_row + 1
            
            # Converter data de referência
            data = datetime.strptime(data_ref, '%d/%m/%Y')
            
            
            # Configurar células
            ws.cell(row=proxima_linha, column=1, value=data)  # Data
            ws.cell(row=proxima_linha, column=1).number_format = 'DD/MM/YYYY'
            
            ws.cell(row=proxima_linha, column=2, value=7)  # Tipo 7 para taxa ADM
            ws.cell(row=proxima_linha, column=3, value=cnpj_cpf)  # CNPJ/CPF
            ws.cell(row=proxima_linha, column=4, value=nome_fornecedor)  # Nome

            # Determinar quinzena
            if data.day == 5:
                referencia = f'ADM. OBRA REF. 1ª QUINZ. {data.strftime("%m/%Y")}'
            else:  # day == 20
                referencia = f'ADM. OBRA REF. 2ª QUINZ. {data.strftime("%m/%Y")}'
            

            ws.cell(row=proxima_linha, column=5, value=referencia)  # Referência
            
            ws.cell(row=proxima_linha, column=6, value=valor)  # Valor unitário
            ws.cell(row=proxima_linha, column=6).number_format = '#,##0.00'
            
            ws.cell(row=proxima_linha, column=7, value=1)  # Dias
            
            ws.cell(row=proxima_linha, column=8, value=valor)  # Valor total
            ws.cell(row=proxima_linha, column=8).number_format = '#,##0.00'
            
            # Data de vencimento (sempre dia 5 do mês seguinte)
            dt_vencto = (data.replace(day=1) + relativedelta(months=1)).replace(day=5)
            ws.cell(row=proxima_linha, column=9, value=dt_vencto)  # Data vencimento
            ws.cell(row=proxima_linha, column=9).number_format = 'DD/MM/YYYY'
            
            ws.cell(row=proxima_linha, column=10, value='ADM')  # Categoria
            ws.cell(row=proxima_linha, column=11, value='')  # Dados bancários
            ws.cell(row=proxima_linha, column=12, value='LANÇAMENTO AUTOMÁTICO')  # Observação
            

            wb.save(PASTA_CLIENTES / f"{cliente}.xlsx")
  
            
        except Exception as e:
            print(f"Erro durante o lançamento: {str(e)}")
            raise Exception(f"Erro ao lançar taxa: {str(e)}")


    def obter_dados_fornecedor(self, cliente):
        """Obtém CNPJ/CPF e nome do administrador principal da planilha de clientes"""
        try:
            
            wb = load_workbook(ARQUIVO_CLIENTES)
            ws = wb['Clientes']
            
            maior_percentual = 0
            dados_adm_principal = None
            
            # Procura por todos os administradores do cliente
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] == cliente and row[6]:  # Nome do cliente e tem percentual
                    percentual = float(row[6]) if row[6] else 0
                    
                    if percentual > maior_percentual:
                        maior_percentual = percentual
                        dados_adm_principal = {
                            'cnpj_cpf': row[4],  # CNPJ/CPF do administrador
                            'nome': row[5]        # Nome do administrador
                        }
            
            wb.close()
            
            if dados_adm_principal:
              
                return dados_adm_principal['cnpj_cpf'], dados_adm_principal['nome']
                
            
            return None, None
                
        except Exception as e:
            print(f"Erro ao obter dados do administrador: {str(e)}")
            raise Exception(f"Erro ao obter dados do administrador: {str(e)}")
        

    def processar_clientes(self):
        """Processa os clientes selecionados que ainda não têm lançamento"""
        selecionados = self.tree_clientes.selection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Selecione pelo menos um cliente!")
            return

        data_ref = datetime.strptime(self.data_ref_entry.get(), '%d/%m/%Y')
        processados = []
        ignorados = []

        for item in selecionados:
            valores = self.tree_clientes.item(item)['values']
            cliente = valores[0]
            tipo = valores[1]
            status = valores[3]

            if status == "JÁ LANÇADO":
                ignorados.append(f"{cliente} - Já possui lançamento no período")
                continue

            try:
                if tipo == "FIXO":
                    # Processar taxa fixa
                    if self.gestao_taxas.processar_lancamentos_fixos(cliente, data_ref):
                        processados.append(f"{cliente} - Taxa Fixa")
                    else:
                        ignorados.append(f"{cliente} - Erro ao processar taxa fixa")
                else:
                    # Processar taxa percentual
                    valor = self.calcular_taxa_adm(cliente, data_ref)
                    if valor > 0:
                        self.lancar_taxa_adm(cliente, data_ref, valor)
                        processados.append(f"{cliente} - Taxa Percentual")
                    else:
                        ignorados.append(f"{cliente} - Valor zero")

            except Exception as e:
                ignorados.append(f"{cliente} - {str(e)}")

        self.carregar_clientes()  # Atualiza a lista
        
        # Mostrar resultados
        if processados:
            mensagem = "Clientes processados com sucesso:\n" + "\n".join(processados)
            if ignorados:
                mensagem += "\n\nClientes ignorados ou com avisos:\n" + "\n".join(ignorados)
            messagebox.showinfo("Processamento Concluído", mensagem)
        else:
            mensagem = "Nenhum cliente processado."
            if ignorados:
                mensagem += "\n\nClientes ignorados ou com avisos:\n" + "\n".join(ignorados)
            messagebox.showwarning("Aviso", mensagem)

    def mostrar_resultado_processamento(self, processados, ignorados):
        """Mostra janela com resultado do processamento"""
        resultado = tk.Toplevel(self.root)
        resultado.title("Resultado do Processamento")
        resultado.geometry("600x400")
        
        # Frame para processados
        frame_proc = ttk.LabelFrame(resultado, text="Clientes Processados")
        frame_proc.pack(fill='both', expand=True, padx=5, pady=5)
        
        texto_proc = tk.Text(frame_proc, height=5)
        texto_proc.pack(fill='both', expand=True)
        for cliente in processados:
            texto_proc.insert(tk.END, f"{cliente}\n")
        texto_proc.config(state='disabled')
        
        # Frame para ignorados
        frame_ign = ttk.LabelFrame(resultado, text="Clientes Ignorados")
        frame_ign.pack(fill='both', expand=True, padx=5, pady=5)
        
        texto_ign = tk.Text(frame_ign, height=5)
        texto_ign.pack(fill='both', expand=True)
        for cliente in ignorados:
            texto_ign.insert(tk.END, f"{cliente}\n")
        texto_ign.config(state='disabled')
        
        # Botão fechar
        ttk.Button(resultado, 
                  text="Fechar", 
                  command=resultado.destroy).pack(pady=10)


    
    def voltar_menu(self):
        """Fecha a janela e retorna ao menu principal"""
        self.root.destroy()
        if self.parent:
            self.parent.deiconify()  # Mostra a janela principal
            self.parent.lift()  # Traz a janela principal para frente

        
##    def sair_sistema(self):
##        """Fecha a janela"""
##        if self.tree.get_children():
##            if messagebox.askyesno("Confirmação", 
##                                 "Existem lançamentos não confirmados. Deseja sair mesmo assim?"):
##                self.root.destroy()
##        else:
##            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = FinalizacaoQuinzena(root)
    root.mainloop()
