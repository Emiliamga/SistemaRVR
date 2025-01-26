import tkinter as tk
from tkinter import ttk, PhotoImage, messagebox
import os
import sys
import importlib
import types
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class SistemaPrincipal:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Sistema de Gestão Financeira")
        self.root.geometry("800x600")
        self.root.lift()

        self.setup_style()
        self.create_main_content()

    def setup_style(self):
        """Configura o estilo visual do aplicativo"""
        style = ttk.Style()
        style.configure('Menu.TFrame', background='#f0f0f0')
        style.configure('Card.TFrame', background='white')
        style.configure('CardTitle.TLabel', 
                       font=('Helvetica', 14, 'bold'),
                       background='white')
        style.configure('CardDesc.TLabel',
                       font=('Helvetica', 10),
                       background='white',
                       wraplength=300)
        style.configure('Action.TButton',
                       font=('Helvetica', 12),
                       padding=10)

    def create_main_content(self):
        """Cria o conteúdo principal da interface"""
        # Frame principal
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill="both", padx=20, pady=20)

        # Logo
        self.logo_path = resource_path("logo.png")
        self.logo = PhotoImage(file=self.logo_path)
        logo_label = ttk.Label(main_frame, image=self.logo)
        logo_label.pack(pady=10)

        # Título
        title_label = ttk.Label(
            main_frame,
            text="Sistema de Gestão Financeira",
            font=('Helvetica', 24, 'bold'),
            background='#f0f0f0'
        )
        title_label.pack(pady=(0, 30))

        # Grid para botões
        grid = ttk.Frame(main_frame)
        grid.pack(expand=True, pady=20)

        self.create_card(grid, "Entrada de Dados", "Cadastro e gestão de dados", 
                        self.abrir_entrada_dados, 0, 0)
        self.create_card(grid, "Finalizar Quinzena", "Processar Taxa de Administração", 
                        self.finalizar_quinzena, 0, 1)
        self.create_card(grid, "Geração de Relatórios", "Visualização de relatórios", 
                        self.abrir_relatorios, 0, 2)
        
        # Botão Sair
        sair_btn = ttk.Button(main_frame, text="Sair", command=self.sair_sistema)
        sair_btn.pack(pady=20)

    def create_card(self, parent, title, description, command, row, col):
        """Cria um card na interface"""
        card = ttk.Frame(parent, style='Card.TFrame')
        card.grid(row=row, column=col, padx=10, pady=10, sticky='nsew')
        
        title_label = ttk.Label(
            card,
            text=title,
            style='CardTitle.TLabel'
        )
        title_label.pack(pady=(20, 10), padx=20)

        desc_label = ttk.Label(
            card,
            text=description,
            style='CardDesc.TLabel'
        )
        desc_label.pack(pady=(0, 20), padx=20)

        button = ttk.Button(
            card,
            text="Acessar",
            style='Action.TButton',
            command=command
        )
        button.pack(pady=(0, 20))

    def reload_module(self, module_name):
        """
        Recarrega um módulo e retorna a versão atualizada
        Args:
            module_name (str): Nome do módulo a ser recarregado
        Returns:
            module: Módulo recarregado
        """
        try:
            # Remover todas as referências ao módulo e seus submódulos
            for key in list(sys.modules.keys()):
                if key == module_name or key.startswith(f"{module_name}."):
                    del sys.modules[key]
            
            # Importar o módulo novamente
            module = importlib.import_module(module_name)
            return module
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar módulo {module_name}: {str(e)}")
            return None

    def abrir_entrada_dados(self):
        """Abre o sistema de entrada de dados"""
        try:
            print("Iniciando abertura do sistema de entrada de dados...")  # Debug
            
            # Limpar todos os módulos relacionados
            modulos_para_limpar = ['Sistema_Entrada_Dados', 'finalizacao_quinzena']
            for mod in modulos_para_limpar:
                if mod in sys.modules:
                    print(f"Removendo módulo {mod}")  # Debug
                    del sys.modules[mod]

            # Recarrega o módulo Sistema_Entrada_Dados
            print("Carregando Sistema_Entrada_Dados...")  # Debug
            modulo = importlib.import_module('Sistema_Entrada_Dados')
            
            if not modulo:
                return

            self.root.attributes('-topmost', False)
            self.root.withdraw()
            
            print("Criando instância do SistemaEntradaDados...")  # Debug
            # Cria nova instância da classe atualizada, passando o root como parent
            app = modulo.SistemaEntradaDados(parent=self.root)
            
            # Configura a referência ao menu principal
            app.menu_principal = self.root
            
            print("Configurando janela...")  # Debug
            app.root.lift()
            app.root.focus_force()
            app.root.protocol("WM_DELETE_WINDOW", lambda: self.finalizar_sistema(app.root))
            
            print("Iniciando mainloop...")  # Debug
            app.root.mainloop()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir sistema de entrada de dados: {str(e)}")
            print(f"Erro detalhado: {str(e)}")  # Debug detalhado
            self.root.deiconify()

    def finalizar_quinzena(self):
        """Abre o sistema de finalização de quinzena"""
        try:
            # Recarrega o módulo de finalização de quinzena
            modulo = self.reload_module('finalizacao_quinzena')
            if not modulo:
                return

            self.root.attributes('-topmost', False)
            self.root.withdraw()
            
            # Cria nova instância da classe atualizada
            app = modulo.FinalizacaoQuinzena(parent=self.root)
            app.root.lift()
            app.root.focus_force()
            app.root.mainloop()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir finalização de quinzena: {str(e)}")
            self.root.deiconify()

            

    def abrir_relatorios(self):
        try:
            modulo = self.reload_module('relatorio_despesas_aprimorado')
            if not modulo:
                return

            self.root.attributes('-topmost', False)
            self.root.withdraw()
            
            # Criar nova janela Toplevel em vez de None
            relatorio_window = tk.Toplevel(self.root)
            relatorio_window.withdraw()  # Esconder inicialmente
            
            # Inicializar o RelatorioUI com a nova janela
            app = modulo.RelatorioUI(relatorio_window)
            app.menu_principal = self.root
            
            # Configurar protocolo de fechamento
            relatorio_window.protocol("WM_DELETE_WINDOW", 
                lambda: self.finalizar_sistema(relatorio_window))
            
            # Mostrar a janela e configurar foco
            relatorio_window.deiconify()
            relatorio_window.lift()
            relatorio_window.focus_force()
            
            # Não chamar mainloop aqui
            
        except Exception as e:
            messagebox.showerror("Erro", 
                f"Erro ao abrir sistema de relatórios: {str(e)}")
            self.root.deiconify()

    def sair_sistema(self):
        """Fecha o sistema após confirmação"""
        if messagebox.askyesno("Confirmar Saída", "Deseja realmente sair do sistema?"):
            self.root.destroy()

    def finalizar_sistema(self, janela):
        """Fecha a janela do sistema e mostra a janela principal"""
        janela.destroy()
        self.root.deiconify()
        self.root.lift()

    def run(self):
        """Inicia a execução do sistema"""
        self.root.mainloop()


if __name__ == '__main__':
    app = SistemaPrincipal()
    app.run()
