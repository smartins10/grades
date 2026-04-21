import flet as ft
import pandas as pd
from playwright.sync_api import sync_playwright
import threading
import subprocess
import sys
import os

def check_and_install_chromium():
    """
    Verifica e instala o Chromium do Playwright, caso não esteja instalado.
    """
    try:
        # Tenta rodar a instalação do Chromium silenciosamente
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
    except Exception as e:
        print(f"Erro ao tentar instalar o Chromium: {e}")

class ESchoolingAutomation:
    def __init__(self, page_ui: ft.Page):
        self.page_ui = page_ui
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        
        self.excel_data = None
        self.should_process = False
        
        # Removido o flet FilePicker por questões de compatibilidade com versões antigas
        
    def __init__(self, page_ui: ft.Page):
        self.page_ui = page_ui
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        
        self.excel_data = None
        self.should_process = False
        
        self.btn_select_file = ft.ElevatedButton(
            "Selecionar Grelha Excel",
            on_click=self.on_file_selected
        )
        self.file_label = ft.Text("Nenhum arquivo selecionado")
        
        self.btn_open_portal = ft.ElevatedButton(
            "1. Abrir Portal eSchooling", 
            on_click=self.open_portal
        )
        
        self.btn_start_input = ft.ElevatedButton(
            "2. Iniciar Lançamento", 
            disabled=True,
            on_click=self.start_input
        )
        
        self.log_view = ft.ListView(expand=1, spacing=5, auto_scroll=True)
        
        self.setup_ui()

    def setup_ui(self):
        self.page_ui.title = "eSchooling - Lançamento Automático"
        
        try:
            self.page_ui.window.width = 600
            self.page_ui.window.height = 650
            self.page_ui.window.position = ft.WindowPosition.CENTER
    
        except:
            pass

        self.page_ui.add(
            ft.Text("eSchooling Lançamento Automático", size=24, color="blue"),
            ft.Text("Automação profissional de lançamento de notas", size=14, color="grey"),
            
            ft.Divider(),
            
            ft.Text("Passo 1: Seleção da Grelha", size=18, color="blue"),
            ft.Text("A grelha (.xlsx) precisa conter 'Nome' e 'Nota' (ou 'Notas'). Opcional: 'Saber Fazer', 'Saber Ser'.", size=12, color="grey"),
            self.btn_select_file,
            self.file_label,
            
            ft.Divider(),
            
            ft.Text("Passo 2: Execução da Automação", size=18, color="blue"),
            ft.Text("Abra o portal, faça login, acesse a janela de notas e inicie o robô.", size=12, color="grey"),
            self.btn_open_portal,
            self.btn_start_input,
            
            ft.Divider(),
            
            ft.Text("Terminal de Logs", size=18, color="blue"),
            self.log_view
        )
        self.log("ℹ️ Sistema carregado. Ambiente pronto para operação.")

    def log(self, message):
        """Adiciona uma mensagem ao painel de log."""
        self.log_view.controls.append(ft.Text(message, size=13))
        self.page_ui.update()

    def on_file_selected(self, e):
        """Lida com a seleção do arquivo Excel usando Tkinter."""
        import tkinter as tk
        from tkinter import filedialog
        
        # Cria uma janela oculta do Tkinter apenas para o popup nativo
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        
        file_path = filedialog.askopenfilename(
            title="Selecione a grelha Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        root.destroy()
        
        if file_path:
            import os
            self.file_label.value = os.path.basename(file_path)
            try:
                # Ler todas as folhas para procurar a correta (ex: 'Lançamento')
                all_sheets = pd.read_excel(file_path, sheet_name=None)
                
                valid_df = None
                found_sheet_name = ""
                
                for sheet_name, df in all_sheets.items():
                    has_nome = False
                    has_nota = False
                    for col in df.columns:
                        col_norm = str(col).strip().lower()
                        if col_norm in ["nome", "aluno", "nome do aluno", "nomes"]:
                            has_nome = True
                        if col_norm in ["nota", "notas", "nota final"]:
                            has_nota = True
                            
                    if has_nome and has_nota:
                        valid_df = df
                        found_sheet_name = sheet_name
                        break
                
                if valid_df is not None:
                    self.excel_data = valid_df
                    self.log(f"✅ Arquivo carregado (Folha '{found_sheet_name}'): {len(self.excel_data)} registros encontrados.")
                    self.file_label.color = "green"
                else:
                    self.log(f"❌ ERRO: Nenhuma folha na grelha contém as colunas necessárias ('Nome' e 'Nota').")
                    self.excel_data = None
                    self.file_label.color = "red"
                    
            except Exception as ex:
                self.log(f"❌ Erro ao ler o arquivo: {str(ex)}")
            self.page_ui.update()

    def open_portal(self, e):
        """Inicia a thread que rodará o Playwright."""
        # Previne cliques múltiplos
        self.btn_open_portal.disabled = True
        self.page_ui.update()
        threading.Thread(target=self._playwright_thread, daemon=True).start()

    def start_input(self, e):
        """Ativa a flag para que a thread do Playwright inicie o processamento."""
        if self.excel_data is None:
            self.log("⚠️ ERRO: Nenhum dado de Excel válido foi carregado.")
            return
        
        self.btn_start_input.disabled = True
        self.page_ui.update()
        self.should_process = True  # Sinaliza a thread do Playwright

    def _playwright_thread(self):
        """Thread principal de automação que mantém o Playwright ativo e responsivo."""
        self.log("⚙️ Verificando e Instalando dependências do navegador (pode demorar na 1ª vez)...")
        check_and_install_chromium()
        
        self.log("🌐 Abrindo o navegador...")
        
        try:
            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.launch(headless=False)
            self.context = self.browser.new_context()
            self.page = self.context.new_page()
            
            self.log("🚀 Acessando https://cmc.eschoolingserver.com/Login.aspx...")
            self.page.goto("https://cmc.eschoolingserver.com/Login.aspx")
            
            self.log("✅ Portal aberto. Faça login e abra o popup de lançamento de notas.")
            self.log("👉 Quando o popup de notas estiver aberto, clique no botão '2. Iniciar Lançamento'.")
            
            self.btn_start_input.disabled = False
            self.page_ui.update()
            
            # Loop de evento do Playwright (mantém a janela viva e monitora a flag de processo)
            while not self.page.is_closed():
                self.page.wait_for_timeout(500) # Cede tempo para não travar a thread
                
                if self.should_process:
                    self.should_process = False
                    self._execute_grades_processing()
                    self.btn_start_input.disabled = False
                    self.page_ui.update()
                    
        except Exception as ex:
            if "Target closed" not in str(ex) and "Browser closed" not in str(ex):
                self.log(f"❌ Ocorreu um erro no navegador: {str(ex)}")
        finally:
            self.log("🚪 Navegador encerrado.")
            if self.playwright:
                try:
                    self.playwright.stop()
                except:
                    pass
            self.btn_open_portal.disabled = False
            self.btn_start_input.disabled = True
            self.page_ui.update()

    def _execute_grades_processing(self):
        """Lógica real de preenchimento de notas, roda na thread do Playwright."""
        self.log("⏳ Localizando a janela de lançamento (popup)...")
        
        popup_page = None
        # Procurar páginas recém abertas pelo contexto
        for p in self.context.pages:
            if "EditCourseEvaluationProposal" in p.url:
                popup_page = p
                break
        
        # Fallback caso a URL seja diferente ou seja a única janela
        if not popup_page:
            if len(self.context.pages) > 1:
                popup_page = self.context.pages[-1]
            else:
                popup_page = self.page
                
        try:
            title = popup_page.title()
            self.log(f"🎯 Janela detectada: {title}")
        except:
            self.log("🎯 Janela detectada.")
            
        self.log("▶️ Iniciando o lançamento de notas...")
        
        success_count = 0
        error_count = 0
        
        for index, row in self.excel_data.iterrows():
            nome = ""
            nota = ""
            saber_fazer = ""
            saber_ser = ""
            
            # Busca de colunas resiliente a espaços e diferenças de maiúsculas/minúsculas
            for col in row.index:
                col_norm = str(col).strip().lower()
                val = row[col]
                
                if pd.isna(val):
                    continue
                    
                val_str = str(val).strip()
                if not val_str or val_str.lower() == 'nan':
                    continue
                    
                if col_norm in ["nome", "aluno", "nome do aluno", "nomes"]:
                    nome = val_str
                elif col_norm in ["nota", "notas", "nota final"]:
                    nota = val_str
                    if nota.endswith('.0'):
                        nota = nota[:-2]
                elif "saber fazer" in col_norm:
                    saber_fazer = val_str
                elif "saber ser" in col_norm or "saber estar" in col_norm:
                    saber_ser = val_str
                    
            if not nome: 
                continue
                
            if not nota and not saber_fazer and not saber_ser:
                continue

            msg = f"Processando: {nome}"
            if nota: msg += f" | Nota: {nota}"
            if saber_fazer: msg += f" | Saber Fazer: {saber_fazer}"
            if saber_ser: msg += f" | Saber Ser: {saber_ser}"
            self.log(msg)
            
            try:
                # 1. Localização do Aluno: localizar a célula <td> que contém o nome exato
                td_locator = popup_page.locator(f"xpath=//td[normalize-space(text())=\"{nome}\"]")
                
                if td_locator.count() == 0:
                    self.log(f"  ❌ Falha: Nome '{nome}' não encontrado na tabela.")
                    error_count += 1
                    continue
                
                # 2. Navegação na Tabela: subir para a linha pai (<tr>) principal do aluno
                tr_locator = td_locator.first.locator("xpath=ancestor::tr[1]")
                
                # 3. Preenchimento: Nota Principal
                if nota:
                    select_locator = tr_locator.locator("css=select[id$='_ddFinalValue']")
                    if select_locator.count() > 0:
                        select_locator.first.select_option(label=nota)
                        self.log(f"  ✅ Nota {nota} lançada com sucesso.")
                    else:
                        self.log(f"  ⚠️ Aviso: Dropdown de nota principal não encontrado.")
                
                # 4. Preenchimento: Saber Fazer
                if saber_fazer:
                    saber_fazer_upper = saber_fazer.upper()
                    sf_span = tr_locator.locator("span", has_text="saber fazer")
                    
                    if sf_span.count() > 0:
                        sf_select = sf_span.first.locator("xpath=ancestor::tr[1]//select")
                        if sf_select.count() > 0:
                            try:
                                sf_select.first.select_option(label=saber_fazer_upper)
                                self.log(f"  ✅ Saber Fazer '{saber_fazer_upper}' lançado com sucesso.")
                            except Exception as e:
                                self.log(f"  ❌ Erro ao selecionar '{saber_fazer_upper}' no Saber Fazer. Valor existe na lista?")
                        else:
                            self.log(f"  ⚠️ Aviso: Dropdown de 'Saber Fazer' não encontrado na linha.")
                    else:
                        self.log(f"  ⚠️ Aviso: Texto 'Saber fazer' não encontrado para este aluno.")
                        
                # 5. Preenchimento: Saber Ser
                if saber_ser:
                    saber_ser_upper = saber_ser.upper()
                    ss_span = tr_locator.locator("span", has_text="saber estar")
                    
                    if ss_span.count() > 0:
                        ss_select = ss_span.first.locator("xpath=ancestor::tr[1]//select")
                        if ss_select.count() > 0:
                            try:
                                ss_select.first.select_option(label=saber_ser_upper)
                                self.log(f"  ✅ Saber Ser '{saber_ser_upper}' lançado com sucesso.")
                            except Exception as e:
                                self.log(f"  ❌ Erro ao selecionar '{saber_ser_upper}' no Saber Ser. Valor existe na lista?")
                        else:
                            self.log(f"  ⚠️ Aviso: Dropdown de 'Saber Ser' não encontrado na linha.")
                    else:
                        self.log(f"  ⚠️ Aviso: Texto 'Saber estar' não encontrado para este aluno.")
                
                success_count += 1
                
                # Pequena pausa visual
                popup_page.wait_for_timeout(100)
                
            except Exception as e:
                self.log(f"  ❌ Erro geral ao lançar dados para {nome}: {str(e)}")
                error_count += 1
                
        self.log("--- 🏁 Concluído ---")
        self.log(f"📊 Resumo: {success_count} Sucessos | {error_count} Erros")


def main(page: ft.Page):
    app = ESchoolingAutomation(page)

if __name__ == "__main__":
    ft.app(target=main)
