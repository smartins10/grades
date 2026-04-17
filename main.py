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
        
        self.btn_select_file = ft.ElevatedButton(
            "Selecionar Excel",
            on_click=self.on_file_selected
        )
        self.file_label = ft.Text("Nenhum arquivo selecionado")
        
        self.btn_open_portal = ft.ElevatedButton(
            "1. Abrir Portal", 
            on_click=self.open_portal
        )
        
        self.btn_start_input = ft.ElevatedButton(
            "2. Iniciar Lançamento", 
            on_click=self.start_input,
            disabled=True
        )
        
        self.log_view = ft.ListView()
        
        self.setup_ui()

    def setup_ui(self):
        self.page_ui.title = "eSchooling - Automação de Notas"
        try:
            self.page_ui.window.width = 600
            self.page_ui.window.height = 600
        except:
            pass

        self.page_ui.add(
            ft.Text("eSchooling Lançamento Automático", size=20),
            ft.Divider(),
            self.btn_select_file,
            self.file_label,
            ft.Divider(),
            self.btn_open_portal,
            self.btn_start_input,
            ft.Divider(),
            ft.Text("Logs do Sistema"),
            self.log_view
        )
        self.log("ℹ️ Sistema iniciado. Por favor, selecione a planilha de notas (.xlsx).")

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
            title="Selecione a Planilha Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        root.destroy()
        
        if file_path:
            import os
            self.file_label.value = os.path.basename(file_path)
            try:
                self.excel_data = pd.read_excel(file_path)
                
                # Validação das colunas requeridas
                required_cols = ['Nome', 'Nota']
                missing = [col for col in required_cols if col not in self.excel_data.columns]
                
                if missing:
                    self.log(f"❌ ERRO: A planilha deve conter as colunas exatas: 'Nome' e 'Nota'. Faltando: {missing}")
                    self.excel_data = None
                    self.file_label.color = "red"
                else:
                    self.log(f"✅ Arquivo carregado: {len(self.excel_data)} registros encontrados.")
                    self.file_label.color = "green"
                    
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
            if pd.isna(row.get('Nome')) or pd.isna(row.get('Nota')):
                continue
                
            nome = str(row['Nome']).strip()
            nota = str(row['Nota']).strip()
            
            # Formata notas do tipo 4.0 para 4 para casar com o label exato do select
            if nota.endswith('.0'):
                nota = nota[:-2]
                
            self.log(f"Processando: {nome} -> Nota a lançar: {nota}")
            
            try:
                # 1. Localização do Aluno: localizar a célula <td> que contém o nome exato
                td_locator = popup_page.locator(f"xpath=//td[normalize-space(text())=\"{nome}\"]")
                
                if td_locator.count() == 0:
                    self.log(f"  ❌ Falha: Nome '{nome}' não encontrado na tabela.")
                    error_count += 1
                    continue
                
                # 2. Navegação na Tabela: subir para a linha pai (<tr>)
                tr_locator = td_locator.first.locator("xpath=ancestor::tr[1]")
                
                # Procurar o elemento <select> cujo atributo id termina com _ddFinalValue
                select_locator = tr_locator.locator("css=select[id$='_ddFinalValue']")
                
                if select_locator.count() == 0:
                    self.log(f"  ❌ Falha: Dropdown de nota não encontrado para o aluno.")
                    error_count += 1
                    continue
                
                # 3. Preenchimento: selecionar o valor pela label
                select_locator.first.select_option(label=nota)
                self.log(f"  ✅ Nota {nota} lançada com sucesso para {nome}")
                success_count += 1
                
                # Pequena pausa visual (opcional)
                popup_page.wait_for_timeout(100)
                
            except Exception as e:
                self.log(f"  ❌ Erro ao lançar nota para {nome}: {str(e)}")
                error_count += 1
                
        self.log("--- 🏁 Concluído ---")
        self.log(f"📊 Resumo: {success_count} Sucessos | {error_count} Erros")


def main(page: ft.Page):
    app = ESchoolingAutomation(page)

if __name__ == "__main__":
    ft.app(target=main)
