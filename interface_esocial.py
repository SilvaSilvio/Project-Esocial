import flet as ft
import importador_dados
import verificar_dados
import os
from datetime import datetime
import subprocess
import re

def main(page: ft.Page):
    # Configure the page
    page.title = "eSocial Data Manager"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.padding = 20
    page.window_width = 800
    page.window_height = 600
    page.window_min_width = 800
    page.window_min_height = 600

    # Status message
    status_message = ft.Text("", color=ft.colors.BLUE)
    
    # Progress bar
    progress_bar = ft.ProgressBar(width=400, visible=False)
    
    # File picker
    file_picker = ft.FilePicker()
    page.overlay.append(file_picker)
    
    # Store selected files
    selected_files = {
        "ficha_financeira": None,
        "esocial": None
    }
    
    # Store file name texts
    file_name_texts = {
        "ficha_financeira": ft.Text("Nenhum arquivo selecionado", size=14),
        "esocial": ft.Text("Nenhum arquivo selecionado", size=14)
    }
    
    # Store current file type being selected
    current_file_type = None
    
    # Store the generated Excel file path
    generated_excel_file = None
    
    # Function to show status messages
    def show_status(message, color=ft.colors.BLUE):
        status_message.value = message
        status_message.color = color
        page.update()
    
    # Function to handle file selection
    def on_file_selection(e: ft.FilePickerResultEvent):
        if current_file_type and e.files:
            selected_files[current_file_type] = e.files[0].path
            file_name = os.path.basename(selected_files[current_file_type])
            file_name_texts[current_file_type].value = file_name
            show_status(f"Arquivo {current_file_type} selecionado: {file_name}")
            page.update()
        else:
            if current_file_type:
                selected_files[current_file_type] = None
                file_name_texts[current_file_type].value = "Nenhum arquivo selecionado"
                show_status(f"Nenhum arquivo {current_file_type} selecionado", ft.colors.RED)
                page.update()

    # Function to import data
    def import_data(e):
        try:
            # Check if both files are selected
            if not all(selected_files.values()):
                show_status("Por favor, selecione ambos os arquivos para importação", ft.colors.RED)
                return
                
            progress_bar.visible = True
            page.update()
            
            # Import data using the selected files
            importador_dados.main(
                ficha_financeira_path=selected_files["ficha_financeira"],
                esocial_path=selected_files["esocial"]
            )
            
            progress_bar.visible = False
            show_status("Importação concluída!", ft.colors.GREEN)
            
        except Exception as e:
            progress_bar.visible = False
            show_status(f"Erro durante a importação: {str(e)}", ft.colors.RED)

    # Function to open Excel file
    def open_excel_file(e):
        if generated_excel_file and os.path.exists(generated_excel_file):
            try:
                if os.name == 'nt':  # Windows
                    os.startfile(generated_excel_file)
                else:  # Linux/Mac
                    subprocess.call(('xdg-open', generated_excel_file))
                show_status("Arquivo Excel aberto!", ft.colors.GREEN)
            except Exception as e:
                show_status(f"Erro ao abrir o arquivo Excel: {str(e)}", ft.colors.RED)
        else:
            show_status(f"Arquivo Excel não encontrado: {generated_excel_file}", ft.colors.RED)
        page.dialog.open = False
        page.update()

    # Function to verify data
    def verify_data(e):
        nonlocal generated_excel_file
        try:
            progress_bar.visible = True
            page.update()
            
            show_status("Verificando dados...")
            verificar_dados.verificar_dados()
            
            # Captura a saída da função consultar_dados_join
            import sys
            from io import StringIO
            
            # Redireciona a saída padrão para capturar as mensagens
            old_stdout = sys.stdout
            sys.stdout = mystdout = StringIO()
            
            verificar_dados.consultar_dados_join()
            
            # Restaura a saída padrão
            sys.stdout = old_stdout
            
            # Obtém as mensagens capturadas
            output = mystdout.getvalue()
            
            # Procura pelo nome do arquivo na saída
            match = re.search(r'Resultados salvos em: (resultado_consulta_\d{8}_\d{6}\.xlsx)', output)
            if match:
                generated_excel_file = match.group(1)
                if os.path.exists(generated_excel_file):
                    # Cria o diálogo de confirmação
                    page.dialog = ft.AlertDialog(
                        title=ft.Text("Arquivo Excel Gerado"),
                        content=ft.Text(f"O arquivo {generated_excel_file} foi gerado com sucesso. Deseja abri-lo agora?"),
                        actions=[
                            ft.TextButton("Sim", on_click=open_excel_file),
                            ft.TextButton("Não", on_click=lambda e: setattr(page.dialog, 'open', False))
                        ],
                        actions_alignment=ft.MainAxisAlignment.END
                    )
                    page.dialog.open = True
                    page.update()
                else:
                    show_status(f"Arquivo Excel não encontrado: {generated_excel_file}", ft.colors.RED)
            else:
                show_status("Não foi possível encontrar o nome do arquivo Excel gerado", ft.colors.RED)
            
            progress_bar.visible = False
            show_status("Verificação concluída com sucesso!", ft.colors.GREEN)
            show_status("Planilhas geradas com sucesso!", ft.colors.GREEN)
            
        except Exception as e:
            progress_bar.visible = False
            show_status(f"Erro durante a verificação: {str(e)}", ft.colors.RED)

    # Function to handle button click
    def handle_file_picker_click(file_type):
        nonlocal current_file_type
        current_file_type = file_type
        file_picker.pick_files(
            allow_multiple=False,
            allowed_extensions=["csv", "xlsx", "xls"]
        )

    # Create buttons with their respective file types
    ficha_financeira_button = ft.ElevatedButton(
        "Selecionar Arquivo",
        icon=ft.icons.UPLOAD_FILE,
        on_click=lambda e: handle_file_picker_click("ficha_financeira")
    )

    esocial_button = ft.ElevatedButton(
        "Selecionar Arquivo",
        icon=ft.icons.UPLOAD_FILE,
        on_click=lambda e: handle_file_picker_click("esocial")
    )

    # Set up file picker callback
    file_picker.on_result = on_file_selection

    # Create the main layout
    page.add(
        ft.Column(
            controls=[
                ft.Text("eSocial Data Manager", size=30, weight=ft.FontWeight.BOLD),
                ft.Divider(),
                
                # Import Section
                ft.Text("Importar Dados", size=20, weight=ft.FontWeight.BOLD),
                
                # Ficha Financeira Section
                ft.Row(
                    controls=[
                        ft.Text("Ficha Financeira:", size=16),
                        ficha_financeira_button,
                        file_name_texts["ficha_financeira"]
                    ],
                    alignment=ft.MainAxisAlignment.START,
                    spacing=10
                ),
                
                # eSocial Section
                ft.Row(
                    controls=[
                        ft.Text("eSocial:", size=16),
                        esocial_button,
                        file_name_texts["esocial"]
                    ],
                    alignment=ft.MainAxisAlignment.START,
                    spacing=10
                ),
                
                # Import Button
                ft.ElevatedButton(
                    "Importar para Banco de Dados",
                    icon=ft.icons.SAVE,
                    on_click=import_data
                ),
                
                ft.Divider(),
                
                # Verify Section
                ft.Text("Verificar Dados", size=20, weight=ft.FontWeight.BOLD),
                ft.ElevatedButton(
                    "Verificar Banco de Dados",
                    icon=ft.icons.VERIFIED,
                    on_click=verify_data
                ),
                
                ft.Divider(),
                
                # Status Section
                progress_bar,
                status_message
            ],
            spacing=20,
            alignment=ft.MainAxisAlignment.START,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
        )
    )

if __name__ == "__main__":
    ft.app(target=main) 