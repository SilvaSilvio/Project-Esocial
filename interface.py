import flet as ft
from read_excel import transform_data
import pandas as pd
from datetime import datetime
import os

class FilePicker(ft.UserControl):
    def __init__(self, title, on_result):
        super().__init__()
        self.title = title
        self.on_result = on_result
        self.file_path = None

    def pick_files_result(self, e: ft.FilePickerResultEvent):
        if e.files:
            file_path = e.files[0].path
            # Verifica se a extensão é xlsx ou csv
            if file_path.lower().endswith(('.xlsx', '.csv')):
                self.file_path = file_path
                self.on_result(self.file_path)
                self.update()
            else:
                self.file_path = None
                self.update()
                # Mostra mensagem de erro
                self.page.show_snack_bar(
                    ft.SnackBar(
                        content=ft.Text("Por favor, selecione apenas arquivos .xlsx ou .csv"),
                        bgcolor=ft.colors.RED
                    )
                )

    def build(self):
        self.file_picker = ft.FilePicker(
            on_result=self.pick_files_result
        )

        return ft.Column([
            ft.Text(self.title, size=16, weight=ft.FontWeight.BOLD),
            ft.ElevatedButton(
                "Selecionar Arquivo",
                icon=ft.icons.UPLOAD_FILE,
                on_click=lambda _: self.file_picker.pick_files(
                    allow_multiple=False
                )
            ),
            ft.Text(
                self.file_path if self.file_path else "Nenhum arquivo selecionado",
                size=12,
                color=ft.colors.GREY_700
            ),
            self.file_picker
        ])

class MainApp(ft.UserControl):
    def __init__(self):
        super().__init__()
        self.esocial_file = None
        self.lg_file = None
        self.depara_file = None
        self.status_text = ft.Text("", color=ft.colors.GREY_700)
        self.group_checkbox = ft.Checkbox(
            label="Agrupar registros iguais",
            value=False
        )

    def on_esocial_file_selected(self, file_path):
        self.esocial_file = file_path
        self.update_status()

    def on_lg_file_selected(self, file_path):
        self.lg_file = file_path
        self.update_status()

    def on_depara_file_selected(self, file_path):
        self.depara_file = file_path
        self.update_status()

    def update_status(self):
        status = []
        if self.esocial_file:
            status.append("✓ Arquivo eSocial selecionado")
        if self.lg_file:
            status.append("✓ Arquivo LG selecionado")
        if self.depara_file:
            status.append("✓ Arquivo Depara selecionado")
        
        self.status_text.value = "\n".join(status) if status else "Selecione os arquivos necessários"
        self.status_text.update()

    def group_and_sum(self, df):
        """
        Agrupa os registros iguais e soma os valores, mantendo todas as informações
        """
        try:
            print("\nIniciando agrupamento...")
            print("Colunas originais:", df.columns.tolist())
            
            # Garante que a coluna Valor está em formato numérico
            df['Valor'] = pd.to_numeric(df['Valor'].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
            
            # Define as colunas para agrupamento (todas exceto Valor)
            # Importante: precisamos agrupar por todas as colunas que identificam unicamente um registro
            group_columns = [col for col in df.columns if col != 'Valor']
            
            print(f"\nAgrupando por: {group_columns}")
            print("Exemplo de dados antes do agrupamento:")
            print(df[['Código', 'Descrição', 'Valor']].head())
            
            # Agrupa e soma os valores
            df_grouped = df.groupby(group_columns, as_index=False).agg({
                'Valor': 'sum'
            })
            
            print("\nExemplo de dados após agrupamento:")
            print(df_grouped[['Código', 'Descrição', 'Valor']].head())
            
            print(f"\nTotal de registros antes do agrupamento: {len(df)}")
            print(f"Total de registros após agrupamento: {len(df_grouped)}")
            
            # Ordena o resultado
            if 'Código' in df_grouped.columns:
                df_grouped = df_grouped.sort_values(by=['Código', 'Valor'], na_position='last')
            
            # Formata o valor para duas casas decimais
            df_grouped['Valor'] = df_grouped['Valor'].round(2)
            
            return df_grouped
            
        except Exception as e:
            print(f"\nERRO no agrupamento: {str(e)}")
            print("Detalhes do erro:")
            print("Tipos de dados das colunas:")
            print(df.dtypes)
            print("\nAmostra dos dados:")
            print(df.head())
            raise e

    def process_files(self):
        if not all([self.esocial_file, self.lg_file, self.depara_file]):
            self.status_text.value = "Por favor, selecione todos os arquivos necessários"
            self.status_text.color = ft.colors.RED
            self.status_text.update()
            return

        try:
            # Processa o arquivo eSocial
            df_esocial = pd.read_excel(self.esocial_file)
            df_transformed = transform_data(df_esocial)
            
            # Aplica agrupamento se selecionado
            if self.group_checkbox.value:
                df_transformed = self.group_and_sum(df_transformed)
            
            # Gera timestamp para o nome do arquivo
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"transformado_{timestamp}.xlsx"
            
            # Salva o resultado
            df_transformed.to_excel(output_file, index=False)
            
            # Atualiza o status com informações do processamento
            status = f"✓ Processamento concluído!\nArquivo salvo como: {output_file}"
            if self.group_checkbox.value:
                status += "\n✓ Registros agrupados e valores somados"
            self.status_text.value = status
            self.status_text.color = ft.colors.GREEN
            self.status_text.update()
            
        except Exception as e:
            self.status_text.value = f"Erro ao processar arquivos: {str(e)}"
            self.status_text.color = ft.colors.RED
            self.status_text.update()

    def build(self):
        return ft.Container(
            content=ft.Column([
                ft.Text("Processador de Arquivos eSocial", 
                       size=24, 
                       weight=ft.FontWeight.BOLD,
                       text_align=ft.TextAlign.CENTER),
                
                ft.Divider(),
                
                FilePicker(
                    "Arquivo eSocial",
                    self.on_esocial_file_selected
                ),
                
                ft.Divider(),
                
                FilePicker(
                    "Arquivo LG",
                    self.on_lg_file_selected
                ),
                
                ft.Divider(),
                
                FilePicker(
                    "Arquivo Depara Eventos",
                    self.on_depara_file_selected
                ),
                
                ft.Divider(),
                
                self.group_checkbox,
                
                ft.Divider(),
                
                ft.ElevatedButton(
                    "Processar Arquivos",
                    icon=ft.icons.PLAY_ARROW,
                    on_click=lambda _: self.process_files()
                ),
                
                self.status_text
            ]),
            padding=20
        )

def main(page: ft.Page):
    page.title = "Processador de Arquivos eSocial"
    page.window_width = 600
    page.window_height = 800
    page.window_resizable = False
    page.window_center()
    
    app = MainApp()
    page.add(app)

if __name__ == "__main__":
    ft.app(target=main) 