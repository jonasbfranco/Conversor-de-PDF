from flet import(
    app, Page, Text, ThemeMode, Container, Row, Column, padding, ElevatedButton,
    FilePicker, FilePickerResultEvent, ProgressBar, FontWeight
)
import os
import win32com.client
import pythoncom

def main(page: Page):
    page.title = "Conversor de DOCX para PDF"
    page.window.width = 420
    page.window.height = 780
    page.theme_mode = ThemeMode.DARK
    page.window.always_on_top = True
    page.window.maximizable = False
    page.window.resizable = False

    def on_diaolg_result(e: FilePickerResultEvent):
        if not e.files:
            return
        
        selected_file = e.files[0].path
        source_text.value = f"Arquivo selecionado: {selected_file}"
        page.update()

    
    def convert_docx_top_pdf(e):
        if not source_text.value:
            result_text.value = "Por favor, selecione um arquivo Word (DOCX)."
            page.update()
            return
        
        result_text.value = ""
        page.update()

        progress_bar.visible = True
        page.update()


        try:
            # Inicializo o ambiente COM
            pythoncom.CoInitialize()

            # Inicializando o Word
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False

            docx_path = source_text.value.split(": ")[1]
            pdf_path = docx_path.replace(".docx", ".pdf")

            # caminho absoluto dos arquivos
            docx_path = os.path.abspath(docx_path)
            pdf_path = os.path.abspath(pdf_path)

            # Define o formato de saida (pdf)
            pdf_format = 17

            # Abre o arquivo Word
            doc = word.Documents.Open(docx_path)

            # Salva o arquivo Word como pdf
            doc.SaveAs(pdf_path, FileFormat=pdf_format)

            # Fechar o documento
            doc.Close()

            # Fecha o word
            word.Quit()

            result_text.value = f"Documento convertido com sucesso: {pdf_path}"

        except Exception as e:
            result_text.value = f"Erro na conversão: {e}"
        finally:
            pythoncom.CoUninitialize()
            progress_bar.visible = False
            page.update()
        

    progress_bar = ProgressBar(width=300, visible=False, color="#EB06FF")

    file_picker = FilePicker(on_result=on_diaolg_result)
    page.overlay.append(file_picker)


    choose_file_button = ElevatedButton(
        width=260,
        height=50,
        bgcolor="#041955",
        content=Text(
            "Ecolher o arquivo...",
            weight=FontWeight.BOLD,
            color="#EB06FF",
            size=16,
        ),
        on_click=lambda _: file_picker.pick_files(allow_multiple=False),
    )

    convert_button = ElevatedButton(
        width=260,
        height=50,
        bgcolor="#041955",
        content=Text(
            "Converter .docx para .pdf",
            weight=FontWeight.BOLD,
            color="#EB06FF",
            size=16,
        ),
        on_click=convert_docx_top_pdf
    )

    source_text = Text(
                    value="",
                    size=16,
                    color="#18DCFF",
                    weight="normal",
                    max_lines=4,
                    width=300,
                    opacity=0.5,
                  )

    result_text = Text(
                    value="",
                    size=16,
                    color="#18DCFF",
                    weight="normal",
                    max_lines=4,
                    width=340,
                    opacity=0.5,
                  )



    page.add(
        Container(
            width=1200,
            height=720,
            bgcolor="#341F97",
            border_radius=35,
            padding=padding.only(left=20, top=60, right=20),
            content=Column(
                spacing=20,
                controls=[
                    Row(
                        controls=[
                            Text(
                                value="CONVERSOR .DOCX PARA .PDF",
                                size=30,
                                color="#EB06FF",
                                weight="bold",
                                max_lines=2,
                                width=300,
                            ),
                        ]
                    ),
                    Column(),
                    Column(),
                    Column(
                        controls=[
                            Text(
                                value="Selecione um arquivo .docx",
                                size=18,
                                color="#97B4FF",
                                weight="bold",
                            ),
                            choose_file_button
                        ]
                    ),
                    Row([source_text]),
                    Row([convert_button]),
                    Row([result_text],),
                    Row([progress_bar])
                ]
            )
        ),
    )

app(target=main)