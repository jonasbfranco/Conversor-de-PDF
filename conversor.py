import os
import win32com.client

# Caminho para o arquivo .docx e o caminhio para o arquivo .pdf
word_path = "Canal_PyFlet.docx"
pdf_path = "Canal_PyFlet.pdf"


# Inicializar o aplicativo Word
word = win32com.client.Dispatch("Word.Application")
word.Visible = False

# caminho absoluto dos arquivos
docx_path = os.path.abspath(word_path)
pdf_path = os.path.abspath(pdf_path)

# Define o formato de saida (pdf)
pdf_format = 17

try:
    # Abre o arquivo Word
    doc = word.Documents.Open(docx_path)

    # Salva o arquivo Word como pdf
    doc.SaveAs(pdf_path, pdf_format)

    # Fechar o documento
    doc.Close()

    print(f"Documento convertido com sucesso: {pdf_path}")
except Exception as e:
    print(f"Erro na conversão: {e}")
finally:
    # Fecha o word
    word.Quit()


