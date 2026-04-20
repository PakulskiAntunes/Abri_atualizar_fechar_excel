import openpyxl
import os
import win32com.client
import time

caminho = r"INSERE O CAMINHO DO ARQUIVO"

def main():
	if not os.path.exists(caminho):
		print("Arquivo não encontrado!")
		return
	try:
		# Abre o Excel VISÍVEL
		excel = win32com.client.Dispatch("Excel.Application")
		excel.Visible = True  # Excel fica visível na tela
		excel.ScreenUpdating = True  # Mostra as atualizações
		
		# Abre o workbook
		print("Abrindo arquivo...")
		wb = excel.Workbooks.Open(caminho, UpdateLinks=2)  # UpdateLinks=2 atualiza links externos
		
		# Aguarda um pouco para processamento
		time.sleep(3)
		
		# Força recálculo
		print("Recalculando fórmulas...")
		excel.CalculateFull()
		
		# Aguarda recálculo completar
		time.sleep(2)
		
		# Salva o arquivo
		print("Salvando arquivo...")
		wb.Save()
		
		# Aguarda
		time.sleep(1)
		
		# Fecha o workbook
		wb.Close(SaveChanges=False)
		
		# Fecha o Excel
		excel.Quit()
		
		print("Arquivo aberto, atualizado e fechado com sucesso!")
	
	except Exception as e:
		print(f"Erro: {e}")

if __name__ == "__main__":
	main()
