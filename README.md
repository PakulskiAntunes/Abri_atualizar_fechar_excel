# 📊 Excel Auto-Updater

Script Python que abre automaticamente um arquivo Excel, força o recálculo de todas as fórmulas (incluindo links externos), salva e fecha — sem intervenção manual.

---

## 🧩 Funcionalidades

- Abre o arquivo Excel de forma visível usando a API COM do Windows
- Atualiza links externos automaticamente
- Força o recálculo completo de todas as fórmulas (`CalculateFull`)
- Salva e fecha o arquivo de forma segura

---

## ✅ Pré-requisitos

- Windows (obrigatório — o script usa `win32com`)
- Python 3.7+
- Microsoft Excel instalado

---

## 📦 Instalação das dependências

```bash
pip install openpyxl pywin32
```

---

## ⚙️ Configuração

Antes de executar, edite a variável `caminho` no início do script com o caminho completo do seu arquivo Excel:

```python
caminho = r"C:\Users\SeuUsuario\Documentos\planilha.xlsx"
```

---

## ▶️ Como usar

```bash
python excel_updater.py
```

O script irá:
1. Verificar se o arquivo existe
2. Abrir o Excel (visível na tela)
3. Atualizar links externos
4. Recalcular todas as fórmulas
5. Salvar o arquivo
6. Fechar o Excel automaticamente

---

## 📁 Estrutura

```
📄 excel_updater.py   # Script principal
📄 README.md
```

---

## ⚠️ Observações

- O script **só funciona no Windows** devido ao uso da biblioteca `pywin32` (interface COM do Excel).
- O Excel ficará visível durante a execução — isso é intencional para garantir que as atualizações de tela sejam processadas corretamente.
- Os `time.sleep()` existem para garantir que o Excel conclua operações assíncronas antes de avançar para o próximo passo. Ajuste os valores se necessário.

---

## 📄 Licença

MIT
