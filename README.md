# Automação com Python e AutoIt

Este projeto é um exemplo básico de automação de interface gráfica no Windows utilizando Python e a biblioteca COM do AutoIt (`AutoItX3.Control`).

O script de exemplo (`exemplo_basico.py`) realiza as seguintes ações:
1. Abre a Calculadora do Windows.
2. Realiza uma soma simples.
3. Valida o resultado através de um checksum de pixels.
4. Fecha a aplicação.

## Pré-requisitos

1. **Python 3.x** instalado.
2. **AutoIt Full Installation** instalado no Windows (necessário para registrar a DLL `AutoItX3.dll` no sistema, permitindo o acesso via `win32com`).
   - Você pode baixar em: AutoIt Downloads

## Instalação

1. Clone este repositório ou baixe os arquivos.

2. Crie um ambiente virtual (recomendado):
   ```bash
   python -m venv venv
   ```

3. Ative o ambiente virtual:
   - **Windows (PowerShell):**
     ```powershell
     .\venv\Scripts\Activate
     ```
   - **Windows (CMD):**
     ```cmd
     .\venv\Scripts\activate.bat
     ```

4. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

## Como Executar

Com o ambiente virtual ativo, execute o script:
```bash
python exemplo_basico.py
```

## Notas

- O cálculo de **checksum** de pixels pode variar dependendo da resolução da tela, tema do Windows ou versão da Calculadora. Se a validação falhar, pode ser necessário recalcular o checksum esperado usando a função `PixelChecksum` do AutoIt na sua máquina.