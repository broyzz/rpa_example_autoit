"""
Exemplo básico de uso do AutoIt via DLL com pywin32
"""
import win32com.client
import time

def exemplo_basico():
    """
    Executa uma demonstração básica de automação utilizando a calculadora do Windows.

    Abre o aplicativo, realiza uma soma simples (123 + 123), valida o resultado
    via checksum de pixels da área de resultado e, por fim, fecha o programa.
    """
    # Instancia o objeto de controle do AutoItX
    autoit = win32com.client.Dispatch("AutoItX3.Control")

    # Exibe tooltip informando o início
    autoit.Tooltip(" === INICIANDO DEMONSTRAÇÃO BÁSICA === ", 1, 1)

    # --- Abertura da Calculadora ---
    autoit.Tooltip(" === ABRINDO PROGRAMA CALCULADORA === ", 1, 1)
    autoit.Send("#r")  # Pressiona Win+R para abrir o Executar
    autoit.WinWaitActive("Executar")  # Aguarda a janela Executar
    autoit.Send("calc{ENTER}")  # Digita calc e pressiona Enter
    autoit.WinActivate("Calculadora", "")  # Garante que a Calculadora está ativa

    # --- Realização do Cálculo ---
    autoit.Tooltip(" === REALIZANDO CALCULO === ", 1, 1)
    autoit.Send("123")
    time.sleep(0.5)  # Pausa para garantir processamento da interface
    autoit.Send("{+}")
    time.sleep(0.5)
    autoit.Send("123")
    time.sleep(0.5)
    autoit.Send("=")

    # --- Validação do Resultado ---
    autoit.Tooltip(" === VALIDANDO CALCULO === ", 1, 1)
    # Checksum esperado para o resultado (pode variar conforme resolução/tema do Windows)
    num_calc = 284645482.0
    checksum = obter_checksum_janela("Calculadora", 270, 92, 307, 121)

    if num_calc == checksum:
        print('Calculo Efetuado com sucesso')
    else: 
        print('Erro no calculo')

    # --- Fechamento da Calculadora ---
    autoit.Tooltip(" === FECHANDO CALCULADORA === ", 1, 1)
    autoit.WinClose("Calculadora")
    autoit.Tooltip("", 0, 0)  # Limpa o tooltip


def obter_checksum_janela(titulo: str, x1: int, y1: int, x2: int, y2: int):    
    """Ativa uma janela e obtém o checksum de uma área específica RELATIVA a ela.

    Para obter o checksum relativo à janela, a função primeiro encontra a posição
    absoluta (x, y) da janela no monitor e ajusta as coordenadas da área.

    Args:
        titulo (str): Título da janela alvo.
        x1 (int): Coordenada X inicial relativa à janela.
        y1 (int): Coordenada Y inicial relativa à janela.
        x2 (int): Coordenada X final relativa à janela.
        y2 (int): Coordenada Y final relativa à janela.

    Returns:
        float | None: O valor do checksum ou None se a janela não for encontrada.
    """
    autoit = win32com.client.Dispatch("AutoItX3.Control")
    
    if autoit.WinExists(titulo):
        # Restaura a janela (SW_RESTORE = 9) caso esteja minimizada e a ativa
        autoit.WinSetState(titulo, "", 9)
        autoit.WinActivate(titulo)
        autoit.WinWaitActive(titulo, "", 5)  # Aguarda até 5s pela ativação
        
        # Obtém a posição atual da janela no monitor (absoluta)
        win_x = autoit.WinGetPosX(titulo)
        win_y = autoit.WinGetPosY(titulo)
        
        # Ajusta as coordenadas locais para coordenadas globais (do monitor)
        # somando a posição inicial da janela
        abs_x1 = win_x + x1
        abs_y1 = win_y + y1
        abs_x2 = win_x + x2
        abs_y2 = win_y + y2
        
        # Calcula o checksum da área especificada nas coordenadas absolutas
        checksum = autoit.PixelChecksum(abs_x1, abs_y1, abs_x2, abs_y2)
        print(f"Checksum da área relativa na janela '{titulo}': {checksum}")
        return checksum
    else:
        print(f"Janela '{titulo}' não encontrada.")
        return None
    

if __name__ == "__main__":
    exemplo_basico()