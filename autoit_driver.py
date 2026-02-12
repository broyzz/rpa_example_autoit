import win32com.client
import time
from typing import Optional
from win32api import GetSystemMetrics

class AutoItDriver:
    def __init__(self):
        """Inicializa o driver do AutoItX via interface COM."""
        try:
            # Usamos o Dispatch padrão. Se houver erro de nomes, 
            # o __getattr__ cuidará da busca dinâmica.
            self.autoit = win32com.client.Dispatch("AutoItX3.Control")
        except Exception as e:
            raise Exception(f"Erro ao carregar AutoItX3: {e}")

    def __getattr__(self, name):
        """
        Permite chamar qualquer método nativo (ex: driver.Send, driver.WinClose).
        Busca dinamicamente no objeto COM o nome solicitado.
        """
        return getattr(self.autoit, name)

    def tooltip(self, text: str, duration: float = 0, x: int = 1, y: int = 1):
        """Exibe um balão de ajuda (ToolTip)."""
        # Correção: O método na DLL é 'ToolTip' (com T maiúsculo)
        self.autoit.ToolTip(text, x, y)
        if duration > 0:
            time.sleep(duration)
            self.autoit.ToolTip("", 0, 0)

    def wait_and_activate(self, title: str, timeout: int = 5) -> bool:
        """Aguarda a existência e ativa uma janela."""
        if self.autoit.WinWait(title, "", timeout):
            self.autoit.WinActivate(title)
            self.autoit.WinWaitActive(title, "", timeout)
            return True
        return False

    def get_window_checksum(self, title: str, x1: int, y1: int, x2: int, y2: int) -> Optional[float]:
        """Calcula o checksum de uma área relativa à janela informada."""
        if not self.autoit.WinExists(title):
            print(f"Janela '{title}' não encontrada.")
            return None

        try:
            sw = GetSystemMetrics(0)
            sh = GetSystemMetrics(1)

            win_w = self.autoit.WinGetPosWidth(title)
            win_h = self.autoit.WinGetPosHeight(title)

            # Centralização da janela
            pos_x = (sw - win_w) // 2
            pos_y = (sh - win_h) // 2

            self.autoit.WinMove(title, "", pos_x, pos_y)
            self.autoit.WinActivate(title)
            self.autoit.WinWaitActive(title, "", 3)

            # Coordenadas Absolutas (Janela + Offset)
            win_x = self.autoit.WinGetPosX(title)
            win_y = self.autoit.WinGetPosY(title)

            time.sleep(0.3) # Pausa para renderização
            
            # Chama o método nativo da DLL
            return float(self.autoit.PixelChecksum(
                win_x + x1, 
                win_y + y1, 
                win_x + x2, 
                win_y + y2
            ))
        except Exception as e:
            print(f"Erro ao calcular checksum: {e}")
            return None
        