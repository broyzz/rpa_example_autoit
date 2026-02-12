from autoit_driver import AutoItDriver


if __name__ == "__main__":
    driver = AutoItDriver()

    # Início da Automação
    driver.tooltip("Iniciando Calculadora...", duration=1)
    driver.Send("#r")
    driver.wait_and_activate("Executar")
    driver.Send("calc{ENTER}")

    if driver.wait_and_activate("Calculadora"):
        driver.Send("123{+}")
        driver.Send("123{ENTER}")

        resultado_check = driver.get_window_checksum("Calculadora", 230, 120, 300, 155)
        print(f"Checksum obtido: {resultado_check}")

        driver.close_window("Calculadora")