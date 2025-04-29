import time
from pynput import keyboard, mouse

# --- Configuração ---
# Escolha a tecla que você quer usar para registrar as coordenadas.
# Exemplos:
# LOG_KEY = keyboard.Key.f1  # Tecla F1
# LOG_KEY = keyboard.KeyCode.from_char('l') # Tecla 'l' (minúscula)
LOG_KEY = keyboard.KeyCode.from_char('l') # <<-- Mude aqui se quiser outra tecla

# Tecla para sair do programa
EXIT_KEY = keyboard.Key.esc
# ---------------------

print(f"Programa de Mapeamento de Mouse Iniciado.")
print(f"Pressione a tecla '{LOG_KEY}' para registrar as coordenadas do mouse.")
print(f"Pressione a tecla '{EXIT_KEY}' para sair.")
print("-" * 30)

# Cria uma instância do controlador do mouse para obter a posição
mouse_controller = mouse.Controller()

# Função que será chamada quando uma tecla for pressionada
def on_press(key):
    try:
        # Verifica se a tecla pressionada é a tecla de log
        if key == LOG_KEY:
            # Pega a posição atual do mouse
            current_position = mouse_controller.position
            # Imprime as coordenadas no console (log)
            print(f"LOG: Coordenadas = {current_position}")

        # Verifica se a tecla pressionada é a tecla de saída
        elif key == EXIT_KEY:
            print("\nTecla de saída pressionada. Encerrando...")
            # Retorna False para parar o listener do teclado
            return False

    except AttributeError:
        # Algumas teclas especiais (como shift, ctrl) não têm o atributo 'char'
        # Se LOG_KEY for uma tecla especial (como F1), a comparação direta funciona.
        # Se LOG_KEY for um caractere, a comparação direta também funciona.
        # Este except é mais um fallback geral.
        # Precisamos verificar a tecla de saída aqui também caso ela seja especial
        if key == EXIT_KEY:
            print("\nTecla de saída pressionada. Encerrando...")
            return False


# Cria um listener para eventos de teclado
# 'on_press=on_press' diz qual função chamar quando uma tecla é pressionada
listener_teclado = keyboard.Listener(on_press=on_press)

# Inicia o listener em uma thread separada
listener_teclado.start()

# Mantém o script principal rodando enquanto o listener estiver ativo
listener_teclado.join()

print("-" * 30)
print("Programa terminado.")