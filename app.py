import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import sys # Usado para sair do script em caso de erro crítico

# --- Configuração ---
NOME_ARQUIVO_EXCEL = 'clientes.xlsx'
NOME_PLANILHA = 'Planilha1'

# Defina quantas linhas VÁLIDAS você quer processar (após a validação).
# Coloque 0 ou um número negativo para processar todas as linhas válidas encontradas.
MAX_LINHAS_A_PROCESSAR = 10  # Exemplo: processar as primeiras 10 linhas válidas

# --- Carregar Planilha ---
try:
    workbook = openpyxl.load_workbook(NOME_ARQUIVO_EXCEL)
    pagina_clientes = workbook[NOME_PLANILHA]
    print(f"Arquivo '{NOME_ARQUIVO_EXCEL}' e planilha '{NOME_PLANILHA}' carregados com sucesso.")
except FileNotFoundError:
    print(f"Erro: Arquivo '{NOME_ARQUIVO_EXCEL}' não encontrado.")
    print("Verifique se o nome do arquivo está correto e se ele está na mesma pasta do script.")
    sys.exit(1) # Sai do script se o arquivo não existe
except KeyError:
    print(f"Erro: Planilha '{NOME_PLANILHA}' não encontrada no arquivo '{NOME_ARQUIVO_EXCEL}'.")
    print("Verifique o nome da planilha dentro do arquivo Excel.")
    sys.exit(1) # Sai do script se a planilha não existe
except Exception as e:
    print(f"Ocorreu um erro inesperado ao carregar a planilha: {e}")
    sys.exit(1)

# --- Processamento ---
linhas_validas_processadas = 0
print(f"\nIniciando processamento...")
if MAX_LINHAS_A_PROCESSAR > 0:
    print(f"Limite de processamento definido para {MAX_LINHAS_A_PROCESSAR} linhas válidas.")
else:
    print("Processando todas as linhas válidas encontradas (sem limite).")

# Iterar pelas linhas começando da segunda linha (min_row=2 para pular cabeçalho)
for linha_idx, linha in enumerate(pagina_clientes.iter_rows(min_row=2), start=2): # start=2 para número da linha correto

    # --- Controle de Limite ---
    # Verifica se já atingimos o limite de processamento (se um limite > 0 foi definido)
    if MAX_LINHAS_A_PROCESSAR > 0 and linhas_validas_processadas >= MAX_LINHAS_A_PROCESSAR:
        print(f"\nLimite de {MAX_LINHAS_A_PROCESSAR} linhas válidas processadas atingido.")
        break # Sai do loop 'for'

    # --- Extração e Validação ---
    # Coluna A (índice 0) -> telefone
    # Coluna B (índice 1) -> mensagem
    telefone_valor = linha[0].value
    mensagem_valor = linha[1].value
    numero_linha_excel = linha[0].row # Pega o número da linha diretamente do objeto cell

    # Valida se os valores são nulos ou vazios (após converter para string e remover espaços)
    telefone = None
    if telefone_valor is not None:
        telefone_str = str(telefone_valor).strip()
        if telefone_str: # Verifica se a string não está vazia após strip
            telefone = telefone_str

    mensagem = None
    if mensagem_valor is not None:
        mensagem_str = str(mensagem_valor).strip()
        if mensagem_str: # Verifica se a string não está vazia após strip
            mensagem = mensagem_str

    # Se telefone ou mensagem não são válidos (continuam None), pula para a próxima linha
    if not telefone or not mensagem:
        print(f"Linha {numero_linha_excel}: Ignorando - Telefone (A) ou Mensagem (B) inválido/vazio. (A='{telefone_valor}', B='{mensagem_valor}')")
        continue # Pula para a próxima iteração do loop 'for'

    # --- Processar Linha Válida ---
    linhas_validas_processadas += 1 # Incrementa o contador APENAS para linhas válidas
    print(f"\nProcessando Linha Válida #{linhas_validas_processadas} (Excel Linha: {numero_linha_excel})")
    print(f"  Telefone: {telefone}")
    # print(f"  Mensagem: {mensagem}") # Descomente se quiser ver a mensagem no console

    try:
        # Montar o link do WhatsApp com a mensagem da Coluna B
        # A função quote garante que caracteres especiais na mensagem sejam formatados corretamente para URL
        link_msg_zap = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        print(f"  Abrindo link: {link_msg_zap}") # Mostra o link que será aberto

        # Abrir o link no navegador padrão
        webbrowser.open(link_msg_zap)
        print("  -> Link aberto. Aguardando WhatsApp Web carregar...")
        sleep(10) # Tempo para WhatsApp Web carregar e encontrar o contato. Pode precisar de ajuste.

        # Clicar no botão de enviar
        # !!! ATENÇÃO: Estas coordenadas (x, y) são MUITO específicas para a sua tela e layout !!!
        # !!! Podem precisar de ajuste se a resolução, tamanho da janela ou layout do WhatsApp mudar. !!!
        # !!! Considere usar reconhecimento de imagem (pyautogui.locateCenterOnScreen) para mais robustez !!!
        x_coord_enviar = 1329
        y_coord_enviar = 688
        print(f"  -> Tentando clicar no botão enviar em ({x_coord_enviar}, {y_coord_enviar})...")
        pyautogui.click(x=x_coord_enviar, y=y_coord_enviar)
        sleep(2) # Espera um pouco para garantir que o clique registrou e a mensagem foi enviada.

        # Fechar a aba do navegador (Ctrl+W)
        print("  -> Fechando a aba (Ctrl+W)...")
        pyautogui.hotkey('ctrl', 'w')
        sleep(2) # Espera a aba fechar.

        print(f"  -> Linha {numero_linha_excel} processada com sucesso.")

    except Exception as e:
        print(f"Erro ao processar a linha {numero_linha_excel} (Telefone: {telefone}): {e}")
        print("  -> Tentando continuar para a próxima linha...")
        # Você pode decidir parar o script aqui se um erro for crítico:
        # break
        # Ou apenas continuar para a próxima linha (comportamento atual)
        continue

# --- Finalização ---
print(f"\n--- Processamento Concluído ---")
print(f"Total de linhas válidas encontradas e processadas: {linhas_validas_processadas}")
if MAX_LINHAS_A_PROCESSAR > 0 and linhas_validas_processadas < MAX_LINHAS_A_PROCESSAR:
    print(f"O limite de {MAX_LINHAS_A_PROCESSAR} não foi atingido (ou não havia mais linhas válidas).")
