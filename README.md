# WhatsApp Bulk Messenger from Excel

Este script Python automatiza o envio de mensagens via WhatsApp Web para contatos listados em uma planilha Excel. Ele lê números de telefone da Coluna A e mensagens personalizadas da Coluna B, valida os dados e, em seguida, usa automação web e de interface gráfica (GUI) para enviar as mensagens.

## Funcionalidades

*   **Leitura de Excel:** Lê dados de arquivos `.xlsx` usando a biblioteca `openpyxl`.
*   **Mapeamento de Colunas:**
    *   Coluna A: Números de telefone dos destinatários (formato internacional recomendado, ex: `5511999998888`).
    *   Coluna B: Conteúdo da mensagem a ser enviada para o respectivo número.
*   **Validação de Dados:** Verifica se as colunas de telefone e mensagem não estão vazias ou nulas. Linhas inválidas são ignoradas, e uma mensagem é exibida no console.
*   **Limite de Processamento:** Permite definir um número máximo de linhas válidas a serem processadas, útil para testes ou envio em lotes menores.
*   **Automação Web:** Abre automaticamente o link `web.whatsapp.com` formatado com o telefone e a mensagem.
*   **Automação de GUI:** Usa `pyautogui` para simular cliques do mouse (para enviar a mensagem) e atalhos de teclado (para fechar a aba do navegador).
*   **Feedback no Console:** Exibe informações sobre o carregamento do arquivo, o progresso do processamento, linhas ignoradas e possíveis erros.

## Pré-requisitos

1.  **Python 3:** Certifique-se de ter o Python 3.x instalado em seu sistema. Você pode baixá-lo em [python.org](https://www.python.org/).
2.  **Bibliotecas Python:** Instale as dependências necessárias:
    ```bash
    pip install openpyxl pyautogui
    ```
3.  **Navegador Web:** Um navegador web padrão (Chrome, Firefox, Edge, etc.).
4.  **WhatsApp Web/Desktop:** Você precisa estar **logado** no WhatsApp Web ou no aplicativo WhatsApp Desktop no navegador que será aberto pelo script *antes* de executá-lo.

## Configuração

Antes de executar o script, você **precisa** ajustar algumas variáveis diretamente no código Python (`.py`):

1.  **Nome do Arquivo e Planilha:**
    *   `NOME_ARQUIVO_EXCEL = 'clientes.xlsx'`: Altere se o seu arquivo Excel tiver um nome diferente.
    *   `NOME_PLANILHA = 'pagina1'`: Altere se a sua planilha dentro do arquivo Excel tiver um nome diferente.

2.  **Limite de Linhas:**
    *   `MAX_LINHAS_A_PROCESSAR = 10`: Defina quantas linhas **válidas** (com telefone e mensagem preenchidos) você deseja processar. Defina como `0` ou um número negativo para processar todas as linhas válidas encontradas.

3.  **Coordenadas do PyAutoGUI (MUITO IMPORTANTE):**
    *   `x_coord_enviar = 1329`
    *   `y_coord_enviar = 688`
    *   **Estas coordenadas são específicas para a resolução de tela, tamanho da janela do navegador e layout do WhatsApp Web no momento em que foram definidas.** Elas **QUASE CERTAMENTE** precisarão ser ajustadas para a *sua* tela.
    *   **Como encontrar suas coordenadas:**
        *   Execute o script uma vez para abrir uma janela do WhatsApp Web.
        *   Pause o script (ou abra um terminal Python separado).
        *   Use o comando `python -m pyautogui` e depois `pyautogui.displayMousePosition()` no terminal, ou execute um pequeno script como:
            ```python
            import pyautogui
            import time
            print("Mova o mouse para o botão ENVIAR e aguarde...")
            time.sleep(5)
            x, y = pyautogui.position()
            print(f"Coordenadas: x={x}, y={y}")
            ```
        *   Posicione o cursor do mouse exatamente sobre o centro do botão "Enviar" na janela do WhatsApp Web e anote os valores de X e Y exibidos.
        *   Atualize `x_coord_enviar` e `y_coord_enviar` no script com os seus valores.

4.  **Tempos de Espera (`sleep`):**
    *   Os valores `sleep(15)`, `sleep(3)`, `sleep(2)` definem pausas para permitir que o navegador carregue, o clique seja registrado e a aba seja fechada.
    *   Se o script estiver falhando (não enviando, fechando cedo demais), tente **aumentar** esses valores, especialmente o primeiro `sleep(15)` após abrir o link, pois o WhatsApp Web pode demorar para carregar. Ajuste conforme a velocidade do seu computador e da sua conexão com a internet.

## Formato do Arquivo Excel (`clientes.xlsx`)

*   O arquivo deve estar no formato `.xlsx`.
*   Use a planilha especificada em `NOME_PLANILHA` (padrão: `pagina1`).
*   A **primeira linha (Linha 1)** é considerada como cabeçalho e será ignorada pelo script.
*   **Coluna A:** Deve conter os números de telefone dos destinatários. **Inclua o código do país e DDD sem espaços ou caracteres especiais** (Exemplo para Brasil/São Paulo: `5511999998888`).
*   **Coluna B:** Deve conter a mensagem de texto completa que você deseja enviar para o número da Coluna A na mesma linha.

**Exemplo:**

| Coluna A (Telefone) | Coluna B (Mensagem)                                    |
| :------------------ | :----------------------------------------------------- |
| 5511999998888       | Olá Cliente A, esta é uma mensagem de teste.           |
| 5521988887777       | Prezado Cliente B, temos uma novidade para você!       |
|                     | Esta linha será ignorada (telefone vazio)              |
| 5531977776666       |                                                        | <-- Esta linha será ignorada (mensagem vazia)
| 5541966665555       | Olá Cliente C, tudo bem? Confira nossa promoção. |

## Como Usar

1.  **Prepare o Arquivo Excel:** Crie ou ajuste seu arquivo `clientes.xlsx` (ou o nome configurado) com os dados no formato especificado acima. Salve-o na mesma pasta do script Python.
2.  **Configure o Script:** Abra o arquivo `.py` e ajuste as variáveis na seção "Configuração" conforme descrito acima (especialmente as coordenadas X, Y!).
3.  **Login no WhatsApp:** Certifique-se de que você está logado no WhatsApp Web (web.whatsapp.com) ou no WhatsApp Desktop no seu navegador padrão.
4.  **Execute o Script:** Abra um terminal ou prompt de comando, navegue até a pasta onde salvou o script e o Excel, e execute o comando:
    ```bash
    python seu_nome_de_script.py
    ```
    (Substitua `seu_nome_de_script.py` pelo nome real do seu arquivo Python).
5.  **Não Interfira:** Enquanto o script estiver rodando, **não use o mouse ou o teclado**. O `pyautogui` precisa controlar o cursor e enviar comandos. Mantenha a janela do navegador que o script abre visível e em foco quando esperado. O script irá abrir e fechar abas do navegador automaticamente.
6.  **Acompanhe o Console:** Observe as mensagens no terminal para ver o progresso, quais linhas estão sendo processadas ou ignoradas, e se ocorrem erros.

## Considerações Importantes e Avisos

*   **Fragilidade do `pyautogui`:** A automação baseada em coordenadas é muito sensível a mudanças na resolução da tela, tamanho da janela, zoom do navegador ou atualizações no layout do site do WhatsApp Web. Se parar de funcionar, a primeira coisa a verificar e ajustar são as coordenadas `x_coord_enviar`, `y_coord_enviar`. Uma alternativa mais robusta (mas mais complexa de implementar) seria usar reconhecimento de imagem com `pyautogui`.
*   **Termos de Serviço do WhatsApp:** O uso excessivo de automação para envio em massa pode violar os Termos de Serviço do WhatsApp e potencialmente levar ao bloqueio da sua conta. **Use este script com responsabilidade e moderação.** Evite enviar spam.
*   **Foco da Janela:** O `pyautogui` interage com a janela que tem foco. Certifique-se de que a janela do navegador aberta pelo script permanece ativa quando o clique/envio deve ocorrer.
*   **Bloqueio de Tela / Modo de Espera:** Desative o bloqueio de tela ou o modo de espera do seu computador durante a execução do script, pois eles interromperão a automação do `pyautogui`.
*   **Erros:** O script possui tratamento básico de erros, mas falhas na conexão, mudanças inesperadas no WhatsApp Web ou outros problemas podem ocorrer. Verifique o console para mensagens de erro.

## Solução de Problemas Comuns

*   **Erro "Arquivo não encontrado" ou "Planilha não encontrada":** Verifique se `NOME_ARQUIVO_EXCEL` e `NOME_PLANILHA` estão corretos no script e se o arquivo `.xlsx` está na mesma pasta do script `.py`.
*   **Script clica no lugar errado ou não clica:** As coordenadas `x_coord_enviar`, `y_coord_enviar` estão incorretas para sua tela. Siga as instruções na seção "Configuração" para encontrá-las e atualizá-las.
*   **Mensagem não é enviada / Aba não fecha / Script parece rápido demais:** Aumente os valores de `sleep()` no script, principalmente o primeiro após `webbrowser.open()`. Seu computador ou internet podem precisar de mais tempo. Verifique também se o layout do WhatsApp Web mudou (posição do botão Enviar).
*   **Aparece mensagem "Telefone inválido" no WhatsApp Web:** Verifique se os números na Coluna A estão no formato internacional correto (código do país + DDD + número) e não contêm caracteres extras.

---
