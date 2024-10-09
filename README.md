# Automa-o-whatsApp
Aqui está um exemplo de README para o seu código:

---

# Automação de Envio de Mensagens no WhatsApp via Excel

Este script Python automatiza o envio de mensagens personalizadas pelo WhatsApp para uma lista de contatos contida em um arquivo Excel. Utilizando as bibliotecas `openpyxl`, `urllib`, `webbrowser` e `pyautogui`, o script lê os dados do Excel, monta uma mensagem de texto para cada contato e interage automaticamente com o WhatsApp Web para enviar a mensagem.

## Requisitos

- Python 3.x
- Bibliotecas Python:
  - `openpyxl`
  - `pyautogui`
- Navegador web (preferencialmente Google Chrome)
- Arquivo Excel contendo os contatos e seus números de telefone
- Imagem de referência da seta usada para iniciar a conversa no WhatsApp Web (`image.png`)

## Como funciona

1. O script lê um arquivo Excel (`clientes.xlsx`) que deve conter:
   - Na primeira coluna: o nome dos contatos
   - Na segunda coluna: o número de telefone dos contatos (incluindo o código do país, sem o "+")

2. Para cada contato válido (nome e telefone presentes), o script monta uma mensagem personalizada e gera um link para abrir o WhatsApp Web com o número de telefone do contato e a mensagem pré-preenchida.

3. O navegador abre automaticamente o link do WhatsApp Web, e o script usa a biblioteca `pyautogui` para localizar a seta de envio na tela e clicar nela, simulando o envio da mensagem.

4. Após o envio, o script fecha a aba do navegador e continua para o próximo contato.

## Como usar

1. Certifique-se de que todos os requisitos estão instalados e de que o arquivo Excel com os contatos está no mesmo diretório do script.

2. O arquivo Excel deve seguir o formato:
   - Coluna A: Nome
   - Coluna B: Número de telefone (com código de país)

3. Execute o script no terminal ou IDE de sua escolha:

   ```bash
   python enviar_mensagens.py
   ```

4. Antes de rodar o script, certifique-se de que a imagem de referência da seta (`image.png`) está salva corretamente no mesmo diretório. O script usará essa imagem para localizar o botão de envio de mensagens no WhatsApp Web.

## Observações

- Certifique-se de que o WhatsApp Web está configurado e autenticado no navegador antes de executar o script.
- O tempo de espera (`sleep`) pode ser ajustado conforme a velocidade da sua internet ou a performance do sistema.
- A imagem da seta (`image.png`) é um elemento essencial para que o `pyautogui` possa clicar no botão de envio no WhatsApp. Caso a interface do WhatsApp Web mude, pode ser necessário atualizar essa imagem.
  
## Melhorias Futuras

- Tratamento de exceções para garantir que erros de conexão ou problemas no carregamento da página do WhatsApp Web não interrompam o fluxo de execução.
- Implementação de logs para monitorar o andamento do envio das mensagens.
- Adicionar uma funcionalidade para ler e formatar corretamente números de telefone em diferentes formatos.

---

Isso oferece uma explicação detalhada sobre o funcionamento do script e como utilizá-lo.
