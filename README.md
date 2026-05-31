# RPA Protheus - Automação de Cadastro de Históricos de Cobrança 

## 📋 Contexto e Problema (O Cenário Manual)
Antes desta automação, o processo de registrar o histórico de cobrança e negociação dos títulos era 100% manual. 
Cada usuário da equipe precisava:
1. Abrir o Protheus.
2. Entrar na rotina específica 
3. Filtra pelo Grupo Econômico e a Carteira.
4. Filtra título por título manualmente.
5. Digitar ou copiar/colar o histórico do que foi conversado com o cliente.
6. Alterar os status de Retirada, Bloqueio ou Negativação individualmente.

**Resultado:** Um processo extremamente lento, sujeito a erros de digitação e que consumia horas produtivas que poderiam ser usadas para a negociação real com o cliente.

## 💡 A Solução: O Processo Automático
Com este RPA, o fluxo de trabalho foi redesenhado. Agora, o usuário foca apenas na parte humana: a cobrança.
* **Entrada de Dados:** O usuário utilizava uma planilha personalizada, que cria os modelos de envio de email para o cliente, após cada tratativa, o registro é enviado para uma planilha na nuvem (Google Sheets).
* **Execução:** O robô lê essas informações e realiza todo o trabalho pesado de navegação e preenchimento no ERP Protheus.
* **Ganho de Eficiência:** O que levava minutos por cliente agora é feito em segundos, com precisão cirúrgica e sem fadiga humana.

---

## 🚀 A Evolução Arquitetural (Duas Abordagens)

Este repositório contém duas versões distintas do projeto, demonstrando a evolução na arquitetura da automação para lidar com os bloqueios de segurança e instabilidades do ERP Totvs Protheus:

### 1. Versão Desktop (PyAutoGUI + OpenCV)
* **Local:** Pasta `google.py/`
* **Abordagem:** O robô "enxerga" o Protheus através de Visão Computacional (`wait_for_image`) e usa o teclado/mouse físicos do sistema operacional. Ele gerencia textos longos utilizando a Área de Transferência (Clipboard) para copiar da planilha e colar no sistema.
* **Desafio:** Exige que a tela fique 100% visível, bloqueia o uso do computador pelo usuário durante a execução e exige o bloqueio automático da máquina (`LockWorkStation`) ao final por segurança.

### 2. Versão Web Avançada (Playwright) - *Versão de Produção*
* **Local:** Pasta `RegistroHistoricoWeb/`
* **Abordagem:** Opera diretamente na camada de rede e no DOM (HTML) do navegador. 
* **Vantagem Absoluta:** Roda em segundo plano, não sequestra o mouse/teclado do usuário e fura as blindagens de segurança avançadas do Protheus Web (como múltiplos Iframes e Shadow DOM).

---
🧠 Lógica de Funcionamento e Passo a Passo (Versão PyAutoGui)

### 1. Sincronização Híbrida de Dados
Download da Nuvem: O script faz o download dos dados mais recentes da planilha do Google Sheets (utilizando um mecanismo de quebra de cache para garantir dados em tempo real).

Validação Local: Ele cruza os dados da nuvem com o arquivo fila_automacao.xlsx local. O robô verifica a coluna PROCESSADO e filtra a lista para gerar o DataFrame dados_a_processar, garantindo que nenhum registro seja duplicado se o script for reiniciado.

### 2. Inclusão do Atendimento (Etapa 1)
Abertura: O robô localiza e clica no botão Incluir Atendimento via reconhecimento de imagem.

Preenchimento: Digita o código do Grupo Econômico, navega pelos campos usando a tecla TAB de forma cadenciada e insere a Carteira.

Sincronismo Visial (Smart Wait): Após pressionar ENTER, o robô utiliza a função wait_for_image para monitorar a tela e aguardar a aparição do botão Confirmar. O código fica retido até que o Protheus termine de carregar os dados.

### 3. Filtragem Dinâmica de Títulos (Etapa 2)
Abertura do Filtro: O robô abre a janela "Filtro através de perguntas".

Foco e Limpeza Inteligente: O script aguarda o cabeçalho da janela aparecer, clica no primeiro campo para dar foco e inicia o preenchimento.

Otimização de Preenchimento: O robô digita os títulos da lista nos primeiros campos disponíveis. Em seguida, ele calcula quantos campos restaram dos 23 totais e dispara o comando DELETE + TAB apenas nas linhas que sobraram em branco, economizando tempo de execução.

### 4. Atualização de Status e Histórico (Etapa 3)
Seleção: Clica em Marcar Todos e avança para a tela de alteração de status.

Regra de Negócio (Dropdown Tipo): O robô identifica o Tipo de atendimento. Se for um dos códigos problemáticos (como 22, 33, 44, 55), ele executa uma lógica de setas para baixo (down) para vencer o bug de travamento do Protheus.

Checkboxes Visuais: O robô lê as colunas de status da planilha (retirada_status, bloqueado_status, negativado_status) e move o mouse para clicar nas coordenadas exatas de SIM ou NAO na tela.

Injeção do Histórico: O texto que o usuário cobrou do cliente é copiado para a Área de Transferência do Windows (Clipboard) e colado via Ctrl + V dentro do campo de histórico do ERP, preservando caracteres especiais e textos longos.

### 5. Gravação e Fechamento de Ciclo
Confirmação: O robô clica no botão final de salvar e aguarda o fechamento da janela.

Gravação do Sucesso: A linha correspondente daquele Grupo Econômico recebe o status "SIM" na coluna PROCESSADO e o arquivo fila_automacao.xlsx é salvo imediatamente no HD.

Segurança: Ao finalizar todas as linhas de dados pendentes, o script executa um comando de sistema que realiza o bloqueio imediato da estação de trabalho do Windows (LockWorkStation).
---

## 🧠 Lógica de Funcionamento e Passo a Passo (Versão Web)

### 1. Orquestração e Inteligência de Fila
Antes de abrir o navegador, o script realiza uma chamada que burla o cache do Google Sheets para baixar os dados da nuvem em tempo real. Em seguida, ele cruza essas informações com uma planilha local (`fila_automacao.xlsx`). O robô separa apenas os clientes que estão com a coluna `PROCESSADO` marcada como `'NAO'`. Isso garante a **Idempotência**: se faltar luz ou a internet cair no meio do processo, o robô sabe exatamente de onde parou sem duplicar dados.

### 2. Autenticação e Navegação
O navegador Chromium é iniciado com um ritmo compassado (`slow_mo`) para respeitar o tempo de resposta do ERP. O robô foca no Iframe correto de autenticação, insere as credenciais cadastradas em variáveis de ambiente seguras (`.env`) e navega pelo menu de favoritos até carregar por completo a "Tela Atendimento".

### 3. Filtros Macro (Grupo e Carteira)
Para cada linha pendente da fila, o robô clica no botão principal de ID único (`Incluir Atend.`) e preenche os campos iniciais de **Grupo Econômico** e **Carteira** usando seletores diretos por ID (como `#COMP6003`), avançando para carregar a base de dados do cliente.

### 4. Filtro Cirúrgico por Títulos
Para garantir que apenas os títulos corretos sejam afetados, o robô acessa a janela de filtros secundários. Ele executa uma rotina de **higienização de slots**, apagando qualquer resquício de títulos digitados em execuções anteriores (aplicando *backspaces* via teclado virtual), distribui dinamicamente os novos títulos nas caixas vazias do Protheus, confirma o filtro,  marca todos os registros localizados e abre a tela de Alteração final.

### 5. Preenchimento Seguro e Fura-Blindagem (Combobox de Status)
Esta é a etapa de maior complexidade técnica, projetada para vencer os comportamentos de interface da Totvs:
* **Digitação Anti-Buffer:** Listas suspensas do Protheus esquecem o que foi digitado se houver pausas. O código resolve isso aplicando digitação contínua via `page.keyboard.type()` com micro-delays (variando de 10ms a 70ms entre as tentativas) para que a combobox absorva a dezena inteira (ex: "41") de uma vez só,Se for um dos códigos problemáticos (como 22, 33, 44, 55), ele executa uma lógica de setas para baixo (down) para vencer o bug de travamento do Protheus.
* **Leitura Shadow DOM:** O sistema esconde o texto selecionado em propriedades nativas inacessíveis por seletores visuais comuns. O robô injeta um script JavaScript (`evaluate_all`) que fura a blindagem e varre o valor `.value` de todas as caixas de texto da tela para validar se o status esperado foi realmente gravado.
* **Self-Healing (Auto-Recuperação):** Se a validação detectar que o Protheus rejeitou o status, o robô cancela a ação enviando um comando `Escape` para fechar a tela suja, clica no botão "Alterar" (`#COMP4576`) para reabrir o registro do zero com o campo limpo, e tenta novamente. O loop repete esse reset por até 5 tentativas escalonadas antes de reportar erro.
* **Consolidação:** Após fixar o status, o robô altera os Radio Buttons (Retirada, Bloqueado, Negativado) e preenche a caixa do **Histórico** com o texto que o usuraio colocou na tratativa.

### 6. Gravação e Auditoria
O robô clica em confirmar para enviar as informações ao banco de dados do ERP. No exato segundo em que o Protheus valida a gravação, o script atualiza a planilha local para `PROCESSADO = 'SIM'`. Ao final do lote, o robô gera relatórios incrementais em arquivos `.txt` (usando o modo *append* para acumular o histórico eterno das operações), separando os IDs processados com sucesso e as falhas técnicas para auditoria.

---

## 🛠️ Tecnologias Utilizadas

* **Python 3.10+**
* **Playwright:** Controle, injeção de scripts e automação web direta.
* **PyAutoGUI & OpenCV:** Automação de interface baseada em coordenadas e visão computacional (versão desktop).
* **Pandas:** Biblioteca de análise de dados usada como o cérebro lógico para cruzamento e sincronização de tabelas.
* **Requests:** Comunicação assíncrona com a API pública do Google Sheets.
* **Python-dotenv:** Gerenciamento seguro de credenciais locais para prevenir vazamento de senhas no código fonte.


