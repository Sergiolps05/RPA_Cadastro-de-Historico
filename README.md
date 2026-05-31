# RPA Protheus - Automação de Cadastro de Históricos de Cobrança 🤖

## 📋 Contexto e Problema (O Cenário Manual)
Antes desta automação, o processo de registrar o histórico de cobrança e negociação dos títulos era 100% manual. 
Cada usuário da equipe precisava:
1. Abrir o Protheus.
2. Localizar o Grupo Econômico e a Carteira.
3. Filtrar título por título manualmente.
4. Digitar ou copiar/colar o histórico do que foi conversado com o cliente.
5. Alterar os status de Retirada, Bloqueio ou Negativação individualmente.

**Resultado:** Um processo extremamente lento, sujeito a erros de digitação e que consumia horas produtivas que poderiam ser usadas para a negociação real com o cliente.

## 💡 A Solução: O Processo Automático
Com este RPA, o fluxo de trabalho foi redesenhado. Agora, o usuário foca apenas na parte humana: a cobrança.
* **Entrada de Dados:** O usuário registra o resumo da conversa (histórico) e os status decididos em uma planilha centralizada no Google Sheets.
* **Execução:** O robô lê essas informações e realiza todo o trabalho pesado de navegação e preenchimento no ERP Protheus.
* **Ganho de Eficiência:** O que levava minutos por cliente agora é feito em segundos, com precisão cirúrgica e sem fadiga humana.

---

## 🚀 A Evolução Arquitetural (Duas Abordagens)

Este repositório contém duas versões distintas do projeto, demonstrando a evolução na arquitetura da automação para lidar com os bloqueios de segurança e instabilidades do ERP Totvs Protheus:

### 1. Versão Desktop (PyAutoGUI + OpenCV)
* **Local:** Pasta `pyautogui_version/`
* **Abordagem:** O robô "enxerga" o Protheus através de Visão Computacional (`wait_for_image`) e usa o teclado/mouse físicos do sistema operacional. Ele gerencia textos longos utilizando a Área de Transferência (Clipboard) para copiar da planilha e colar no sistema.
* **Desafio:** Exige que a tela fique 100% visível, bloqueia o uso do computador pelo usuário durante a execução e exige o bloqueio automático da máquina (`LockWorkStation`) ao final por segurança.

### 2. Versão Web Avançada (Playwright) - *Versão de Produção*
* **Local:** Pasta `playwright_version/`
* **Abordagem:** Opera diretamente na camada de rede e no DOM (HTML) do navegador. 
* **Vantagem Absoluta:** Roda em segundo plano, não sequestra o mouse/teclado do usuário e fura as blindagens de segurança avançadas do Protheus Web (como múltiplos Iframes e Shadow DOM).

---

## 🧠 Lógica de Funcionamento e Passo a Passo (Versão Web)

O robô da versão de produção opera seguindo uma esteira de processamento rigorosa e resiliente. Abaixo está o ciclo de vida completo de uma execução:

### 1. Orquestração e Inteligência de Fila
Antes de abrir o navegador, o script realiza uma chamada que burla o cache do Google Sheets para baixar os dados da nuvem em tempo real. Em seguida, ele cruza essas informações com uma planilha local (`fila_automacao.xlsx`). O robô separa apenas os clientes que estão com a coluna `PROCESSADO` marcada como `'NAO'`. Isso garante a **Idempotência**: se faltar luz ou a internet cair no meio do processo, o robô sabe exatamente de onde parou sem duplicar dados.

### 2. Autenticação e Navegação
O navegador Chromium é iniciado com um ritmo compassado (`slow_mo`) para respeitar o tempo de resposta do ERP. O robô foca no Iframe correto de autenticação, insere as credenciais cadastradas em variáveis de ambiente seguras (`.env`) e navega pelo menu de favoritos até carregar por completo a "Tela Atendimento".

### 3. Filtros Macro (Grupo e Carteira)
Para cada linha pendente da fila, o robô clica no botão principal de ID único (`Incluir Atend.`) e preenche os campos iniciais de **Grupo Econômico** e **Carteira** usando seletores diretos por ID (como `#COMP6003`), avançando para carregar a base de dados do cliente.

### 4. Filtro Cirúrgico por Títulos
Para garantir que apenas os boletos corretos sejam afetados, o robô acessa a janela de filtros secundários. Ele executa uma rotina de **higienização de slots**, apagando qualquer resquício de títulos digitados em execuções anteriores (aplicando *backspaces* via teclado virtual), distribui dinamicamente os novos títulos nas caixas vazias do Protheus, marca todos os registros localizados e abre a tela de Alteração final.

### 5. Preenchimento Seguro e Fura-Blindagem (Combobox de Status)
Esta é a etapa de maior complexidade técnica, projetada para vencer os comportamentos de interface da Totvs:
* **Digitação Anti-Buffer:** Listas suspensas do Protheus esquecem o que foi digitado se houver pausas. O código resolve isso aplicando digitação contínua via `page.keyboard.type()` com micro-delays (variando de 10ms a 70ms entre as tentativas) para que a combobox absorva a dezena inteira (ex: "41") de uma vez só.
* **Leitura Shadow DOM:** O sistema esconde o texto selecionado em propriedades nativas inacessíveis por seletores visuais comuns. O robô injeta um script JavaScript (`evaluate_all`) que fura a blindagem e varre o valor `.value` de todas as caixas de texto da tela para validar se o status esperado foi realmente gravado.
* **Self-Healing (Auto-Recuperação):** Se a validação detectar que o Protheus rejeitou o status, o robô cancela a ação enviando um comando `Escape` para fechar a tela suja, clica no botão "Alterar" (`#COMP4576`) para reabrir o registro do zero com o campo limpo, e tenta novamente. O loop repete esse reset por até 5 tentativas escalonadas antes de reportar erro.
* **Consolidação:** Após fixar o status, o robô altera os Radio Buttons (Retirada, Bloqueado, Negativado) e preenche a caixa do **Histórico**.

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

## 💻 Como Rodar o Projeto

**1. Clone o repositório:**
```bash
git clone [https://github.com/seu-usuario/seu-repositorio.git](https://github.com/seu-usuario/seu-repositorio.git)
cd seu-repositorio
