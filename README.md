RPA Protheus - Automação de Cadastro de Históricos de Cobrança 🤖
📋 Contexto e Problema (O Cenário Manual)
Antes desta automação, o processo de registrar o histórico de cobrança e negociação dos títulos era 100% manual.

Cada usuário da equipe precisava:

Abrir o Protheus.

Localizar o Grupo Econômico e a Carteira.

Filtrar título por título manualmente.

Digitar ou copiar/colar o histórico do que foi conversado com o cliente.

Alterar os status de Retirada, Bloqueio ou Negativação individualmente.

Resultado: Um processo extremamente lento, sujeito a erros de digitação e que consumia horas produtivas que poderiam ser usadas para a negociação real com o cliente.

💡 A Solução: O Processo Automático
Com este RPA, o fluxo de trabalho foi redesenhado. Agora, o usuário foca apenas na parte humana: a cobrança.

Entrada de Dados: O usuário registra o resumo da conversa (histórico) e os status decididos em uma planilha centralizada no Google Sheets.

Execução: O robô lê essas informações e realiza todo o trabalho pesado de navegação e preenchimento no ERP Protheus.

Ganho de Eficiência: O que levava minutos por cliente agora é feito em segundos, com precisão cirúrgica e sem fadiga humana.

🧠 Lógica de Funcionamento do Código
1. Inteligência de Dados (Sincronização)
O script utiliza uma Sincronização Híbrida. Ele baixa os dados da nuvem mas mantém uma fila_automacao.xlsx local. Isso garante que, se o processo for interrompido, o robô saiba exatamente de onde parou, consultando a coluna PROCESSADO.

2. Visão Computacional e "Smart Wait"
O robô "enxerga" o Protheus através da biblioteca OpenCV. Em vez de usar esperas fixas que atrasam o processo, ele usa a função wait_for_image.

Exemplo: Após inserir a carteira e dar Enter, o robô monitora a tela. Assim que o botão "Confirmar" aparece (sinal de que o Protheus carregou os dados), o robô clica e segue. Se o sistema estiver lento, ele espera; se estiver rápido, ele voa.

3. Preenchimento Automático de Histórico
Para evitar erros de codificação de texto, o robô utiliza o Clipboard (Área de Transferência). Ele copia o histórico da planilha e "cola" dentro do Protheus, garantindo que o texto longo e detalhado da cobrança entre perfeitamente no ERP.

4. Segurança e Encerramento
Ao processar todos os clientes da fila, o robô salva o relatório final e executa o bloqueio da estação de trabalho (LockWorkStation). Isso garante que a conta do usuário não fique exposta após o término da tarefa.

🛠️ Tecnologias Utilizadas
Python: Linguagem base.

PyAutoGUI & OpenCV: Para a "mão" e os "olhos" do robô.

Pandas: Para o "cérebro" que gerencia as planilhas.

Requests: Para a comunicação com a planilha online.
