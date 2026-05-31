import os
import sys
import io
import time
import random
import requests
import pandas as pd
from playwright.sync_api import Playwright, sync_playwright
from dotenv import load_dotenv

# ==============================================================================
# 1. CONFIGURAÇÕES GERAIS
# ==============================================================================
# Carrega usuário e senha do arquivo oculto .env para não vazar credenciais
load_dotenv()
USUARIO = os.getenv("USUARIO")
SENHA = os.getenv("SENHA")

# Links da nuvem e caminhos locais
URL_BASE = 'https://docs.google.com/spreadsheets/d/e/URLDOGOOGLESHEETS/pub?output=csv'
CAMINHO_FILA_LOCAL = "fila_automacao.xlsx"

# Dicionário para corrigir a mira da combobox caso o Protheus tenha itens duplicados (ex: descer a seta 1x)
REGRAS_SETA_STATUS = {
    "22": 1, "33": 2, "44": 3, "55": 4
}

# Tempos de espera para blindar o robô contra a lerdeza do servidor da Totvs
TIMEOUT_PADRAO_GLOBAL   = 90000   # Dá 1 minuto e meio de folga para telas comuns
TIMEOUT_LOGISTICA_LOTES = 180000  # Dá 3 minutos de folga quando a busca for pesada
TIMEOUT_ESTABILIZACAO   = 2000    # Respiro de 2 segundos para o visual da tela 

print(">>> Inicializando Motor de Automação com Fila de Controle Ativa...")

# ==============================================================================
# 2. CARGA E LIMPEZA
# ==============================================================================
def carregar_dados_online(url):
    """ Baixa a planilha do Google. """
    try:
        # Cria um número aleatório no final do link. 
        # manda a versão atualizada da planilha, ignorando o cache antigo.
        cache_buster = str(time.time()).replace('.', '') + str(random.randint(1000, 9999))
        url_atualizada = f"{url}&cache_buster={cache_buster}"
        
        print("  Baixando dados da planilha online (Quebra de Cache)...")
        response = requests.get(url_atualizada)
        response.raise_for_status() 

        # Transforma os dados em string na hora de ler.
        data = response.content.decode('utf-8')
        df = pd.read_csv(io.StringIO(data), dtype=str)
        df.columns = df.columns.str.strip() 
        
        # Limpa o campo 'titulo' e garante que tenha 9 dígitos preenchendo com zero.
        if 'titulo' in df.columns:
            df['titulo'] = df['titulo'].str.split('.').str[0].str.zfill(9) 
            
        return df 
    except Exception as e:
        print(f" Erro crítico ao acessar a planilha online: {e}")
        return None

# ==============================================================================
# 3. EXECUÇÃO PRINCIPAL
# ==============================================================================
def run(playwright: Playwright, df_dados, df_sincronizado, caminho_fila) -> None: 
    # Abre o navegador em modo visível. 
    #  atrasa todos os cliques do robô para dar tempo pro Protheus.
    browser = playwright.chromium.launch(headless=False, slow_mo=1400)
    context = browser.new_context()
    page = context.new_page()
    page.set_default_timeout(TIMEOUT_PADRAO_GLOBAL)

    lista_sucesso = []
    lista_erro = []

    try:
        # --- ETAPA DE LOGIN E AMBIENTE ---
        print("\n--- Acessando o Protheus ---")
        page.goto("https://URLDOPROTHEUSWEB")
        page.get_by_role("group", name="Ambiente no servidor").get_by_role("combobox").select_option("CS86R7_PROD")
        page.get_by_role("button", name="Ok").click()

        time.sleep(5) 
        
        #foca nessa "janela interna" antes de digitar.
        frame = page.frame_locator("iframe").first
        frame.get_by_role("textbox", name="Insira seu usuário").fill(USUARIO)
        frame.get_by_role("textbox", name="Insira sua senha").fill(SENHA)
        frame.get_by_role("button", name="Entrar").click()
        frame.get_by_role("button", name="Entrar").click()

        # --- NAVEGAÇÃO ATÉ A ROTINA ---
        page.wait_for_selector("text=Recentes", timeout=TIMEOUT_PADRAO_GLOBAL)
        page.get_by_text("Recentes", exact=True).click()
        page.get_by_title("Tela Atendimento", exact=True).click()
        
        print("  Aguardando carregamento da rotina...")
        page.get_by_role("button", name="Confirmar").click()
        time.sleep(12) # Pausa longa porque o carregamento inicial da rotina é pesado
        
        # Se aparecer aquela telinha de filtro indesejada, já fecha ela
        if page.get_by_role("button", name="Cancelar").is_visible():
            page.get_by_role("button", name="Cancelar").click()

        # ======================================================================
        # LOOP DE PROCESSAMENTO (Lê linha por linha da planilha)
        # ======================================================================
        for i, linha in df_dados.iterrows():
            chaveID_csv    = linha['ChaveUnica']
            titulo_csv     = linha['titulo']
            grupo_csv      = linha['grupo_economico']
            carteira_csv   = linha['carteira']
            status_csv     = linha['tipo']
            hist_csv       = linha['historico']
            retirada_csv   = linha['retirada_status']
            bloqueado_csv  = linha['bloqueado_status']
            negativado_csv = linha['negativado_status']
            
            print(f"\n  ITEM {i+1}/{len(df_dados)} --> ID: {chaveID_csv} | Grupo: {grupo_csv}")
            
            try:
                # 1. Abre a tela de inclusão do atendimento
                btn_incluir = page.get_by_role("button", name="Incluir Atend.")
                btn_incluir.wait_for(state="visible")
                btn_incluir.click()
        
                print("       Preenchendo parâmetros de Grupo Economico/Carteira...")
                page.locator("#COMP6003").get_by_role("textbox").first.fill(grupo_csv)
                page.locator("#COMP6013").get_by_role("textbox").first.fill(carteira_csv)
                
                btn_ok = page.get_by_role("button", name="Ok")
                btn_ok.wait_for(state="visible")
                btn_ok.click()
                
                # --- BUSCA DOS TÍTULOS ESPECÍFICOS ---
                try:
                    print("       Acessando painel de filtros de títulos...")
                    btn_confirmar = page.get_by_role("button", name="Confirmar")
                    btn_confirmar.wait_for(state="visible", timeout=TIMEOUT_LOGISTICA_LOTES)
                    time.sleep(TIMEOUT_ESTABILIZACAO / 1000) 
                    
                    page.locator("#COMP6046").get_by_role("button", name="Filtrar").first.click()
                    page.get_by_text("Num. Titulo").click()
                    page.locator("#COMP7518").get_by_role("button", name="Aplicar filtros selecionados").first.click()
                    
                    # Quebra os títulos que vieram separados por vírgula na planilha
                    titulos_da_linha = [t.strip() for t in str(linha['titulo']).split(',')]
                    indice_t = 0
                    print(f"        Limpando slots antigos e aplicando {len(titulos_da_linha)} títulos...")
                    
                    # Varre cada id (os IDs andam de 2 em 2: 10503, 10505...)
                    for comp_id in range(10503, 10549, 2):
                        campo = page.locator(f"#COMP{comp_id}").get_by_role("textbox").first
                        # Se ainda tiver título nosso pra preencher, ele digita
                        if indice_t < len(titulos_da_linha):
                            campo.fill(titulos_da_linha[indice_t])
                            indice_t += 1
                        # Se a nossa lista acabou, ele limpa
                        else:
                            if campo.input_value() != "":
                                campo.fill("")  
                                page.keyboard.press("Backspace") 
                                
                    print("       Confirmando lote de filtros aplicados...")
                    page.locator("#COMP10550").get_by_role("button", name="Confirmar").first.click()
                    
                    print("       Selecionando todos os registros localizados...")
                    page.locator("#COMP6075").get_by_role("button", name="Marcar Todos").first.click()
                    page.locator("#COMP6021").get_by_role("button", name="Confirmar").first.click()
                    page.get_by_role("button", name="Sair da página").click()
            
                    print(" Aguardando processamento das tabelas internas (5s)...")
                    page.wait_for_timeout(5000) 
                    
                    # Prepara a tela final de alteração
                    page.locator("#COMP4578").get_by_role("button", name="Marcar Todos").first.click() 
                    page.locator("#COMP4576").get_by_role("button", name="Alterar").click()
                    
                    # ------------------------------------------------------------------
                    # 2. SELEÇÃO DO STATUS NA COMBOBOX 
                    # ------------------------------------------------------------------
                    status_esperado = str(status_csv).strip()
                    campo_status = page.locator("#COMP6002")
                    sucesso_selecao = False
                    
                    # Escalamos até 5 tentativas da combobox
                    for tentativa in range(1, 6):
                        if tentativa == 1:
                            tática_desc = "Digitação Padrão (Delay 50ms)"
                            delay_tecla = 50
                        elif tentativa == 2:
                            tática_desc = "Replay de Segurança (Delay 10ms)"
                            delay_tecla = 10 
                        elif tentativa == 3:
                            tática_desc = "Digitação Rápida (Delay 20ms)"
                            delay_tecla = 20
                        elif tentativa == 4:
                            tática_desc = "Digitação Cadenciada (Delay 30ms)"
                            delay_tecla = 30
                        else:
                            tática_desc = "Digitação Jato (Delay 10ms)"
                            delay_tecla = 10
                            
                        print(f"         [Tentativa {tentativa}/5] {tática_desc} para Status {status_esperado}...")
                        
                        try:
                            campo_status.click()
                            page.wait_for_timeout(400) 
                            
                            #  digita de forma contínua em um só fluxo de dados pra combobox não engasgar
                            page.keyboard.type(status_esperado, delay=delay_tecla)
                            
                            page.wait_for_timeout(600) 
                            
                            # Correção de bugs visuais da Totvs para itens que duplicam
                            if status_esperado in REGRAS_SETA_STATUS:
                                descidas = REGRAS_SETA_STATUS[status_esperado]
                                for _ in range(descidas):
                                    page.keyboard.press("ArrowDown")
                                    page.wait_for_timeout(150) 
                            
                            page.keyboard.press("Enter")
                            page.wait_for_timeout(300) 
                            
                            # Tab consolida a informação
                            page.keyboard.press("Tab")
                            page.wait_for_timeout(800) 
                            
                            # ==========================================================
                            # VALIDAÇÃO CIRÚRGICA 
                            # ==========================================================
                            sucesso_visual = False
                            padrao = f"{status_esperado}-"
                            
                            try:
                                # Pega todos os inputs da tela, mesmo os escondidos
                                todos_inputs = page.locator("input, textarea")
                                
                                # Entra dentro do HTML via JavaScript e extrai os valores já digitados (.value)
                                valores_caixas = todos_inputs.evaluate_all("(elementos) => elementos.map(el => el.value || '')")
                                
                                # Verifica se o nosso status fixou em alguma caixa de texto validamente
                                for valor in valores_caixas:
                                    if str(valor).strip().startswith(padrao):
                                        sucesso_visual = True
                                        break
                            except:
                                sucesso_visual = False
                            
                            if sucesso_visual:
                                print(f"         [OK] Validado Real! O status '{status_esperado}' foi lido dentro da caixa de texto.")
                                sucesso_selecao = True
                                break # Achou o texto, quebra o loop de tentativas e segue a vida
                            else:
                                print(f"         [AVISO] ERRO DETECTADO! O número '{status_esperado}-' não cravou na caixa.")
                                
                                # Fecha a tela de alteração inteira
                                page.keyboard.press("Escape") 
                                page.wait_for_timeout(1000) # Aguarda a janelinha fechar
                                
                                # Clica no botão "Alterar" na grid para abrir tudo do zero
                                print("         [INFO] Reabrindo a tela de Alteração para tentar novamente...")
                                page.locator("#COMP4576").get_by_role("button", name="Alterar").click()
                                page.wait_for_timeout(1500) # Aguarda a tela renderizar antes da nova tentativa
                                
                        except Exception as e:
                                print(f"         [AVISO] Erro técnico na execução: {str(e)}")
                            
                                # Faz o mesmo processo de reset se der erro de código
                                page.keyboard.press("Escape")
                                page.wait_for_timeout(1000)
                                
                                print("         [INFO] Reabrindo a tela de Alteração para tentar novamente...")
                                page.locator("#COMP4576").get_by_role("button", name="Alterar").click()
                                page.wait_for_timeout(1500)

                    if not sucesso_selecao:
                        raise Exception(f"Interface do Protheus rejeitou a seleção do status {status_esperado} após 5 tentativas estruturadas.")
                    # ------------------------------------------------------------------

                    page.keyboard.press("Tab") 
                    
                    # 3. Flags Adicionais e Histórico
                    print(f"       Configurando Flags -> Retirada: {retirada_csv} | Bloqueado: {bloqueado_csv} | Negativado: {negativado_csv}")
                    page.locator("#COMP6004").get_by_role("radio", name=f"-{'Sim' if retirada_csv.upper() == 'SIM' else 'Nao'}").check()
                    page.locator("#COMP6012").get_by_role("radio", name=f"-{'Sim' if bloqueado_csv.upper() == 'SIM' else 'Nao'}").check()
                    page.locator("#COMP6006").get_by_role("radio", name=f"-{'Sim' if negativado_csv.upper() == 'SIM' else 'Nao'}").check()
                        
                    print("       Gravando descrição do Histórico...")
                    page.locator("#COMP6007").get_by_role("textbox").click()
                    page.locator("#COMP6007").get_by_role("textbox").fill(hist_csv)
                    
                    # 4. Confirmação Final pra gravar no banco de dados da Totvs
                    page.get_by_role("button", name="Confirmar").click()
                    page.get_by_role("button", name="Confirmar").click()
                    
                    print(f"         ID: {chaveID_csv} processado com sucesso.")
                    lista_sucesso.append(chaveID_csv) 
                    time.sleep(3) 

                    # ------------------------------------------------------------------
                    # GRAVAÇÃO IMEDIATA NA PLANILHA LOCAL 
                    # ------------------------------------------------------------------
                    # Assim que acaba um cliente, já carimba 'SIM' na planilha do computador. 
                    df_sincronizado.loc[df_sincronizado['ChaveUnica'] == chaveID_csv, 'PROCESSADO'] = 'SIM'
                    df_sincronizado.to_excel(caminho_fila, index=False)
                    print(f"        Planilha local de controle atualizada.")
                    # ------------------------------------------------------------------

                    time.sleep(3) 

                except Exception as e_intern:
                    print(f"         Falha na Grid de Títulos: {str(e_intern)[:50]}")
                    lista_erro.append(f"{chaveID_csv} | Erro na Grid: {str(e_intern)[:50]}")
                    page.pause() # Deixa a tela congelada ver o que deu errado manualmente
                    page.keyboard.press("Escape") 
                    continue # Pula pra próxima linha da planilha

            except Exception as e_titulo:
                print(f"     Erro Crítico no ID {chaveID_csv}: {str(e_titulo)[:50]}")
                lista_erro.append(f"{chaveID_csv} | Erro Geral: {str(e_titulo)[:50]}")
                page.pause()
                page.keyboard.press("Escape")
                continue

    except Exception as e:
        print(f" Erro Crítico Operacional Playwright: {e}")
    
    finally:
            # ==============================================================================
            # ENCERRAMENTO E RELATÓRIOS 
            # ==============================================================================
            print("\n" + "="*40)
            print(" RELATÓRIO DE EXECUÇÃO DO RPA:")
            print(f"     Sucessos confirmados: {len(lista_sucesso)}")
            print(f"     Falhas registradas:  {len(lista_erro)}")
            print("="*40)
            
            # Cria um carimbo de tempo para separar as execuções no bloco de notas
            carimbo_tempo = time.strftime("%d/%m/%Y %H:%M:%S") # <-- APAGAR (Linha Mãe que cria a variável)
            
            # Usa o "a" (append) para não apagar os dados antigos
            try:
                if len(lista_sucesso) > 0:
                    with open("log_sucessos.txt", "a", encoding="utf-8") as f_sucesso:
                        f_sucesso.write(f"\n--- Execução: {carimbo_tempo} ---\n") # <-- APAGAR (Tenta usar a variável)
                        for chave in lista_sucesso: f_sucesso.write(f"{chave}\n")
            except Exception as e: print(f"Erro ao salvar log de sucesso: {e}")

            try:
                if len(lista_erro) > 0:
                    with open("log_erros.txt", "a", encoding="utf-8") as f_erro:
                        f_erro.write(f"\n--- Execução: {carimbo_tempo} ---\n") # <-- APAGAR (Tenta usar a variável)
                        for erro in lista_erro: f_erro.write(f"{erro}\n")
            except Exception as e: print(f"Erro ao salvar log de erros: {e}")

            input("\nPressione Enter para encerrar e fechar o navegador...")
            browser.close()
# ==============================================================================
# 4. ORQUESTRAÇÃO DE INÍCIO 
# ==============================================================================
if __name__ == "__main__":
    # Puxa os dados da nuvem do Google
    dados_fonte = carregar_dados_online(URL_BASE)
    
    if dados_fonte is not None:
        try:
            # Verifica se já tem um arquivo da fila no PC
            if os.path.exists(CAMINHO_FILA_LOCAL):
                print("  Lendo arquivo de controle local (fila_automacao.xlsx)...")
                dados_controle = pd.read_excel(CAMINHO_FILA_LOCAL, dtype=str)
                dados_controle.columns = dados_controle.columns.str.strip()
            # Se for a primeira vez rodando, cria a planilha local do zero
            else:
                print("  Criando novo arquivo de controle local...")
                dados_controle = dados_fonte.copy()
                dados_controle['PROCESSADO'] = 'NAO'

            # Garante que as chaves fiquem no mesmo padrão de texto pra não dar bug de comparação
            dados_fonte['ChaveUnica'] = dados_fonte['ChaveUnica'].astype(str).str.strip()
            dados_controle['ChaveUnica'] = dados_controle['ChaveUnica'].astype(str).str.strip()

            # Cruza a nuvem com o PC (trazendo a coluna de processos pra tabela nova)
            dados_sincronizados = dados_fonte.merge(
                dados_controle[['ChaveUnica', 'PROCESSADO']], 
                on='ChaveUnica', 
                how='left'
            )
            
            # Se entrou item novo que não tem status, carimba como 'NAO'
            dados_sincronizados['PROCESSADO'] = dados_sincronizados['PROCESSADO'].fillna('NAO')
            dados_sincronizados.loc[dados_sincronizados['PROCESSADO'].str.strip() == '', 'PROCESSADO'] = 'NAO'

            # Separa só quem falta trabalhar
            dados_a_processar = dados_sincronizados[dados_sincronizados['PROCESSADO'].str.upper() != 'SIM'].copy()
            
            print(f"  Fila Sincronizada! Total Nuvem: {len(dados_sincronizados)} | Falta processar: {len(dados_a_processar)}")

            # Dispara o Playwright se tiver trabalho a fazer
            if len(dados_a_processar) > 0:
                with sync_playwright() as playwright:
                    run(playwright, dados_a_processar, dados_sincronizados, CAMINHO_FILA_LOCAL)
            else:
                print("  Tudo pronto! Não há lotes novos. Todos os itens da planilha já estão como PROCESSADO = 'SIM'.")

        except Exception as e_fila:
            print(f"  Erro crítico na organização da fila de dados: {e_fila}")
            sys.exit() 
    else:
        print("  Operação cancelada: Falha ao ler banco de dados online.")
