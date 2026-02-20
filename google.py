#pip install pandas openpyxl pyperclip opencv-python

import pyautogui as pa
import time
import pandas as pd
import os
import sys
import pyperclip
import cv2
import requests 
import io      
import random  
import time as tm 

# Configurações globais de segurança e pausa
pa.FAILSAFE = False
pa.PAUSE = 2.2

# --- DEFINIÇÃO DE CAMINHOS ---
# 1. Arquivo FONTE (Onde os usuários inserem dados)
URL_BASE = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTV4GWTD4qpCEnZlaKJtxqHkg3KB3fpITO-tzb54WkeWJqhRWM-jDnQwcx8W5KDyEPAiL7W3XI2mD83/pub?output=csv'

# 2. Encontra o caminho completo da pasta ONDE O SCRIPT ESTÁ ('ProdApp')
caminho_script = os.path.dirname(os.path.abspath(sys.argv[0]))

# 3. Define o caminho COMPLETO para o arquivo FONTE e FILA
CAMINHO_FILA_LOCAL = os.path.join(caminho_script, 'fila_automacao.xlsx')

# 4. Caminho de Imagens (Mantido)
caminho_base_py = os.path.dirname(caminho_script)
ASSETS_PATH = os.path.join(caminho_base_py, 'Prints')

print(f"DEBUG: Fonte de Dados: GOOGLE URL")
print(f"DEBUG: Fila de Controle: {CAMINHO_FILA_LOCAL}")
print(f"DEBUG: O script procurará imagens em: {ASSETS_PATH}")

#########################################
#Função para clique seguro (mantida)
def click_on_image(image_path, confidence=0.90, action_name="Elemento", num_clicks=1,grayscale=False,region=None):
    """ Busca imagem (mantido para evitar quebras) """
    try:
        full_path = os.path.join(ASSETS_PATH, image_path)
        localizacao = pa.locateCenterOnScreen(full_path, confidence=confidence, grayscale=grayscale, region=region)

        if localizacao is not None:
            pa.moveTo(localizacao); pa.click(localizacao, clicks=num_clicks)
            return True
        else:
            print(f"ALERTA CRÍTICO: O elemento '{action_name}' não foi encontrado! Caminho: {full_path}")
            sys.exit(1)

    except Exception as e:
        print(f"ERRO CRÍTICO ao buscar ou clicar em '{action_name}': {e}")
        sys.exit(1)

#########################################

# --- INÍCIO DO BLOCO DE LEITURA E SINCRONIZAÇÃO HÍBRIDA ---
try:
   # 1. GERA UM VALOR ÚNICO PARA QUEBRAR O CACHE
    cache_buster = str(tm.time()).replace('.', '') + str(random.randint(1000, 9999))
    PLANILHA_URL_CACHE = f"{URL_BASE}&cache_buster={cache_buster}"
    
    # 2. FAZ O DOWNLOAD DO CONTEÚDO CSV DA NUVEM (sempre a versão mais atualizada)
    print("Baixando dados da planilha online (Quebra de Cache)...")
    response = requests.get(PLANILHA_URL_CACHE)
    response.raise_for_status() # Garante que o download foi bem-sucedido

    # 3. LÊ O CONTEÚDO CSV ONLINE
    data = response.content.decode('utf-8')
    dados_fonte = pd.read_csv( # <-- MUDOU DE read_excel para read_csv
        io.StringIO(data),
        dtype=str # Lê tudo como string primeiro para segurança
    ).fillna('')

    # ### DIAGNÓSTICO ADICIONADO AQUI ###
    print("### COLUNAS LIDAS DA FONTE (dados_automacao.xlsx):", list(dados_fonte.columns))
    if 'ChaveUnica' not in dados_fonte.columns:
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print("ERRO CRÍTICO: Coluna 'ChaveUnica' NÃO encontrada no arquivo FONTE!")
        print("Verifique o cabeçalho em dados_automacao.xlsx")
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        sys.exit() # Para o script aqui se a coluna chave estiver faltando
    # ### FIM DO DIAGNÓSTICO FONTE ###

    # 2. LÊ O ARQUIVO DE CONTROLE (FILA)
    if os.path.exists(CAMINHO_FILA_LOCAL):
        dados_controle = pd.read_excel(CAMINHO_FILA_LOCAL, dtype=str).fillna('')
        # ### DIAGNÓSTICO ADICIONADO AQUI ###
        print("### COLUNAS LIDAS DO CONTROLE (fila_automacao.xlsx):", list(dados_controle.columns))
        if 'ChaveUnica' not in dados_controle.columns:
            print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            print("ERRO CRÍTICO: Coluna 'ChaveUnica' NÃO encontrada no arquivo de CONTROLE!")
            print("Verifique o cabeçalho em fila_automacao.xlsx")
            print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            # Não paramos aqui, pois o controle pode ser recriado
        # ### FIM DO DIAGNÓSTICO CONTROLE ###
    else:
        # Se a Fila não existe, usa a Fonte como base para a criação
        dados_controle = dados_fonte.copy()
        if 'PROCESSADO' not in dados_controle.columns:
             dados_controle['PROCESSADO'] = '' # Garante que a coluna exista

    # Garante que a ChaveUnica é string em ambos para o merge
    dados_fonte['ChaveUnica'] = dados_fonte['ChaveUnica'].astype(str).str.strip()
    dados_controle['ChaveUnica'] = dados_controle['ChaveUnica'].astype(str).str.strip()

   # --- SUBSTITUA A PARTIR DA LINHA 4. SINCRONIZAÇÃO ---

    # 4. SINCRONIZAÇÃO: Faz o merge e prioriza o STATUS 'PROCESSADO' do arquivo de CONTROLE
    dados_sincronizados = dados_fonte.merge(
        dados_controle[['ChaveUnica', 'PROCESSADO']], 
        on='ChaveUnica', 
        how='left' 
        # Suffixes removidos pois não há conflito na coluna PROCESSADO
    )
    
    # 5. PREPARA O DATAFRAME FINAL
    # Preenche os registros novos (que estão como NaN) com 'NAO'
    dados_sincronizados['PROCESSADO'] = dados_sincronizados['PROCESSADO'].fillna('NAO')

    # Define o DataFrame que será usado no loop
    dados = dados_sincronizados.copy()
    
    # 6. FILTRA: Apenas linhas onde PROCESSADO NÃO é 'SIM'
    if 'PROCESSADO' in dados.columns:
        dados_a_processar = dados[dados['PROCESSADO'].str.upper() != 'SIM'].copy()
    else:
        print("AVISO: Coluna 'PROCESSADO' não encontrada. Processando todas as linhas.")
        dados_a_processar = dados.copy()

    print(f"Iniciando automação. {len(dados_a_processar)} grupos econômicos encontrados para processar.")

    # Loop principal que itera sobre cada linha do DataFrame FILTRADO
    for indice, linha in dados_a_processar.iterrows(): # Usamos o índice original do dados_a_processar
        # --- LEITURA E PREPARAÇÃO DOS DADOS ---
        chave_unica_do_registro = str(linha['ChaveUnica']).strip()

        carteira = str(linha['carteira']).strip()
        grupo_economico = str(linha['grupo_economico']).strip()
        titulos_string = str(linha['titulo']).strip()
        tipo = str(linha['tipo']).strip()
        historico = str(linha['historico']).strip()

        retirada_status = str(linha['retirada_status']).strip().upper().replace(' ', '').replace('ÃO', 'AO')
        bloqueado_status = str(linha['bloqueado_status']).strip().upper().replace(' ', '').replace('ÃO', 'AO')
        negativado_status = str(linha['negativado_status']).strip().upper().replace(' ', '').replace('ÃO', 'AO')

        lista_de_titulos = titulos_string.split(',')

        print(f"\nProcessando Grupo: {grupo_economico} ({len(lista_de_titulos)} títulos), Carteira: {carteira}")

        # --- Etapa 1: Inserir Grupo Econômico e Carteira ---
        time.sleep(2)
        click_on_image('botao_incluiratendimento.png', action_name='Botão Incluir Atendimento') # RESTAURADO
        pa.write(grupo_economico)
        pa.press("tab", presses=4)
        pa.write(carteira)
        pa.press("enter")
        time.sleep(180)   

        # --- Etapa 2: Loop INTERNO para filtrar cada título do grupo ---
        click_on_image('botao_filtrarhistorico.png', action_name='Botão Filtar Histórico') # RESTAURADO
        pa.click(x= 709, y= 480) #clica em num de títul
        pa.click(x= 709, y= 480)
        click_on_image('botao_aplicarfiltros.png',action_name="Botao Aplicar Filtros") # RESTAURADO

        ############################# limpeza 23 itens
        pa.press('delete')

        for _ in range(22):
            pa.press('delete',_pause=False)
            time.sleep(1.0)
            pa.press('tab',_pause=False)

        pa.press('delete')
        pa.press('tab')
        pa.press('tab')
        pa.press('tab')

        #######################
        for titulo in lista_de_titulos:
            pa.write(titulo.strip())
            time.sleep(0.2)

        pa.click(x= 1132, y= 727) # Confirmar
        


        # --- Etapa 3: Marcar, Alterar e Inserir Histórico ---
        click_on_image('botao_marcartodos.png',action_name="Botão Marcar Todos") # RESTAURADO
        click_on_image('botao_confirmar.png',action_name="Botão Confirmar") # RESTAURADO
        pa.moveTo(x= 1184, y= 649)
        pa.click(x= 1184, y= 649) # Sair da pagina (Coordenada Mantida)
        click_on_image('botao_marcartodos.png',action_name="Botão Marcar Todos") # RESTAURADO
        pa.moveTo(x= 1856, y= 398)
        pa.click(x= 1856, y= 398) # Alterar (Coordenada Mantida)

        # Inserir Tipo07
        pa.moveTo(x= 899, y= 387)
        pa.click(x= 899, y= 387)
        pa.write(tipo, interval=0.2)
        pa.press("tab")
        

        # Verificação de status (Retirada, Bloqueado, Negativado)
        pa.press("tab", presses=1)
        

        # 1. RETIRADA
        if retirada_status == 'SIM':
            pa.moveTo(x= 601, y= 453); pa.click(x= 601, y= 453); pa.click(x= 601, y= 453)
        elif retirada_status == 'NAO':
            pa.moveTo( x= 600, y= 476); pa.click( x= 600, y= 476); pa.click( x= 600, y= 476)

        # 2. BLOQUEADO
        if bloqueado_status == 'SIM':
            pa.moveTo(x=868, y= 453); pa.click(x=868,y= 453); pa.click(x=868, y= 453)
        elif bloqueado_status == 'NAO':
            pa.moveTo(x=867, y= 476); pa.click(x=867,y= 476); pa.click(x=867, y= 476)


        # 3. NEGATIVADO
        if negativado_status == 'SIM':
            pa.moveTo(x=1179, y= 450); pa.click(x=1179, y= 450); pa.click(x=1179, y= 450)
        elif negativado_status == 'NAO':
            pa.moveTo(x= 1179, y= 475); pa.click(x= 1179, y= 475); pa.click(x= 1179, y= 475)

        pa.press('tab', presses=2)
        time.sleep(0.5)


        # Inserir Histórico
        pa.click(x= 949, y= 675)
        pyperclip.copy(historico); pa.hotkey('ctrl', 'v');
      

        # Ação Final
        pa.press ("tab")
        pa.click(x= 1248, y= 801); pa.click(x= 1248, y= 801) #Botão de confirmar
        time.sleep(3)

        # --- BLOCO DE MARCAÇÃO E SALVAMENTO (VALIDAÇÃO) ---
        try:
            # 1. Marca a linha no DATAFRAME SINCRONIZADO usando a ChaveUnica
            dados_sincronizados.loc[dados_sincronizados['ChaveUnica'] == chave_unica_do_registro, 'PROCESSADO'] = 'SIM'

            # 2. SALVA O DATAFRAME COMPLETO DE VOLTA NA FILA DE CONTROLE
            dados_sincronizados.to_excel(CAMINHO_FILA_LOCAL, index=False)

            print(f"✅ SUCESSO: Registro '{chave_unica_do_registro}' marcado como PROCESSADO na fila de controle.")

        except Exception as e:
            print(f"ALERTA CRÍTICO: Falha ao salvar na fila de controle. Detalhe: {e}")

        time.sleep(5)

    print("\n\n#####################################################")
    print("✅ Automação concluída com sucesso! ✅")
    print(f"Verifique o arquivo '{CAMINHO_FILA_LOCAL}' para a lista completa de processados.")
    print("#####################################################")

    pa.hotkey('win', 'l')

except FileNotFoundError: # Este erro agora só se aplica à FILA LOCAL
    print(f"ERRO CRÍTICO: O arquivo de Fila Local não foi encontrado: '{CAMINHO_FILA_LOCAL}'")
except requests.exceptions.HTTPError as errh: # <-- ADICIONE ESTE BLOCO
    print(f"\nERRO HTTP na leitura da planilha online: {errh}")
    print("Verifique se a URL está correta e se a planilha está PUBLIADA COMO CSV.")
except requests.exceptions.ConnectionError as errc: # <-- ADICIONE ESTE BLOCO
    print(f"\nERRO DE CONEXÃO: {errc}")
    print("Verifique sua conexão com a internet.")
except Exception as e:
    print(f"\nERRO INESPERADO: O script foi interrompido. Detalhe: {e}")
    sys.exit()