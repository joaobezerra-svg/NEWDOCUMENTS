@app.route('/api/processar', methods=['POST'])
def processar():
    try:
        data = request.json
        link = data.get('link')
        aba = data.get('aba')
        letra_escola = data.get('letra_escola')
        filtro_excluir = data.get('filtro_excluir')
        colunas_remover_str = data.get('colunas_remover', '')
        formato = data.get('formato')

        sheets_service, drive_service = get_services()
        sid = extrair_id(link)

        # Lendo todos os dados
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=sid, range=f"'{aba}'!A:ZZ").execute()
        linhas_todas = result.get('values', [])

        if len(linhas_todas) < 4:
            return {"error": "Planilha curta demais (menos de 4 linhas)"}, 400

        cabecalho_original = linhas_todas[3]

        # Índices
        idx_esc = ord(letra_escola.upper()) - ord('A')
        
        # --- MUDANÇA AQUI: Tratando múltiplos termos de exclusão ---
        termos_proibidos = [t.strip().upper() for t in filtro_excluir.split(',')] if filtro_excluir else []

        idx_remover = [int(x) for x in colunas_remover_str.split(',') if x.strip().isdigit()]
        colunas_que_ficam = [(i, c.strip()) for i, c in enumerate(cabecalho_original) if i not in idx_remover and c.strip()]

        grupos = defaultdict(list)

        for linha in linhas_todas[4:]:
            if not linha: continue

            # --- MUDANÇA AQUI: BUSCA GLOBAL EM TODA A LINHA ---
            # Transformamos a linha inteira em texto para verificar os termos proibidos (FALSE, #REF!, etc)
            texto_linha_completa = " ".join([str(celula).upper() for celula in linha])
            
            # Se encontrar qualquer termo proibido em QUALQUER coluna, pula a linha
            if any(termo in texto_linha_completa for termo in termos_proibidos if termo):
                continue

            # Verificação de segurança: Ignorar linhas onde o nome do servidor (ou dados principais) está vazio
            # (Substituindo a trava da coluna 33 que estava excluindo tudo)
            if len(linha) < 3 or not str(linha[0]).strip(): 
                continue

            # Monta a linha apenas com colunas que o usuário quer no documento
            dados_linha = []
            for i, _ in colunas_que_ficam:
                val = str(linha[i]).strip() if i < len(linha) else ""
                dados_linha.append(val)

            escola = str(linha[idx_esc]).strip() if idx_esc < len(linha) else "GERAL"
            grupos[escola].append(dados_linha)

        if not grupos:
             return {"error": f"❌ Nenhum dado encontrado. Verifique se os termos {termos_proibidos} estão excluindo todas as linhas ou se a aba '{aba}' está correta."}, 400

        # --- O RESTANTE DO CÓDIGO (GERAÇÃO WORD) SEGUE IGUAL ---
        # ... (mantenha o código de geração do docx que você já tem)
