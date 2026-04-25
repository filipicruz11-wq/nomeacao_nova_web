import datetime
import re
from openpyxl import Workbook

# =====================================================
# ESCALA FIXA CEJUSC (v6.3 - EQUILÍBRIO GLOBAL)
# =====================================================
escala_fixa = {
    "Segunda": {
        "13:30": ["ÉZIO BARCELOS JÚNIOR", "LIZANDRA GONÇALVES BOTÃO"],
        "14:30": ["LUCA ZUCCARI BOSKOVITZ", "JULIANA THIAGO RODRIGUES"],
        "15:30": ["ÉZIO BARCELOS JÚNIOR", "MARCELA ALVES BRANCO PINTO"]
    },
    "Terça": {
        "13:30": ["ADOLFO BRAGA NETO", "FABIANA FUKASE FLORENCIO"],
        "14:30": ["DANIELLA BOPPRÉ DE A. ABRAM", "JULIANA THIAGO RODRIGUES"],
        "15:30": ["JULIANA BABY MARQUES F. MOLES", "FABIANA FUKASE FLORENCIO"]
    },
    "Quarta": {
        "13:30": ["INGRID TEIXEIRA ANZAI", "DANIELE FRANCISCA B. REIS"],
        "14:30": ["LUCA ZUCCARI BOSKOVITZ", "MARCELA ALVES BRANCO PINTO"],
        "15:30": ["INGRID TEIXEIRA ANZAI", "DANIELE FRANCISCA B. REIS"]
    },
    "Quinta": {
        "13:30": ["LIZANDRA GONÇALVES BOTÃO", "ÉZIO BARCELOS JÚNIOR"],
        "14:30": ["DANIELLA BOPPRÉ DE A. ABRAM", "JULIANA BABY MARQUES F. MOLES"],
        "15:30": ["LIZANDRA GONÇALVES BOTÃO"]
    }
}

def obter_nome_dia(data_s):
    try:
        dt_obj = datetime.datetime.strptime(data_s, "%d/%m/%Y")
        dias = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
        return dias[dt_obj.weekday()]
    except: return None

def gerar_nomeacoes_web(texto_existentes, texto_novos):
    # Dicionário para contar audiências DE CADA UM no lote
    placar_geral = {}
    # Controle de quantos processos já ocuparam aquele slot (dia/hora)
    vagas_preenchidas = {}
    
    padrao = r"(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2})\s+([\d.-]+)\s+(\S+)\s+(.*)"
    nomeacoes_finais = []

    # Extrair e limpar as linhas
    linhas = [l.strip() for l in texto_novos.strip().split("\n") if re.search(padrao, l)]
    
    # IMPORTANTE: Processamos na ordem que chegam para manter a lógica de preenchimento
    for linha in linhas:
        match = re.search(padrao, linha)
        d_s, h_s, proc, sen, vara = match.groups()
        dia_txt = obter_nome_dia(d_s)
        
        # 1. Cancelamentos
        if "CANCELAD" in sen.upper() or "CANCELAD" in linha.upper():
            nomeacoes_finais.append([d_s, h_s, proc, sen, vara, "AUDIÊNCIA CANCELADA", "N/A"])
            continue

        mediador_escolhido = "VAGO (SEM ESCALA)"
        tipo = "N/A"

        # 2. Verificar se o dia/hora existe na escala
        if dia_txt in escala_fixa and h_s in escala_fixa[dia_txt]:
            candidatos = escala_fixa[dia_txt][h_s]
            
            # Verificar quantas audiências já colocamos nesse exato minuto/dia
            chave_slot = (d_s, h_s)
            ocupacao_atual = vagas_preenchidas.get(chave_slot, 0)

            # Só agendamos se o número de processos não exceder o número de mediadores na tabela
            if ocupacao_atual < len(candidatos):
                # LÓGICA DE EQUILÍBRIO:
                # Se houver mais de um candidato, vemos quem já tem menos processos NO TOTAL
                # Mas precisamos garantir que não vamos repetir o mesmo mediador no mesmo horário
                
                # Filtrar candidatos que ainda não foram usados NESTE específico (d_s, h_s)
                # (Para o caso de horários com 2 mediadores)
                ja_usados_neste_slot = [n[5] for n in nomeacoes_finais if n[0] == d_s and n[1] == h_s]
                disponiveis_agora = [c for c in candidatos if c not in ja_usados_neste_slot]

                if disponiveis_agora:
                    # ESCOLHA: Quem tem o menor valor no placar_geral
                    escolhido = min(disponiveis_agora, key=lambda m: placar_geral.get(m, 0))
                    
                    mediador_escolhido = escolhido
                    placar_geral[escolhido] = placar_geral.get(escolhido, 0) + 1
                    vagas_preenchidas[chave_slot] = ocupacao_atual + 1
                    tipo = "JEC" if "JEC" in vara.upper() else "PAGA"
                else:
                    mediador_escolhido = "VAGO (EXCEDEU ESCALA)"
            else:
                mediador_escolhido = "VAGO (EXCEDEU ESCALA)"

        nomeacoes_finais.append([d_s, h_s, proc, sen, vara, mediador_escolhido, tipo])

    # Ordenar e Salvar
    nomeacoes_finais.sort(key=lambda x: (datetime.datetime.strptime(x[0], "%d/%m/%Y"), x[1]))
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Nomeações"
    ws.append(["Data", "Horário", "Processo", "Senha", "Vara", "Mediador", "Tipo"])
    for row in nomeacoes_finais: ws.append(row)
    
    # Adiciona um pequeno resumo no final do Excel para você conferir a equidade
    ws.append([])
    ws.append(["RESUMO DE NOMEAÇÕES (EQUILÍBRIO)"])
    for med in sorted(placar_geral.keys()):
        ws.append([med, placar_geral[med]])

    wb.save("NOMEACOES_CEJUSC.xlsx")
