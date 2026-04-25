import datetime
import re
from openpyxl import Workbook

# =====================================================
# ESCALA FIXA CEJUSC (v6.2 - EQUILÍBRIO POR SLOT)
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
    except:
        return None

def gerar_nomeacoes_web(texto_existentes, texto_novos):
    # Contador de audiências por mediador NO LOTE ATUAL
    contagem_lote = {}
    # Controle para não colocar dois mediadores no EXATO mesmo processo
    controle_duplicidade_hora = {}
    
    padrao = r"(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2})\s+([\d.-]+)\s+(\S+)\s+(.*)"
    nomeacoes_finais = []

    linhas = texto_novos.strip().split("\n")
    
    for linha in linhas:
        match = re.search(padrao, linha)
        if not match: continue
            
        d_s, h_s, proc, sen, vara = match.groups()
        dia_txt = obter_nome_dia(d_s)
        
        if "CANCELAD" in sen.upper() or "CANCELAD" in linha.upper():
            nomeacoes_finais.append([d_s, h_s, proc, sen, vara, "AUDIÊNCIA CANCELADA", "N/A"])
            continue

        mediador_escolhido = "VAGO (SEM ESCALA)"
        tipo = "N/A"
        
        if dia_txt and dia_txt in escala_fixa and h_s in escala_fixa[dia_txt]:
            lista_disponiveis = escala_fixa[dia_txt][h_s]
            
            # 1. Identificar quantos processos já temos para este exato dia e hora
            chave_hora = (d_s, h_s)
            tentativa_num = controle_duplicidade_hora.get(chave_hora, 0)

            # Se ainda houver "vagas" na escala para esse horário
            if tentativa_num < len(lista_disponiveis):
                # Se houver mais de uma opção (ex: Ézio ou Lizandra), escolhe quem tem menos no acumulado
                if len(lista_disponiveis) > 1:
                    # Ordena os disponíveis pelo total de audiências que já receberam neste processamento
                    # Quem tem menos vem primeiro
                    opcoes_ordenadas = sorted(lista_disponiveis, key=lambda m: contagem_lote.get(m, 0))
                    
                    # Se for a primeira audiência do slot, pega o que tem menos trabalho
                    # Se for a segunda audiência do slot, pega o outro
                    escolhido = opcoes_ordenadas[tentativa_num]
                else:
                    escolhido = lista_disponiveis[0]

                mediador_escolhido = escolhido
                contagem_lote[escolhido] = contagem_lote.get(escolhido, 0) + 1
                tipo = "JEC" if "JEC" in vara.upper() else "PAGA"
                controle_duplicidade_hora[chave_hora] = tentativa_num + 1
            else:
                mediador_escolhido = "VAGO (EXCEDEU ESCALA)"

        nomeacoes_finais.append([d_s, h_s, proc, sen, vara, mediador_escolhido, tipo])

    # Ordenação e Geração do Excel (Igual ao anterior)
    nomeacoes_finais.sort(key=lambda x: (datetime.datetime.strptime(x[0], "%d/%m/%Y"), x[1]))
    wb = Workbook()
    ws = wb.active
    ws.title = "Nomeações"
    ws.append(["Data", "Horário", "Processo", "Senha", "Vara", "Mediador", "Tipo"])
    for row in nomeacoes_finais: ws.append(row)
    wb.save("NOMEACOES_CEJUSC.xlsx")
