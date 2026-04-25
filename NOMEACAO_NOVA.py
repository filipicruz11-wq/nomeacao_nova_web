import datetime
import re
from openpyxl import Workbook

# =====================================================
# ESCALA FIXA CEJUSC (v6.0 - TRAVA DE EXCEDENTES)
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
        "15:30": ["LIZANDRA GONÇALVES BOTÃO"] # APENAS LIZANDRA. O 2º processo sairá como VAGO.
    }
}

def obter_nome_dia(data_s):
    dt_obj = datetime.datetime.strptime(data_s, "%d/%m/%Y")
    dias = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
    return dias[dt_obj.weekday()]

def gerar_nomeacoes_fixas(texto_novos):
    controle_duplicidade = {}
    padrao = r"(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2})\s+([\d.-]+)\s+(\S+)\s+(.*)"
    
    nomeacoes_finais = []

    for linha in texto_novos.strip().split("\n"):
        match = re.search(padrao, linha)
        if not match: continue
            
        d_s, h_s, proc, sen, vara = match.groups()
        dia_txt = obter_nome_dia(d_s)
        
        # 1. Tratamento de Cancelamento
        if "CANCELAD" in sen.upper() or "CANCELAD" in linha.upper():
            nomeacoes_finais.append([d_s, h_s, proc, sen, vara, "AUDIÊNCIA CANCELADA", "N/A"])
            continue

        mediador_escolhido = "VAGO (SEM ESCALA)"
        tipo = "N/A"
        
        # 2. Verificação da Escala Fixa
        if dia_txt in escala_fixa and h_s in escala_fixa[dia_txt]:
            lista_meds = escala_fixa[dia_txt][h_s]
            
            chave = (d_s, h_s)
            index = controle_duplicidade.get(chave, 0)
            
            # Só atribui se ainda houver mediador disponível na lista daquele horário
            if index < len(lista_meds):
                mediador_escolhido = lista_meds[index]
                tipo = "JEC" if "JEC" in vara.upper() else "PAGA"
            else:
                # Se for o 2º processo da Quinta às 15:30 ou o 3º de qualquer outro horário
                mediador_escolhido = "VAGO (EXCEDEU ESCALA)"
            
            controle_duplicidade[chave] = index + 1

        nomeacoes_finais.append([d_s, h_s, proc, sen, vara, mediador_escolhido, tipo])

    # Salvar Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Nomeações"
    ws.append(["Data", "Horário", "Processo", "Senha", "Vara", "Mediador", "Tipo"])

    # Ordenação por data e depois por hora
    f_list = sorted(nomeacoes_finais, key=lambda x: (datetime.datetime.strptime(x[0], "%d/%m/%Y"), x[1]))

    for row in f_list: ws.append(row)
    wb.save("NOMEACOES_CEJUSC_FIXO.xlsx")
    print("Concluído! Arquivo 'NOMEACOES_CEJUSC_FIXO.xlsx' gerado.")
