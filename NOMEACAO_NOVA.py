import datetime
import re
from openpyxl import Workbook

# =====================================================
# ESCALA FIXA CEJUSC (v6.1 - INTEGRADA COM WEB)
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

# Mantive o nome 'gerar_nomeacoes_web' para o app.py não dar erro
# Adicionei 'texto_existentes' como opcional para manter compatibilidade com o formulário
def gerar_nomeacoes_web(texto_existentes, texto_novos):
    controle_duplicidade = {}
    padrao = r"(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2})\s+([\d.-]+)\s+(\S+)\s+(.*)"
    
    nomeacoes_finais = []

    linhas = texto_novos.strip().split("\n")
    
    for linha in linhas:
        match = re.search(padrao, linha)
        if not match: continue
            
        d_s, h_s, proc, sen, vara = match.groups()
        dia_txt = obter_nome_dia(d_s)
        
        # 1. Tratamento de Cancelamento (Conforme regra de negócio)
        if "CANCELAD" in sen.upper() or "CANCELAD" in linha.upper():
            nomeacoes_finais.append([d_s, h_s, proc, sen, vara, "AUDIÊNCIA CANCELADA", "N/A"])
            continue

        mediador_escolhido = "VAGO (SEM ESCALA)"
        tipo = "N/A"
        
        # 2. Aplicação da Escala Fixa
        if dia_txt and dia_txt in escala_fixa and h_s in escala_fixa[dia_txt]:
            lista_meds = escala_fixa[dia_txt][h_s]
            
            chave = (d_s, h_s)
            index = controle_duplicidade.get(chave, 0)
            
            if index < len(lista_meds):
                mediador_escolhido = lista_meds[index]
                tipo = "JEC" if "JEC" in vara.upper() else "PAGA"
            else:
                mediador_escolhido = "VAGO (EXCEDEU ESCALA)"
            
            controle_duplicidade[chave] = index + 1

        nomeacoes_finais.append([d_s, h_s, proc, sen, vara, mediador_escolhido, tipo])

    # Ordenação cronológica para o Excel
    nomeacoes_finais.sort(key=lambda x: (datetime.datetime.strptime(x[0], "%d/%m/%Y"), x[1]))

    # Geração do Arquivo
    wb = Workbook()
    ws = wb.active
    ws.title = "Nomeações"
    ws.append(["Data", "Horário", "Processo", "Senha", "Vara", "Mediador", "Tipo"])

    for row in nomeacoes_finais:
        ws.append(row)

    # O nome do arquivo deve ser exatamente o que o app.py espera
    wb.save("NOMEACOES_CEJUSC.xlsx")
