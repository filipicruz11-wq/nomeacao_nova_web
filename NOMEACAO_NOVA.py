import datetime
import re
from openpyxl import Workbook

# =====================================================
# ESCALA FIXA CEJUSC (Conforme Tabela)
# =====================================================
# Estrutura: escala[Dia][Horário] = [Lista de Mediadores]
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
        "15:30": ["LIZANDRA GONÇALVES BOTÃO"] # Conforme o "?" na tabela, apenas 1 por enquanto
    }
}

def obter_nome_dia(data_s):
    dt_obj = datetime.datetime.strptime(data_s, "%d/%m/%Y")
    dias = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
    return dias[dt_obj.weekday()]

def gerar_nomeacoes_fixas(texto_novos):
    # Dicionário para controlar quem já foi usado no mesmo horário/dia (para pegar o 1º ou 2º da lista)
    controle_duplicidade = {}

    padrao = r"(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2})\s+([\d.-]+)\s+(\S+)\s+(.*)"
    linhas = texto_novos.strip().split("\n")
    
    nomeacoes_finais = []

    for linha in linhas:
        match = re.search(padrao, linha)
        if not match:
            continue
            
        d_s, h_s, proc, sen, vara = match.groups()
        dia_txt = obter_nome_dia(d_s)
        
        # 1. Verificar Cancelamento
        if "CANCELAD" in sen.upper() or "CANCELAD" in linha.upper():
            nomeacoes_finais.append([d_s, h_s, proc, sen, vara, "AUDIÊNCIA CANCELADA", "N/A"])
            continue

        # 2. Buscar na Escala
        mediador_escolhido = "SEM ESCALA DEFINIDA"
        tipo = "N/A"
        
        if dia_txt in escala_fixa and h_s in escala_fixa[dia_txt]:
            lista_meds = escala_fixa[dia_txt][h_s]
            
            # Lógica para alternar entre o primeiro e o segundo mediador do mesmo horário
            chave = (d_s, h_s)
            index = controle_duplicidade.get(chave, 0)
            
            if index < len(lista_meds):
                mediador_escolhido = lista_meds[index]
                controle_duplicidade[chave] = index + 1
                tipo = "JEC" if "JEC" in vara.upper() else "PAGA"
            else:
                mediador_escolhido = "EXCEDEU LIMITE DO HORÁRIO"

        nomeacoes_finais.append([d_s, h_s, proc, sen, vara, mediador_escolhido, tipo])

    # Geração do Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Nomeações Fixas"
    ws.append(["Data", "Horário", "Processo", "Senha", "Vara", "Mediador", "Tipo"])

    # Ordenar por data e hora antes de salvar
    f_list = sorted(nomeacoes_finais, key=lambda x: (datetime.datetime.strptime(x[0], "%d/%m/%Y"), x[1]))

    for row in f_list:
        ws.append(row)

    wb.save("NOMEACOES_CEJUSC_FIXO.xlsx")
    print("Arquivo gerado com sucesso: NOMEACOES_CEJUSC_FIXO.xlsx")

# Para rodar, basta chamar a função com o seu texto de novos processos
# gerar_nomeacoes_fixas(seu_texto_aqui)
