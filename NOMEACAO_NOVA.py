import datetime
import re
import random
from openpyxl import Workbook
from copy import deepcopy

# =====================================================
# CONFIGURAÇÃO DOS MEDIADORES
# =====================================================

mediadores_config = {
    "ÉZIO BARCELOS JÚNIOR": {"dias": ["Segunda", "Terça", "Quinta", "Sexta"], "somente_1330": False, "nao_1330": False, "max_mes": None},
    "INGRID TEIXEIRA ANZAI": {"dias": ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"], "somente_1330": False, "nao_1330": False, "max_mes": None},
    "JULIANA BABY MARQUES F. MOLES": {"dias": ["Segunda", "Terça", "Quinta"], "somente_1330": False, "nao_1330": False, "max_mes": None},
    "JULIANA THIAGO RODRIGUES": {"dias": ["Segunda", "Terça"], "somente_1330": False, "nao_1330": False, "max_mes": None},
    "LIZANDRA GONÇALVES BOTÃO": {"dias": ["Segunda", "Terça", "Quinta", "Sexta"], "somente_1330": False, "nao_1330": False, "max_mes": None},
    "LUCA ZUCCARI BOSKOVITZ": {"dias": ["Segunda", "Terça", "Quarta"], "somente_1330": False, "nao_1330": False, "max_mes": None},
    "MARCELA ALVES BRANCO PINTO": {"dias": ["Segunda", "Quarta", "Sexta"], "somente_1330": False, "nao_1330": False, "max_mes": None},
    "ADOLFO BRAGA NETO": {"dias": ["Terça"], "somente_1330": True, "nao_1330": False, "max_mes": 2},
    "DANIELE FRANCISCA B. REIS": {"dias": ["Terça", "Quarta", "Sexta"], "somente_1330": True, "nao_1330": False, "max_mes": None},
    "DANIELLA BOPPRÉ DE A. ABRAM": {"dias": ["Terça", "Quinta"], "somente_1330": False, "nao_1330": True, "max_mes": 2},
    "FABIANA FUKASE FLORENCIO": {"dias": ["Terça", "Quarta", "Quinta", "Sexta"], "somente_1330": False, "nao_1330": False, "max_mes": None}
}

# --- FUNÇÕES DE APOIO ---
def dia_semana(data_str):
    data = datetime.datetime.strptime(data_str, "%d/%m/%Y")
    return ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][data.weekday()]

def pode_atuar(nome, horario, data_str, c_pago, c_gratuito, c_dia, c_semana):
    dados = mediadores_config[nome]
    dia = dia_semana(data_str)
    data = datetime.datetime.strptime(data_str, "%d/%m/%Y")
    ano, sem, _ = data.isocalendar()
    
    if dia not in dados["dias"]: return False
    if dados["somente_1330"] and horario != "13:30": return False
    if dados["nao_1330"] and horario == "13:30": return False
    if dados["max_mes"] is not None and (c_pago[nome] + c_gratuito[nome]) >= dados["max_mes"]: return False
    if nome in ["DANIELLA BOPPRÉ DE A. ABRAM", "ADOLFO BRAGA NETO"] and c_semana.get((nome, ano, sem), 0) >= 1: return False
    
    # Regra de 2h de intervalo
    if (nome, data_str) in c_dia:
        novo = datetime.datetime.strptime(horario, "%H:%M")
        for h in c_dia[(nome, data_str)]:
            existente = datetime.datetime.strptime(h, "%H:%M")
            if abs((novo - existente).total_seconds()) / 3600 < 2: return False
    return True

# =====================================================
# MOTOR DE SIMULAÇÃO (v5)
# =====================================================

def gerar_nomeacoes_web(texto_existentes, texto_novos):
    # 1. Preparar Histórico Base
    hist_pago = {n: 0 for n in mediadores_config}
    hist_gratuito = {n: 0 for n in mediadores_config}
    hist_dia = {}
    hist_semana = {}

    if texto_existentes:
        for linha in texto_existentes.strip().split("\n"):
            partes = re.split(r'\t|\s{2,}', linha.strip())
            if len(partes) < 4: continue
            data_s, hora_s, med = partes[0], partes[1], partes[-1]
            if med in mediadores_config:
                if "JEC" in linha.upper(): hist_gratuito[med] += 1
                else: hist_pago[med] += 1
                try:
                    dt = datetime.datetime.strptime(data_s, "%d/%m/%Y")
                    a, s, _ = dt.isocalendar()
                    hist_dia.setdefault((med, data_s), []).append(hora_s)
                    hist_semana[(med, a, s)] = hist_semana.get((med, a, s), 0) + 1
                except: continue

    # 2. Extrair Novas Audiências
    padrao = r"(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2})\s+(\d{7,}-\d{2}\.\d{4})\s+(\S+)\s+(.*)"
    novas_list = []
    for linha in texto_novos.strip().split("\n"):
        res = re.search(padrao, linha)
        if res: novas_list.append(list(res.groups()))

    # 3. Rodar 1000 Simulações
    melhor_resultado = None
    menor_score = float('inf')

    for _ in range(1000):
        # Clonar estados para esta simulação
        curr_pago = deepcopy(hist_pago)
        curr_gratuito = deepcopy(hist_gratuito)
        curr_dia = deepcopy(hist_dia)
        curr_semana = deepcopy(hist_semana)
        sim_nomeacoes = []
        sim_logistica_penalty = 0
        
        # Embaralhar a ordem de processamento das audiências
        audiencias_shuffled = random.sample(novas_list, len(novas_list))
        
        for data_s, hora_s, proc, sen, vara in audiencias_shuffled:
            is_jec = "JEC" in vara.upper()
            aptos = [m for m in mediadores_config if pode_atuar(m, hora_s, data_s, curr_pago, curr_gratuito, curr_dia, curr_semana)]
            if is_jec: aptos = [m for m in aptos if m != "ADOLFO BRAGA NETO"]
            
            if not aptos:
                sim_nomeacoes.append([data_s, hora_s, proc, sen, vara, "SEM DISPONIBILIDADE", "N/A"])
                continue
            
            # Escolha baseada em carga atual da simulação
            if is_jec: aptos.sort(key=lambda x: (curr_gratuito[x], -curr_pago[x]))
            else: aptos.sort(key=lambda x: (curr_pago[x], -curr_gratuito[x]))
            
            escolhido = aptos[0]
            
            # Penalidade de Logística: se o mediador já tem audiência no dia, aumenta o score da simulação
            if (escolhido, data_s) in curr_dia:
                sim_logistica_penalty += 50 # Peso alto para evitar 2 no mesmo dia

            # Atualizar estado da simulação
            if is_jec: curr_gratuito[escolhido] += 1
            else: curr_pago[escolhido] += 1
            dt = datetime.datetime.strptime(data_s, "%d/%m/%Y")
            a, s, _ = dt.isocalendar()
            curr_dia.setdefault((escolhido, data_s), []).append(hora_s)
            curr_semana[(escolhido, a, s)] = curr_semana.get((escolhido, a, s), 0) + 1
            sim_nomeacoes.append([data_s, hora_s, proc, sen, vara, escolhido, "JEC" if is_jec else "PAGA"])

        # 4. Calcular Score de Qualidade da Simulação
        # O desvio padrão simplificado: (max - min)
        diff_jec = max(curr_gratuito.values()) - min(curr_gratuito.values())
        diff_paga = max(curr_pago.values()) - min(curr_pago.values())
        
        total_score = (diff_jec * 10) + (diff_paga * 20) + sim_logistica_penalty
        
        if total_score < menor_score:
            menor_score = total_score
            melhor_resultado = {"nomeacoes": sim_nomeacoes, "pago": curr_pago, "gratuito": curr_gratuito}

    # --- GERAR EXCEL FINAL ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Nomeações Otimizadas"
    ws.append(["Data", "Horário", "Processo", "Senha", "Vara", "Mediador", "Tipo"])

    # Ordenar por data para o Excel ficar bonito
    final_list = sorted(melhor_resultado["nomeacoes"], key=lambda x: (datetime.datetime.strptime(x[0], "%d/%m/%Y"), x[1]))
    for row in final_list: ws.append(row)

    ws.append([])
    ws.append(["RELATÓRIO DE EQUIDADE (V5 - OTIMIZADO)"])
    ws.append(["Mediador", "Remuneradas", "JEC", "Total"])
    for n in sorted(mediadores_config.keys()):
        p, g = melhor_resultado["pago"][n], melhor_resultado["gratuito"][n]
        ws.append([n, p, g, p + g])

    wb.save("NOMEACOES_CEJUSC.xlsx")
