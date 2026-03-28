import datetime
import re
import random
from openpyxl import Workbook
from copy import deepcopy

# =====================================================
# CONFIGURAÇÃO DOS MEDIADORES (v5.3 - TRAVAS DE FERRO)
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
    "FABIANA FUKASE FLORENCIO": {"dias": ["Terça", "Quarta", "Quinta", "Sexta"], "somente_1330": False, "nao_1330": False, "max_mes": None},
    "PATRÍCIA MARIA O. PASSANEZI": {"dias": ["Segunda"], "somente_1330": False, "nao_1330": False, "max_mes": None}
}

def obter_nome_dia(data_s):
    dt_obj = datetime.datetime.strptime(data_s, "%d/%m/%Y")
    return ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][dt_obj.weekday()]

def pode_atuar(nome, horario_s, data_s, c_pago, c_gratuito, c_dia, c_semana, vara):
    config_med = mediadores_config[nome]
    dia_txt = obter_nome_dia(data_s)
    is_jec = "JEC" in vara.upper()
    
    # --- BLOQUEIOS RÍGIDOS PATRÍCIA ---
    if nome == "PATRÍCIA MARIA O. PASSANEZI":
        if not is_jec: return False  # NUNCA faz remunerada
        if dia_txt != "Segunda": return False # NUNCA faz fora de segunda
    
    # --- BLOQUEIOS RÍGIDOS ADOLFO ---
    if nome == "ADOLFO BRAGA NETO" and is_jec: return False

    # --- REGRAS GERAIS ---
    if dia_txt not in config_med["dias"]: return False
    if config_med["somente_1330"] and horario_s != "13:30": return False
    if config_med["nao_1330"] and horario_s == "13:30": return False
    
    dt_obj = datetime.datetime.strptime(data_s, "%d/%m/%Y")
    ano, sem, _ = dt_obj.isocalendar()
    
    if config_med["max_mes"] is not None and (c_pago[nome] + c_gratuito[nome]) >= config_med["max_mes"]: return False
    if nome in ["DANIELLA BOPPRÉ DE A. ABRAM", "ADOLFO BRAGA NETO"] and c_semana.get((nome, ano, sem), 0) >= 1: return False
    
    if (nome, data_s) in c_dia:
        h_novo = datetime.datetime.strptime(horario_s, "%H:%M")
        for h_ex_s in c_dia[(nome, data_s)]:
            h_ex = datetime.datetime.strptime(h_ex_s, "%H:%M")
            if abs((h_novo - h_ex).total_seconds()) / 3600 < 2: return False
            
    return True

# =====================================================
# MOTOR DE SIMULAÇÃO (v5.3)
# =====================================================

def gerar_nomeacoes_web(texto_existentes, texto_novos):
    hist_pago = {n: 0 for n in mediadores_config}; hist_gratuito = {n: 0 for n in mediadores_config}
    hist_dia, hist_semana = {}, {}

    if texto_existentes:
        for linha in texto_existentes.strip().split("\n"):
            partes = re.split(r'\t|\s{2,}', linha.strip())
            if len(partes) < 4: continue
            d_s, h_s, med = partes[0], partes[1], partes[-1]
            if med in mediadores_config:
                if "JEC" in linha.upper(): hist_gratuito[med] += 1
                else: hist_pago[med] += 1
                try:
                    dt = datetime.datetime.strptime(d_s, "%d/%m/%Y"); a, s, _ = dt.isocalendar()
                    hist_dia.setdefault((med, d_s), []).append(h_s)
                    hist_semana[(med, a, s)] = hist_semana.get((med, a, s), 0) + 1
                except: continue

    padrao = r"(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2})\s+(\d{7,}-\d{2}\.\d{4})\s+(\S+)\s+(.*)"
    novas_list = [list(re.search(padrao, l).groups()) for l in texto_novos.strip().split("\n") if re.search(padrao, l)]

    melhor_resultado = None; menor_score = float('inf')

    for _ in range(200):
        c_pago, c_gratuito = deepcopy(hist_pago), deepcopy(hist_gratuito)
        c_dia, c_semana = deepcopy(hist_dia), deepcopy(hist_semana)
        sim_nomeacoes = []; sim_penalty = 0
        
        aud_shuffled = random.sample(novas_list, len(novas_list))
        
        for d_s, h_s, proc, sen, vara in aud_shuffled:
            is_jec = "JEC" in vara.upper(); dia_txt = obter_nome_dia(d_s)
            
            # Filtra aptos já com as travas rígidas de Patrícia e Adolfo dentro do pode_atuar
            aptos = [m for m in mediadores_config if pode_atuar(m, h_s, d_s, c_pago, c_gratuito, c_dia, c_semana, vara)]
            
            if not aptos:
                sim_nomeacoes.append([d_s, h_s, proc, sen, vara, "SEM DISPONIBILIDADE", "N/A"]); continue
            
            # --- LÓGICA DE SELEÇÃO PRIORITÁRIA ---
            if is_jec and dia_txt == "Segunda" and "PATRÍCIA MARIA O. PASSANEZI" in aptos:
                escolhido = "PATRÍCIA MARIA O. PASSANEZI"
            else:
                # Se for Segunda/JEC e a Patrícia NÃO estiver apta (por causa das 2h), aí sim distribuímos
                # Removemos a Patrícia de qualquer outra situação para garantir que ela não pegue remunerada
                aptos_filtrados = [a for a in aptos if a != "PATRÍCIA MARIA O. PASSANEZI" or (is_jec and dia_txt == "Segunda")]
                
                if is_jec: aptos_filtrados.sort(key=lambda x: (c_gratuito[x], -c_pago[x]))
                else: aptos_filtrados.sort(key=lambda x: (c_pago[x], -c_gratuito[x]))
                escolhido = aptos_filtrados[0]
            
            if (escolhido, d_s) in c_dia: sim_penalty += 200 # Aumentei a multa de logística para 200

            if is_jec: c_gratuito[escolhido] += 1
            else: c_pago[escolhido] += 1
            
            dt = datetime.datetime.strptime(d_s, "%d/%m/%Y"); a, s, _ = dt.isocalendar()
            c_dia.setdefault((escolhido, d_s), []).append(h_s)
            c_semana[(escolhido, a, s)] = c_semana.get((escolhido, a, s), 0) + 1
            sim_nomeacoes.append([d_s, h_s, proc, sen, vara, escolhido, "JEC" if is_jec else "PAGA"])

        # Score: Max-Min de JEC e Pagas + Logística
        score = (max(c_gratuito.values()) - min(c_gratuito.values())) * 20 + \
                (max(c_pago.values()) - min(c_pago.values())) * 30 + sim_penalty
        
        if score < menor_score:
            menor_score = score
            melhor_resultado = {"nomeacoes": sim_nomeacoes, "pago": c_pago, "gratuito": c_gratuito}

    # Excel
    wb = Workbook(); ws = wb.active; ws.title = "Nomeações"
    ws.append(["Data", "Horário", "Processo", "Senha", "Vara", "Mediador", "Tipo"])
    f_list = sorted(melhor_resultado["nomeacoes"], key=lambda x: (datetime.datetime.strptime(x[0], "%d/%m/%Y"), x[1]))
    for row in f_list: ws.append(row)
    ws.append([]); ws.append(["RELATÓRIO DE EQUIDADE (V5.3)"])
    ws.append(["Mediador", "Remuneradas", "JEC", "Total"])
    for n in sorted(mediadores_config.keys()):
        p, g = melhor_resultado["pago"][n], melhor_resultado["gratuito"][n]
        ws.append([n, p, g, p + g])
    wb.save("NOMEACOES_CEJUSC.xlsx")
