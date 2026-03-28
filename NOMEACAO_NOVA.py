import datetime
import re
from openpyxl import Workbook

# =====================================================
# CONFIGURAÇÃO DOS MEDIADORES
# =====================================================

mediadores = {
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

# =====================================================
# FUNÇÃO PRINCIPAL V3 - EQUIDADE REAL
# =====================================================

def gerar_nomeacoes_web(texto_existentes, texto_novos):
    controle_pago = {nome: 0 for nome in mediadores}
    controle_gratuito = {nome: 0 for nome in mediadores}
    controle_dia = {}
    controle_semana = {}

    padrao = r"(\d{2}/\d{2}/\d{4})\s+(\d{1,2}:\d{2})\s+(\d{7,}-\d{2}\.\d{4})\s+(\S+)\s+(.*)"

    # --- PROCESSAR EXISTENTES ---
    if texto_existentes:
        linhas_existentes = texto_existentes.strip().split("\n")
        for linha in linhas_existentes:
            partes = linha.strip().split("\t")
            if len(partes) < 3: continue
            
            data_str = partes[0].strip()
            horario = partes[1].strip()
            mediador = partes[-1].strip()
            
            if mediador in mediadores:
                # Como o histórico antigo não diferencia JEC, mantemos como pago para fins de trava mensal
                controle_pago[mediador] += 1
                data = datetime.datetime.strptime(data_str, "%d/%m/%Y")
                ano, semana, _ = data.isocalendar()
                controle_dia.setdefault((mediador, data_str), []).append(horario)
                controle_semana[(mediador, ano, semana)] = controle_semana.get((mediador, ano, semana), 0) + 1

    # --- FUNÇÕES AUXILIARES ---
    def dia_semana(data):
        return ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][data.weekday()]

    def pode_atuar(nome, dia, horario, data_str):
        dados = mediadores[nome]
        data = datetime.datetime.strptime(data_str, "%d/%m/%Y")
        ano, semana, _ = data.isocalendar()
        
        if dia not in dados["dias"]: return False
        if dados["somente_1330"] and horario != "13:30": return False
        if dados["nao_1330"] and horario == "13:30": return False
        if dados["max_mes"] is not None and (controle_pago[nome] + controle_gratuito[nome]) >= dados["max_mes"]: return False
        
        if nome in ["DANIELLA BOPPRÉ DE A. ABRAM", "ADOLFO BRAGA NETO"]:
            if controle_semana.get((nome, ano, semana), 0) >= 1: return False
                
        chave = (nome, data_str)
        if chave in controle_dia:
            novo = datetime.datetime.strptime(horario, "%H:%M")
            for h in controle_dia[chave]:
                existente = datetime.datetime.strptime(h, "%H:%M")
                if abs((novo - existente).total_seconds()) / 3600 < 2: return False
        return True

    def escolher_mediador_v3(dia, horario, data_str, vara):
        is_jec = "JEC" in vara.upper()
        
        # Filtra aptos e aplica a REGRA DO ADOLFO (Não faz JEC)
        aptos = []
        for m in mediadores:
            if pode_atuar(m, dia, horario, data_str):
                if is_jec and m == "ADOLFO BRAGA NETO":
                    continue
                aptos.append(m)
        
        if not aptos: return "SEM DISPONIBILIDADE"
        
        if is_jec:
            # EQUIDADE JEC: Prioriza quem tem MENOS JEC acumulado
            # Desempate: Quem tem MAIS remuneradas (para compensar)
            aptos.sort(key=lambda x: (controle_gratuito[x], -controle_pago[x]))
        else:
            # EQUIDADE REMUNERADA: Prioriza quem tem MENOS remuneradas
            # Desempate: Quem tem MAIS JEC (recompensa pelo trabalho gratuito)
            aptos.sort(key=lambda x: (controle_pago[x], -controle_gratuito[x]))

        escolhido = aptos[0]
        
        if is_jec: controle_gratuito[escolhido] += 1
        else: controle_pago[escolhido] += 1
        
        data = datetime.datetime.strptime(data_str, "%d/%m/%Y")
        ano, semana, _ = data.isocalendar()
        controle_dia.setdefault((escolhido, data_str), []).append(horario)
        controle_semana[(escolhido, ano, semana)] = controle_semana.get((escolhido, ano, semana), 0) + 1
        
        return escolhido

    # --- GERAÇÃO DO EXCEL ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Nomeações"
    ws.append(["Data", "Horário", "Processo", "Senha", "Vara", "Mediador", "Tipo"])

    linhas = texto_novos.strip().split("\n")
    for linha in linhas:
        resultado = re.search(padrao, linha)
        if not resultado: continue
        data_str, horario, processo, senha, vara = resultado.groups()
        
        try:
            data = datetime.datetime.strptime(data_str, "%d/%m/%Y")
            dia = dia_semana(data)
            mediador = escolher_mediador_v3(dia, horario, data_str, vara)
            tipo = "GRATUITA (JEC)" if "JEC" in vara.upper() else "REMUNERADA"
            ws.append([data_str, horario, processo, senha, vara, mediador, tipo])
        except: continue

    ws.append([])
    ws.append(["RELATÓRIO DE EQUIDADE (V3)"])
    ws.append(["Mediador", "Remuneradas", "JEC", "Total"])
    for nome in sorted(mediadores.keys()):
        p, g = controle_pago[nome], controle_gratuito[nome]
        ws.append([nome, p, g, p + g])

    wb.save("NOMEACOES_CEJUSC.xlsx")
