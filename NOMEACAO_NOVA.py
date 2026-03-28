import datetime
import re
from openpyxl import Workbook

# =====================================================
# CONFIGURAÇÃO DOS MEDIADORES (ATUALIZADO)
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
# FUNÇÃO PRINCIPAL PARA WEB
# =====================================================

def gerar_nomeacoes_web(texto_existentes, texto_novos):
    """
    Processa audiências existentes e novas, gerando um Excel com a distribuição.
    """
    controle = {nome: 0 for nome in mediadores}
    controle_dia = {}
    controle_semana = {}

    # =====================================================
    # CARREGAR EXISTENTES
    # =====================================================
    if texto_existentes:
        linhas_existentes = texto_existentes.strip().split("\n")
        for linha in linhas_existentes:
            partes = linha.strip().split("\t")
            if len(partes) < 3:
                continue
            
            data_str = partes[0].strip()
            horario = partes[1].strip()
            mediador = partes[-1].strip()
            
            if mediador == "" or mediador.upper() == "CANCELADA":
                continue
            if mediador not in mediadores:
                continue
            
            try:
                data = datetime.datetime.strptime(data_str, "%d/%m/%Y")
                ano, semana, _ = data.isocalendar()
                controle[mediador] += 1
                chave_dia = (mediador, data_str)
                controle_dia.setdefault(chave_dia, []).append(horario)
                chave_semana = (mediador, ano, semana)
                controle_semana[chave_semana] = controle_semana.get(chave_semana, 0) + 1
            except:
                continue

    # =====================================================
    # FUNÇÕES AUXILIARES
    # =====================================================
    def dia_semana(data):
        dias = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
        return dias[data.weekday()]

    def pode_atuar(nome, dia, horario, data_str):
        dados = mediadores[nome]
        data = datetime.datetime.strptime(data_str, "%d/%m/%Y")
        ano, semana, _ = data.isocalendar()
        
        # Filtro de Dia
        if dia not in dados["dias"]:
            return False
        
        # Filtro de Horário Específico
        if dados["somente_1330"] and horario != "13:30":
            return False
        if dados["nao_1330"] and horario == "13:30":
            return False
            
        # Limite Mensal
        if dados["max_mes"] is not None and controle[nome] >= dados["max_mes"]:
            return False
            
        # Máximo 1 por semana (Regra específica para Daniella e Adolfo)
        if nome in ["DANIELLA BOPPRÉ DE A. ABRAM", "ADOLFO BRAGA NETO"]:
            if controle_semana.get((nome, ano, semana), 0) >= 1:
                return False
                
        # Conflito de horário (Mínimo 2h de intervalo no mesmo dia)
        chave = (nome, data_str)
        if chave in controle_dia:
            novo = datetime.datetime.strptime(horario, "%H:%M")
            for h in controle_dia[chave]:
                existente = datetime.datetime.strptime(h, "%H:%M")
                diferenca = abs((novo - existente).total_seconds()) / 3600
                if diferenca < 2:
                    return False
        return True

    def escolher_mediador(dia, horario, data_str):
        data = datetime.datetime.strptime(data_str, "%d/%m/%Y")
        ano, semana, _ = data.isocalendar()
        
        # Filtra quem pode atuar
        aptos = [m for m in mediadores if pode_atuar(m, dia, horario, data_str)]
        
        if not aptos:
            return "SEM DISPONIBILIDADE"
        
        # Ordena pelo que tem menos nomeações no total (Equilíbrio)
        aptos.sort(key=lambda x: controle[x])
        escolhido = aptos[0]
        
        # Atualiza contadores
        controle[escolhido] += 1
        chave_dia = (escolhido, data_str)
        controle_dia.setdefault(chave_dia, []).append(horario)
        chave_semana = (escolhido, ano, semana)
        controle_semana[chave_semana] = controle_semana.get(chave_semana, 0) + 1
        
        return escolhido

    # =====================================================
    # PROCESSAMENTO DAS NOVAS AUDIÊNCIAS
    # =====================================================
    wb = Workbook()
    ws = wb.active
    ws.title = "Nomeações"
    ws.append(["Data", "Horário", "Processo", "Mediador"])

    linhas = texto_novos.strip().split("\n")
    # Regex para capturar Data, Horário e Número do Processo
    padrao = r"(\d{2}/\d{2}/\d{4}).*?(\d{1,2}:\d{2}).*?(\d{7,}-\d{2}\.\d{4}(?:\.\d\.\d{2}\.\d{4})?)"

    for linha in linhas:
        resultado = re.search(padrao, linha)
        if not resultado:
            continue
            
        data_str = resultado.group(1)
        horario = resultado.group(2)
        processo = resultado.group(3)
        
        try:
            data = datetime.datetime.strptime(data_str, "%d/%m/%Y")
            dia = dia_semana(data)
            mediador = escolher_mediador(dia, horario, data_str)
            ws.append([data_str, horario, processo, mediador])
        except:
            continue

    # Relatório Final no rodapé do Excel
    ws.append([])
    ws.append(["RELATÓRIO FINAL"])
    ws.append(["Mediador", "Total Geral (Existentes + Novos)"])
    for nome, total in sorted(controle.items(), key=lambda x: x[1], reverse=True):
        ws.append([nome, total])

    wb.save("NOMEACOES_CEJUSC.xlsx")
