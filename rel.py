# -*- coding: utf-8 -*-
"""
Created on Tue Feb  1 10:28:43 2022
@author: johannmaschio
"""

import pandas as pd
from pandas import read_excel
from datetime import timedelta, datetime, tzinfo, timezone
import csv 
import datetime

# Load data
data_entrante = read_excel("C:/RELATORIOS/Entrante.xlsx")
data_sainte = read_excel("C:/RELATORIOS/Sainte.xlsx")
data_tabulacao = read_excel("C:/RELATORIOS/Tabulação.xlsx")

# Data preprocessing
data_entrante = data_entrante.drop(columns=['Origem', 'Destino', 'Duração', 'Tempo de espera', 
                                            'Início de Atendimento', 'Fila'])
data_sainte = data_sainte.drop(columns=['Origem', 'Destino', 'Duração', 'Tempo de espera', 
                                            'Início de Atendimento', 'Fila'])
data_tabulacao = data_tabulacao.drop(columns=['Play', 'Telefone', 'Tempo de Espera'])

# Remove toda linha que tiver algum dado NaN
data_entrante = data_entrante.dropna()
data_sainte = data_sainte.dropna()
data_tabulacao = data_tabulacao.dropna()

# Remove Desistencia com menos de 40s e ligações tabuladas como Transferência, Engano e Não Tabulada 
for line in data_tabulacao.index:
    if data_tabulacao['Tabulação'][line] == 'Desistência' and data_tabulacao['Tempo de Atendimento'][line].hour == 0 and data_tabulacao['Tempo de Atendimento'][line].minute == 0 and data_tabulacao['Tempo de Atendimento'][line].second <= 40: 
        data_tabulacao = data_tabulacao.drop(labels = line)
    elif data_tabulacao['Tabulação'][line] == 'Transferência' or data_tabulacao['Tabulação'][line] == 'Engano' or data_tabulacao['Tabulação'][line] == 'Não Tabulada':
        data_tabulacao = data_tabulacao.drop(labels = line)
    elif data_tabulacao['Tabulação'][line] == 'Queda' and data_tabulacao['Tempo de Atendimento'][line].hour == 0 and data_tabulacao['Tempo de Atendimento'][line].minute == 0 and data_tabulacao['Tempo de Atendimento'][line].second <= 40:
        data_tabulacao = data_tabulacao.drop(labels = line)
               
# Remove ligações entrantes com menos de 50s
for line in data_entrante.index:
    if data_entrante['Tempo de Atendimento'][line].hour == 0 and data_entrante['Tempo de Atendimento'][line].minute == 0 and data_entrante['Tempo de Atendimento'][line].second <= 50:
        data_entrante = data_entrante.drop(labels = line)

# Remove ligações saintes com menos de 40s        
for line in data_sainte.index:
    if data_sainte['Tempo de Atendimento'][line].hour == 0 and data_sainte['Tempo de Atendimento'][line].minute == 0 and data_sainte['Tempo de Atendimento'][line].second <= 40:
        data_sainte = data_sainte.drop(labels = line)

# Iniciando a lista final
final_list = {}
final_list['Data'] = []
final_list['Agente'] = []
final_list['Tempo Atendimento'] = []
final_list['Tabulação'] = []

# Carregando os dados na lista ## Entrante e tabulação ##
for line_e in data_entrante.index:
    for line_t in data_tabulacao.index:
        if data_entrante['Data'][line_e] == data_tabulacao['Data'][line_t] and data_tabulacao['Agente'][line_t] == data_entrante['Agente'][line_e]:
            final_list['Data'].append(data_entrante['Data'][line_e])
            final_list['Agente'].append(data_tabulacao['Agente'][line_t])
            final_list['Tempo Atendimento'].append(data_tabulacao['Tempo de Atendimento'][line_t])
            final_list['Tabulação'].append(data_tabulacao['Tabulação'][line_t])
        continue 

# Carregando os dados saintes #
for line in data_sainte.index:
    final_list['Data'].append(data_sainte['Data'][line])
    final_list['Agente'].append(data_sainte['Agente'][line])
    final_list['Tempo Atendimento'].append(data_sainte['Tempo de Atendimento'][line])
    final_list['Tabulação'].append('Sainte')

# Criando o Dataframe e exportando para Excel.xlsx
columns = ['Agente', 'Data', 'Tabulação', 'Tempo Atendimento']
final_df = pd.DataFrame(final_list, columns=columns)
final_df.to_excel('final.xlsx', encoding='utf-8', index=False)

pross = final_df.drop(columns=["Data", "Tabulação", "Tempo Atendimento"])
pross = final_df.groupby(["Agente"]).count()

list_agentes = []
for agente in pross.index:
    list_agentes.append(agente)

tempos = {}
tempos["Agente"] = []
tempos["Tempo total sainte"] = []
tempos["Tempo total entrante"] = []
tempos["contagem sainte"] = []
tempos["contagem entrante"] = []
tempos["media sainte"] = []
tempos["media entrante"] = []

for agente in list_agentes:
    tempos["Agente"].append(agente)
    count_s = 0
    count_e = 0
    tempo_s = timedelta()
    tempo_e = timedelta()
    med_s = timedelta()
    med_s = timedelta()
    
    for line in final_df.index:
        if agente == final_df["Agente"][line] and final_df["Tabulação"][line] == "Sainte":
            count_s += 1
            tempo_s += timedelta(hours = final_df["Tempo Atendimento"][line].hour, minutes = final_df["Tempo Atendimento"][line].minute, seconds = final_df["Tempo Atendimento"][line].second)                      
        elif agente == final_df["Agente"][line] and final_df["Tabulação"][line] != "Sainte":
            count_e += 1
            tempo_e += timedelta(hours = final_df["Tempo Atendimento"][line].hour, minutes = final_df["Tempo Atendimento"][line].minute, seconds = final_df["Tempo Atendimento"][line].second)
    tempos["contagem sainte"].append(count_s)
    tempos["contagem entrante"].append(count_e)
    tempos["Tempo total sainte"].append(tempo_s)
    tempos["Tempo total entrante"].append(tempo_e)
    if tempo_s != 0 and count_s != 0:
        tempos["media sainte"].append(tempo_s / count_s)
    else:
        tempos["media sainte"].append(0)
    if tempo_e != 0 and count_e != 0:
        tempos["media entrante"].append(tempo_e / count_e)
    else:
        tempos["media entrante"].append(0)
        
columns = ['Agente', 'Tempo total sainte', 'contagem sainte', 'media sainte', 'Tempo total entrante', 'contagem entrante', 'media entrante']
final_tempos = pd.DataFrame(tempos, columns=columns)
final_tempos.to_excel('final_tempos.xlsx', encoding='utf-8', index=False)

print("Relatorios gerados!!!")
