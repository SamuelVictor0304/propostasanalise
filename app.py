import pandas as pd
from datetime import datetime

# Caminho completo da planilha
caminho_planilha = r'X:\SERVER_ANTIGO\SETOR DE COBRANÇA\BASE DE DADOS\SISTEMA PROPOSTAS\PROPOSTAS 2025.xlsx'

# Lendo a planilha de propostas, especificando a subplanilha 'PROPOSTAS JUD 1'
df = pd.read_excel(caminho_planilha, sheet_name='PROPOSTAS JUD 1')

# Convertendo a coluna de data para datetime
df['DATA DO ENVIO DA PROPOSTA'] = pd.to_datetime(df['DATA DO ENVIO DA PROPOSTA'])

# Adicionando colunas de mês e ano para facilitar o filtro
df['Mês'] = df['DATA DO ENVIO DA PROPOSTA'].dt.strftime('%m/%Y')

# Mostrando os valores únicos na coluna STATUS para verificar
print("\nValores únicos na coluna STATUS:")
print(df['STATUS'].unique())

# Função para gerar o relatório
def gerar_relatorio(dados):
    # Agrupando por negociador e status
    resultado = dados.groupby(['NEGOCIADOR', 'STATUS']).size().unstack(fill_value=0)
    
    # Verificando quais colunas existem e renomeando
    colunas_renomear = {}
    if 'Aprovada' in resultado.columns:
        colunas_renomear['Aprovada'] = 'Aprovadas'
    if 'Recusada' in resultado.columns:
        colunas_renomear['Recusada'] = 'Negadas'
    
    # Renomeando as colunas existentes
    resultado = resultado.rename(columns=colunas_renomear)
    
    # Garantindo que as colunas existam
    for col in ['Aprovadas', 'Negadas']:
        if col not in resultado.columns:
            resultado[col] = 0
    
    # Calculando o total de propostas por negociador
    resultado['Total de Propostas'] = resultado['Aprovadas'] + resultado['Negadas']
    
    # Calculando a taxa de aprovação
    resultado['Taxa de Aprovação (%)'] = (resultado['Aprovadas'] / resultado['Total de Propostas'] * 100).round(2)
    
    # Ordenando por total de propostas
    resultado = resultado.sort_values('Total de Propostas', ascending=False)
    
    # Removendo a linha 'vazio' se existir
    if 'vazio ' in resultado.index:
        resultado = resultado.drop('vazio ')
    
    return resultado

# Mostrando os negociadores disponíveis
print("\nNegociadores disponíveis:")
print(df['NEGOCIADOR'].unique())

# Perguntando se quer filtrar por negociador
filtrar_negociador = input("\nDeseja filtrar por algum negociador específico? (S/N): ").upper()
if filtrar_negociador == 'S':
    negociador = input("Digite o nome do negociador: ")
    df = df[df['NEGOCIADOR'] == negociador]
    print(f"\nFiltrando dados para o negociador: {negociador}")

# Gerando relatório geral
relatorio_geral = gerar_relatorio(df)

# Gerando relatório por mês
relatorio_mensal = df.groupby('Mês').apply(gerar_relatorio)

# Salvando os resultados com nomes fixos
nome_arquivo_geral = r'C:\Users\samuel.victor\Desktop\sistema propostas\Resultado_Geral.xlsx'
nome_arquivo_mensal = r'C:\Users\samuel.victor\Desktop\sistema propostas\Resultado_Mensal.xlsx'

# Salvando relatório geral
relatorio_geral.to_excel(nome_arquivo_geral)

# Salvando relatório mensal
relatorio_mensal.to_excel(nome_arquivo_mensal)

print(f'\nAnálise concluída!')
print(f'Relatório geral salvo em: {nome_arquivo_geral}')
print(f'Relatório mensal salvo em: {nome_arquivo_mensal}')

# Mostrando os meses disponíveis
print("\nMeses disponíveis para filtro:")
print(df['Mês'].unique())

# Exemplo de como filtrar por mês específico
mes_especifico = input("\nDigite o mês desejado (formato: MM/YYYY) ou pressione Enter para sair: ")
if mes_especifico:
    df_filtrado = df[df['Mês'] == mes_especifico]
    relatorio_mes = gerar_relatorio(df_filtrado)
    print(f"\nResultado para o mês {mes_especifico}:")
    print(relatorio_mes)
    
    # Salvando relatório do mês específico
    nome_arquivo_mes = f'X:\SERVER_ANTIGO\SETOR DE COBRANÇA\BASE DE DADOS\SISTEMA PROPOSTAS\\Resultado_{mes_especifico.replace("/", "_")}.xlsx'
    relatorio_mes.to_excel(nome_arquivo_mes)
    print(f'Relatório do mês salvo em: {nome_arquivo_mes}')
