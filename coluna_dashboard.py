import pandas as pd
import matplotlib.pyplot as plt

# 1. Carregamento e Tratamento Inicial
caminho_arquivo = r'C:\ #mais uma vez você precisa colocar o caminho escolhido para o script achar a pasta com as planilhas
df = pd.read_excel(caminho_arquivo)

df.columns = df.columns.str.strip()
df['Situação'] = df['Situação'].astype(str).str.strip()

coluna_nome = 'NOME DO VENDEDOR' # Certifique-se de que a coluna na planilha está escrita assim

# 2. Definições da Equipe e Dicionário de Cores
consultores_alvo = ['Ana Laura', 'Carlos', 'Michelle', 'Fernanda']
situacoes_alvo = ['Ativo (sem estudar)', 'Ativo (estudando)', 'Cancelado', 'Evadido']

cores_padrao = {
    'Cancelado': '#F1C40F',       # Amarelo
    'Evadido': '#E74C3C',         # Vermelho
    'Ativo (sem estudar)': '#3498DB', # Azul
    'Ativo (estudando)': '#2ECC71'    # Verde
}

# 3. Função do Gráfico de Barras Empilhadas
def gerar_grafico_equipe(ano_referencia, nome_arquivo_saida):
    # Filtra pelo ano
    df_filtrado = df[df['Ano'] == ano_referencia].copy()
    
    if df_filtrado.empty:
        print(f"Atenção: Nenhum registro para o ano {ano_referencia}.")
        return

    # Filtra os dados apenas para os 4 consultores e as 4 situações específicas
    df_filtrado = df_filtrado[df_filtrado[coluna_nome].isin(consultores_alvo)]
    df_filtrado = df_filtrado[df_filtrado['Situação'].isin(situacoes_alvo)]

    # Engine de agrupamento (Linhas: Consultores, Colunas: Situação)
    metricas = df_filtrado.groupby([coluna_nome, 'Situação']).size().unstack(fill_value=0)

    # Garante que as 4 situações apareçam no gráfico, mesmo que algum consultor tenha "0" em alguma delas
    for sit in situacoes_alvo:
        if sit not in metricas.columns:
            metricas[sit] = 0
            
    # Reordena as colunas para o empilhamento seguir a ordem visual do dicionário
    metricas = metricas[situacoes_alvo]

    # Ordenação das barras: do consultor com mais matrículas no total para o menor
    metricas['Total'] = metricas.sum(axis=1)
    metricas = metricas.sort_values(by='Total', ascending=False).drop(columns='Total')

    # Plotagem 16:9
    ax = metricas.plot(
        kind='bar', 
        stacked=True, 
        figsize=(16, 9),
        color=[cores_padrao[col] for col in metricas.columns],
        edgecolor='white',
        width=0.6 # Deixei a barra um pouco mais estreita por serem apenas 4 pessoas
    )
# Filtra os dados apenas para os 4 consultores e as 4 situações específicas
    df_filtrado = df_filtrado[df_filtrado[coluna_nome].isin(consultores_alvo)]
    df_filtrado = df_filtrado[df_filtrado['Situação'].isin(situacoes_alvo)]

    # --- TRAVA DE SEGURANÇA ADICIONADA AQUI ---
    if df_filtrado.empty:
        print(f"Erro: Nenhum dado encontrado após filtrar pelos consultores. Verifique se a coluna '{coluna_nome}' existe e contém os nomes corretos.")
        return

    # Engine de agrupamento (Linhas: Consultores, Colunas: Situação)
    metricas = df_filtrado.groupby([coluna_nome, 'Situação']).size().unstack(fill_value=0)

    # Adiciona os números exatos dentro de cada cor da barra
    for container in ax.containers:
        labels = [f'{int(v)}' if v > 0 else '' for v in container.datavalues]
        ax.bar_label(container, labels=labels, label_type='center', fontsize=13, weight='bold', color='white')

    # Ajustes estéticos
    plt.title(f'Desempenho da Equipe: Retenção e Evasão - {ano_referencia}', fontsize=22, weight='bold', pad=30)
    plt.ylabel('Volume de Matrículas', fontsize=14, weight='bold', labelpad=15)
    
    # Rotação em 0 porque são apenas 4 nomes curtos, fica perfeitamente legível
    plt.xticks(rotation=0, ha='center', fontsize=14, weight='bold') 
    
    plt.legend(title="Status do Aluno", bbox_to_anchor=(1.02, 1), loc='upper left', fontsize=12)
    
    plt.tight_layout()
    plt.savefig(f'{nome_arquivo_saida}.png', bbox_inches='tight', dpi=150)
    plt.close()
    print(f"Gráfico da equipe '{nome_arquivo_saida}.png' gerado com sucesso!")

# 4. Execução
gerar_grafico_equipe(2025, 'Analise_Consultores_Ativos_2025')
