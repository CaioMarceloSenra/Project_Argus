import pandas as pd
import matplotlib.pyplot as plt

# 1. Carregamento e Tratamento Inicial
caminho_arquivo = r'C:\Users\caiom\OneDrive\Anexos\Área de Trabalho\Project_Argus\Matriculas_2024.xlsx'
df = pd.read_excel(caminho_arquivo)

df.columns = df.columns.str.strip()
df['Situação'] = df['Situação'].astype(str).str.strip()
coluna_nome = 'nome do usuário'

# 2. Dicionário de Cores
cores_padrao = {
    'Cancelado': '#F1C40F',
    'Evadido': '#E74C3C',
    'Ativo (sem estudar)': '#3498DB',
    'Ativo (estudando)': '#2ECC71'
}

# 3. Função para encurtar nomes (Primeiro + Segundo nome)
def encurtar_nome(nome_completo):
    if pd.isna(nome_completo): return "N/I"
    partes = str(nome_completo).split()
    if len(partes) >= 2:
        return f"{partes[0]} {partes[1]}"
    return partes[0]

# 4. Função do Gráfico de Barras
def gerar_grafico_barras_por_ano(ano_referencia, nome_arquivo_saida):
    df_filtrado = df[df['Ano'] == ano_referencia].copy()
    
    if df_filtrado.empty:
        print(f"Atenção: Nenhum registro para {ano_referencia}.")
        return

    df_filtrado['Nome_Curto'] = df_filtrado[coluna_nome].apply(encurtar_nome)

    # Engine de agrupamento (Linhas: Nomes, Colunas: Situação)
    metricas = df_filtrado.groupby(['Nome_Curto', 'Situação']).size().unstack(fill_value=0)

    # Ordenação por volume total
    metricas['Total'] = metricas.sum(axis=1)
    metricas = metricas.sort_values(by='Total', ascending=False).drop(columns='Total')

    # Plotagem 16:9
    ax = metricas.plot(
        kind='bar', 
        stacked=True, 
        figsize=(16, 9),
        color=[cores_padrao.get(col, '#BDC3C7') for col in metricas.columns],
        edgecolor='white',
        width=0.7 
    )

    # Adiciona os números exatos dentro de cada cor da barra
    for container in ax.containers:
        labels = [f'{int(v)}' if v > 0 else '' for v in container.datavalues]
        ax.bar_label(container, labels=labels, label_type='center', fontsize=11, weight='bold', color='white')

    plt.title(f'Perspectiva por Consultor - {ano_referencia}', fontsize=22, weight='bold', pad=30)
    plt.xticks(rotation=0, ha='center', fontsize=12, weight='bold') 
    plt.legend(title="Status", bbox_to_anchor=(1.0, 1), loc='upper left', fontsize=12)
    
    plt.tight_layout()
    plt.savefig(f'{nome_arquivo_saida}.png', bbox_inches='tight', dpi=150)
    plt.close()
    print(f"Gráfico de consultores '{nome_arquivo_saida}.png' gerado com sucesso!")

# 5. Execução
gerar_grafico_barras_por_ano(2024, 'Analise_Consultores_Composto_2024')
