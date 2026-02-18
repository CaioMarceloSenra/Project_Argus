import pandas as pd
import matplotlib.pyplot as plt

# 1. Carregamento e Tratamento Inicial
caminho_arquivo = r'C:\Users\caiom\OneDrive\Anexos\Área de Trabalho\Project_Argus\Matriculas_2024.xlsx'
df = pd.read_excel(caminho_arquivo)

df.columns = df.columns.str.strip()
df['Situação'] = df['Situação'].astype(str).str.strip()

# 2. Dicionário de Cores
cores_padrao = {
    'Cancelado': '#F1C40F',
    'Evadido': '#E74C3C',
    'Ativo (sem estudar)': '#3498DB',
    'Ativo (estudando)': '#2ECC71'
}

# 3. Função do Gráfico de Rosca
def gerar_grafico_pizza_por_ano(ano_referencia, nome_arquivo_saida):
    df_filtrado = df[df['Ano'] == ano_referencia]
    
    if df_filtrado.empty:
        print(f"Atenção: Nenhum registro encontrado para o ano {ano_referencia}.")
        return

    distribuicao = df_filtrado['Situação'].value_counts()
    cores_aplicadas = [cores_padrao.get(status, '#BDC3C7') for status in distribuicao.index]
    
    fig, ax = plt.subplots(figsize=(16, 9))
    
    # Parâmetros ajustados para não cortar o texto
    wedges, texts, autotexts = ax.pie(
        distribuicao, 
        labels=distribuicao.index, 
        autopct='%1.1f%%',
        startangle=140,
        colors=cores_aplicadas,
        shadow=True,
        radius=0.8,           # Reduz o tamanho da pizza
        pctdistance=0.85,     # Distância da porcentagem
        labeldistance=1.1,    # Distância do rótulo
        textprops={'fontsize': 14, 'weight': 'bold'}
    )

    plt.setp(autotexts, size=12, color="white", weight='bold')

    # Transforma em Rosca (Donut Chart) para um visual mais limpo
    centro = plt.Circle((0,0), 0.50, fc='white')
    fig.gca().add_artist(centro)

    plt.title(f'Status da Carteira de Alunos - {ano_referencia}', fontsize=18, weight='bold', pad=30)
    
    plt.tight_layout()
    plt.savefig(f'{nome_arquivo_saida}.png', bbox_inches='tight', dpi=150)
    plt.close()
    print(f"Gráfico geral '{nome_arquivo_saida}.png' gerado com sucesso!")

# 4. Execução
gerar_grafico_pizza_por_ano(2024, 'Distribuicao_Status_2024_16x9')
