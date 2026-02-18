import pandas as pd
import matplotlib.pyplot as plt

# 1. Carregamento e Tratamento Inicial
caminho_arquivo = r'C: #Nesse campo em específico você coloca a rota de entrada para a planilha
df = pd.read_excel(caminho_arquivo)

df.columns = df.columns.str.strip()
df['Situação'] = df['Situação'].astype(str).str.strip()

# 2. Dicionário de Cores
cores_padrao = {
    'Cancelado': '#F1C40F',       # Amarelo
    'Evadido': '#E74C3C',         # Vermelho
    'Ativo (sem estudar)': '#3498DB', # Azul
    'Ativo (estudando)': '#2ECC71'    # Verde
}

# 3. Função do Gráfico de Pizza
def gerar_grafico_pizza_geral(ano_referencia, nome_arquivo_saida):
    df_filtrado = df[df['Ano'] == ano_referencia].copy()
    
    if df_filtrado.empty:
        print(f"Atenção: Nenhum registro encontrado para o ano {ano_referencia}.")
        return

    # Conta o total de alunos em cada Situação
    contagem = df_filtrado['Situação'].value_counts()

    # Garante que as cores aplicadas sejam as do dicionário, na mesma ordem da contagem
    cores_pizza = [cores_padrao.get(sit, '#BDC3C7') for sit in contagem.index]

    # Plotagem da figura com fundo transparente/branco
    plt.figure(figsize=(10, 8))
    
    # Criando o gráfico de pizza
    plt.pie(
        contagem, 
        labels=contagem.index, 
        autopct='%1.1f%%', # Mostra a porcentagem com 1 casa decimal
        startangle=140,    # Rotaciona levemente para um design mais harmonioso
        colors=cores_pizza,
        textprops={'fontsize': 13, 'weight': 'bold', 'color': '#333333'}, 
        wedgeprops={'edgecolor': 'white', 'linewidth': 2} # Adiciona uma borda branca entre as fatias
    )

    plt.title(f'Panorama Geral de Retenção e Evasão - {ano_referencia}', fontsize=20, weight='bold', pad=20)
    
    plt.tight_layout()
    plt.savefig(f'{nome_arquivo_saida}.png', bbox_inches='tight', dpi=150)
    plt.close()
    print(f"Gráfico de pizza '{nome_arquivo_saida}.png' gerado com sucesso!")

# 4. Execução
gerar_grafico_pizza_geral(2025, 'Analise_Pizza_Geral_2025')
