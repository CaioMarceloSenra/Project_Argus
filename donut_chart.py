import pandas as pd
import matplotlib.pyplot as plt

# 1. Carregamento e Tratamento Inicial
caminho_arquivo = r #É necessário ajustar o caminho do arquivo para o local onde ele está salvo
df = pd.read_excel(caminho_arquivo)

df.columns = df.columns.str.strip()
df['Situação'] = df['Situação'].astype(str).str.strip()

# Ajuste dos nomes para um visual mais limpo no gráfico
df['Situação'] = df['Situação'].replace({
    'Ativo (estudando)': 'Ativos',
    'Ativo (sem estudar)': 'Sem Estudar'
})

# 2. Dicionário de Cores #Você pode ajustar com as cores que desejar!
cores_padrao = {
    'Cancelado': '#F1C40F',       # Amarelo
    'Evadido': "#E7783C",         # Vermelho
    'Ativos': '#2ECC71',          # Verde
    'Sem Estudar': '#3498DB'      # Azul
}

# 3. Função do Gráfico de Rosca (Donut Chart)
def gerar_grafico_glamour(ano_referencia, nome_arquivo_saida):
    df_filtrado = df[df['Ano'] == ano_referencia].copy()
    
    if df_filtrado.empty:
        print(f"Atenção: Nenhum registro encontrado para o ano {ano_referencia}.")
        return

    # Contagem
    contagem = df_filtrado['Situação'].value_counts()
    cores_pizza = [cores_padrao.get(sit, '#BDC3C7') for sit in contagem.index]

    # Configuração da figura com fundo sóbrio
    fig, ax = plt.subplots(figsize=(10, 7), facecolor='#F8F9F9')
    ax.set_facecolor('#F8F9F9')
    
    # Destacar levemente todas as fatias
    explode = [0.03] * len(contagem)

    # Gráfico Base
    wedges, texts, autotexts = ax.pie(
        contagem, 
        autopct='%1.1f%%', 
        startangle=90, 
        colors=cores_pizza,
        explode=explode, 
        pctdistance=0.75, 
        textprops={'fontsize': 12, 'weight': 'bold', 'color': 'white'}, 
        wedgeprops={'edgecolor': 'white', 'linewidth': 1.5, 'antialiased': True}
    )

    # Transformando em Rosca (Donut Chart)
    centro_branco = plt.Circle((0,0), 0.55, fc='#F8F9F9')
    fig.gca().add_artist(centro_branco)

    # Ajuste dos números para melhor visibilidade
    for autotext in autotexts:
        autotext.set_color('#2C3E50') 
        autotext.set_fontsize(13)

    # Legenda à direita, fora do gráfico
    ax.legend(
        wedges, contagem.index,
        title="Status do Aluno",
        title_fontsize=13,
        loc="center left",
        bbox_to_anchor=(1, 0, 0.5, 1),
        fontsize=12,
        frameon=False 
    )

    plt.title(f'Panorama de Retenção e Evasão\nAno Letivo {ano_referencia}', 
              fontsize=18, weight='bold', color='#2C3E50', loc='center', pad=20)
    
    # Salvando em alta resolução para garantir a qualidade do design
    plt.tight_layout()
    plt.savefig(f'{nome_arquivo_saida}.png', bbox_inches='tight', dpi=300) 
    plt.close()
    print(f"Gráfico de rosca '{nome_arquivo_saida}.png' gerado com sucesso!")

# 4. Execução
gerar_grafico_glamour(2024, 'Analise_Rosca') #Aqui você pode colocar o nome do arquivo que quiser!
