import pandas as pd
from faker import Faker
import random
from datetime import date

# Inicializa o Faker com dados no padrão brasileiro (pt_BR)
fake = Faker('pt_BR')

def gerar_dados_alunos(quantidade=200):
    dados = []
    
    # Listas de opções
    cursos = [
        'Matemática', 'Ciência de Dados', 'Letras', 
        'Engenharia de Software', 'Engenharia Elétrica'
    ]
    situacoes = ['Ativo (sem estudar)', 'Ativo (estudando)', 'Cancelado', 'Evadido']
    consultores = ['Ana Laura', 'Carlos', 'Michelle', 'Fernanda']
    concursos = ['Vestibular 2025', 'Nota do ENEM', 'Transferência', 'Segunda Graduação']
    arquivos_origem = ['base_matriculas_geral.csv', 'export_erp.xlsx', 'crm_leads.csv']

    # Definindo as configurações do ano desejado 
    inicio_2025 = date(2025, 1, 1)
    fim_2025 = date(2025, 12, 31)

    for _ in range(quantidade):
        nome_pessoa = fake.name()
        
        # Sorteia a data de ingresso restrita ao ano de 2025
        data_ingresso = fake.date_between_dates(date_start=inicio_2025, date_end=fim_2025)
        
        registro = {
            'ID USUÁRIO': fake.unique.random_int(min=10000, max=99999),
            'NOME DO USUÁRIO': nome_pessoa,
            'USUÁRIO CPF': fake.cpf(),
            'CANDIDATO': fake.unique.numerify('CAND-######'),
            'CONCURSO': random.choice(concursos),
            'RA': fake.unique.numerify('#########'),
            'NOME ALUNO': nome_pessoa,
            'POLO': f"{fake.city()} - {fake.state_abbr()}",
            'DATA INGRESSO': data_ingresso.strftime('%Y-%m-%d'),
            'CURSO': random.choice(cursos),
            # A data de pagamento agora ocorre entre o ingresso e o último dia de 2025
            'DATA PAGAMENTO': fake.date_between_dates(date_start=data_ingresso, date_end=fim_2025).strftime('%Y-%m-%d'),
            'INDICADO NO EU INDICO CANDIDATO': random.choice(['Sim', 'Não']),
            'ARQUIVO_ORIGEM': random.choice(arquivos_origem),
            'Situação': random.choice(situacoes),
            'Ano': data_ingresso.year # Sempre será o ano escolhido
        }
        dados.append(registro)
        
    return pd.DataFrame(dados)

# Gerando o lote de alunos fictícios
df_alunos = gerar_dados_alunos()

# Exportando o DataFrame diretamente para o Excel
df_alunos.to_excel('dataset_alunos_2025.xlsx', index=False, engine='openpyxl')

print("Arquivo 'dataset_alunos_2025.xlsx' gerado com sucesso e restrito ao ano escolhido!")
