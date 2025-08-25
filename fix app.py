import pandas as pd
import warnings
import sys
from collections import Counter
import re

warnings.filterwarnings('ignore')

def encontrar_coluna(df, possiveis_nomes):
    for nome in possiveis_nomes:
        if nome in df.columns:
            return nome
    return None

def analisar_palavras(texto):
    if pd.isna(texto) or str(texto).strip() == '':
        return 'Neutro', []

    texto_limpo = str(texto).lower().strip()

    positivas = [
        'ótimo', 'excelente', 'bom', 'muito bom', 'perfeito', 'maravilhoso',
        'satisfeito', 'recomendo', 'gostei', 'adorei', 'fantástico', 'incrível',
        'eficiente', 'rápido', 'qualidade', 'profissional', 'atencioso', 'legal',
        'top', 'show', 'massa', 'bacana', 'nota 10', 'amei', 'super',
        'excepcional', 'sensacional', 'parabéns', 'sucesso', 'aprovado',
        'recomendaria', 'voltaria', 'melhor'
    ]

    negativas = [
        'ruim', 'péssimo', 'terrível', 'horrível', 'insatisfeito', 'decepcionado',
        'problema', 'demora', 'lento', 'caro', 'desorganizado', 'mal atendido',
        'não gostei', 'não recomendo', 'frustrante', 'irritante', 'confuso',
        'desastre', 'zero', 'lixo', 'odiei', 'pior', 'decepção', 'falha',
        'erro', 'defeito'
    ]

    palavras_pos = [p for p in positivas if p in texto_limpo]
    palavras_neg = [p for p in negativas if p in texto_limpo]

    if len(palavras_pos) > len(palavras_neg):
        return 'Positivo', palavras_pos
    elif len(palavras_neg) > len(palavras_pos):
        return 'Negativo', palavras_neg
    else:
        return 'Neutro', []

def classificar_por_nota(nota):
    try:
        if pd.isna(nota) or str(nota).strip().lower() in ['', 'n/a', 'na', 'nan']:
            return None

        nota_str = str(nota).strip().replace(',', '.')
        numeros = re.findall(r'\d+\.?\d*', nota_str)

        if not numeros:
            return None

        nota_num = float(numeros[0])

        if 0 <= nota_num <= 6:
            return 'Negativo'
        elif 7 <= nota_num <= 10:
            return 'Positivo'
        else:
            return 'Neutro'

    except:
        return None

def analisar_sentimento(texto, nota):
    sentimento_palavras, palavras_encontradas = analisar_palavras(texto)

    if len(palavras_encontradas) > 0:
        return {
            'sentimento': sentimento_palavras,
            'criterio': 'Palavras-chave',
            'palavras': palavras_encontradas
        }

    sentimento_nota = classificar_por_nota(nota)

    if sentimento_nota is not None:
        return {
            'sentimento': sentimento_nota,
            'criterio': 'Nota',
            'palavras': []
        }

    return {
        'sentimento': 'Neutro',
        'criterio': 'Sem indicadores',
        'palavras': []
    }

def detectar_sugestoes(texto):
    if pd.isna(texto) or str(texto).strip() == '':
        return False, []

    texto_limpo = str(texto).lower().strip()
    sugestivas = [
        'sugiro', 'acho que', 'deveria', 'mas', 'possibilidade',
        'minha opinião', 'sugestão', 'deveriam', 'seria melhor', 'poderiam'
    ]

    palavras_encontradas = [p for p in sugestivas if p in texto_limpo]

    return bool(palavras_encontradas), palavras_encontradas

# Carregar planilha
try:
    df = pd.read_excel("avaliacoesclientes.xlsx", sheet_name="My new form", header=0)
    print(f"Planilha carregada: {df.shape[0]} linhas, {df.shape[1]} colunas")
except FileNotFoundError:
    print("ERRO: Arquivo 'avaliacoesclientes.xlsx' não encontrado!")
    sys.exit(1)
except Exception as e:
    print(f"ERRO ao carregar planilha: {e}")
    sys.exit(1)

df = df.dropna(how="all")

mapeamentos = {
    'Feedback': ["Dê a sua opinião sobre o serviço:", "Feedback", "Opinião"],
    'Nota': ["Como você nos avalia?", "Nota", "Avaliação"],
    'Nome': ["Seu nome:", "Nome", "Cliente"],
    'Email': ["Seu e-mail:", "E-mail", "Email"],
    'Telefone': ["Seu telefone:", "Telefone", "Celular"],
    'Sugestao_Original': ["Gostaria de deixar alguma sugestão?", "Sugestão", "Deixe sua sugestão", "eixar algures"]
}

for nova_coluna, possiveis_nomes in mapeamentos.items():
    coluna_encontrada = encontrar_coluna(df, possiveis_nomes)
    if coluna_encontrada and coluna_encontrada != nova_coluna:
        df = df.rename(columns={coluna_encontrada: nova_coluna})

df = df.drop_duplicates(subset=['Nome', 'Email', 'Telefone', 'Feedback'], keep='first')

if 'Feedback' not in df.columns:
    print("ERRO: Coluna de feedback não encontrada!")
    sys.exit(1)

feedbacks_validos = df[df['Feedback'].notna() & (df['Feedback'].astype(str).str.strip() != '')]

if len(feedbacks_validos) == 0:
    print("ERRO: Nenhum feedback válido encontrado!")
    sys.exit(1)

print(f"Analisando {len(feedbacks_validos)} feedbacks...")

resultados = []

for i, (_, row) in enumerate(feedbacks_validos.iterrows(), 1):
    print("-" * 50)

    nome = row.get('Nome')
    email = row.get('Email', 'N/A')
    telefone = row.get('Telefone', 'N/A')
    feedback = str(row['Feedback']).strip()
    nota = row.get('Nota', 'N/A')
    sugestao_bruta = row.get("Sugestao_Original", "")
    sugestao_texto = str(sugestao_bruta).strip() if pd.notna(sugestao_bruta) and str(sugestao_bruta).strip() != '' else ''

    print(f"Comentário: {feedback}")
    print(f"Nome: {nome}")
    print(f"Email: {email}")
    print(f"Telefone: {telefone}")
    print(f"Sugestão original: {sugestao_texto}")

    resultado = analisar_sentimento(feedback, nota)
    tem_sugestao, sugestoes_encontradas = detectar_sugestoes(sugestao_texto)

    resultados.append({
    'ID': i,
    'Nome': nome,
    'Email': email,
    'Telefone': telefone,
    'Feedback': feedback,
    'Nota': nota,
    'Sentimento': resultado['sentimento'],
    'Criterio': resultado['criterio'],
    'Palavras Chave Encontradas': ', '.join(resultado['palavras']) if resultado['palavras'] else 'Nenhuma',
    'Gostaria de deixar alguma sugestão?': sugestao_texto,

})


    print(f"Feedback {i}: {resultado['sentimento']} ({resultado['criterio']})")
    if tem_sugestao:
        print(f"→ Contém sugestão: {', '.join(sugestoes_encontradas)}")

try:
    df_resultado = pd.DataFrame(resultados)
    df_resultado.to_excel('analise_feedbacks_limpa.xlsx', index=False)
    print(f"\nArquivo salvo: analise_feedbacks_limpa.xlsx")

    contagem = df_resultado['Sentimento'].value_counts()
    print("\nEstatísticas:")
    for sentimento, qtd in contagem.items():
        perc = (qtd / len(df_resultado)) * 100
        print(f"  {sentimento}: {qtd} ({perc:.1f}%)")

    todos_feedbacks = ' '.join(df_resultado['Feedback'].dropna().astype(str).tolist()).lower()
    palavras = re.findall(r'\b\w{4,}\b', todos_feedbacks)
    contador = Counter(palavras)
    mais_comuns = contador.most_common(10)

    print("\n🔍 Palavras mais comuns nos feedbacks:")
    for palavra, freq in mais_comuns:
        print(f"  {palavra}: {freq}x")

    notas_validas = []
    for nota in df_resultado['Nota']:
        try:
            nota_str = str(nota).strip().replace(',', '.')
            numeros = re.findall(r'\d+\.?\d*', nota_str)
            if numeros:
                nota_num = float(numeros[0])
                if 0 <= nota_num <= 10:
                    notas_validas.append(nota_num)
        except:
            continue
            
except Exception as e:
    print(f"ERRO: Ocorreu um erro inesperado durante a análise final: {e}")
    sys.exit(1)
