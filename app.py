import os
import time
import re
from datetime import datetime
from collections import Counter

import gspread
from google.oauth2.service_account import Credentials

# === CONFIGURAÇÕES ===
NOME_PLANILHA_ORIGEM = "Feedbackss"
NOME_ABA_FEEDBACKS = "Respostas ao formulário 1"

NOME_PLANILHA_DESTINO = "analise_feedbacks_resultado"
NOME_ABA_ANALISE = "Resultados"

INTERVALO_VERIFICACAO = 30  # segundos

# Caminho automático do JSON (mesma pasta do app.py)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAMINHO_CREDENCIAIS = r"C:\Users\ldowr\Desktop\automacao_planilhas\credenciais.json"



print("🔍 Verificando caminho:", CAMINHO_CREDENCIAIS)
print("📂 Existe?", os.path.exists(CAMINHO_CREDENCIAIS))
print("📂 Arquivos na pasta:", os.listdir(os.path.dirname(CAMINHO_CREDENCIAIS)))



  # ajuste conforme necessário

class AnalisadorFeedbacks:
    def __init__(self):
        self.client = None
        self.aba_feedbacks = None
        self.aba_analise = None

    def conectar(self):
        """Conecta ao Google Sheets"""
        print(f"🔑 Usando credenciais em: {CAMINHO_CREDENCIAIS}")
        if not os.path.exists(CAMINHO_CREDENCIAIS):
            raise FileNotFoundError(f"Arquivo de credenciais não encontrado: {CAMINHO_CREDENCIAIS}")

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_file(CAMINHO_CREDENCIAIS, scopes=scopes)
        self.client = gspread.authorize(creds)

        # Conecta à planilha de origem
        planilha_origem = self.client.open(NOME_PLANILHA_ORIGEM)
        try:
            self.aba_feedbacks = planilha_origem.worksheet(NOME_ABA_FEEDBACKS)
        except:
            self.aba_feedbacks = planilha_origem.get_worksheet(0)
            print(f"⚠️ Aba '{NOME_ABA_FEEDBACKS}' não encontrada, usando primeira: {self.aba_feedbacks.title}")

        # Conecta à planilha de destino
        planilha_destino = self.client.open(NOME_PLANILHA_DESTINO)
        try:
            self.aba_analise = planilha_destino.worksheet(NOME_ABA_ANALISE)
        except:
            print(f"📊 Criando aba de análise '{NOME_ABA_ANALISE}'...")
            self.aba_analise = planilha_destino.add_worksheet(title=NOME_ABA_ANALISE, rows=1000, cols=12)
            cabecalhos = [
                'Timestamp_Processamento', 'ID_Feedback', 'Nome_Cliente', 'Email_Cliente', 'Telefone_Cliente',
                'Nota_Avaliacao', 'Texto_Feedback', 'Sentimento_Final', 'Criterio_Classificacao',
                'Palavras_Chave', 'Status_Processamento', 'Data_Feedback'
            ]
            self.aba_analise.insert_row(cabecalhos, 1)

        print(f"✅ Conectado às planilhas: '{NOME_PLANILHA_ORIGEM}' e '{NOME_PLANILHA_DESTINO}'")
        return True

    def analisar_sentimento(self, feedback, nota):
        """Classifica sentimento baseado em nota e palavras-chave"""
        if nota:
            try:
                nota = float(re.findall(r'\d+\.?\d*', str(nota).replace(',', '.'))[0])
                if 1 <= nota <= 6:
                    return "Negativo", f"nota_{nota}"
                elif 7 <= nota <= 10:
                    return "Positivo", f"nota_{nota}"
            except:
                pass

        texto = str(feedback).lower()
        positivos = ["ótimo", "excelente", "bom", "perfeito", "gostei", "adorei", "amei", "top"]
        negativos = ["ruim", "péssimo", "horrível", "insatisfeito", "problema", "não gostei", "caro", "lento"]

        encontrados_pos = [p for p in positivos if p in texto]
        encontrados_neg = [p for p in negativos if p in texto]

        if encontrados_pos and not encontrados_neg:
            return "Positivo", ", ".join(encontrados_pos)
        elif encontrados_neg and not encontrados_pos:
            return "Negativo", ", ".join(encontrados_neg)
        elif encontrados_pos and encontrados_neg:
            return "Neutro", ", ".join(encontrados_pos + encontrados_neg)
        else:
            return "Neutro", "Nenhuma palavra-chave"

    def processar_feedbacks(self):
        """Lê feedbacks e grava análise"""
        dados = self.aba_feedbacks.get_all_values()
        if len(dados) <= 1:
            print("📝 Nenhum feedback encontrado")
            return 0

        cabecalhos = [h.lower() for h in dados[0]]
        linhas = dados[1:]
        resultados = []

        for idx, linha in enumerate(linhas, 1):
            try:
                nome = linha[cabecalhos.index("nome")] if "nome" in cabecalhos else "N/A"
                email = linha[cabecalhos.index("email")] if "email" in cabecalhos else "N/A"
                telefone = linha[cabecalhos.index("telefone")] if "telefone" in cabecalhos else "N/A"
                nota = linha[cabecalhos.index("nota")] if "nota" in cabecalhos else "N/A"
                feedback = linha[cabecalhos.index("feedback")] if "feedback" in cabecalhos else "N/A"
                timestamp = linha[cabecalhos.index("data")] if "data" in cabecalhos else "N/A"

                sentimento, criterio = self.analisar_sentimento(feedback, nota)
                resultados.append([
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    idx, nome, email, telefone, nota, feedback[:200],
                    sentimento, criterio, criterio, "Sucesso", timestamp
                ])
            except Exception as e:
                print(f"⚠️ Erro ao processar feedback {idx}: {e}")

        if resultados:
            self.aba_analise.clear()
            cabecalhos = [
                'Timestamp_Processamento', 'ID_Feedback', 'Nome_Cliente', 'Email_Cliente', 'Telefone_Cliente',
                'Nota_Avaliacao', 'Texto_Feedback', 'Sentimento_Final', 'Criterio_Classificacao',
                'Palavras_Chave', 'Status_Processamento', 'Data_Feedback'
            ]
            self.aba_analise.insert_row(cabecalhos, 1)
            self.aba_analise.insert_rows(resultados, 2)

            contagem = Counter([r[7] for r in resultados])
            print("📈 Distribuição de sentimentos:", dict(contagem))

        return len(resultados)

    def monitorar(self, intervalo=INTERVALO_VERIFICACAO):
        """Loop contínuo de monitoramento"""
        print(f"🚀 Monitorando feedbacks a cada {intervalo}s...\n")
        while True:
            self.processar_feedbacks()
            print(f"⏳ Próxima atualização em {intervalo}s...\n")
            time.sleep(intervalo)

def main():
    print("🎯 ANALISADOR AUTOMÁTICO DE FEEDBACKS - GOOGLE SHEETS\n" + "=" * 60)
    analisador = AnalisadorFeedbacks()
    if analisador.conectar():
        analisador.monitorar(INTERVALO_VERIFICACAO)

if __name__ == "__main__":
    main()
