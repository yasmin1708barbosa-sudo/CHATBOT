"""
NutriBot — Chatbot de Alimentação Saudável e Qualidade de Vida

O sistema:
  1. conversa sobre alimentação, hidratação, sono, atividade física e bem-estar;
  2. faz uma avaliação rápida dos hábitos do usuário e dá uma nota de 0 a 100;
  3. calcula o IMC;
  4. gera um relatório completo em Word (.docx) com o resultado e as dicas.

Como rodar:
    pip install -r requirements.txt
    streamlit run app.py

Observação: as orientações são gerais e educativas, baseadas no Guia Alimentar
para a População Brasileira (Ministério da Saúde) e em recomendações da OMS.
Não substituem a avaliação de um nutricionista ou médico.
"""

import io
import random
import re
import unicodedata
from datetime import datetime

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor


# ----------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------
def normalizar(texto: str) -> str:
    """Deixa o texto minúsculo e sem acentos para facilitar a comparação."""
    texto = texto.lower().strip()
    texto = unicodedata.normalize("NFD", texto)
    return "".join(c for c in texto if unicodedata.category(c) != "Mn")


# ----------------------------------------------------------------------
# Base de conhecimento: cada tema tem palavras-chave e respostas
# ----------------------------------------------------------------------
TEMAS = {
    "saudacao": {
        "palavras": ["oi", "ola", "bom dia", "boa tarde", "boa noite", "eai", "e ai", "hello"],
        "respostas": [
            "Olá! 🥗 Sou o **NutriBot**, seu assistente de alimentação saudável e qualidade de vida. "
            "Pode me perguntar sobre água, frutas, café da manhã, lanches, sono, exercícios, estresse... "
            "Digite **ajuda** para ver tudo o que eu sei!",
            "Oi! Que bom te ver por aqui 😊 Quer dicas sobre alimentação, sono ou atividade física? "
            "Digite **ajuda** para ver os temas.",
        ],
    },
    "ajuda": {
        "palavras": ["ajuda", "menu", "opcoes", "temas", "o que voce sabe", "o que voce faz", "help"],
        "respostas": [
            "Posso conversar sobre estes temas:\n\n"
            "- 💧 **água** e hidratação\n"
            "- 🍎 **frutas, verduras** e legumes\n"
            "- ☕ **café da manhã**\n"
            "- 🍽️ **prato saudável** (almoço/jantar)\n"
            "- 🥪 **lanches** saudáveis\n"
            "- 🍬 **açúcar**, 🧂 **sal** e 🍟 **ultraprocessados**\n"
            "- 😴 **sono**\n"
            "- 🏃 **exercício** físico\n"
            "- 🧘 **estresse** e saúde mental\n"
            "- 📏 **IMC** (calculo pra você!)\n"
            "- 💡 **dica** do dia\n"
            "- 📋 **avaliação** dos seus hábitos e **relatório** em Word (veja as abas acima!)\n\n"
            "É só escrever uma dessas palavras ou fazer uma pergunta!"
        ],
    },
    "agua": {
        "palavras": ["agua", "hidrat", "beber", "sede", "litro"],
        "respostas": [
            "💧 **Hidratação** é essencial! Uma referência comum para adultos é cerca de **35 ml de água por kg** "
            "de peso por dia (ex.: 70 kg ≈ 2,4 litros), mas a necessidade varia com o calor e a atividade física.\n\n"
            "Dicas:\n- Tenha sempre uma garrafinha por perto\n- Beba antes de sentir sede\n"
            "- Urina clara é sinal de boa hidratação\n- Prefira água a sucos e refrigerantes",
        ],
    },
    "frutas": {
        "palavras": ["fruta", "verdura", "legume", "vegeta", "salada", "fibra"],
        "respostas": [
            "🍎 **Frutas, verduras e legumes** são fontes de vitaminas, minerais e fibras.\n\n"
            "- Tente incluir esses alimentos em **todas as refeições**\n"
            "- Varie as cores do prato: cada cor traz nutrientes diferentes 🌈\n"
            "- Prefira a fruta inteira ao suco — ela tem mais fibras e sacia mais\n"
            "- Escolha alimentos da época: são mais baratos e saborosos",
        ],
    },
    "cafe": {
        "palavras": ["cafe da manha", "desjejum", "manha", "acordar"],
        "respostas": [
            "☕ Ideias de **café da manhã** saudável:\n\n"
            "- Pão integral com ovo mexido + uma fruta\n"
            "- Iogurte natural com aveia e banana\n"
            "- Tapioca com queijo branco e tomate\n"
            "- Cuscuz com ovo + café sem muito açúcar\n\n"
            "Dica: combine um **carboidrato**, uma **proteína** e uma **fruta**.",
        ],
    },
    "prato": {
        "palavras": ["almoco", "jantar", "prato", "refeicao", "comida", "cardapio"],
        "respostas": [
            "🍽️ Um jeito simples de montar um **prato equilibrado**:\n\n"
            "- **Metade do prato**: verduras e legumes 🥦\n"
            "- **Um quarto**: arroz, batata, mandioca ou macarrão 🍚\n"
            "- **Um quarto**: feijão, carne, frango, peixe ou ovo 🍗\n\n"
            "O bom e velho **arroz com feijão** é uma combinação excelente!",
        ],
    },
    "lanche": {
        "palavras": ["lanche", "lanchinho", "beliscar", "fome", "petisco"],
        "respostas": [
            "🥪 Opções de **lanches saudáveis**:\n\n"
            "- Frutas (banana, maçã, laranja)\n- Castanhas ou amendoim (um punhado)\n"
            "- Iogurte natural\n- Pipoca feita na panela com pouco óleo\n"
            "- Sanduíche de pão integral com frango desfiado\n\n"
            "Evite trocar refeições por salgadinhos e biscoitos recheados.",
        ],
    },
    "acucar": {
        "palavras": ["acucar", "doce", "refrigerante", "refri", "chocolate", "sobremesa"],
        "respostas": [
            "🍬 **Açúcar**: o excesso está ligado a ganho de peso, diabetes e cáries.\n\n"
            "- Diminua aos poucos o açúcar no café e nos sucos\n"
            "- Troque refrigerante por água com limão ou água com gás\n"
            "- Fique de olho em rótulos: açúcar aparece com vários nomes (xarope, maltodextrina...)\n"
            "- Doce não é proibido! Só deixe para ocasiões especiais 😉",
        ],
    },
    "sal": {
        "palavras": ["sal", "sodio", "pressao"],
        "respostas": [
            "🧂 **Sal**: o consumo em excesso aumenta o risco de pressão alta.\n\n"
            "- A recomendação da OMS é menos de **5 g de sal por dia** (uma colher de chá)\n"
            "- Use temperos naturais: alho, cebola, ervas, limão\n"
            "- Evite temperos prontos em cubo e embutidos (salsicha, presunto)\n"
            "- Tire o saleiro da mesa!",
        ],
    },
    "ultraprocessados": {
        "palavras": ["ultraprocessad", "industrializad", "fast food", "salgadinho", "miojo", "embutido", "processad"],
        "respostas": [
            "🍟 O **Guia Alimentar para a População Brasileira** dá uma regra de ouro:\n\n"
            "> Prefira sempre alimentos *in natura* ou minimamente processados e preparações culinárias "
            "a alimentos ultraprocessados.\n\n"
            "Ultraprocessados (salgadinhos, refrigerantes, nuggets, macarrão instantâneo) costumam ter muito "
            "sal, açúcar, gordura e aditivos. Quanto mais **comida de verdade**, melhor!",
        ],
    },
    "sono": {
        "palavras": ["sono", "dormir", "insonia", "cansado", "cansaco", "descanso"],
        "respostas": [
            "😴 **Sono** de qualidade é parte da saúde! Adultos precisam, em geral, de **7 a 9 horas** por noite.\n\n"
            "- Durma e acorde sempre no mesmo horário\n"
            "- Evite celular e telas 1 hora antes de deitar 📱\n"
            "- Evite café e energéticos à noite\n"
            "- Jantar leve ajuda a dormir melhor\n"
            "- Deixe o quarto escuro, silencioso e fresco",
        ],
    },
    "exercicio": {
        "palavras": ["exercicio", "atividade fisica", "academia", "treino", "correr", "caminhar", "caminhada", "esporte", "sedentar"],
        "respostas": [
            "🏃 **Atividade física**: a OMS recomenda para adultos de **150 a 300 minutos** por semana de "
            "atividade moderada (ex.: 30 min de caminhada, 5 dias por semana).\n\n"
            "- Comece devagar e aumente aos poucos\n"
            "- Escolha algo que você goste: dança, bike, futebol, caminhada\n"
            "- Use a escada em vez do elevador\n"
            "- Levante e se alongue a cada hora se ficar muito tempo sentado",
        ],
    },
    "estresse": {
        "palavras": ["estresse", "stress", "ansiedade", "ansioso", "triste", "saude mental", "relaxar", "nervoso"],
        "respostas": [
            "🧘 **Qualidade de vida** também é cuidar da mente:\n\n"
            "- Faça pausas durante o dia e respire fundo (inspire 4s, segure 4s, solte 4s)\n"
            "- Passe tempo com amigos e família\n"
            "- Tenha um hobby e momentos de lazer\n"
            "- Atividade física ajuda muito a reduzir o estresse\n\n"
            "Se estiver se sentindo mal com frequência, converse com alguém de confiança ou procure um "
            "profissional de saúde. O **CVV** atende 24h pelo telefone **188**.",
        ],
    },
    "dica": {
        "palavras": ["dica", "conselho", "sugestao"],
        "respostas": [
            "💡 Dica do dia: **descasque mais e desembale menos!** Comida de verdade é a base de uma boa alimentação.",
            "💡 Dica do dia: **coma devagar e sem telas**. Mastigar bem ajuda o corpo a perceber a saciedade.",
            "💡 Dica do dia: **cozinhe mais em casa**. Você controla o sal, o óleo e o açúcar.",
            "💡 Dica do dia: **planeje as refeições da semana**. Economiza tempo e dinheiro!",
            "💡 Dica do dia: **leve uma garrafinha de água** pra onde for.",
            "💡 Dica do dia: **refeições em companhia** são mais prazerosas. Coma com a família ou amigos quando puder.",
        ],
    },
    "agradecimento": {
        "palavras": ["obrigad", "valeu", "vlw", "brigad", "agradeco"],
        "respostas": [
            "De nada! 😊 Cuide-se bem!",
            "Por nada! Qualquer dúvida sobre alimentação ou qualidade de vida, estou aqui 🥦",
        ],
    },
    "despedida": {
        "palavras": ["tchau", "ate mais", "ate logo", "flw", "adeus", "sair"],
        "respostas": [
            "Até mais! 👋 Lembre-se: beba água, coma comida de verdade e durma bem!",
            "Tchau! Foi bom conversar. Cuide da sua saúde 💚",
        ],
    },
}

NAO_ENTENDI = [
    "Hmm, não entendi muito bem 🤔 Tente perguntar sobre **água**, **frutas**, **sono**, **exercício** "
    "ou digite **ajuda** para ver todos os temas.",
    "Desculpe, ainda não sei responder isso. Digite **ajuda** para ver o que eu sei!",
]


# ----------------------------------------------------------------------
# IMC
# ----------------------------------------------------------------------
def classificar_imc(imc: float) -> str:
    if imc < 18.5:
        return "abaixo do peso"
    if imc < 25:
        return "peso adequado ✅"
    if imc < 30:
        return "sobrepeso"
    return "obesidade"


def tentar_calcular_imc(texto: str):
    """Procura peso e altura na mensagem, ex.: '70 kg 1,75 m' ou '70 1.75'."""
    numeros = [float(n.replace(",", ".")) for n in re.findall(r"\d+(?:[.,]\d+)?", texto)]
    peso = altura = None
    for n in numeros:
        if 0.5 <= n <= 2.5 and altura is None:
            altura = n
        elif 100 <= n <= 250 and altura is None:  # altura em cm
            altura = n / 100
        elif 20 <= n <= 300 and peso is None:
            peso = n
    if peso and altura:
        imc = round(peso / (altura ** 2), 1)
        return (
            f"📏 Com **{peso:g} kg** e **{altura:.2f} m**, seu IMC é **{imc:.1f}** → {classificar_imc(imc)}.\n\n"
            "Lembre-se: o IMC é só uma referência geral e não considera massa muscular, idade ou sexo. "
            "Para uma avaliação completa, procure um nutricionista ou médico."
        )
    return None


# ----------------------------------------------------------------------
# Lógica principal do chatbot
# ----------------------------------------------------------------------
def responder(mensagem: str) -> str:
    texto = normalizar(mensagem)
    if not texto:
        return "Digite alguma coisa pra gente conversar! 😊"

    if any(p in texto for p in ["relatorio", "avaliacao", "avaliar", "nota", "pontuacao", "meus habitos"]):
        return (
            "📋 Na aba **Avaliação de hábitos** você responde 7 perguntas rápidas e recebe uma nota de 0 a 100. "
            "Depois, na aba **Relatório**, dá pra baixar tudo em Word, com as dicas personalizadas!"
        )

    # IMC: se a pessoa falou de IMC ou mandou peso + altura
    if "imc" in texto or "massa corporal" in texto or ("kg" in texto and ("m" in texto or "cm" in texto)):
        resultado = tentar_calcular_imc(texto)
        if resultado:
            return resultado
        return (
            "📏 Posso calcular seu **IMC**! Me mande seu peso e altura assim:\n\n"
            "`70 kg 1,75 m`"
        )

    # Procura o tema com mais palavras-chave encontradas
    melhor_tema, melhor_pontos = None, 0
    for nome, tema in TEMAS.items():
        pontos = 0
        for palavra in tema["palavras"]:
            # palavras curtas precisam bater inteiras (ex.: "oi", "sal")
            if len(palavra) <= 3:
                if re.search(rf"\b{re.escape(palavra)}\b", texto):
                    pontos += 1
            elif palavra in texto:
                pontos += len(palavra.split())  # expressões maiores valem mais
        if pontos > melhor_pontos:
            melhor_tema, melhor_pontos = nome, pontos

    if melhor_tema:
        return random.choice(TEMAS[melhor_tema]["respostas"])
    return random.choice(NAO_ENTENDI)



# ----------------------------------------------------------------------
# Avaliação de hábitos
# Cada hábito recebe 0 (precisa de atenção), 1 (regular) ou 2 (bom).
# ----------------------------------------------------------------------
STATUS = {2: "Bom", 1: "Regular", 0: "Precisa de atenção"}
ICONE_STATUS = {2: "🟢", 1: "🟡", 0: "🔴"}

HABITOS = [
    {
        "chave": "agua",
        "titulo": "💧 Hidratação",
        "pergunta": "Quantos copos de água (250 ml) você bebe por dia?",
        "tipo": "numero", "min": 0, "max": 20, "padrao": 4,
        "nota": lambda v: 2 if v >= 8 else 1 if v >= 5 else 0,
        "dicas": {
            2: "Ótimo! Você está bem hidratado. Continue assim.",
            1: "Quase lá! Tente chegar a 8 copos por dia, levando uma garrafinha com você.",
            0: "Beba mais água: comece com um copo ao acordar e outro em cada refeição.",
        },
    },
    {
        "chave": "frutas",
        "titulo": "🍎 Frutas, verduras e legumes",
        "pergunta": "Quantas porções de frutas, verduras e legumes você come por dia?",
        "tipo": "numero", "min": 0, "max": 15, "padrao": 2,
        "nota": lambda v: 2 if v >= 5 else 1 if v >= 3 else 0,
        "dicas": {
            2: "Excelente! Continue variando as cores do prato.",
            1: "Bom começo! Inclua uma fruta no lanche e salada no almoço para chegar a 5 porções.",
            0: "Inclua frutas, verduras ou legumes em todas as refeições — a OMS recomenda pelo menos 5 porções por dia.",
        },
    },
    {
        "chave": "caseira",
        "titulo": "🍲 Comida de verdade",
        "pergunta": "Em quantos dias da semana você come comida caseira (feita na hora, com alimentos frescos)?",
        "tipo": "numero", "min": 0, "max": 7, "padrao": 4,
        "nota": lambda v: 2 if v >= 5 else 1 if v >= 3 else 0,
        "dicas": {
            2: "Muito bem! Comida caseira é a base de uma alimentação saudável.",
            1: "Tente cozinhar ou levar marmita mais vezes na semana.",
            0: "Planeje as refeições da semana e cozinhe em maior quantidade para congelar porções.",
        },
    },
    {
        "chave": "ultra",
        "titulo": "🍟 Ultraprocessados e refrigerante",
        "pergunta": "Em quantos dias da semana você consome refrigerante, salgadinho, fast food ou doces industrializados?",
        "tipo": "numero", "min": 0, "max": 7, "padrao": 3,
        "nota": lambda v: 2 if v <= 1 else 1 if v <= 4 else 0,
        "dicas": {
            2: "Ótimo controle! Deixe esses alimentos só para ocasiões especiais.",
            1: "Tente diminuir: troque o refrigerante por água com limão e o salgadinho por fruta ou pipoca caseira.",
            0: "O consumo está alto. Comece trocando um item por dia por uma opção natural.",
        },
    },
    {
        "chave": "sono",
        "titulo": "😴 Sono",
        "pergunta": "Quantas horas você dorme por noite, em média?",
        "tipo": "numero", "min": 0, "max": 14, "padrao": 6,
        "nota": lambda v: 2 if 7 <= v <= 9 else 1 if v in (6, 10) else 0,
        "dicas": {
            2: "Seu sono está na faixa recomendada (7 a 9 horas). Continue!",
            1: "Tente ajustar para 7 a 9 horas, com horários fixos para dormir e acordar.",
            0: "Seu sono precisa de atenção: evite telas 1 hora antes de deitar e mantenha uma rotina de horários.",
        },
    },
    {
        "chave": "exercicio",
        "titulo": "🏃 Atividade física",
        "pergunta": "Quantos minutos de atividade física você faz por semana?",
        "tipo": "numero", "min": 0, "max": 1500, "padrao": 60, "passo": 10,
        "nota": lambda v: 2 if v >= 150 else 1 if v >= 60 else 0,
        "dicas": {
            2: "Parabéns! Você atinge a recomendação da OMS (150 min por semana).",
            1: "Falta pouco para os 150 min semanais: some caminhadas curtas no dia a dia.",
            0: "Comece com 10 minutos de caminhada por dia e aumente aos poucos.",
        },
    },
    {
        "chave": "estresse",
        "titulo": "🧘 Estresse",
        "pergunta": "De 1 a 5, qual o seu nível de estresse na maior parte dos dias? (1 = tranquilo, 5 = muito estressado)",
        "tipo": "escala", "min": 1, "max": 5, "padrao": 3,
        "nota": lambda v: 2 if v <= 2 else 1 if v == 3 else 0,
        "dicas": {
            2: "Que bom! Continue reservando tempo para lazer e descanso.",
            1: "Faça pausas durante o dia e experimente exercícios de respiração.",
            0: "Cuide da sua saúde mental: converse com pessoas de confiança e, se precisar, procure um profissional. CVV: 188.",
        },
    },
]

EXEMPLO = {"nome": "Maria (exemplo)", "peso": 68.0, "altura": 1.65,
           "agua": 5, "frutas": 2, "caseira": 4, "ultra": 5, "sono": 6, "exercicio": 40, "estresse": 4}


def avaliar_habitos(respostas: dict) -> dict:
    """Recebe as respostas e devolve nota de 0 a 100 + detalhes de cada hábito."""
    detalhes = []
    for h in HABITOS:
        valor = respostas[h["chave"]]
        nota = h["nota"](valor)
        detalhes.append({
            "titulo": h["titulo"], "valor": valor, "nota": nota,
            "status": STATUS[nota], "dica": h["dicas"][nota],
        })
    pontos = sum(d["nota"] for d in detalhes)
    pontuacao = round(pontos / (2 * len(HABITOS)) * 100)
    if pontuacao >= 80:
        resumo = "Excelente! Seus hábitos estão muito bons. 🎉"
    elif pontuacao >= 50:
        resumo = "Bom caminho! Alguns ajustes vão fazer diferença na sua qualidade de vida. 💪"
    else:
        resumo = "Atenção! Vale a pena começar a mudar alguns hábitos, um passo de cada vez. 🌱"
    return {"pontuacao": pontuacao, "resumo": resumo, "detalhes": detalhes}


def calcular_imc(peso: float, altura: float):
    if not peso or not altura:
        return None
    imc = round(peso / (altura ** 2), 1)
    return {"valor": imc, "classe": classificar_imc(imc).replace(" ✅", "")}


# ----------------------------------------------------------------------
# Relatório em Word
# ----------------------------------------------------------------------
VERDE = RGBColor(0x2E, 0x7D, 0x32)


def gerar_relatorio_docx(nome: str, resultado: dict, imc, historico: list) -> bytes:
    doc = Document()
    estilo = doc.styles["Normal"]
    estilo.font.name = "Calibri"
    estilo.font.size = Pt(11)

    titulo = doc.add_heading("NutriBot — Relatório de Hábitos Saudáveis", level=0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    info = doc.add_paragraph()
    info.add_run("Nome: ").bold = True
    info.add_run(nome or "Não informado")
    info.add_run("\nData: ").bold = True
    info.add_run(datetime.now().strftime("%d/%m/%Y às %H:%M"))

    # Resultado geral
    doc.add_heading("1. Resultado geral", level=1)
    p = doc.add_paragraph()
    r = p.add_run(f"Pontuação: {resultado['pontuacao']}/100")
    r.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = VERDE
    doc.add_paragraph(resultado["resumo"])

    if imc:
        p = doc.add_paragraph()
        p.add_run("IMC: ").bold = True
        p.add_run(f"{imc['valor']} ({imc['classe']}). ")
        p.add_run("O IMC é apenas uma referência geral e não considera massa muscular, idade ou sexo.").italic = True

    # Tabela de hábitos
    doc.add_heading("2. Avaliação por hábito", level=1)
    tabela = doc.add_table(rows=1, cols=3)
    tabela.style = "Light Grid Accent 6"
    cab = tabela.rows[0].cells
    cab[0].text, cab[1].text, cab[2].text = "Hábito", "Sua resposta", "Situação"
    for d in resultado["detalhes"]:
        linha = tabela.add_row().cells
        linha[0].text = d["titulo"]
        linha[1].text = str(d["valor"])
        linha[2].text = d["status"]

    # Dicas
    doc.add_heading("3. Dicas personalizadas", level=1)
    ordenados = sorted(resultado["detalhes"], key=lambda d: d["nota"])  # piores primeiro
    for d in ordenados:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(f"{d['titulo']}: ").bold = True
        p.add_run(d["dica"])

    # Conversa
    if historico:
        doc.add_heading("4. Conversa com o NutriBot", level=1)
        for msg in historico:
            quem = "Você" if msg["role"] == "user" else "NutriBot"
            texto = re.sub(r"[*`>]", "", msg["content"])
            p = doc.add_paragraph()
            p.add_run(f"{quem}: ").bold = True
            p.add_run(texto)

    doc.add_paragraph()
    aviso = doc.add_paragraph(
        "Aviso: este relatório tem caráter educativo e foi baseado no Guia Alimentar para a "
        "População Brasileira (Ministério da Saúde) e em recomendações da OMS. "
        "Não substitui a avaliação de um nutricionista ou médico."
    )
    aviso.runs[0].italic = True
    aviso.runs[0].font.size = Pt(9)

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


# ----------------------------------------------------------------------
# Interface (Streamlit)
# ----------------------------------------------------------------------
CSS = """
<style>
.block-container {padding-top: 2rem; max-width: 1000px;}
.nb-header {display:flex; justify-content:space-between; align-items:center;
            border-bottom:1px solid #2a3240; padding-bottom:.6rem; margin-bottom:1rem;}
.nb-titulo {font-size:1.6rem; font-weight:700;}
.nb-sub {color:#9aa4b2; font-size:.95rem;}
.nb-badge {background:#163a22; color:#4ade80; border:1px solid #2f6f45;
           padding:.2rem .6rem; border-radius:999px; font-size:.75rem;}
.nb-card {background:#161b22; border:1px solid #2a3240; border-radius:10px;
          padding:1rem 1.2rem; margin-bottom:.8rem;}
.nb-nota {font-size:3rem; font-weight:800; color:#4ade80; line-height:1;}
.nb-rodape {color:#6b7280; font-size:.8rem; text-align:center; margin-top:2rem;}
</style>
"""


def iniciar_estado():
    padrao = {"nome": "", "peso": 0.0, "altura": 0.0}
    padrao.update({h["chave"]: h["padrao"] for h in HABITOS})
    for chave, valor in padrao.items():
        st.session_state.setdefault(f"in_{chave}", valor)
    st.session_state.setdefault("mensagens", [
        {"role": "assistant", "content": TEMAS["saudacao"]["respostas"][0]}
    ])
    st.session_state.setdefault("resultado", None)


def carregar_exemplo():
    for chave, valor in EXEMPLO.items():
        st.session_state[f"in_{chave}"] = valor
    st.session_state.resultado = None


def aba_chat():
    st.markdown("Converse comigo sobre **alimentação, hidratação, sono, exercícios e bem-estar**.")
    exemplos = ["ajuda", "quanto de água devo beber?", "ideias de lanche", "como dormir melhor?", "imc 70 kg 1,75 m"]
    colunas = st.columns(len(exemplos))
    clicado = None
    for col, ex in zip(colunas, exemplos):
        if col.button(ex, use_container_width=True):
            clicado = ex

    for msg in st.session_state.mensagens:
        with st.chat_message(msg["role"], avatar="🥗" if msg["role"] == "assistant" else "🙂"):
            st.markdown(msg["content"])

    pergunta = st.chat_input("Pergunte sobre alimentação ou qualidade de vida...") or clicado
    if pergunta:
        st.session_state.mensagens.append({"role": "user", "content": pergunta})
        st.session_state.mensagens.append({"role": "assistant", "content": responder(pergunta)})
        st.rerun()


def aba_avaliacao():
    topo = st.columns([3, 1])
    topo[0].markdown("Responda as perguntas e descubra sua **nota de hábitos saudáveis** de 0 a 100.")
    topo[1].button("📂 Carregar exemplo", on_click=carregar_exemplo, use_container_width=True)

    with st.container(border=True):
        st.markdown("**Seus dados** (opcional — usados para o IMC e o relatório)")
        c1, c2, c3 = st.columns(3)
        c1.text_input("Nome", key="in_nome")
        c2.number_input("Peso (kg)", min_value=0.0, max_value=300.0, step=0.5, key="in_peso")
        c3.number_input("Altura (m)", min_value=0.0, max_value=2.5, step=0.01, key="in_altura")

    with st.container(border=True):
        st.markdown("**Seus hábitos**")
        colunas = st.columns(2)
        for i, h in enumerate(HABITOS):
            col = colunas[i % 2]
            if h["tipo"] == "escala":
                col.slider(h["pergunta"], h["min"], h["max"], key=f"in_{h['chave']}")
            else:
                col.number_input(h["pergunta"], h["min"], h["max"], step=h.get("passo", 1), key=f"in_{h['chave']}")

    if st.button("✅ Avaliar meus hábitos", type="primary", use_container_width=True):
        respostas = {h["chave"]: st.session_state[f"in_{h['chave']}"] for h in HABITOS}
        st.session_state.resultado = avaliar_habitos(respostas)

    res = st.session_state.resultado
    if res:
        st.divider()
        c1, c2 = st.columns([1, 2])
        with c1:
            st.markdown(f"<div class='nb-card'><div class='nb-sub'>Sua nota</div>"
                        f"<div class='nb-nota'>{res['pontuacao']}</div><div class='nb-sub'>de 100</div></div>",
                        unsafe_allow_html=True)
            imc = calcular_imc(st.session_state.in_peso, st.session_state.in_altura)
            if imc:
                st.metric("IMC", imc["valor"], imc["classe"], delta_color="off")
        with c2:
            st.progress(res["pontuacao"] / 100)
            st.markdown(f"**{res['resumo']}**")
            for d in res["detalhes"]:
                st.markdown(f"{ICONE_STATUS[d['nota']]} **{d['titulo']}** — {d['status']}  \n"
                            f"<span class='nb-sub'>{d['dica']}</span>", unsafe_allow_html=True)
        st.info("📄 Vá até a aba **Relatório** para baixar o resultado completo em Word.")


def aba_relatorio():
    res = st.session_state.resultado
    if not res:
        st.warning("Faça primeiro a **Avaliação de hábitos** para gerar o relatório. "
                   "Dica: use o botão **Carregar exemplo** para testar rapidinho!")
        return
    imc = calcular_imc(st.session_state.in_peso, st.session_state.in_altura)
    incluir_chat = st.checkbox("Incluir a conversa com o NutriBot no relatório", value=True)
    historico = st.session_state.mensagens if incluir_chat else []

    texto_imc = f"{imc['valor']} ({imc['classe']})" if imc else "não informado"
    st.markdown(f"<div class='nb-card'><b>Resumo</b><br>Nome: {st.session_state.in_nome or 'Não informado'}<br>"
                f"Nota: <b>{res['pontuacao']}/100</b><br>IMC: {texto_imc}<br>"
                f"Mensagens na conversa: {len(historico)}</div>", unsafe_allow_html=True)

    arquivo = gerar_relatorio_docx(st.session_state.in_nome, res, imc, historico)
    nome_arq = f"relatorio_nutribot_{datetime.now():%Y%m%d_%H%M}.docx"
    st.download_button("⬇️ Baixar relatório em Word", arquivo, file_name=nome_arq, type="primary",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       use_container_width=True)


def aba_sobre():
    st.markdown(
        """
### Sobre o NutriBot
O **NutriBot** é um chatbot que incentiva hábitos de **alimentação saudável e qualidade de vida**.

**O que ele faz**
- 💬 Responde dúvidas sobre alimentação, hidratação, sono, exercícios e estresse
- 📋 Avalia 7 hábitos do dia a dia e dá uma nota de 0 a 100
- 📏 Calcula o IMC
- 📄 Gera um relatório em Word com dicas personalizadas

**Como funciona**
O chatbot usa **processamento de linguagem natural baseado em regras**: o texto é normalizado
(minúsculas e sem acentos) e comparado com palavras-chave de cada tema. O tema com mais
correspondências define a resposta.

**Fontes**
- Guia Alimentar para a População Brasileira — Ministério da Saúde
- Recomendações da Organização Mundial da Saúde (OMS) sobre atividade física, sal e frutas/verduras

⚠️ *As informações são educativas e não substituem um nutricionista ou médico.*
        """
    )


def main():
    st.set_page_config(page_title="NutriBot - Alimentação Saudável", page_icon="🥗", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    iniciar_estado()

    st.markdown(
        "<div class='nb-header'><div><div class='nb-titulo'>🥗 NutriBot</div>"
        "<div class='nb-sub'>Chatbot de alimentação saudável e qualidade de vida</div></div>"
        "<span class='nb-badge'>v1.0 · protótipo</span></div>",
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.header("🥗 NutriBot")
        st.write("Dicas de alimentação, sono, atividade física e bem-estar.")
        st.warning("⚠️ Informações educativas. Não substituem nutricionista ou médico.")
        if st.button("🗑️ Limpar conversa", use_container_width=True):
            st.session_state.mensagens = [{"role": "assistant", "content": TEMAS["saudacao"]["respostas"][0]}]
            st.rerun()

    chat, avaliacao, relatorio, sobre = st.tabs(["💬 Chat", "📋 Avaliação de hábitos", "📄 Relatório", "ℹ️ Sobre"])
    with chat:
        aba_chat()
    with avaliacao:
        aba_avaliacao()
    with relatorio:
        aba_relatorio()
    with sobre:
        aba_sobre()

    st.markdown("<div class='nb-rodape'>FMU · Tecnologia em Inteligência Artificial · NutriBot</div>",
                unsafe_allow_html=True)


if __name__ == "__main__":
    main()
