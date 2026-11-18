# NutriBot 🥗

Chatbot educativo sobre alimentação saudável e qualidade de vida, desenvolvido com Streamlit.

## Funcionalidades

- Conversa sobre alimentação, hidratação, sono, atividade física e bem-estar.
- Avaliação de sete hábitos, com pontuação de 0 a 100.
- Cálculo informativo do IMC.
- Geração de relatório em Word (`.docx`) com os resultados e dicas.

## Executar localmente

É recomendado utilizar Python 3.12.

```bash
pip install -r requirements.txt
streamlit run app.py
```

O Streamlit abrirá o aplicativo no navegador.

## Publicar no Streamlit Community Cloud

1. Envie `app.py`, `requirements.txt`, `README.md` e a pasta `.streamlit` para a branch `main` do GitHub.
2. Acesse https://share.streamlit.io/ e entre com sua conta GitHub.
3. Selecione **Create app**, escolha o repositório, a branch `main` e o arquivo `app.py`.
4. Clique em **Deploy**.

## Aviso

O NutriBot oferece informações gerais e educativas. Não substitui a avaliação de nutricionista ou médico. O cálculo do IMC é apenas uma referência geral e tem limitações.
