# 🚀 Publicar o Robô no Google Cloud Run

Guia para colocar o **Robô de Consulta Processual** online, com um link que
seu pai abre no navegador (upload do CSV → botão → download do Excel).

Sem GitHub. O deploy é feito **direto desta pasta**.

---

## Pré-requisitos (só na sua máquina, uma vez)

1. **Conta Google / projeto no Google Cloud** (você já tem: `ia@harpia.tech`).
2. **gcloud CLI** instalado: https://cloud.google.com/sdk/docs/install
3. **Faturamento ativado** no projeto (o Cloud Run tem free tier generoso;
   um uso ocasional como esse normalmente **não gera custo**).

---

## Passo a passo

Abra o terminal **nesta pasta** e rode, na ordem:

```bash
# 1) Login (abre o navegador)
gcloud auth login

# 2) Escolha/defina o projeto (troque pelo ID do seu projeto)
gcloud config set project SEU_PROJECT_ID

# 3) Habilita os serviços necessários (uma vez por projeto)
gcloud services enable run.googleapis.com cloudbuild.googleapis.com

# 4) Deploy direto da pasta local (o Google monta o container e publica)
gcloud run deploy consulta-processual \
  --source . \
  --region southamerica-east1 \
  --allow-unauthenticated \
  --memory 512Mi \
  --cpu 1 \
  --timeout 3600 \
  --concurrency 4 \
  --max-instances 2
```

No fim ele imprime uma linha assim:

```
Service URL: https://consulta-processual-xxxxxxxx-rj.a.run.app
```

**Esse é o link.** Manda pro seu pai. 🎉

---

## Por que cada parâmetro

| Parâmetro | Motivo |
|---|---|
| `--source .` | Monta o container a partir dos arquivos desta pasta (usa o `Dockerfile`). |
| `--region southamerica-east1` | Servidor em São Paulo — menor latência no Brasil. |
| `--allow-unauthenticated` | Deixa o link **público** (qualquer um com o link acessa). Veja "Proteger" abaixo. |
| `--timeout 3600` | A consulta pode levar ~10 min; o padrão do Cloud Run é 5 min. 3600 = 60 min. |
| `--concurrency 4` | Poucos usuários simultâneos — evita estourar o rate limit do CNJ. |
| `--max-instances 2` | Trava de segurança de custo (não sobe dezenas de instâncias). |

> Ao rodar pela 1ª vez, o gcloud pode perguntar se quer criar um repositório
> no **Artifact Registry** — responda **Y** (é onde a imagem fica guardada).

---

## Atualizar depois de mexer no código

É só rodar de novo o comando do **Passo 4**. Ele republica no mesmo link.

---

## 🔒 (Opcional) Proteger com senha

O link acima é público. Como é dado de cliente (SEBRAE), vale proteger.
Duas opções simples:

**A) Senha simples dentro do app** (mais fácil pro pai)
Peça pra eu adicionar uma telinha de senha no `app.py` (via `st.secrets`).
Aí o link é público, mas só entra quem sabe a senha.

**B) Login Google (IAM)** — mais seguro, porém o pai precisa ter conta Google
autorizada. Troque `--allow-unauthenticated` por `--no-allow-unauthenticated`
e depois:

```bash
gcloud run services add-iam-policy-binding consulta-processual \
  --region southamerica-east1 \
  --member "user:email-do-seu-pai@gmail.com" \
  --role roles/run.invoker
```

---

## 💰 Custo

- Cloud Run cobra **só enquanto processa** e **escala a zero** quando ocioso.
- Free tier mensal (2 milhões de requisições, tempo de CPU generoso) cobre
  folgado um uso esporádico. Na prática: **~R$ 0** para esse cenário.
- O `--max-instances 2` evita surpresas.

---

## Testar localmente antes (opcional)

```bash
pip install -r requirements.txt
streamlit run app.py
# abre em http://localhost:8501
```

---

## Arquivos que compõem o app web

| Arquivo | Função |
|---|---|
| `app.py` | Interface web (Streamlit): upload, progresso, download. |
| `core.py` | Lógica de consulta ao DataJud/CNJ + geração do Excel. |
| `requirements.txt` | Bibliotecas Python. |
| `Dockerfile` | Receita do container para o Cloud Run. |
| `.dockerignore` | Evita subir CSVs/resultados/venv no build. |

> Os arquivos antigos (`robo_consulta_gui.py`, `Consultar_Processos.bat`)
> continuam funcionando como versão desktop e **não atrapalham** o deploy
> (estão no `.dockerignore`).
