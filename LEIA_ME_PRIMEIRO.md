# 📋 Como usar o Robô de Consulta Processual

Um programa simples que verifica na base oficial do CNJ (DataJud) **quais
processos já estão arquivados definitivamente** e quais continuam em andamento.

---

## ✅ Como usar (bem simples)

1. **Dê dois cliques** no arquivo **`Consultar_Processos.bat`**
   - Na **primeira vez**, ele instala sozinho tudo o que precisa
     (pode levar alguns minutos — é normal, basta aguardar).
   - Se aparecer uma janelinha do Windows pedindo permissão, clique em **Sim**.

2. Vai abrir uma **janela azul**. Nela:
   - Clique em **📂 Escolher CSV...** e selecione a planilha de processos
     (por exemplo `processos_para_consulta_v2.csv`).
   - Clique em **▶ Iniciar consulta**.

3. Acompanhe a **barra de progresso** e os contadores
   (arquivados / em andamento / não encontrados).

4. Quando terminar, clique em **📊 Abrir resultado**.
   - O relatório em Excel é salvo **na mesma pasta**, com nome
     `Resultado_Consulta_AAAA-MM-DD_HHMM.xlsx`.

Pode fechar e abrir quantas vezes quiser. Da segunda vez em diante,
abre na hora (não precisa instalar nada de novo).

---

## 📊 O que aparece no resultado (Excel com 5 abas)

| Aba | O que mostra |
|---|---|
| Todos os Resultados | Lista completa, com o status de cada processo |
| ✅ Arquivados Definitivamente | Processos que podem sair do acompanhamento |
| ❌ Em Andamento | Processos que continuam ativos |
| ❓ Verificar Manualmente | Sem número CNJ novo ou não localizados |
| 📊 Resumo | Totais e portais para consulta manual |

---

## ❓ Perguntas comuns

**Preciso instalar alguma coisa?**
Não manualmente. O `Consultar_Processos.bat` instala tudo automaticamente
na primeira vez.

**A chave de acesso ao CNJ pode vencer?**
Não é problema. O robô **busca a chave oficial atual do CNJ automaticamente**
toda vez que roda. Se o CNJ trocar a chave, o robô continua funcionando.

**Deu erro. E agora?**
Tire uma foto da janela preta com a mensagem e envie para o suporte.

**Quanto tempo demora?**
Cerca de **0,7 segundo por processo** (respeitando o limite do CNJ).
Uma lista de ~500 processos leva de **6 a 10 minutos**.

---

## 📁 Arquivos desta pasta

| Arquivo | Para que serve |
|---|---|
| **`Consultar_Processos.bat`** | ▶️ É este que você abre (dois cliques) |
| `robo_consulta_gui.py` | O robô em si (não precisa abrir) |
| `processos_para_consulta_v2.csv` | Lista de processos a consultar |
| `Resultado_Consulta_*.xlsx` | Relatórios gerados (criados após rodar) |
| `.venv` (pasta) | Ambiente isolado criado automaticamente — **não mexer** |

> ℹ️ Tudo é instalado **dentro desta pasta** (numa subpasta `.venv`), sem
> mexer no resto do computador e sem precisar de senha de administrador.
> Se um dia quiser remover o robô por completo, basta apagar esta pasta.
