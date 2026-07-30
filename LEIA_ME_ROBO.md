# Robô de Consulta Processual — SEBRAE Contencioso

## O que este robô faz

Consulta automaticamente a **API pública do CNJ (DataJud)** para cada um dos
**519 processos com número CNJ válido** da lista de ausentes, identificando
quais estão **arquivados definitivamente** e quais seguem em andamento.

Os **86 processos com formato de número antigo** (pré-2006) são registrados
para consulta manual nos portais dos respectivos tribunais.

---

## Pré-requisitos

### Python 3.8 ou superior
Verificar versão instalada:
```
python --version
```
Se não tiver Python, baixar em: https://www.python.org/downloads/

### Bibliotecas necessárias
```
pip install requests pandas openpyxl tqdm
```

---

## Arquivos necessários (mesma pasta)

| Arquivo | Descrição |
|---|---|
| `robo_consulta_processos.py` | Script principal do robô |
| `processos_para_consulta.csv` | Lista de processos gerada pelo sistema |

---

## Como executar

### Windows
1. Abrir o **Prompt de Comando** (`cmd`) ou **PowerShell`
2. Navegar até a pasta onde estão os arquivos:
   ```
   cd C:\caminho\para\a\pasta
   ```
3. Executar:
   ```
   python robo_consulta_processos.py
   ```

### Mac / Linux
```bash
cd /caminho/para/a/pasta
python3 robo_consulta_processos.py
```

---

## O que esperar durante a execução

- O robô exibe o progresso em tempo real (barra de progresso)
- Cada processo leva ~0,7 segundos (respeitando o limite da API)
- **519 processos ≈ 6 a 10 minutos** de execução
- Logs são salvos em `robo_log.txt`

---

## Resultado gerado

Arquivo: **`Resultado_Consulta_Processual.xlsx`** com 5 abas:

| Aba | Conteúdo |
|---|---|
| Todos os Resultados | Lista completa com status de cada processo |
| Arquivados Definitivamente | Processos confirmados como arquivados |
| Em Andamento | Processos ativos confirmados |
| Verificar Manualmente | Processos sem CNJ padrão ou com status indefinido |
| Resumo | Totais e estatísticas da consulta |

---

## Possíveis resultados por processo

| Status | Significado | Ação |
|---|---|---|
| **ARQUIVADO** | API confirmou arquivamento | Remover do acompanhamento ativo |
| **EM ANDAMENTO** | Processo ativo no DataJud | Manter no acompanhamento |
| **NÃO ENCONTRADO** | Não consta no DataJud | Verificar se tribunal integrado ao CNJ |
| **ACESSO NEGADO** | Tribunal restringe acesso | Consultar portal do tribunal |
| **SEM CNJ PADRÃO** | Número em formato antigo | Consultar manualmente (ver tabela abaixo) |
| **TIMEOUT / ERRO** | Falha de conexão | Tentar novamente |

---

## Consulta manual — portais por tribunal

Para os **86 processos sem CNJ padrão**, consultar diretamente:

| Estado | Tribunal | Portal |
|---|---|---|
| AD, DF | TJDFT / TRF1 | https://www.tjdft.jus.br / https://processual.trf1.jus.br |
| MG | TJMG / TRF1 | https://www.tjmg.jus.br / https://processual.trf1.jus.br |
| SP | TJSP / TRF3 | https://esaj.tjsp.jus.br / https://www.trf3.jus.br |
| RJ | TJRJ / TRF2 | https://www3.tjrj.jus.br / https://eproc.trf2.jus.br |
| PR, SC, RS | TRF4 | https://eproc.trf4.jus.br |
| BA, AM, GO, ES, MS | TRF1 | https://processual.trf1.jus.br |

---

## Limitações conhecidas

1. **Integração DataJud**: Nem todos os tribunais alimentam o DataJud com
   movimentações em tempo real. Um processo pode estar arquivado sem que
   a API reflita essa informação.

2. **Processos antigos**: Processos de 2000–2006 têm numeração anterior ao
   padrão CNJ (adotado em 2010) e não são localizáveis pela API.

3. **Rate limiting**: A API tem limite de ~100 req/min. O robô respeita
   automaticamente com pausa de 0,7s entre chamadas.

4. **CAPTCHA**: Alguns portais de tribunais estaduais exigem CAPTCHA e não
   podem ser consultados automaticamente.

---

## Suporte

Em caso de dúvidas ou erros, verificar o arquivo `robo_log.txt` gerado
durante a execução para diagnóstico detalhado.
