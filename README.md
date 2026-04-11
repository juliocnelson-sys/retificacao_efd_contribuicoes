# EFD Contribuições — Retificação

Aplicativo web para retificação de arquivos EFD Contribuições.  
Desenvolvido em Python com Streamlit.

---

## Como usar

### Opção A — Streamlit Cloud (recomendado para empresa, grátis)

1. Crie uma conta em https://streamlit.io/cloud
2. Crie um repositório no GitHub e suba os arquivos desta pasta
3. Em Streamlit Cloud clique em **New app** → selecione o repositório → `app.py`
4. Clique em **Deploy** — em 1 minuto o link estará disponível para todos os usuários

### Opção B — Rodar localmente (Windows)

1. Instale o Python 3.11: https://www.python.org/downloads/
2. Abra o Prompt de Comando na pasta do projeto
3. Execute:
```
pip install -r requirements.txt
streamlit run app.py
```
4. O navegador abrirá automaticamente em http://localhost:8501

### Opção C — Servidor interno da empresa

```bash
pip install -r requirements.txt
streamlit run app.py --server.port 8080 --server.address 0.0.0.0
```

Acesse de qualquer computador da rede: http://IP_DO_SERVIDOR:8080

---

## Como preencher a planilha

| Coluna           | Formato      | Obrigatoriedade |
|------------------|--------------|-----------------|
| CNPJ_ESTAB       | 14 dígitos   | Obrigatório     |
| PERÍODO (MMAAAA) | Ex: 012024   | Obrigatório     |
| Demais campos    | Ver planilha | Conforme layout |

- Campos **laranja** = obrigatórios
- Campos **azul claro** = opcionais
- Campos **cinza** = calculados automaticamente pelo sistema

---

## Regras aplicadas automaticamente

- **0140**: CNPJ já existente no TXT é reutilizado. Novo registro criado apenas se não encontrado.
- **0200**: Itens centralizados — inseridos uma vez no bloco 0 geral, sem duplicidade.
- **A100/C100**: Totais (VL_DOC, VL_PIS, VL_COFINS) calculados somando os itens A170/C170.
- **Bloco 9**: Todos os contadores 9900 recalculados automaticamente.
- **0000**: Campo IND_RET marcado como `1` (retificador).

---

## Arquivos do projeto

```
app.py              ← Aplicativo principal
requirements.txt    ← Dependências Python
retifica_efd.py     ← Script de linha de comando (uso avançado)
EFD_Contribuicoes_Modelo_v3.xlsx  ← Planilha modelo para preenchimento
README.md           ← Este arquivo
```
