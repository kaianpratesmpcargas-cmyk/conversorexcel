# MP CARGAS — Conversor SSW

Aplicativo desktop profissional em Python para converter dados copiados diretamente do sistema **SSW** em formato de tabela Markdown/texto delimitado para planilhas **Excel (.xlsx)** com relatórios gerenciais, indicadores de atraso e agrupamento por destino.

---

## 🚀 Funcionalidades

- **Área de Colagem Ampla**: Permite colar tabelas diretamente com `Ctrl + V` ou através do botão dedicado **"Colar da Área de Transferência"**.
- **Parser Confiável e Inteligente**:
  - Limpeza automática de links Markdown `[CTRC](URL)`.
  - Remoção de tags HTML (`<br>`, `&#xA0;`, `&nbsp;`, etc.), negritos (`**`) e espaços duplicados.
  - Normalização de valores monetários no formato brasileiro (`R$ 1.250,50` ou `8.500,00`) e pesos em decimais.
  - Interpretação inteligente da coluna **Atraso**:
    - Maior que 0: **ATRASADO**
    - Igual a 0: **NO PRAZO**
    - Vazio ou ausente: **SEM INFORMAÇÃO**
  - Descarte seguro de separadores Markdown, cabeçalhos residuais, botões de paginação (`×`) e rodapés de sistema.
- **Dashboard e Tabela Interativa (Tkinter)**:
  - Cards com indicadores: **Total de CT-RCs**, **Atrasados**, **No Prazo**, **Sem Informação**, **Peso Total** e **Frete Total**.
  - Tabela com rolagem vertical e horizontal.
  - Busca rápida em tempo real por CTRC, Nota Fiscal, Remetente, Destinatário ou Destino.
  - Filtros combinados por **Situação** e **Destino**.
  - Cores dinâmicas para fácil visualização de entregas em atraso ou no prazo.
- **Exportação Profissional para Excel (.xlsx)**:
  - **Aba Acompanhamento**: Cabeçalho azul marinho estilizado, congelamento da 1ª linha (`freeze_panes`), autofiltro ativado, largura de coluna automática, formatação de moeda brasileira (`R$ #,##0.00`), peso numérico e destaque condicional de linhas/situações.
  - **Aba RESUMO**: Indicadores consolidados e tabela gerencial agrupada por **Destino** (Quantidade, Atrasados e Peso).
  - Diálogo padrão de salvamento com nome sugerido `MP_CARGAS_SSW_YYYY-MM-DD.xlsx` e confirmação antes de sobrescrever.
- **100% Local**: Não utiliza IA, não depende de banco de dados.

---

## 📂 Estrutura do Projeto

```text
CONVERSOR SSW/
├── main.py             # Interface gráfica principal e fluxo do aplicativo
├── parser_ssw.py       # Limpeza profunda, extração e validação dos dados do SSW
├── excel_export.py     # Geração e estilização profissional da planilha Excel (OpenPyXL)
├── dashboard.py        # Dashboard de métricas, cards e tabela filtrável (Tkinter)
├── requirements.txt    # Dependências mínimas do projeto
└── README.md           # Instruções e documentação
```

---

## 🛠️ Requisitos e Instalação

### 1. Pré-requisitos
- **Python 3.10** ou superior instalado no Windows.
- As bibliotecas `tkinter` já vêm instaladas por padrão no instalador oficial do Python para Windows.

### 2. Instalação das Dependências

Abra o prompt de comando (PowerShell ou CMD) na pasta do projeto e execute:

```powershell
pip install -r requirements.txt
```

---

## ▶️ Como Executar

Execute o comando:

```powershell
python main.py
```

### Passo a Passo de Uso:
1. Abra a consulta no sistema SSW e copie os dados da tabela.
2. No aplicativo **MP CARGAS**, clique no botão **"Colar da Área de Transferência"** ou pressione `Ctrl + V` na área de texto.
3. Clique em **"PROCESSAR TABELA"**.
4. Veja os indicadores nos cards e visualize a tabela com filtros de busca.
5. Clique em **"EXPORTAR EXCEL (.xlsx)"** e selecione o local desejado para salvar a planilha.


## Distribuição de tabelas por base e WhatsApp

A versão atual também permite montar uma fila de tabelas do SSW. Para cada tabela, o operador escolhe manualmente a base, evitando qualquer tentativa de identificar a base pelo conteúdo da tabela. O sistema então vincula automaticamente a base ao motorista e ao WhatsApp cadastrados em `motoristas.json`.

### Fluxo
1. Cole uma tabela do SSW.
2. Escolha a base.
3. Clique em **PROCESSAR E SALVAR NA FILA**.
4. Repita para as demais bases.
5. Clique em **ENVIAR TODAS NO WHATSAPP**.

As tabelas são enviadas separadamente, cada uma para o número da respectiva base. A fila é persistida em `fila_tabelas.json`.

### Integração com o bot WhatsApp

O conversor não cria um segundo WhatsApp. Ele chama o bot existente por HTTP. Configure no ambiente: `WHATSAPP_API_URL` e, se necessário, `WHATSAPP_API_TOKEN`.

O conversor envia ao endpoint um JSON com `to`, `message`, `base`, `motorista`, `queue_id` e `idempotency_key`. O endpoint do bot deve retornar HTTP 2xx quando o envio for aceito.

Exemplo de configuração no Windows:
```bat
set WHATSAPP_API_URL=http://127.0.0.1:PORTA/api/whatsapp/send
set WHATSAPP_API_TOKEN=SE_HOUVER_TOKEN
python web_app.py
```

> A URL acima é apenas um exemplo de configuração. Ela deve ser substituída pelo endpoint real do bot existente.

### Arquivos adicionados
- `motoristas.json`: cadastro Base → Motorista → WhatsApp.
- `fila_tabelas.json`: fila persistente de tabelas.
- `web_app.py`: interface e API da fila/envio.
- `web_app_original.py`: cópia da versão anterior.


## GERAR O EXE NO WINDOWS

1. Extraia esta pasta.
2. Dê duplo clique em `GERAR_EXE.bat`.
3. Aguarde a instalação e a compilação.
4. O EXE estará em `dist\MP_CARGAS_Conversor_SSW.exe`.
5. O EXE abre o sistema no navegador e o botão `ABRIR WHATSAPP` abre a conversa com o número da base e a mensagem preenchida.

Não é necessário WhatsApp API, Baileys ou token. O envio é manual: o WhatsApp abre e você clica em Enviar.
