# Revinn Estratégias - Gerador de Relatórios e Auditoria Hoteleira BI

![Python Version](https://img.shields.io/badge/python-3.11%2B-blue.svg)
![Framework](https://img.shields.io/badge/UI-CustomTkinter-159F92.svg)
![Docker](https://img.shields.io/badge/docker-ready-2496ED.svg)

Aplicação desktop profissional de **Business Intelligence e Auditoria Hoteleira** desenvolvida para a **Revinn Estratégias**. O sistema automatiza a ingestão, limpeza, processamento e visualização de dados operacionais (check-outs, reservas emitidas e dados de pickup/evolução YoY), consolidando relatórios gerenciais em múltiplos formatos (**Excel**, **Word** e **PDF Lâmina de Performance**).

---

## 📌 Funcionalidades Principais

- **Limpeza e Normalização de Dados de Auditoria (CSV):**
  - Leitura inteligente com tolerância a múltiplos encodigs (`utf-8`, `latin1`).
  - Tratamento de delimitadores residuais e remoção automatizada de linhas totais/inválidas.
  - Formatação monetária (BRL), cálculo de *Length of Stay* (LOS) e *Lead Time*.

- **Análise de Pickup e Evolução YoY (Excel):**
  - Leitura dinâmica de planilhas de evolução (`evolution_table`) e fechamento (`pickup_table`).
  - Cálculo de métricas-chave: Taxa de Ocupação, Diária Média (ADR) e RevPAR (Comparativo Ano Atual vs. Ano Anterior).

- **Geração de Gráficos e Visualizações:**
  - Gráfico de linha interativo de tendência diária YoY.
  - Gráficos de barras comparativos para fechamento de Ocupação e ADR.

- **Exportação Multiformato:**
  - **Excel (`.xlsx`):** Tabelas estruturadas de check-outs por categoria, por canal e reservas emitidas.
  - **Word (`.docx`):** Documento formatado no padrão corporativo da empresa (capa personalizada, tipografia Inter/Poppins, tabelas estilizadas, rodapés com numeração dinâmica).
  - **PDF (`.pdf`):** Lâmina de Performance executiva estilizada via HTML/CSS (WeasyPrint), com a seguinte estrutura:
    - **Capa** com margem fina na cor da marca e termo de confidencialidade.
    - **Subcapa "Análise de Check-outs"**, seguida dos KPIs (Receita, Ocupação, ADR, RevPAR) e das tabelas *Check-outs por Categoria*, *Check-outs por Canal* e *Distribuição por Data de Emissão (Check-outs)*.
    - **Subcapa "Análise de Emissões"**, com as tabelas *Emissão por Canal* e *Distribuição de Reservas por Previsto de Hospedagem*.
    - **Seções de canais** (Booking.com, Expedia), que podem ser ativadas/desativadas por switch na interface.
    - **Glossário de Termos** para leitores não familiarizados com hotelaria (ADR, RevPAR, OTB, etc.).
    - **Contracapa** apenas com margem fina (sem preenchimento sólido).

- **Distribuições por Mês:** leitura flexível de datas (`DD/MM/YYYY` com fallback automático), garantindo a distribuição correta entre todos os meses.

- **Interface Gráfica (GUI) Moderna:**
  - Interface desenvolvida com **CustomTkinter**.
  - Suporte a alternância entre **Modo Escuro (Dark Mode)** e **Modo Claro (Light Mode)**.
  - Seções expansíveis para inclusão de notas e capturas de tela dos canais.
  - **Switch por canal (Booking e Expedia):** desativar o switch omite a seção de análise de performance do canal no relatório.
  - Feedback em tempo real com barra de progresso e console de log detalhado.

- **Relatório Semanal (Dropdown no canto superior direito: `Semanal` / `Mensal`):**
  - Automação a partir de planilha modelo semanal (ex: `Modelo Hotel - ... - Venda Semanal.xlsx`).
  - Geração de PDF contendo as 3 tabelas de emissões (*Emissões por Canal*, *Emissões por Categoria* e *Evolução Semanal*) na mesma identidade visual corporativa do relatório mensal.
  - **Capa Preta Exclusiva** no padrão Revinn (preto `#050707` + verde `#159F92`/`#005151`).
  - **Gráficos de Benchmark (Exploratórios):** suíte com 10 visualizações distribuídas em 3 páginas ao final do PDF (evolução com média móvel e variação semanal, share de canal, mix por categoria, dia da semana de compra, antecedência ADS, permanência LOS, mix de tarifários, heatmap canal × categoria, mês de check-in e dispersão ADR × LOS) para suporte a novas análises.

---

## 🏗️ Arquitetura do Projeto

O projeto segue princípios de **Clean Architecture**, dividindo a aplicação em módulos bem definidos e desacoplados:

```text
report_generator/
├── main.py                     # Ponto de entrada da aplicação
├── requirements.txt            # Dependências Python
├── Dockerfile                  # Imagem Docker com dependências de sistema (GTK3, Tkinter)
├── docker-compose.yml          # Orquestração do container Docker
├── .dockerignore               # Arquivos ignorados no build do Docker
├── build_exe.ps1               # Script que gera o executável (PyInstaller) e o LEIA-ME.txt
├── README.md                   # Documentação do projeto
└── src/
    ├── core/
    │   ├── models.py           # Schemas e constantes (AuditSchema, PickupSchema, MONTH_MAP)
    │   ├── readers.py          # Leitura de arquivos (AuditCsvReader, PickupExcelReader)
    │   ├── processors.py       # Orquestração do processamento (DataProcessor, PickupProcessor)
    │   ├── weekly_reader.py    # Leitor otimizado da planilha modelo semanal (WeeklyWorkbookReader)
    │   └── utils/              # Funções puras reutilizáveis
    │       ├── currency.py     # clean_currency
    │       ├── dates.py        # parse_date_flex
    │       └── files.py        # read_lines_tolerant
    ├── infrastructure/
    │   ├── exporters.py        # Exportadores de arquivos mensais (Excel, Word docx, WeasyPrint PDF)
    │   └── weekly_exporter.py  # Exportador de PDF semanal com capa preta (WeasyPrint)
    ├── visualization/
    │   ├── visualizers.py      # Geradores de gráficos mensais e relatórios tabulares
    │   └── weekly_charts.py    # Gráficos benchmark semanais (Matplotlib puro, thread-safe)
    └── ui/
        ├── gui.py              # Interface gráfica principal e seletor Semanal/Mensal
        └── weekly_panel.py     # Painel e formulário do relatório semanal
```

---

## 🛠️ Requisitos do Sistema

### Módulos Python Necessários
- `pandas` (>= 2.0.0)
- `openpyxl` (>= 3.1.0)
- `python-docx` (>= 1.1.0)
- `matplotlib` (>= 3.7.0)
- `customtkinter` (>= 5.2.0)
- `weasyprint` (>= 60.0)
- `Pillow` (>= 10.0.0)

### Requisito para Windows (Exportação de PDF)
A geração de PDFs utiliza o **WeasyPrint**, que depende das bibliotecas C **GTK+3**.
- Se for executar diretamente no Windows (sem Docker), é necessário instalar o **GTK3 Runtime**:
  - Baixe e instale o [GTK3-Runtime Win64](https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases).

---

## 🚀 Como Executar o Projeto

### Opção 1: Execução Local (Python Native)

1. **Clone o repositório ou navegue até a pasta do projeto:**
   ```bash
   cd report_generator
   ```

2. **Crie e ative um ambiente virtual:**
   - **Windows (PowerShell):**
     ```powershell
     python -m venv venv
     .\venv\Scripts\Activate.ps1
     ```
   - **Linux / macOS:**
     ```bash
     python3 -m venv venv
     source venv/bin/activate
     ```

3. **Instale as dependências:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Execute a aplicação:**
   ```bash
   python main.py
   ```

---

### Opção 2: Executável Windows (.exe)

Gera uma versão standalone **com as DLLs do GTK3/WeasyPrint embutidas** — quem receber o executável **não precisa instalar Python nem o GTK3 Runtime**.

1. Na máquina de build, instale uma vez o [GTK3 Runtime](https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases) e as dependências (`pip install -r requirements.txt pyinstaller` no `venv`).
2. Gere o executável:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
   ```
3. O resultado fica em `dist\Gerador_Relatorios_Revinn\`, contendo:
   - `Gerador_Relatorios_Revinn.exe`
   - `_internal\` (bibliotecas, WeasyPrint, GTK3, fontes, logo)
   - `LEIA-ME.txt` (instruções de uso, gerado automaticamente pelo script)
4. Para distribuir, compacte e envie a **pasta completa**. O `.exe` deve ser executado de **dentro** dela, nunca copiado sozinho.

---

### Opção 3: Execução via Docker e Docker Compose

O Docker é ideal para rodar o projeto em um ambiente totalmente isolado e pré-configurado com as bibliotecas do GTK+3 e Tkinter.

#### No Linux / WSL2 (com servidor X11 habilitado):

1. **Permita a conexão da interface gráfica ao servidor X11:**
   ```bash
   xhost +local:docker
   ```

2. **Suba o container com o Docker Compose:**
   ```bash
   docker-compose up --build
   ```

#### No Windows (usando Docker Desktop + X-Server):
Para abrir a interface gráfica do Tkinter dentro de um container Docker no Windows:
1. Inicie um servidor X no Windows (ex: **VcXsrv** ou **Xming**).
2. Execute o comando:
   ```cmd
   docker-compose up --build
   ```

---

## 📊 Estrutura e Formato dos Dados de Entrada

### 1. Relatórios de Auditoria (CSV)
O sistema exige dois arquivos CSV principais:
- **Check-outs:** Contém o histórico de reservas finalizadas.
- **Reservas Emitidas:** Contém o volume de reservas captadas no período.

**Mapeamento das Colunas Esperadas:**
| Nome da Coluna no CSV | Mapeamento no Sistema | Descrição |
| :--- | :--- | :--- |
| `textBox17` | Canal | Nome do Canal de Venda (OTAs, Direto, etc.) |
| `textBox13` | Categoria | Categoria do Quarto/Acomodação |
| `textBox20` | CheckIn | Data de Entrada (`DD/MM/YYYY`) |
| `textBox19` | CheckOut | Data de Saída (`DD/MM/YYYY`) |
| `textBox22` | Receita | Valor monetário da reserva (ex: `1.250,00 BRL`) |
| `textBox3` | Status | Status da reserva |
| `textBox11` | DataReserva | Data de emissão da reserva |

### 2. Planilha de Pickup (Excel `.xlsx`) - *Opcional*
Permite comparar a evolução diária e o fechamento do mês contra o ano anterior (YoY):
- Deve possuir uma aba com `evolution_table` no nome para análise diária.
- Deve possuir uma aba com `pickup_table` no nome para fechamento consolidado.

---

## 📈 Relatórios Gerados

Ao clicar no botão **GERAR RELATÓRIOS**, a aplicação salvará no diretório de destino:

1. **`<Nome_Arquivo>.xlsx`**:
   - `Cat_CO`: Resumo de check-outs agrupados por Categoria de Acomodação (ADR, LOS, Receita).
   - `Canal_CO`: Resumo de check-outs agrupados por Canal de Venda (Receita, Reservas, LOS, ADR, ADS, Share %).
   - `Canal_RE`: Resumo de reservas emitidas agrupadas por Canal de Venda.

2. **`<Nome_Arquivo>.docx`**:
   - Documento no formato Microsoft Word pronto para apresentação gerencial ao cliente, contendo tabelas estilizadas, resumo de métricas e gráficos de tendência YoY.

3. **`<Nome_Arquivo>.pdf`**:
   - **Lâmina de Performance Executive PDF**, layout responsivo com visual corporativo moderno, cartões de KPIs principais e lâminas extras dos canais (Booking.com, Expedia).

---

## 📄 Licença e Créditos

Desenvolvido para **Revinn Estratégias** - Consultoria e Estratégia Hoteleira. Todos os direitos reservados.
