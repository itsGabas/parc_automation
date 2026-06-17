# 📄 Automação de Parcelamentos Estaduais

Automação desenvolvida para um escritório de contabilidade com o objetivo de eliminar o processo manual de download mensal de guias de parcelamentos tributários estaduais — tarefa que consumia horas da equipe e estava sujeita a erros humanos.

O sistema acessa os portais fiscais dos estados de **São Paulo**, **Minas Gerais** e **Rio Grande do Sul**, baixa as guias do mês vigente, organiza os arquivos em pastas e atualiza uma planilha de controle com o status de cada parcelamento, colorindo as linhas conforme a situação encontrada.

---

## 🚀 Funcionalidades

- Acesso automatizado aos portais da **SEFAZ-RS**, **Fazenda-MG** e **PGE-SP (Dívida Ativa)**
- Resolução automática de captcha via integração com a API do **2Captcha**
- Download e renomeação automática dos PDFs das guias, organizados por mês e tipo de parcelamento
- Leitura de planilha Excel com os dados dos parcelamentos do mês corrente
- Atualização automática da planilha de controle com:
  - ✅ Verde → guia baixada com sucesso
  - 🟡 Amarelo → parcelamento com parcelas atrasadas
  - 🔴 Vermelho → erro, parcelamento desistente ou situação irregular
- Log completo de toda a execução no terminal
- Empacotável como `.exe` via PyInstaller para uso sem instalação de Python

---

## 🗂️ Estrutura do Projeto

```
📁 automacao-parcelamentos/
├── main.py              # Orquestrador — executa os três módulos em sequência
├── da_sp_att.py         # Módulo SP — Portal da Dívida Ativa (PGE-SP)
├── mg_atual.py          # Módulo MG — Fazenda de Minas Gerais
├── sefaz_rs_att.py      # Módulo RS — SEFAZ Rio Grande do Sul
├── adm_sp.py            # Módulo SP ADM — Portal CAWeb Fazenda SP
├── utils.py             # Funções utilitárias compartilhadas entre os módulos
├── config.py            # Leitura centralizada das variáveis de ambiente
├── .env.example         # Modelo do arquivo de configuração (sem dados reais)
├── .gitignore
└── README.md
```

---

## ⚙️ Tecnologias Utilizadas

| Tecnologia | Uso |
|---|---|
| Python 3.11+ | Linguagem principal |
| Selenium + undetected-chromedriver | Automação e navegação web |
| pandas | Leitura e manipulação da planilha de entrada |
| openpyxl | Escrita e formatação da planilha de controle |
| 2Captcha API | Resolução automática de captchas nos portais |
| python-dotenv | Gerenciamento de variáveis de ambiente |
| WebDriver Manager | Gerenciamento automático do ChromeDriver |
| PyInstaller | Geração de executável `.exe` |

---

## 🔧 Como Configurar

### Pré-requisitos

- Python 3.11 ou superior
- Google Chrome instalado
- Conta ativa no [2Captcha](https://2captcha.com) com saldo disponível

### Instalação

```bash
# Clone o repositório
git clone https://github.com/itsGabas/parc_automation.git
cd automacao-parcelamentos

# Crie e ative um ambiente virtual
python -m venv venv
venv\Scripts\activate  # Windows

# Instale as dependências
pip install -r requirements.txt
```

### Configuração do ambiente

```bash
# Copie o arquivo de exemplo
copy .env.example .env
```

Edite o arquivo `.env` com seus dados:

```env
CAMINHO_PLANILHA=C:\caminho\para\PARCELAMENTO ESTADUAL.xlsx
PASTA_DOWNLOADS=C:\caminho\para\downloads
PASTA_BASE=C:\caminho\para\pasta\de\saida
TWOCAPTCHA_API_KEY=sua_chave_aqui
```

### Execução

```bash
python main.py
```

---

## 📋 Como Funciona

### Planilha de entrada

O sistema lê a aba correspondente ao mês atual (ex: `JUNHO 2025`) de uma planilha Excel com os parcelamentos cadastrados. Identifica automaticamente as colunas de tipo e número de parcelamento, filtrando apenas os registros de cada estado.

### Fluxo de execução

```
main.py
  ├── da_sp_att.py  →  Acessa PGE-SP → resolve captcha → baixa GARE → atualiza planilha
  ├── mg_atual.py   →  Acessa Fazenda MG → resolve Turnstile → baixa guia → atualiza planilha
  └── sefaz_rs_att.py → Acessa SEFAZ RS → resolve captcha → baixa DAT → atualiza planilha
```

### Organização dos arquivos baixados

Os PDFs são salvos automaticamente na estrutura:

```
📁 PASTA_BASE/
└── 06 - JUN 2025/
    ├── Parcelamentos D.A. - SP/
    │   └── DARE_EMPRESA XYZ - 123456.pdf
    ├── Parcelamentos MG/
    │   └── GNRE_EMPRESA XYZ.pdf
    └── Parcelamentos RS/
        └── DAT_EMPRESA XYZ.pdf
```

---

## 🔐 Segurança

- Nenhuma credencial, caminho de rede ou chave de API está hardcoded no código
- Todos os dados sensíveis são carregados exclusivamente via arquivo `.env` (não versionado)
- O arquivo `.env.example` serve como guia de configuração sem expor dados reais
- A integração com serviço de captcha é mencionada na documentação, mas a chave permanece local

---

## 💡 Contexto e Motivação

Este projeto nasceu da observação de um processo repetitivo: todo mês, a equipe administrativa precisava acessar manualmente os portais de três estados diferentes, baixar dezenas de guias e registrar o status de cada uma numa planilha — um trabalho que consumia horas e estava sujeito a esquecimentos e inconsistências.

A automação eliminou esse processo por completo. Hoje o sistema roda com um clique duplo no executável e entrega os arquivos organizados e a planilha atualizada em minutos.

---

## 📌 Observações

- O sistema foi desenvolvido e testado em **Windows**
- O módulo `adm_sp.py` depende de `pywinauto` e `pyautogui` para interagir com o popup nativo do Windows de seleção de certificado digital — por isso, **esse módulo específico não é portável para Linux/Mac**
- Os portais governamentais podem alterar sua estrutura HTML sem aviso, exigindo manutenção pontual dos seletores XPath
- A resolução de captcha depende de saldo ativo na conta 2Captcha

---

## 👨‍💻 Autor

**Gabas** — Desenvolvedor em formação, com foco em automação, back-end Python e análise de dados.

[![LinkedIn](https://img.shields.io/badge/LinkedIn-blue?style=flat&logo=linkedin)](www.linkedin.com/in/gabriel-bonamichi-aab43b2a8)
[![GitHub](https://img.shields.io/badge/GitHub-black?style=flat&logo=github)](https://github.com/itsGabas)
