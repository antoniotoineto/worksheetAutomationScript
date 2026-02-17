# ℹ️ README
## Selecione a linguagem abaixo | Select the language below

<p align="center">
  <a href="#portugues">🇧🇷 Português</a> | 
  <a href="#english">🇺🇸 English</a>
</p>

---

<a id="portugues"></a>

# 📊 Automação de planilhas

Este repositório contém ferramentas de automação desenvolvidas para gerar planilhas padronizadas a partir da extração de horas realizada na Intranet da empresa (Zeus).

Atualmente, existem dois executáveis de automação:

- **DetailedWorkSheetAutomation** → Gera planilhas detalhadas de atividades por profissional.

- **ConsolidatedWorkSheetAutomation** → Gera planilhas consolidadas (em breve).

# 🚀 Como Funciona

Ambos os executáveis automatizam a transformação dos relatórios brutos de horas em planilhas Excel estruturadas e formatadas.

## 🧾 Passo a Passo de Utilização

### 1️⃣ Extraia as horas no Zeus

- Acesse a Intranet da empresa (Zeus).

- Exporte a planilha contendo as horas registradas pelos profissionais.

- Após o download, renomeie o arquivo para "**vilt.xlsx**"

### 2️⃣ Adicione o arquivo ao Projeto

- Mova o arquivo renomeado para o seguinte diretório: `/base`

- A estrutura deve ficar assim:

```
('detailed' ou 'consolidated')WorkSheetGenerator/
├── base/
│   ├── template.xlsx
│   ├── infos.xlsx
│   └── vilt.xlsx   ← (arquivo adicionado)
│
└── ('detailed' ou 'consolidated')WorkSheetGenerator.exe 
```

### 3️⃣ Execute a Automação

- Entre na pasta correspondente:

    - Para geração da planilha **DETALHADA** → abra a pasta detailedWorkSheetGenerator

    - Para geração da planilha **CONSOLIDADA** → abra a pasta consolidatedWorkSheetGenerator

- Dê duplo clique no arquivo executável: `detailedWorkSheetGenerator.exe` ou `consolidatedWorkSheetGenerator.exe`

## 🎯 Resultado Esperado

O script irá automaticamente:

- Validar os arquivos obrigatórios
- Processar os dados de entrada
- Gerar a planilha formatada
- Salvar o resultado dentro da pasta `/output`

## 🛡️ Aviso do Windows Defender

Em alguns ambientes, o Windows Defender pode exibir um aviso de segurança ao executar o arquivo `.exe`.

Isso acontece porque o executável foi gerado localmente e não possui assinatura digital.

Caso apareça o aviso:

- Clique em **Mais Informações**

- Selecione **Executar assim mesmo**

O arquivo é seguro e foi desenvolvido internamente para uso corporativo.

# 📁 Local do Arquivo Gerado

As planilhas geradas estarão disponíveis dentro de: `output`

# ⚙ Notas Técnicas

- Desenvolvido em Python

- Compilado utilizando PyInstaller

- Suporta execução como:

    - Script Python (.py)

    - Executável standalone (.exe)

# 📌 Manutenção

Para atualizações ou melhorias, modifique o código-fonte dentro do diretório `/src` e gere novamente o executável utilizando PyInstaller:

```
cd .\src\

python -m PyInstaller --onefile --name detailedWorkSheetGenerator main.py
```

<a id="english"></a>

# 📊 Worksheet Automation Tools

This repository contains automation tools designed to generate standardized worksheets based on raw hour extraction files from the company intranet system (Zeus).

There are currently two automation executables:

- **DetailedWorkSheetAutomation** → Generates activities detailed worksheets per professional.

- **ConsolidatedWorkSheetAutomation** → Generates consolidated worksheets (coming soon).

# 🚀 How It Works

Both executables automate the transformation of raw hour reports into structured and formatted Excel worksheets.

## 🧾 Step-by-Step Usage
### 1️⃣ Extract Hours from Zeus

- Access the company intranet system (Zeus).

- Export the worksheet containing professionals' logged hours.

- After downloading, rename the file to "**vilt.xlsx**"

### 2️⃣ Add the File to the Project

- Move the renamed file into the following directory: `/base`
- The structure should look like this:

```
('detailed' or 'consolidated')WorkSheetGenerator/
├── base/
│   ├── template.xlsx
│   ├── infos.xlsx
│   └── vilt.xlsx   ← (added file)
│
└── ('detailed' or 'consolidated')WorkSheetGenerator.exe 
```

### 3️⃣ Execute the Automation
- Enter the corresponding folder:

    - For **DETAILED** worksheet generation → open the detailedWorkSheetGenerator folder

    - For **CONSOLIDATED** worksheet generation → open the consolidatedWorkSheetGenerator folder

- Double click on the executable file: `detailedWorkSheetGenerator.exe` or `consolidatedWorkSheetGenerator.exe`

## 🎯 Expected Result

The script will automatically:

- Validate required files
- Process the input data
- Generate the output worksheet
- Save the result inside the /output folder

## 🛡️ Windows Defender Warning

In some environments, Windows Defender may display a security warning when executing the .exe file.

This happens because the executable is locally generated and not digitally signed.

If prompted:

- Click More Info

- Select Run Anyway

The file is safe and internally developed for company use.

# 📁 Output Location

Generated worksheets will be available inside: `output`

# ⚙ Technical Notes

- Developed in Python

- Compiled using PyInstaller

- Supports execution both as:

    - Python script (.py)

    - Standalone executable (.exe)

# 📌 Maintenance

For updates or improvements, modify the source code inside the /src directory and regenerate the executable using PyInstaller:

```
cd .\src\

python -m PyInstaller --onefile --name detailedWorkSheetGenerator main.py
```