# 🦁 Organizador de IR — LION APP SYSTEM

Planilha Excel (`organizador_IR.xlsx`) criada para **organizar, em um só lugar, as informações necessárias para a declaração de Imposto de Renda da Pessoa Física (IRPF)**: dados do titular, saldos informados pelos bancos e entradas mensais de receita.

A planilha funciona como um pequeno "aplicativo" dentro do Excel: possui menu de navegação clicável, campos protegidos contra edição acidental, listas suspensas, formatação automática de documentos (CPF, CEP, telefone) e soma automática dos valores.

> ⚠️ **Os dados presentes no arquivo são fictícios** (nome, CPF, banco, valores etc.) e servem apenas como exemplo de preenchimento.

---

## 📑 Sumário

1. [Visão geral](#-visão-geral)
2. [Estrutura do workbook](#-estrutura-do-workbook)
3. [Detalhamento por aba](#-detalhamento-por-aba)
4. [Recursos do Excel aplicados](#-recursos-do-excel-aplicados)
5. [Fórmulas](#-fórmulas)
6. [Validação de dados](#-validação-de-dados)
7. [Proteção e segurança](#-proteção-e-segurança)
8. [Navegação (menu e botões)](#-navegação-menu-e-botões)
9. [Design e identidade visual](#-design-e-identidade-visual)
10. [Formatos numéricos personalizados](#-formatos-numéricos-personalizados)
11. [Compatibilidade](#-compatibilidade)
12. [Como usar](#-como-usar)
13. [Limitações conhecidas e ideias de evolução](#-limitações-conhecidas-e-ideias-de-evolução)
14. [Estrutura sugerida do repositório](#-estrutura-sugerida-do-repositório)

---

## 🎯 Visão geral

| Item | Descrição |
|---|---|
| **Arquivo** | `organizador_IR.xlsx` |
| **Finalidade** | Coletar e organizar dados para a declaração anual de IRPF |
| **Abas visíveis** | `TITULAR`, `INFORMES`, `NOTAS` |
| **Abas ocultas** | `BANCOS` (base de dados das listas suspensas) |
| **Fórmulas** | 1 (soma do total dos informes bancários) |
| **Tabelas estruturadas** | 1 (`Tabela2`, na aba `NOTAS`) |
| **Listas suspensas** | 5 áreas de validação em 3 abas |
| **Macros (VBA)** | ❌ Nenhuma — arquivo `.xlsx` puro |
| **Fluxo** | `1. TITULAR` → `2. INFORMES` → `3. NOTAS` |

---

## 🗂️ Estrutura do workbook

```
organizador_IR.xlsx
├── TITULAR    → 1. Dados do titular           (C2:E18)
├── INFORMES   → 2. Informes de rendimentos    (C2:F1000 formatada)
├── NOTAS      → 3. Notas bancárias/holerites  (C2:E35, com Tabela2)
└── BANCOS     → (oculta) lista de 50 bancos   (A1:A51)
```

Cada aba visível segue o mesmo padrão de layout:

- **Coluna A**: faixa lateral com o menu/logotipo (imagem + botões flutuantes);
- **Coluna B**: margem em branco;
- **Colunas C–E**: área de conteúdo (título, subtítulo, rótulos e campos de preenchimento).

---

## 🔍 Detalhamento por aba

### 1️⃣ `TITULAR` — Dados do titular

Formulário de cadastro da pessoa física. Rótulos na coluna **C**, campos de preenchimento na coluna **D**.

| Campo | Célula | Tipo / Formato | Observação |
|---|---|---|---|
| Nome | `D5` | Texto | |
| CPF | `D6` | Número com máscara `000.000.000-00` | Digita-se apenas os 11 dígitos |
| Nascimento | `D7` | Texto (`dd/mm/aaaa`) | Armazenado como texto |
| Título de eleitor | `D8` | Número | |
| Cônjuge | `D9` | Texto | |
| Rua | `D10` | Texto | |
| Rua abreviada | `D11` | Texto | Versão curta do endereço |
| CEP | `D12` | Número com máscara `00000-000` | |
| Telefone | `D13` | Número com máscara `(00)0000-0000` | |
| Celular | `D14` | Número com máscara `(00)00000-0000` | |
| E-mail | `D15` | Texto com **hiperlink `mailto:`** | Abre o cliente de e-mail com assunto pré-definido |
| Houve alterações da entrega anterior | `D16` | Lista suspensa `SIM / NÃO` | |
| Dependente cônjuge | `D17` | Lista suspensa `SIM / NÃO` | |
| Residente do exterior | `D18` | Lista suspensa `SIM / NÃO` | |

### 2️⃣ `INFORMES` — Informes de rendimentos bancários

Registra o saldo/valor informado por até **3 bancos** e calcula o total.

| Bloco | Células | Conteúdo |
|---|---|---|
| Total geral | `C5:C6` | Rótulo `TOTAL` + fórmula de soma em `C6` |
| 1º Banco | `C8:D11` | Banco (`D9`), Valor atual (`D10`), Anexo (`D11`) |
| 2º Banco | `C13:D16` | Banco (`D14`), Valor atual (`D15`), Anexo (`D16`) |
| 3º Banco | `C18:D21` | Banco (`D19`), Valor atual (`D20`), Anexo (`D21`) |

- O campo **BANCO** é uma **lista suspensa** alimentada pela aba oculta `BANCOS`.
- O campo **ANEXO 🖇️** recebe o nome do arquivo comprobatório (ex.: `informe_2025.pdf`).
- A aba tem formatação aplicada até a linha **1000**, preparada para crescimento futuro.

### 3️⃣ `NOTAS` — Notas bancárias ou extrato de holerites

Registro mês a mês de todas as entradas de receita.

- **Faixa "ENTRADAS"** em `C6:E6` (célula mesclada, fundo rosa `#EE37BF`, texto branco).
- **`Tabela2`** (`C7:E35`) com 3 colunas e **28 linhas de dados** disponíveis (linhas 8 a 35):

| Coluna | Formato | Validação |
|---|---|---|
| `DATA` | `mmmm-yyyy` (ex.: *janeiro-2026*) | — |
| `CATEGORIA` | Texto | Lista suspensa: `Holerite`, `CNPJ`, `Freelancer` |
| `VALOR` | `R$ #,##0.00` | — |

- Estilo da tabela: `TableStyleMedium2` com linhas zebradas e filtro automático.
- Título (`C3:E3`) mesclado para o subtítulo explicativo.

### 🔒 `BANCOS` — Base de dados (aba oculta)

- Intervalo `A1:A51`: cabeçalho `bancos` + **50 instituições** no formato `código - nome` (ex.: `33 - Banco Santander`, `260 - Nubank`, `77 - Banco Inter`).
- Serve exclusivamente como **fonte da lista suspensa** da aba `INFORMES`.
- Para editar a lista, é preciso reexibir a aba (clique direito em qualquer aba → *Reexibir*; como a barra de abas está minimizada, veja a seção [Limitações](#-limitações-conhecidas-e-ideias-de-evolução)).

---

## 🧰 Recursos do Excel aplicados

| # | Recurso | Onde é usado |
|---|---|---|
| 1 | **Fórmula com `SUM`** | `INFORMES!C6` |
| 2 | **Validação de dados – lista fixa** | `TITULAR!D16:D18`, `NOTAS!D8:D35` |
| 3 | **Validação de dados – lista referenciando outra aba** (extensão x14) | `INFORMES!D9`, `D14`, `D19` → `BANCOS!$A$2:$A$51` |
| 4 | **Mensagens de entrada e de erro personalizadas** | Listas de banco em `INFORMES` |
| 5 | **Proteção de planilha** (células bloqueadas × desbloqueadas) | `TITULAR`, `INFORMES`, `NOTAS` |
| 6 | **Aba oculta** | `BANCOS` |
| 7 | **Tabela estruturada** (`Tabela2`) com estilo e filtro | `NOTAS!C7:E35` |
| 8 | **Células mescladas** | `NOTAS!C3:E3` e `NOTAS!C6:E6` |
| 9 | **Formatos numéricos personalizados** (CPF, CEP, telefone, moeda, data) | `TITULAR`, `INFORMES`, `NOTAS` |
| 10 | **Hiperlink de e-mail** (`mailto:` com assunto) | `TITULAR!D15` |
| 11 | **Formas (retângulos arredondados) com hiperlink interno** | Menu de navegação das 3 abas |
| 12 | **Imagem (logotipo PNG 259×259)** | Cabeçalho lateral das 3 abas |
| 13 | **Ocultação de linhas de grade** | Todas as abas |
| 14 | **Ocultação de cabeçalhos de linha/coluna** | `TITULAR`, `INFORMES`, `NOTAS` |
| 15 | **Barra de abas minimizada** (`tabRatio=0`) | Workbook |
| 16 | **Larguras de coluna personalizadas** | Todas as abas |
| 17 | **Formatação de células** (preenchimentos, fontes, cores, alinhamento) | Todas as abas |
| 18 | **Emoji em rótulos** (🖇️ e 💜) | `INFORMES` e menu |

---

## ➕ Fórmulas

A planilha possui **uma única fórmula**, na aba `INFORMES`:

| Célula | Fórmula | Função |
|---|---|---|
| `INFORMES!C6` | `=(SUM(D10,D15,D20))` | Soma os valores atuais dos três bancos informados |

Formato de exibição: `"R$" #,##0.00;[Red]-"R$" #,##0.00` — valores negativos aparecem em **vermelho**.

> 💡 Como a soma referencia células individuais (`D10`, `D15`, `D20`) e não um intervalo, **novos bancos adicionados abaixo do 3º não entram no total** automaticamente — veja as [ideias de evolução](#-limitações-conhecidas-e-ideias-de-evolução).

---

## ✅ Validação de dados

| Aba | Intervalo | Tipo | Fonte / Opções | Mensagens |
|---|---|---|---|---|
| `TITULAR` | `D16:D18` | Lista | `SIM,NÃO` | Alerta de erro padrão |
| `INFORMES` | `D9`, `D14`, `D19` | Lista (referência a outra aba) | `BANCOS!$A$2:$A$51` | **Entrada:** *"informe um banco"* → *"Informe um banco vinculado ao seu CPF"*<br>**Erro:** *"Banco não encontrado"* → *"Informe um banco da lista"* |
| `NOTAS` | `D8:D35` | Lista | `Holerite,CNPJ,Freelancer` | Alerta de erro padrão |

Observações técnicas:

- A validação em `INFORMES` usa o **recurso estendido x14** do Excel (`x14:dataValidations`), que permite referenciar intervalos de **outra planilha** de forma nativa.
- Todas as listas têm `allowBlank` ativo, ou seja, o campo pode permanecer vazio.

---

## 🔐 Proteção e segurança

- As abas `TITULAR`, `INFORMES` e `NOTAS` estão com **proteção de planilha ativada**, bloqueando objetos e cenários.
- **Apenas as células de preenchimento estão desbloqueadas** (`locked = False`):
  - `TITULAR!D5:D18`
  - `INFORMES!D9:D11`, `D14:D16`, `D19:D21`
  - `NOTAS!C7:E35` (incluindo o cabeçalho da tabela)
- **Rótulos, títulos, fórmula de total, logotipo e botões ficam bloqueados**, evitando que o usuário apague ou desconfigure a estrutura.
- A opção *"Selecionar células bloqueadas"* está **desabilitada**: o cursor só consegue navegar pelas células de entrada.
- A proteção é aplicada **sem senha** (pode ser removida em *Revisão → Desproteger Planilha*). Para uso mais restritivo, defina uma senha.
- A aba `BANCOS` está **oculta**, mas **sem proteção**.

---

## 🧭 Navegação (menu e botões)

Cada aba visível possui, na lateral esquerda, um painel de navegação construído com **imagem + formas do Excel**:

| Elemento | Descrição |
|---|---|
| **Logotipo** | Imagem PNG (259×259 px) |
| **Cabeçalho "LION APP"** | Retângulo azul-marinho `#0B1F4B` |
| **Botão `TITULAR`** | Retângulo arredondado → hiperlink para `TITULAR!C1` |
| **Botão `INFORMES`** | Retângulo arredondado → hiperlink para `INFORMES!C1` |
| **Botão `NOTAS`** | Retângulo arredondado → hiperlink para `NOTAS!C1` |
| **Assinatura** | Retângulo com o texto **"SYSTEM BY DYADIX 💜"** |
| **Botões `PRÓXIMO =>` / `<= ANTERIOR`** | Navegação sequencial entre as etapas (1 → 2 → 3) |

Distribuição dos botões sequenciais:

| Aba | `<= ANTERIOR` | `PRÓXIMO =>` |
|---|:-:|:-:|
| `TITULAR` | — | ✅ |
| `INFORMES` | ✅ | ✅ |
| `NOTAS` | ✅ | — |

Todos os links são **internos** (`#ABA!C1`), sem macros, funcionando apenas com hiperlinks aplicados às formas.

---

## 🎨 Design e identidade visual

### Paleta de cores

| Uso | Cor | Hex |
|---|---|---|
| Títulos das seções | Roxo | `#7030A0` |
| Subtítulos (fundo) | Rosa claro | `#F1CEEE` |
| Texto dos subtítulos | Cinza escuro | `#3A3A3A` |
| **Campos de preenchimento** (fundo) | Amarelo claro | `#FFEB9C` |
| **Campos de preenchimento** (texto) | Marrom | `#9C5700` |
| Cabeçalho da tabela | Azul-marinho | `#0B1F4B` |
| Faixa "ENTRADAS" | Rosa/magenta | `#EE37BF` |
| Botões do menu | Roxo / azul-marinho | `#7030A0` / `#002060` |
| Botão de destaque | Azul acinzentado / roxo escuro | `#7187AF` / `#3F0065` |

> 🟨 **Convenção visual:** todo campo em **amarelo** (`#FFEB9C`) é um campo que o usuário deve preencher.

### Tipografia

| Elemento | Fonte | Tamanho |
|---|---|---|
| Títulos de seção | Aptos Narrow (negrito) | 15 pt |
| Rótulos | Quattrocento Sans | 11 pt |
| Campos de preenchimento | Aptos Narrow | 11 pt |
| Total (`INFORMES!C6`) | Aptos Narrow (negrito) | 12 pt |
| Cabeçalho da tabela | Quattrocento Sans (negrito, branco) | 10 pt |
| Linhas da tabela | Calibri | 11 pt |

### Layout

- Linhas de grade **ocultas** e cabeçalhos de linha/coluna **ocultos** para aparência de aplicativo.
- Larguras de coluna ajustadas por aba:

| Aba | A | C | D | E | F–I |
|---|---|---|---|---|---|
| `TITULAR` | 27 | 60,7 | 60,7 | 25,7 | — |
| `INFORMES` | 27 | 60,7 | 60,7 | 25,7 | 37,3 |
| `NOTAS` | 27 | 26,9 | 60,7 | 38,6 | — |

---

## 🔢 Formatos numéricos personalizados

| Campo | Código de formato | Entrada → Exibição |
|---|---|---|
| CPF | `000"."000"."000"-"00` | `12312312315` → `123.123.123-15` |
| CEP | `00000\-000` | `347415` → `00347-415` |
| Telefone | `"("00")"0000"-"0000` | `1126856965` → `(11)2685-6965` |
| Celular | `"("00")"00000"-"0000` | `11975683265` → `(11)97568-3265` |
| Valor (campos) | `"R$"\ #,##0.00` | `5333333.33` → `R$ 5.333.333,33` |
| Total | `"R$"\ #,##0.00;[Red]\-"R$"\ #,##0.00` | Negativos em vermelho |
| Data (tabela) | `mmmm\-yyyy` | `01/01/2026` → `janeiro-2026` |

> O uso de máscaras via formato numérico mantém o valor armazenado como **número puro**, facilitando conferências e eventuais exportações.

---

## 💻 Compatibilidade

| Software | Situação |
|---|---|
| **Microsoft Excel 2010+ / Microsoft 365** | ✅ Totalmente compatível (recomendado) |
| **Excel Online** | ⚠️ Funciona, com possíveis limitações em formas e hiperlinks de formas |
| **LibreOffice Calc** | ⚠️ Parcial — listas de outra aba e formas podem não funcionar como no Excel |
| **Google Planilhas** | ❌ Não recomendado — formas/botões e validação estendida não são preservados |
| **Bibliotecas Python (`openpyxl` etc.)** | ⚠️ **Atenção:** ao abrir e salvar o arquivo com `openpyxl`, as **formas (botões de navegação) são perdidas** e a **validação de dados estendida (x14) é removida** |

> ⚠️ Se você pretende automatizar ou editar o arquivo por script, trabalhe sobre uma cópia e valide o resultado no Excel.

---

## ▶️ Como usar

1. **Baixe** o arquivo `organizador_IR.xlsx` e abra no Excel.
2. Em **1. TITULAR**, preencha todos os campos amarelos. Use as listas suspensas para `SIM/NÃO`.
3. Clique em **PRÓXIMO =>** para ir a **2. INFORMES**:
   - Escolha o banco na lista suspensa;
   - Informe o valor atual;
   - Registre o nome do arquivo anexo (PDF do informe).
4. Clique em **PRÓXIMO =>** para ir a **3. NOTAS**:
   - Registre cada entrada com **data**, **categoria** (`Holerite`, `CNPJ` ou `Freelancer`) e **valor**.
5. Retorne a **INFORMES** para conferir o **TOTAL** calculado.
6. Salve o arquivo e utilize os dados para preencher a declaração no programa da Receita Federal.

---


## 📝 Créditos

**LION APP SYSTEM** by **DYADIX** 💜

> Este material é apenas uma ferramenta de organização pessoal e **não substitui** orientação contábil ou a conferência dos dados no programa oficial da Receita Federal.
