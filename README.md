# 🏭 Apontamento Automático de Produção - PCP (v4)

Sistema web leve e direto ao ponto desenvolvido em **Python (Flask)** para o apontamento de produção na linha fabril, geração de etiquetas de palete no layout de romaneio com códigos de barras **GS1-128** (GTIN, Lote, Quantidade e SSCC), controle de selos por produto e exportação de relatórios para Excel.

---

## ✨ Funcionalidades Principais

* **Interface Web para Operadores:** Cartões organizados por linha de produção com seleção rápida de turno, quantidade (com suporte a ponta de lote) e opção de selo (`SEM SELO` / `COM SELO`).
* **Etiqueta Econômica & Layout Romaneio:** Gera PDF em formato A4 paisagem replicando fielmente o layout da planilha `ROMANEIO_DE_PALETE.xlsx` (com destaque visual hachurado para produtos com selo, economizando toner).
* **Padrão Logístico GS1-128:** Geração vetorial de códigos de barras contendo Application Identifiers `(01)` GTIN-14, `(10)` Lote, `(37)` Quantidade e `(00)` SSCC sequencial único com cálculo automático de dígito verificador (Módulo 10).
* **Integração com Planilha Base:** Lê automaticamente descrições, validade e pesos (líquido e bruto proporcional à quantidade do palete) diretamente da aba `Base` do Excel, sem alterá-la.
* **Turno Noturno Inteligente:** Paletes apontados após a meia-noite são contabilizados automaticamente no **dia de produção** em que o turno iniciou.
* **Impressão Direta e Silenciosa:** Suporte nativo ao **SumatraPDF** no Windows para envio imediato à impressora sem abrir janelas de visualização.
* **Painel PCP & Exportação:** Acompanhamento em tempo real do resumo diário, histórico de reimpressões/falhas e download completo da base em `.xlsx`.

---

## 📋 Pré-requisitos

* **Python 3.10+**
* **Bibliotecas Python:**
  * `Flask`
  * `openpyxl`
  * `reportlab`
* **Impressão Silenciosa (Opcional, recomendado para Windows):**
  * [SumatraPDF](https://www.sumatrapdfreader.org/) instalado ou o executável `SumatraPDF.exe` na mesma pasta do script.

---

## 🚀 Instalação e Execução

1. **Clone o repositório:**
   ```bash
   git clone https://github.com/seu-usuario/seu-repositorio.git
   cd seu-repositorio
   ```

2. **Crie um ambiente virtual e instale as dependências:**
   ```bash
   python -m venv venv
   # No Windows:
   venv\Scripts\activate
   # No Linux/macOS:
   source venv/bin/activate

   pip install flask openpyxl reportlab
   ```

3. **Arquivos obrigatórios na pasta raiz:**
   Certifique-se de que os seguintes arquivos estejam no mesmo diretório do script principal:
   * `ROMANEIO_DE_PALETE.xlsx` (contendo a aba `Base` com as colunas `CÓD`, `DESCRIÇÃO - 01`, `DESCRIÇÃO - 02`, `VALIDADE`, `QNT`, `PESO LIQUIDO DOS PRODUTOS` e `PESO BRUTO - PALLET`).
   * `logo_oasis.png` *(opcional — caso não exista, o sistema imprime o texto "OASIS" no cabeçalho)*.

4. **Inicie o servidor:**
   ```bash
   python apontamento_selo.py
   ```
   O sistema estará acessível na rede local através de `http://localhost:5000` (ou pelo IP da máquina na porta `5000`).

---

## ⚙️ Configuração do Dia a Dia

Todas as regras de negócio foram centralizadas no bloco **`>>> ÁREA DE CONFIGURAÇÃO`** dentro do código-fonte para facilitar a manutenção sem precisar alterar a lógica do sistema:

* **`TURNOS`:** Definição de nome, horário de início e fim de cada turno (turnos que viram a meia-noite são tratados automaticamente).
* **`CATALOGO`:** Cadastro dos produtos por linha, unidade de medida, quantidade padrão, código de vínculo (`cod`) com a aba `Base`, `gtin` real e flag `usa_selo=False` para itens que não utilizam selo.
* **`CORES_LINHA`:** Identidade visual (cores hexadecimais) para diferenciar cada linha de produção na tela do operador.
* **Layout e Etiqueta:** Ajuste de colunas por linha, textos padrão (`TITULO_ETIQUETA`, `SELOS`) e ativação/desativação dos códigos GS1-128 no rodapé (`ETIQUETA_COM_CODIGO_DE_BARRAS`).

---

## 🔧 Variáveis de Ambiente (Opcionais)

Você pode personalizar o comportamento do servidor em produção definindo as seguintes variáveis de ambiente:

| Variável | Descrição | Padrão |
| :--- | :--- | :--- |
| `PCP_DB` | Caminho para o arquivo do banco de dados SQLite. | `./pcp_producao.db` |
| `PCP_ROMANEIO` | Caminho completo para a planilha `ROMANEIO_DE_PALETE.xlsx`. | `./ROMANEIO_DE_PALETE.xlsx` |
| `PCP_SECRET_KEY` | Chave secreta para assinatura de sessões do Flask. | `troque-esta-chave-em-producao` |
| `PCP_HOST` | Endereço IP de escuta do servidor web. | `0.0.0.0` |
| `PCP_PORT` | Porta de escuta do servidor web. | `5000` |
| `PRINTER_NAME` | Nome **exato** da impressora instalada no Windows. | *Impressora padrão do Windows* |
| `SUMATRA_PATH` | Caminho completo para o executável `SumatraPDF.exe`. | *Busca automática no sistema* |
| `PRINT_SETTINGS` | Parâmetros extras repassados ao SumatraPDF (ex.: `fit`, `noscale`). | *Nenhum* |

---

## 🗂️ Estrutura de Arquivos Gerados

* `pcp_producao.db`: Banco de dados SQLite (em modo WAL) contendo todo o histórico de apontamentos, reimpressões e status.
* `etiquetas/`: Pasta criada automaticamente onde os PDFs gerados são armazenados. Arquivos com mais de **30 dias** (`DIAS_GUARDAR_PDF`) são limpos automaticamente na inicialização do sistema.

---

## 🖥️ Telas do Sistema

1. **Painel do Operador (`/`):**
   * Seleção global de turno.
   * Cartões por produto com cálculo dinâmico do lote do dia (`AAMMDD-LXX`).
   * Botão **Imprimir Novo Palete** (gera novo SSCC e contabiliza na produção) e **Reimprimir Última** (reimprime o último palete válido daquele SKU com marcação de *REIMPRESSÃO* sem duplicar a contagem).
   * Botão de **Página de Teste** no rodapé para validar a comunicação com a impressora sem sujar o banco de dados.
2. **Painel PCP (`/pcp`):**
   * Tabela consolidada com total de paletes e unidades produzidas no dia atual.
   * Tabela de resumo diário com separação clara entre paletes **com selo** e **sem selo** por produto e linha.
   * Histórico detalhado dos últimos 200 registros com badges de status.
   * Botão **Baixar Relatório Completo em Excel (`/exportar-excel`)** gerando a planilha em memória.
