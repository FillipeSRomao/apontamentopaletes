# Sistema PCP – Apontamento de Produção e Etiqueta de Palete

Aplicação web (Flask) que roda no computador da produção. O operador abre a tela no navegador,
escolhe o produto e a quantidade e o sistema:

1. **registra o palete** num banco SQLite;
2. **gera a etiqueta em PDF**, no layout do `ROMANEIO_DE_PALETE.xlsx`;
3. **abre o PDF numa nova aba do navegador**, de onde a etiqueta é impressa (Ctrl+P).

O PCP consulta um **relatório** com a contagem do dia separada em **COM SELO** e **SEM SELO** e
exporta tudo para Excel.

> Versão atual: **5**. Instalação passo a passo para Windows 10: veja `GUIA_INSTALACAO_PCP.pdf`.

---

## Funcionalidades

- Tela do operador com um cartão por produto, agrupados por linha de produção.
- Botão **SEM SELO / COM SELO** em cada produto que usa selo. Produtos que nunca levam selo mostram
  "Produto sem selo" e a etiqueta sai com **NÃO SE APLICA**.
- Etiqueta A4 paisagem no layout do romaneio (logo, código, descrição, quantidade, fabricação,
  validade, lote, pesos e número da linha).
- **COM SELO** é reconhecível de longe e gasta pouco toner: moldura grossa em volta da etiqueta e
  quadro do selo com listras (sem fundo preto).
- **Lote** calculado pela mesma regra da planilha do romaneio.
- Validade e pesos calculados a partir da aba **Base** do `ROMANEIO_DE_PALETE.xlsx`.
- **Reimpressão** idêntica à original (aviso REIMPRESSÃO; não soma no relatório).
- **Relatório do PCP** (`/pcp`): totais do dia por selo, tabela por produto e últimos registros.
- **Exportação para Excel** com as abas *Apontamento* (todos os registros) e *Resumo por selo*.
- Atualização automática do banco de dados de versões anteriores.

## Estrutura da pasta

```text
C:\PCP\
├── apontamento_producao_pcp.py   # o sistema (arquivo único)
├── ROMANEIO_DE_PALETE.xlsx       # cadastro dos produtos (aba Base) – o seu arquivo
├── logo_oasis.png                # logo impresso na etiqueta
├── requirements.txt              # bibliotecas Python
├── instalar_pcp.bat              # instalação (uma vez)
├── iniciar_pcp.bat               # liga o sistema
├── GUIA_INSTALACAO_PCP.pdf       # guia completo
├── README.md                     # este arquivo
├── pcp_producao.db               # banco de dados (criado sozinho)
├── etiquetas\                    # PDFs gerados (criada sozinha)
└── venv\                         # ambiente Python (criado pelo instalar_pcp.bat)
```

Não coloque o sistema em Área de Trabalho ou Documentos: se o OneDrive estiver ativo, ele pode
corromper o banco de dados.

## Requisitos

- Windows 10 (64 bits) com **Python 3.13** (o sistema foi testado também com Python 3.12).
- Bibliotecas: `flask`, `reportlab`, `openpyxl` (instaladas pelo `instalar_pcp.bat`).
- Navegador **Microsoft Edge** ou **Google Chrome**.
- Impressora com papel A4.
- Internet apenas na instalação (existe procedimento offline no guia).

## Instalação rápida

1. Instale o Python 3.13 (64 bits) marcando **Add python.exe to PATH**.
2. Extraia o pacote em `C:\PCP` e copie para lá o seu `ROMANEIO_DE_PALETE.xlsx`.
3. Dê dois cliques em `instalar_pcp.bat` (aguarde "Bibliotecas OK").
4. Dê dois cliques em `iniciar_pcp.bat`. O navegador abre em `http://localhost:5000`.
5. Clique em **Gerar etiqueta de teste**, imprima com Ctrl+P e confira o papel.

Detalhes, firewall, início automático com o Windows e solução de problemas: `GUIA_INSTALACAO_PCP.pdf`.

## Uso diário

1. Confira o **turno** no seletor do topo (o número do turno entra no lote).
2. No cartão do produto, confira a **quantidade** (padrão da linha; ajuste só em ponta de lote).
3. Nos produtos com selo, marque **SEM SELO** ou **COM SELO**.
4. Clique em **Imprimir Novo Palete**. O PDF abre em outra aba: **Ctrl+P**, imprima e feche a aba.
5. **Reimprimir Última** abre de novo a última etiqueta do produto. No relatório, o link
   **Abrir PDF** mostra a etiqueta de qualquer registro (PDFs guardados por 30 dias).

> O palete é contado quando a etiqueta é **gerada**, não quando a impressora termina. Se a aba for
> fechada sem imprimir, use *Reimprimir Última* ou *Abrir PDF*. Atualizar (F5) a aba do PDF não cria
> palete novo.

## Regras de negócio

### Lote

Mesma regra da planilha do romaneio:

```text
DDD + AA + linha + código + turno + " - " + H:MM
```

| Parte | Origem |
|---|---|
| `DDD` | dia do ano (001 a 366) do **dia de produção** |
| `AA` | ano com 2 dígitos |
| `linha` | número da linha sem zero à esquerda (06 vira 6) |
| `código` | coluna CÓD da aba Base, com no mínimo 3 dígitos (3 vira 003) |
| `turno` | número do turno (Turno 2 → 2, Turno 3 → 3) |
| `H:MM` | hora e minuto da impressão (hora sem zero à esquerda) |

Exemplo: `31926646662 - 10:58` = dia 319, ano 26, linha 6, código 4666, turno 2, impresso às 10:58.

O lote completo é gravado no banco e repetido nas reimpressões. Na tela do operador aparece o lote
sem a hora; o último dígito muda junto com o seletor de turno.

### Dia de produção

Paletes do Turno 3 feitos depois da meia-noite contam no dia em que o turno começou. Isso vale para
o lote, para a FABRICAÇÃO e para o relatório. Os horários vêm do bloco `TURNOS` do script.

### Validade e pesos

- **Validade** = dia de produção + dias de validade da aba Base (a coluna VALIDADE traz a fórmula
  `=$E$1+365`; vale o número somado).
- **Peso líquido** = peso do palete cheio (aba Base) proporcional à quantidade impressa.
- **Peso bruto** = peso líquido + tara do palete (bruto − líquido da aba Base).

### Selo

- Produtos que usam selo: **COM SELO** ou **SEM SELO**, escolhido no cartão (padrão: SEM SELO).
- Produtos com `usa_selo=False`: o quadro mostra **NÃO SE APLICA**.
- A escolha de selo fica marcada só no cartão usado.

### Número do palete

Número sequencial (id do registro original). Aparece no rodapé da etiqueta e é o mesmo nas
reimpressões.

### Outros campos

RESERVADO, Nº PEDIDO e CLIENTE saem sempre `N/D`.

## Produtos cadastrados

| Linha | Produto | CÓD | Qtd. padrão | Selo |
|---|---|---|---|---|
| 02 | Copo 200mL | 5 | 80 | não usa |
| 03 | Garrafa 5L | 3 | 126 | sim |
| 05 | Pacote 1,5L Com Gás | 4626 | 84 | sim |
| 05 | Pacote 1,5L Sem Gás | 4625 | 84 | sim |
| 05 | Pacote Saborizada Limão | 4671 | 126 | não usa |
| 05 | Pacote Saborizada Tangerina | 4673 | 126 | não usa |
| 05 | Pacote Saborizada Frutas Vermelhas | 4670 | 126 | não usa |
| 05 | Pacote Saborizada Maçã Verde | 4672 | 126 | não usa |
| 05 | Pacote Saborizada Abacaxi com Hortelã | 4422 | 126 | não usa |
| 05 | Pacote Saborizada Uva | 4674 | 126 | não usa |
| 06 | Pacote 500mL Com Gás | 4668 | 140 | sim |
| 06 | Pacote 500mL Sem Gás | 4666 | 140 | sim |
| 06 | Pacote 355mL Com Gás | 4669 | 160 | sim |
| 06 | Pacote 355mL Sem Gás | 4667 | 160 | sim |

## Manutenção

Toda configuração fica na seção **`>>> ÁREA DE CONFIGURAÇÃO`** do `apontamento_producao_pcp.py`.
Salve o arquivo em **UTF-8** (para não estragar os acentos) e reinicie o `iniciar_pcp.bat` depois de
qualquer alteração.

### Produto novo

1. **Aba Base** do `ROMANEIO_DE_PALETE.xlsx` (aba oculta: botão direito em uma aba → *Reexibir*): preencha
   uma linha nova da tabela com CÓD, DESCRIÇÃO, DESCRIÇÃO - 01, DESCRIÇÃO - 02, VALIDADE
   (`=$E$1+365`), QNT e os pesos do palete cheio. Salve.
2. **Script**: acrescente uma linha ao `CATALOGO`:

   ```python
   produto("06", "Pacote 330mL Sem Gás", "Pacotes", 100, cod="4700"),
   ```

   Os valores são: número da linha, nome na tela, unidade, quantidade padrão e o CÓD da Base.
   Se o produto nunca leva selo, acrescente `usa_selo=False`. Opcional: `qtd_max=` (a tela aceita
   até 3× a quantidade padrão).

Se o CÓD não existir na aba Base, o programa não inicia e informa qual produto está errado.

### Retirar produto

Apague a linha do `CATALOGO`. O histórico continua no banco e no Excel exportado.

### Outras alterações

| O que mudar | Onde |
|---|---|
| Validade, pesos, textos da etiqueta | linha do produto na aba Base |
| Quantidade padrão | 4º valor da linha no `CATALOGO` |
| Horário dos turnos | bloco `TURNOS` (o nome **precisa** ter o número, ex.: `"Turno 2"`) |
| Título da faixa do topo da etiqueta | `TITULO_ETIQUETA` |
| Texto do selo nos produtos que não usam | `TEXTO_SELO_NAO_APLICA` |
| Cartões por linha na tela | `COLUNAS_POR_LINHA` |
| Linhas do painel do PCP | `LIMITE_PAINEL` |
| Dias que os PDFs são guardados | `DIAS_GUARDAR_PDF` (padrão 30) |

### Atualizar o sistema

Feche a janela preta, guarde uma cópia do `.py` antigo, troque pelo novo e abra o `iniciar_pcp.bat`.
Não precisa reinstalar. O banco é atualizado automaticamente.

### Backup

Copie com frequência `pcp_producao.db`, `ROMANEIO_DE_PALETE.xlsx` e `apontamento_producao_pcp.py`.
A pasta `etiquetas` não precisa de backup.

## Variáveis de ambiente (opcionais)

| Variável | Para que serve | Padrão |
|---|---|---|
| `PCP_DB` | caminho do banco SQLite | `pcp_producao.db` na pasta do sistema |
| `PCP_ROMANEIO` | caminho do `ROMANEIO_DE_PALETE.xlsx` | mesma pasta do sistema |
| `PCP_HOST` | endereço de escuta | `0.0.0.0` |
| `PCP_PORT` | porta | `5000` |
| `PCP_SECRET_KEY` | chave de sessão do Flask (mensagens da tela) | valor fixo de exemplo |

## Rotas

| Método | Rota | Função |
|---|---|---|
| GET | `/` | tela do operador |
| POST | `/imprimir` | registra o palete, gera o PDF e redireciona para ele |
| POST | `/reimprimir/<produto>` | gera a reimpressão da última etiqueta do produto |
| POST | `/teste-impressao` | gera uma etiqueta de teste (não grava no banco) |
| GET | `/etiqueta/<id>` | abre o PDF de um registro |
| GET | `/pcp` | relatório do PCP |
| GET | `/exportar-excel` | baixa o relatório completo em Excel |

## Banco de dados

SQLite, tabela `producao`:

| Coluna | Conteúdo |
|---|---|
| `id` | chave do registro |
| `data_hora` | data e hora da impressão (`dd/mm/aaaa hh:mm:ss`) |
| `dia_producao` | dia de produção (`dd/mm/aaaa`) |
| `turno`, `linha`, `sku` | turno, linha e nome do produto |
| `cod` | CÓD da aba Base |
| `lote` | lote completo (com a hora) |
| `quantidade` | unidades no palete |
| `palete` | nº do palete (mesmo nas reimpressões) |
| `validade` | data de validade impressa |
| `peso_liquido`, `peso_bruto` | pesos impressos (kg) |
| `selo` | `COM SELO`, `SEM SELO` ou vazio (produto sem selo) |
| `status` | `Em Produção`, `Reimpressão`, `Pendente` ou `Erro de Impressão` |

<<<<<<< HEAD
Somente registros com status **Em Produção** entram nas somas do relatório. Reimpressões e falhas
ao gerar a etiqueta não somam.

Bancos criados por versões anteriores são migrados na primeira execução. As colunas antigas `gtin`
e `sscc` continuam nesses bancos apenas como histórico, e cada SSCC antigo vira o número do palete
do primeiro registro dele.

## O que o sistema não faz (por enquanto)

- **GTIN, códigos de barras GS1-128 e SSCC:** foram retirados. A etiqueta não tem código de barras.
  A versão com GS1 está em pacotes anteriores e pode ser reativada quando a empresa implementar.
- **Impressão silenciosa:** foi retirada. A impressão é feita pelo navegador a partir do PDF.
- **Login:** não há tela de login; qualquer computador da rede com o endereço acessa o sistema e os
  relatórios. Use somente em rede interna e fechada (o tráfego é HTTP, sem criptografia).
- **RESERVADO / Nº PEDIDO / CLIENTE:** saem fixos como `N/D`.

## Histórico de versões

| Versão | Principais mudanças |
|---|---|
| 3 | Etiqueta própria com GS1-128 (GTIN, lote, quantidade e SSCC), impressão pelo SumatraPDF, dia de produção, painel e Excel |
| 4 | Etiqueta no layout do romaneio, leitura da aba Base do Excel, validade e pesos, botão de selo por produto e visual COM SELO econômico |
| 5 | Sem GTIN/GS1/SSCC (nº sequencial de palete), lote no padrão da planilha (com o número do turno), relatório com contagem por selo e aba *Resumo por selo* no Excel, PDF aberto no navegador para imprimir |

## Problemas comuns

| Sintoma | O que fazer |
|---|---|
| "Planilha do romaneio não encontrada" | Copie o `ROMANEIO_DE_PALETE.xlsx` para a pasta do sistema, com esse nome. |
| "o código X não existe na aba Base" | Cadastre o CÓD na aba Base (ou corrija o `cod=` no `CATALOGO`) e reinicie. |
| "O nome do turno precisa ter o número" | No bloco `TURNOS`, use nomes como `"Turno 2"`. |
| Navegador não abre o sistema | O `iniciar_pcp.bat` não está rodando. |
| O PDF é baixado em vez de abrir | Abra o arquivo baixado e imprima dele, ou ajuste "Documentos PDF" nas configurações do navegador. |
| Etiqueta cortada ou girada | Na janela de impressão: Layout *Paisagem* e Escala *Ajustar à área de impressão*. |
| Lote ou turno diferente do esperado | Confira data, hora e fuso do Windows e o turno marcado no topo. |
| Sistema parou de responder | Um clique dentro da janela preta pausa o programa: clique nela e aperte Enter. |

Mais detalhes no `GUIA_INSTALACAO_PCP.pdf`.
=======
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
>>>>>>> 046faee555a9d9a21b5f79a9d9592824e99fb7d7
