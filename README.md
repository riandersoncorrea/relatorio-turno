# Relatório de Turno — CCM

Aplicação client-side (HTML + CSS + JavaScript puro) para geração do **Relatório de
Turno do Centro de Controle de Manutenção (CCM)**. Não possui backend, não usa
Power Automate, não usa banco de dados e não depende de Node.js em produção —
todo o processamento (leitura do Excel, cálculo de indicadores, geração do
gráfico e exportação em PNG/PDF) acontece no navegador do usuário.

Identidade visual construída a partir das cores da logo da **Gerência de
Serviços Operacionais** (teal institucional + acentos dourado, azul, magenta e
verde extraídos dos ícones da logo).

---

## 1. Estrutura de arquivos

```
relatorio-ccm/
├── index.html          # Página única da aplicação
├── css/
│   └── style.css        # Toda a estilização (identidade visual + folha do relatório)
├── js/
│   └── script.js         # Toda a lógica: leitura do Excel, KPIs, gráfico, paginação, exportação
├── assets/
│   ├── logo-gerencia.png # Logo da Gerência de Serviços Operacionais
│   └── logo-pcm.png      # Logo do PCM
├── lib/                  # Bibliotecas de terceiros, baixadas localmente (sem CDN em produção)
│   ├── xlsx.full.min.js       # SheetJS — leitura de .xlsx/.xls no navegador
│   ├── chart.umd.min.js       # Chart.js — gráficos de chamados por disciplina
│   ├── html2canvas.min.js     # Captura do relatório em imagem
│   └── jspdf.umd.min.js       # Geração do PDF a partir da imagem capturada
└── README.md
```

Nenhuma dessas bibliotecas é carregada via CDN em tempo de execução — todas
ficam em `lib/` e sobem junto com o restante dos arquivos para o SharePoint,
evitando problemas de bloqueio de CDN, CORS ou indisponibilidade externa.

---

## 2. Como usar

1. Abra `index.html` no navegador (ou acesse a página publicada no SharePoint).
2. Preencha **Data**, **Turno** e, opcionalmente, a **Observação do Turno**.
3. Clique em **Selecionar Arquivo Excel** e escolha um arquivo `.xlsx` ou `.xls`.
   O campo **Responsável** (somente leitura) é preenchido automaticamente a
   partir da coluna "Criado por".
4. Clique em **Gerar Relatório**. O sistema irá:
   - Ler a primeira planilha do arquivo;
   - Identificar as colunas oficiais do layout "Rel turno.xlsx";
   - Calcular chamados recebidos, concluídos e pendentes ("Status usuário");
   - Classificar a localidade (EFC / São Luís) e a disciplina de cada chamado
     e montar os dois gráficos "Chamados por Disciplina";
   - Montar a tabela de chamados pendentes, paginando automaticamente em
     folhas adicionais de 816×1056 quando necessário.
5. Revise a pré-visualização do relatório.
6. Clique em **Gerar Imagem (PNG)** para baixar o relatório em alta resolução
   (folhas paginadas são empilhadas em uma única imagem final).
7. Clique em **Gerar PDF** para baixar o mesmo conteúdo em PDF (uma página
   por folha do relatório).
8. Clique em **Novo Relatório** para voltar ao formulário (os campos
   preenchidos são mantidos, prontos para gerar um novo relatório com outro
   Excel).

Nenhuma etapa recarrega a página.

---

## 3. Estrutura esperada do Excel

A primeira planilha do arquivo deve seguir o layout "Rel turno.xlsx", com
todas as colunas abaixo (cabeçalhos comparados sem diferenciar maiúsculas,
acentos, espaços e pontuação). Se alguma faltar, o erro informa qual.

| Coluna           | Uso                                                        |
|------------------|------------------------------------------------------------|
| Ordem            | Coluna ORDEM da tabela de pendências                       |
| texto breve      | Coluna TEXTO BREVE da tabela de pendências                 |
| CenTrab respon.  | Localidade: começa com `RG` → EFC; qualquer outro → São Luís |
| Status usuário   | `CONCLUÍDO` ou `PENDENTE` → KPIs e tabela de pendências     |
| Criado por       | Chave do responsável pelo relatório                        |
| Grp.plnj.PM      | Disciplina do chamado                                      |

**Disciplinas:** MCR → Refrigeração · MCI → Cívil · MHI → Hidráulica ·
MEL → Elétrica · CPV → Pragas e Vetores.

**Responsáveis:** C0731117 → Maria Eduarda · C0731491 → Lívia Cunha ·
C0730918 → Antônio Ribeiro · C0711275 → Weslly Braga.

A geração é bloqueada (com mensagem indicando valor e linha) quando o arquivo
está vazio, falta coluna obrigatória, há status diferente de
CONCLUÍDO/PENDENTE, código de Grp.plnj.PM não mapeado, chave de "Criado por"
desconhecida ou mais de um responsável no mesmo arquivo.

---

## 4. Configuração (sem precisar mexer na lógica)

Todos os pontos de customização ficam centralizados no objeto `CONFIG`, no
topo do arquivo `js/script.js`:

- **`COLUMN_ALIASES`** — nomes (normalizados) aceitos para cada coluna do Excel.
- **`DISCIPLINAS`** — mapeamento código de Grp.plnj.PM → disciplina.
- **`RESPONSAVEIS`** — mapeamento chave de "Criado por" → responsável.
- **`LOCALIDADE_EFC_PREFIX`** — prefixo de "CenTrab respon." que indica EFC.
- **`CHART_PALETTE`** — cores das disciplinas nos gráficos (mesma cor por
  disciplina nos dois gráficos).
- **`KPI_COLORS`** — cores dos três indicadores (recebidos/concluídos/pendentes).
- **`COMPACT_THRESHOLD`** — a partir de quantos chamados pendentes a tabela
  passa a usar fonte/espaçamento reduzidos.
- **`MOCK_DATA_ENABLED`** — **mantenha `false` em produção**. Sirva apenas
  para testar o layout do relatório sem precisar de um Excel real durante o
  desenvolvimento (função `gerarDadosMockParaTeste`, claramente isolada no
  final do arquivo e fácil de remover).

---

## 5. O relatório como imagem (816×1056)

Cada folha do relatório é renderizada com dimensão fixa de **816×1056px**
(proporção de página corporativa em retrato). Quando a tabela de chamados
pendentes não cabe em uma única folha, o sistema **pagina automaticamente**:
mede o espaço realmente disponível na tela (não corta conteúdo) e cria folhas
adicionais de continuação, cada uma ainda com 816×1056px, cabeçalho e
numeração de página.

Na exportação:

- **Imagem (PNG)** — cada folha é capturada via `html2canvas` com escala 3x
  (alta resolução). Se houver mais de uma folha, elas são empilhadas
  verticalmente em **uma única imagem final**, conforme solicitado.
- **PDF** — as mesmas capturas em alta resolução são inseridas em um PDF via
  `jsPDF`, uma página por folha.

O gráfico (Chart.js) é renderizado com `animation: false` e
`devicePixelRatio` elevado, e a captura só é disparada depois que a página
confirma fontes carregadas (`document.fonts.ready`) e o layout estabilizado —
garantindo que o gráfico e a tabela já estejam completamente desenhados antes
da exportação.

---

## 6. Publicação no SharePoint

### 6.1. Opção simples — Biblioteca de Documentos + Web Part "Incorporar" (Embed)

Esta é a forma recomendada para a maioria dos sites modernos do SharePoint,
pois **não exige habilitar scripts customizados** no site:

1. Faça upload de toda a pasta `relatorio-ccm/` (mantendo a estrutura de
   subpastas) para uma Biblioteca de Documentos do site.
2. Abra `index.html` na biblioteca, copie sua URL direta (ex.:
   `.../Documentos Compartilhados/relatorio-ccm/index.html`).
3. Em uma página moderna do SharePoint, adicione a Web Part **Incorporar**
   (Embed) e cole a URL do `index.html`.
4. A página roda dentro de um `<iframe>` isolado, com todo o HTML/CSS/JS
   funcionando normalmente — inclusive a leitura do Excel e a exportação de
   imagem/PDF, que acontecem no navegador do usuário, dentro do iframe.

### 6.2. Opção alternativa — páginas clássicas com scripts customizados habilitados

Se o site for clássico (ou moderno com scripts customizados habilitados via
`Set-SPOSite -DenyAddAndCustomizePages 0`), é possível usar a Web Part
**Editor de Script/Conteúdo** apontando para o mesmo `index.html`, ou
incorporar o conteúdo diretamente. Essa opção é desencorajada pela Microsoft
para sites modernos por motivos de segurança — prefira a opção 6.1.

### ⚠️ Limitação importante do SharePoint moderno

Páginas modernas do SharePoint **bloqueiam a execução de `<script>` e HTML
customizado inserido diretamente na página** quando "permitir scripts
customizados" está desabilitado no site/tenant (configuração padrão em
muitos ambientes corporativos, incluindo Microsoft 365 mais recentes). Por
isso a Web Part **Incorporar (Embed)** é a abordagem recomendada: ela carrega
a aplicação em um `iframe` apontando para o arquivo hospedado na própria
biblioteca de documentos, contornando essa restrição sem precisar de
permissões elevadas.

### 6.3. Adaptação para SPFx (abordagem 100% suportada pela Microsoft)

Se a organização exigir uma solução totalmente integrada (sem iframe) e com
suporte oficial da Microsoft, a aplicação pode ser portada para um **SPFx Web
Part**. Principais mudanças necessárias:

1. Criar o projeto com o Yeoman generator (`@microsoft/generator-sharepoint`),
   escolhendo um Web Part do tipo "No framework" ou "React".
2. Mover `css/style.css` para os estilos do componente (SCSS module do SPFx).
3. Portar a lógica de `js/script.js` para o método `render()` (ou para um
   componente React), mantendo as mesmas funções — a lógica de negócio
   (leitura do Excel, cálculo de KPIs, paginação, exportação) não precisa
   mudar, apenas como/quando o DOM é montado.
4. Trocar as bibliotecas de `lib/*.js` (carregadas via `<script>`) por
   dependências npm equivalentes (`xlsx`, `chart.js`, `html2canvas`, `jspdf`),
   importadas via `import` e empacotadas pelo webpack do SPFx.
5. Mover `assets/logo-gerencia.png` e `assets/logo-pcm.png` para a pasta
   `assets/` do Web Part (referenciadas via `require()`/import de imagem).
6. Empacotar (`gulp bundle --ship` / `gulp package-solution --ship`) e
   publicar o `.sppkg` no App Catalog do tenant.

Essa é uma tarefa de portabilidade de empacotamento, não de reescrita: toda a
lógica de leitura de Excel, cálculo de indicadores, montagem do relatório e
exportação em imagem/PDF permanece a mesma.

---

## 7. Compatibilidade

- Testado em navegadores modernos baseados em Chromium (Edge/Chrome), que são
  o padrão em ambientes Microsoft 365/SharePoint.
- Não requer Node.js, servidor próprio ou build step em produção — basta
  abrir/hospedar os arquivos estáticos.
- Todas as bibliotecas de terceiros estão vendorizadas em `lib/`, sem
  dependência de CDN externo em tempo de execução.
