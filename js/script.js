/**
 * Relatório de Turno — CCM (Centro de Controle de Manutenção)
 * Aplicação 100% client-side: leitura de Excel, cálculo de indicadores,
 * geração de gráfico, montagem do relatório (folha 816x1056) e exportação
 * em PNG / PDF. Sem backend, sem Power Automate, sem banco de dados.
 *
 * Compatível com hospedagem em biblioteca de documentos do SharePoint
 * (arquivos estáticos HTML/CSS/JS + bibliotecas locais em /lib).
 */
(function () {
  "use strict";

  // ==========================================================================
  // CONFIGURAÇÃO — ajuste aqui nomes de colunas, status considerados
  // "concluído" e demais parâmetros, sem precisar mexer na lógica abaixo.
  // ==========================================================================
  const CONFIG = {
    SHEET_WIDTH: 816,
    SHEET_HEIGHT: 1056,
    EXPORT_SCALE: 3, // resolução da imagem exportada (3x = alta qualidade)

    // Colunas oficiais do layout "Rel turno.xlsx". Os cabeçalhos são
    // comparados normalizados (minúsculas, sem acento, sem espaço/pontuação).
    COLUMN_ALIASES: {
      ordem: ["ordem"],
      textoBreve: ["textobreve"],
      centrab: ["centrabrespon"],
      status: ["statususuario"],
      criadoPor: ["criadopor"],
      grpPlnj: ["grpplnjpm"],
    },

    // Rótulos exibidos nas mensagens de validação (nome exato da coluna no Excel).
    COLUMN_LABELS: {
      ordem: "Ordem",
      textoBreve: "texto breve",
      centrab: "CenTrab respon.",
      status: "Status usuário",
      criadoPor: "Criado por",
      grpPlnj: "Grp.plnj.PM",
    },

    REQUIRED_FIELDS: [
      "ordem",
      "textoBreve",
      "centrab",
      "status",
      "criadoPor",
      "grpPlnj",
    ],

    // Valores aceitos em "Status usuário" (já normalizados: minúsculo, sem acento).
    STATUS_CONCLUIDO: "concluido",
    STATUS_PENDENTE: "pendente",

    // "CenTrab respon." iniciando com este prefixo → EFC; demais → São Luís.
    LOCALIDADE_EFC_PREFIX: "RG",
    LOCALIDADES: ["EFC", "São Luís"],

    // Código de "Grp.plnj.PM" → disciplina exibida nos gráficos (ordem de exibição).
    DISCIPLINAS: {
      MCR: "Refrigeração",
      MCI: "Cívil",
      MHI: "Hidráulica",
      MEL: "Elétrica",
      CPV: "Pragas e Vetores",
    },

    // Chave de "Criado por" → responsável pelo relatório.
    RESPONSAVEIS: {
      C0731117: "Maria Eduarda",
      C0731491: "Lívia Cunha",
      C0730918: "Antônio Ribeiro",
      C0711275: "Weslly Braga",
    },

    // Máximo de exemplos listados em cada mensagem de validação.
    MAX_VALIDATION_EXAMPLES: 5,

    // Paleta institucional extraída da logo da Gerência de Serviços Operacionais.
    CHART_PALETTE: [
      "#0f6f63",
      "#e8a431",
      "#3b8fce",
      "#b8316b",
      "#3fa671",
      "#0a5348",
      "#c9832a",
      "#2c6489",
    ],

    KPI_COLORS: {
      recebidos: "#3b8fce",
      concluidos: "#3fa671",
      pendentes: "#b8316b",
    },

    // Acima deste número de pendências, a tabela passa a usar fonte/espaçamento reduzidos.
    COMPACT_THRESHOLD: 12,

    // Ative apenas durante desenvolvimento para testar o layout sem precisar
    // de um arquivo Excel real. NUNCA deixe ativo em produção.
    MOCK_DATA_ENABLED: false,
  };

  const state = {
    chartInstances: [],
    reportData: null,
    // Resultado do processamento do Excel selecionado (responsável, registros, erros).
    processed: null,
    processedFile: null,
  };

  const dom = {};

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    cacheDom();
    bindEvents();
    setDefaultDate();
  }

  function cacheDom() {
    dom.formPanel = document.getElementById("formPanel");
    dom.reportPanel = document.getElementById("reportPanel");

    dom.inputData = document.getElementById("inputData");
    dom.inputResponsavel = document.getElementById("inputResponsavel");
    dom.inputTurno = document.getElementById("inputTurno");
    dom.inputObservacao = document.getElementById("inputObservacao");
    dom.inputExcel = document.getElementById("inputExcel");

    dom.btnSelectFile = document.getElementById("btnSelectFile");
    dom.fileNameLabel = document.getElementById("fileNameLabel");
    dom.btnGerarRelatorio = document.getElementById("btnGerarRelatorio");

    dom.messages = document.getElementById("messages");

    dom.reportContainer = document.getElementById("reportContainer");
    dom.btnGerarImagem = document.getElementById("btnGerarImagem");
    dom.btnGerarPDF = document.getElementById("btnGerarPDF");
    dom.btnNovoRelatorio = document.getElementById("btnNovoRelatorio");
  }

  function bindEvents() {
    dom.btnSelectFile.addEventListener("click", () => dom.inputExcel.click());

    dom.inputExcel.addEventListener("change", async () => {
      const file = dom.inputExcel.files[0];
      clearMessages();
      state.processed = null;
      state.processedFile = null;
      dom.inputResponsavel.value = "";

      if (file) {
        dom.fileNameLabel.textContent = file.name;
        dom.fileNameLabel.classList.add("file-input__name--set");

        // Identifica o responsável assim que o Excel é carregado.
        if (/\.(xlsx|xls)$/i.test(file.name)) {
          const processed = await getProcessedExcel(file);
          if (processed.errors.length > 0) {
            processed.errors.forEach((msg) => showMessage(msg, "error"));
          }
        }
      } else {
        dom.fileNameLabel.textContent = "Nenhum arquivo selecionado";
        dom.fileNameLabel.classList.remove("file-input__name--set");
      }
    });

    dom.btnGerarRelatorio.addEventListener("click", handleGerarRelatorio);
    dom.btnGerarImagem.addEventListener("click", handleGerarImagem);
    dom.btnGerarPDF.addEventListener("click", handleGerarPDF);
    dom.btnNovoRelatorio.addEventListener("click", resetToForm);
  }

  function setDefaultDate() {
    const today = new Date();
    const iso = today.toISOString().slice(0, 10);
    dom.inputData.value = iso;
  }

  // ==========================================================================
  // 1. LEITURA DO EXCEL (SheetJS)
  // ==========================================================================

  /**
   * Lê um arquivo .xlsx/.xls no navegador e retorna os registros da primeira
   * planilha como um array de objetos (chave = cabeçalho da coluna).
   */
  function readExcelFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (event) => {
        try {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: "array" });

          if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
            reject(new Error("O arquivo Excel não contém nenhuma planilha."));
            return;
          }

          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const rows = XLSX.utils.sheet_to_json(worksheet, {
            defval: "",
            raw: false,
          });
          resolve(rows);
        } catch (err) {
          reject(
            new Error(
              "Não foi possível ler o arquivo Excel. Verifique se o arquivo não está corrompido.",
            ),
          );
        }
      };

      reader.onerror = () =>
        reject(new Error("Erro ao carregar o arquivo selecionado."));
      reader.readAsArrayBuffer(file);
    });
  }

  // ==========================================================================
  // 2. NORMALIZAÇÃO DE TEXTO / COLUNAS / REGISTROS
  // ==========================================================================

  /** Remove acentos, converte para minúsculas e recorta espaços nas bordas. */
  function normalizeText(value) {
    return String(value == null ? "" : value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .replace(/\s+/g, " ")
      .toLowerCase();
  }

  /** Chave de cabeçalho normalizada (sem espaços/pontuação) para comparação com os aliases. */
  function normalizeHeaderKey(value) {
    return normalizeText(value).replace(/[^a-z0-9]/g, "");
  }

  /**
   * Identifica, a partir dos cabeçalhos reais do Excel, qual coluna corresponde
   * a cada campo oficial (ordem, textoBreve, centrab, status, criadoPor, grpPlnj).
   */
  function buildColumnMap(headerKeys) {
    const map = {};

    Object.keys(CONFIG.COLUMN_ALIASES).forEach((field) => {
      const aliases = CONFIG.COLUMN_ALIASES[field];
      const found = headerKeys.find((header) =>
        aliases.includes(normalizeHeaderKey(header)),
      );
      if (found) {
        map[field] = found;
      }
    });

    return map;
  }

  /**
   * Converte as linhas cruas do Excel para o formato usado no relatório e
   * acrescenta as classificações internas (localidade, disciplina, responsável).
   * Os valores originais da planilha são preservados.
   */
  function normalizeRecords(rawRows, columnMap) {
    const cell = (row, field) => String(row[columnMap[field]] || "").trim();

    return (
      rawRows
        .map((row, index) => {
          const record = {
            linha: index + 2, // +2: cabeçalho ocupa a linha 1 do Excel
            ordem: cell(row, "ordem"),
            textoBreve: cell(row, "textoBreve"),
            centrab: cell(row, "centrab"),
            status: cell(row, "status"),
            criadoPor: cell(row, "criadoPor"),
            grpPlnj: cell(row, "grpPlnj"),
          };
          record.localidade = classifyLocalidade(record.centrab);
          record.disciplina = mapDisciplina(record.grpPlnj);
          record.statusKey = normalizeStatus(record.status);
          return record;
        })
        // descarta linhas totalmente vazias (comuns ao final de planilhas exportadas)
        .filter(
          (r) =>
            r.ordem ||
            r.textoBreve ||
            r.centrab ||
            r.status ||
            r.criadoPor ||
            r.grpPlnj,
        )
    );
  }

  // ==========================================================================
  // 3. CLASSIFICAÇÕES — status, localidade, disciplina e responsável
  // ==========================================================================

  /** "CONCLUÍDO" / "PENDENTE" (tolerante a maiúsculas, acentos e espaços) → chave interna. */
  function normalizeStatus(status) {
    const key = normalizeText(status);
    if (key === CONFIG.STATUS_CONCLUIDO) return "concluido";
    if (key === CONFIG.STATUS_PENDENTE) return "pendente";
    return null;
  }

  /** Duas primeiras letras de "CenTrab respon." = RG → EFC; qualquer outra → São Luís. */
  function classifyLocalidade(centrab) {
    const prefix = String(centrab || "").trim().slice(0, 2).toUpperCase();
    return prefix === CONFIG.LOCALIDADE_EFC_PREFIX ? "EFC" : "São Luís";
  }

  /** Código de "Grp.plnj.PM" → nome da disciplina (null se não mapeado). */
  function mapDisciplina(grpPlnj) {
    const code = String(grpPlnj || "").trim().toUpperCase();
    return CONFIG.DISCIPLINAS[code] || null;
  }

  /** Chave de "Criado por" → { chave, nome } (nome null se a chave não for conhecida). */
  function mapResponsavel(criadoPor) {
    const value = String(criadoPor || "").trim().toUpperCase();
    const chave =
      Object.keys(CONFIG.RESPONSAVEIS).find((key) => value.includes(key)) ||
      value;
    return { chave, nome: CONFIG.RESPONSAVEIS[chave] || null };
  }

  // ==========================================================================
  // 4. CÁLCULO DOS KPIs
  // ==========================================================================

  function calculateKPIs(records) {
    const recebidos = records.length;
    const concluidos = records.filter((r) => r.statusKey === "concluido").length;
    const pendentes = records.filter((r) => r.statusKey === "pendente").length;
    return { recebidos, concluidos, pendentes };
  }

  // ==========================================================================
  // 5. AGRUPAMENTO POR LOCALIDADE × DISCIPLINA (para os dois gráficos)
  // ==========================================================================

  /**
   * Retorna, para cada localidade, a contagem de chamados por disciplina.
   * Todas as cinco disciplinas aparecem (com 0 quando não houver chamados),
   * na mesma ordem, para que os dois gráficos tenham a mesma estrutura.
   */
  function groupByLocalidadeDisciplina(records) {
    const disciplinas = Object.values(CONFIG.DISCIPLINAS);
    const result = {};

    CONFIG.LOCALIDADES.forEach((localidade) => {
      const doLocal = records.filter((r) => r.localidade === localidade);
      result[localidade] = disciplinas.map((disciplina) => ({
        disciplina,
        total: doLocal.filter((r) => r.disciplina === disciplina).length,
      }));
    });

    return result;
  }

  // ==========================================================================
  // 6. IDENTIFICAÇÃO DAS PENDÊNCIAS
  // ==========================================================================

  function identifyPendingRecords(records) {
    return records.filter((r) => r.statusKey === "pendente");
  }

  // ==========================================================================
  // 7. VALIDAÇÕES
  // ==========================================================================

  function validateFormValues(values) {
    const errors = [];

    if (!values.data) errors.push("Informe a data do turno.");
    if (!values.turno) errors.push("Selecione o turno.");
    if (!values.file) {
      errors.push("Selecione um arquivo Excel para continuar.");
    } else if (!/\.(xlsx|xls)$/i.test(values.file.name)) {
      errors.push(
        "O arquivo selecionado não é um Excel válido (.xlsx ou .xls).",
      );
    }

    return errors;
  }

  function validateColumnMap(columnMap) {
    const errors = [];
    CONFIG.REQUIRED_FIELDS.forEach((field) => {
      if (!columnMap[field]) {
        errors.push(
          `Não foi possível encontrar a coluna "${CONFIG.COLUMN_LABELS[field]}" na planilha.`,
        );
      }
    });
    return errors;
  }

  /** Formata uma lista de ocorrências "valor (linha N)" limitada a poucos exemplos. */
  function describeExamples(records, valueFn) {
    const max = CONFIG.MAX_VALIDATION_EXAMPLES;
    const examples = records
      .slice(0, max)
      .map((r) => `"${valueFn(r) || "(vazio)"}" (linha ${r.linha})`)
      .join(", ");
    const extra = records.length > max ? ` e mais ${records.length - max}` : "";
    return examples + extra;
  }

  /**
   * Valida as regras de conteúdo do novo layout e identifica o responsável.
   * Retorna { responsavel, errors }. Qualquer erro impede a geração.
   */
  function validateRecords(records) {
    const errors = [];

    const invalidStatus = records.filter((r) => !r.statusKey);
    if (invalidStatus.length > 0) {
      errors.push(
        `"Status usuário" deve ser CONCLUÍDO ou PENDENTE. Valores inválidos: ` +
          describeExamples(invalidStatus, (r) => r.status) +
          ".",
      );
    }

    const unmappedDisciplina = records.filter((r) => !r.disciplina);
    if (unmappedDisciplina.length > 0) {
      errors.push(
        `"Grp.plnj.PM" não mapeado para nenhuma disciplina ` +
          `(esperado: ${Object.keys(CONFIG.DISCIPLINAS).join(", ")}). Valores encontrados: ` +
          describeExamples(unmappedDisciplina, (r) => r.grpPlnj) +
          ".",
      );
    }

    const unknownCriador = records.filter(
      (r) => !mapResponsavel(r.criadoPor).nome,
    );
    if (unknownCriador.length > 0) {
      errors.push(
        `"Criado por" com chave desconhecida: ` +
          describeExamples(unknownCriador, (r) => r.criadoPor) +
          ".",
      );
    }

    // Responsável: todas as linhas precisam apontar para a mesma pessoa.
    const responsaveis = new Map();
    records.forEach((r) => {
      const { chave, nome } = mapResponsavel(r.criadoPor);
      if (nome) responsaveis.set(chave, nome);
    });

    let responsavel = null;
    if (responsaveis.size > 1) {
      const lista = Array.from(responsaveis.entries())
        .map(([chave, nome]) => `${nome} (${chave})`)
        .join(", ");
      errors.push(
        `Foram encontrados diferentes responsáveis no arquivo: ${lista}. ` +
          "O relatório deve conter chamados de um único responsável.",
      );
    } else if (responsaveis.size === 1 && unknownCriador.length === 0) {
      responsavel = Array.from(responsaveis.values())[0];
    }

    return { responsavel, errors };
  }

  // ==========================================================================
  // 7.1 PROCESSAMENTO DO EXCEL (leitura + validação + classificação)
  // ==========================================================================

  /**
   * Lê e valida o Excel, preenchendo automaticamente o campo Responsável.
   * O resultado fica em cache para o mesmo arquivo (evita reler ao gerar).
   */
  async function getProcessedExcel(file) {
    if (state.processed && state.processedFile === file) {
      return state.processed;
    }

    let processed;
    try {
      processed = await processExcel(file);
    } catch (err) {
      console.error(err);
      processed = {
        records: [],
        responsavel: null,
        errors: [err.message || "Não foi possível processar o arquivo Excel."],
      };
    }

    state.processed = processed;
    state.processedFile = file;
    dom.inputResponsavel.value = processed.responsavel || "";
    return processed;
  }

  async function processExcel(file) {
    const rawRows = CONFIG.MOCK_DATA_ENABLED
      ? gerarDadosMockParaTeste()
      : await readExcelFile(file);

    if (!rawRows || rawRows.length === 0) {
      throw new Error("O arquivo Excel está vazio: nenhum registro encontrado.");
    }

    const columnMap = buildColumnMap(Object.keys(rawRows[0]));
    const columnErrors = validateColumnMap(columnMap);
    if (columnErrors.length > 0) {
      return { records: [], responsavel: null, errors: columnErrors };
    }

    const records = normalizeRecords(rawRows, columnMap);
    if (records.length === 0) {
      throw new Error("O arquivo Excel está vazio: nenhum registro válido encontrado.");
    }

    const { responsavel, errors } = validateRecords(records);
    return { records, responsavel, errors };
  }

  // ==========================================================================
  // 8. MENSAGENS DE FEEDBACK NA INTERFACE
  // ==========================================================================

  function showMessage(text, type) {
    const box = document.createElement("div");
    box.className = `message message--${type}`;
    box.textContent = text;
    dom.messages.appendChild(box);
  }

  function clearMessages() {
    dom.messages.innerHTML = "";
  }

  // ==========================================================================
  // 9. ORQUESTRAÇÃO — GERAR RELATÓRIO
  // ==========================================================================

  async function handleGerarRelatorio() {
    clearMessages();

    const values = {
      data: dom.inputData.value,
      turno: dom.inputTurno.value,
      observacao: dom.inputObservacao.value.trim(),
      file: dom.inputExcel.files[0],
    };

    const formErrors = validateFormValues(values);
    if (formErrors.length > 0) {
      formErrors.forEach((msg) => showMessage(msg, "error"));
      return;
    }

    setButtonLoading(dom.btnGerarRelatorio, true, "Processando...");

    try {
      const processed = await getProcessedExcel(values.file);

      // Qualquer inconsistência (colunas, status, disciplina, responsável)
      // bloqueia a geração para não produzir um relatório incorreto.
      if (processed.errors.length > 0) {
        processed.errors.forEach((msg) => showMessage(msg, "error"));
        return;
      }

      const records = processed.records;

      const reportData = {
        data: values.data,
        responsavel: processed.responsavel,
        turno: values.turno,
        observacao: values.observacao,
        kpis: calculateKPIs(records),
        grouped: groupByLocalidadeDisciplina(records),
        pending: identifyPendingRecords(records),
      };

      state.reportData = reportData;

      // O painel precisa estar visível ANTES de montar o relatório: a
      // paginação e o gráfico dependem de medidas reais de layout (getBoundingClientRect/
      // clientWidth), que retornam zero dentro de um elemento com "hidden" (display:none).
      dom.formPanel.hidden = true;
      dom.reportPanel.hidden = false;

      await renderReport(reportData);

      dom.reportPanel.scrollIntoView({ behavior: "smooth", block: "start" });

      showMessage("Relatório gerado com sucesso.", "success");
    } catch (err) {
      console.error(err);
      showMessage(
        err.message || "Não foi possível processar o arquivo Excel.",
        "error",
      );
      // Garante que o formulário volte a ficar visível caso o erro ocorra
      // depois de o painel do relatório já ter sido revelado.
      dom.reportPanel.hidden = true;
      dom.formPanel.hidden = false;
    } finally {
      setButtonLoading(dom.btnGerarRelatorio, false, "Gerar Relatório");
    }
  }

  function resetToForm() {
    dom.reportPanel.hidden = true;
    dom.formPanel.hidden = false;
    clearMessages();

    destroyCharts();
    dom.reportContainer.innerHTML = "";
  }

  function destroyCharts() {
    state.chartInstances.forEach((chart) => chart.destroy());
    state.chartInstances = [];
  }

  function setButtonLoading(button, isLoading, label) {
    button.disabled = isLoading;
    button.textContent = label;
  }

  // ==========================================================================
  // 10. RENDERIZAÇÃO DO RELATÓRIO (folhas de 816x1056 + paginação automática)
  // ==========================================================================

  /** Pequeno helper para criar elementos DOM sem usar innerHTML (evita XSS com dados do Excel). */
  function h(tag, className, textContent) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (textContent !== undefined && textContent !== null)
      node.textContent = textContent;
    return node;
  }

  async function renderReport(reportData) {
    destroyCharts();
    dom.reportContainer.innerHTML = "";

    const compactMode = reportData.pending.length > CONFIG.COMPACT_THRESHOLD;

    // ---- Folha principal (cabeçalho, info, KPIs, observação, gráfico, pendências) ----
    const mainSheet = buildSheetSkeleton(reportData, {
      continuation: false,
      compact: compactMode,
    });
    dom.reportContainer.appendChild(mainSheet.root);

    const pages = [mainSheet];

    if (reportData.pending.length === 0) {
      const empty = h(
        "p",
        "sheet-table-empty",
        "Nenhum chamado pendente identificado neste turno.",
      );
      mainSheet.tableSectionEl.appendChild(empty);
    } else {
      const { fitted, remaining } = fitRowsToPage(
        mainSheet,
        reportData.pending,
        compactMode,
      );
      updateTableNote(
        mainSheet.noteEl,
        fitted.length,
        reportData.pending.length,
        remaining.length > 0,
      );

      let pendingRemaining = remaining;
      let pageIndex = 2;
      let safety = 0;

      while (pendingRemaining.length > 0 && safety < 50) {
        safety++;
        const continuationSheet = buildSheetSkeleton(reportData, {
          continuation: true,
          compact: compactMode,
          pageIndex,
        });
        dom.reportContainer.appendChild(continuationSheet.root);
        pages.push(continuationSheet);

        const result = fitRowsToPage(
          continuationSheet,
          pendingRemaining,
          compactMode,
        );
        let fittedCount = result.fitted.length;
        let remainingAfter = result.remaining;

        // Salvaguarda: garante progresso mesmo se uma única linha não couber
        // sozinha na folha (ex.: descrição muito longa).
        if (fittedCount === 0 && pendingRemaining.length > 0) {
          appendTableRow(continuationSheet.tbodyEl, pendingRemaining[0]);
          fittedCount = 1;
          remainingAfter = pendingRemaining.slice(1);
        }

        const startIdx = reportData.pending.length - pendingRemaining.length;
        updateTableNote(
          continuationSheet.noteEl,
          fittedCount,
          reportData.pending.length,
          remainingAfter.length > 0,
          startIdx,
        );

        pendingRemaining = remainingAfter;
        pageIndex++;
      }
    }

    // Numeração final de páginas (só é conhecida após montar todas as folhas).
    pages.forEach((page, idx) => {
      page.pageNumberEl.textContent = `Página ${idx + 1} de ${pages.length}`;
    });

    // Gráficos são desenhados por último, apenas na folha principal.
    // Mesma escala (eixo X) nos dois, para comparação visual direta.
    const maxTotal = Math.max(
      1,
      ...CONFIG.LOCALIDADES.flatMap((loc) =>
        reportData.grouped[loc].map((g) => g.total),
      ),
    );
    for (const localidade of CONFIG.LOCALIDADES) {
      // eslint-disable-next-line no-await-in-loop
      await renderChartWhenReady(
        mainSheet.chartCanvases[localidade],
        reportData.grouped[localidade],
        maxTotal,
      );
    }
  }

  /**
   * Monta o esqueleto de uma folha do relatório (816x1056).
   * `continuation: true` gera uma folha simplificada (sem KPIs/observação/gráfico),
   * dedicada a continuar a tabela de chamados pendentes.
   */
  function buildSheetSkeleton(reportData, options) {
    const root = h(
      "div",
      "report-sheet" + (options.compact ? " report-sheet--compact" : ""),
    );

    // ---- Cabeçalho com as duas logos ----
    const header = h("div", "sheet-header");
    const chipGerencia = h(
      "div",
      "sheet-header__chip sheet-header__chip--light",
    );
    const imgGerencia = document.createElement("img");
    imgGerencia.src = "assets/logo-gerencia.png";
    imgGerencia.alt = "Gerência de Serviços Operacionais";
    chipGerencia.appendChild(imgGerencia);

    const divider = h("div", "sheet-header__divider");

    const chipPcm = h("div", "sheet-header__chip sheet-header__chip--pcm");
    const imgPcm = document.createElement("img");
    imgPcm.src = "assets/logo-pcm.png";
    imgPcm.alt = "PCM — Serviços Operacionais";
    chipPcm.appendChild(imgPcm);

    header.appendChild(chipGerencia);
    header.appendChild(divider);
    header.appendChild(chipPcm);
    root.appendChild(header);

    // ---- Informações do turno ----
    const info = h("div", "sheet-info");
    info.appendChild(buildInfoItem("Data", formatDateBR(reportData.data)));
    info.appendChild(buildInfoItem("Responsável", reportData.responsavel));
    info.appendChild(buildInfoItem("Turno", reportData.turno));
    root.appendChild(info);

    // ---- Título ----
    const titleWrap = h("div", "sheet-title");
    const titleText = options.continuation
      ? `Chamados Pendentes — Continuação (${options.pageIndex}ª folha)`
      : "Acompanhamento de Chamados - CCM";
    titleWrap.appendChild(h("h2", null, titleText));
    titleWrap.appendChild(h("div", "sheet-title__bar"));
    root.appendChild(titleWrap);

    const chartCanvases = {};

    if (!options.continuation) {
      // ---- KPIs ----
      const kpis = h("div", "sheet-kpis");
      kpis.appendChild(
        buildKpiCard(
          "Chamados Recebidos",
          reportData.kpis.recebidos,
          CONFIG.KPI_COLORS.recebidos,
        ),
      );
      kpis.appendChild(
        buildKpiCard(
          "Chamados Concluídos",
          reportData.kpis.concluidos,
          CONFIG.KPI_COLORS.concluidos,
        ),
      );
      kpis.appendChild(
        buildKpiCard(
          "Chamados Pendentes",
          reportData.kpis.pendentes,
          CONFIG.KPI_COLORS.pendentes,
        ),
      );
      root.appendChild(kpis);

      // ---- Observação do turno ----
      const obsSection = h("div", "sheet-observacao");
      obsSection.appendChild(
        h("p", "sheet-section-label", "Observação do Turno"),
      );
      obsSection.appendChild(
        h(
          "div",
          "sheet-observacao__box",
          reportData.observacao ||
            "Sem observações registradas para este turno.",
        ),
      );
      root.appendChild(obsSection);

      // ---- Gráficos (um por localidade, lado a lado, mesma estrutura visual) ----
      const chartsRow = h("div", "sheet-charts-row");
      root.appendChild(chartsRow);
      CONFIG.LOCALIDADES.forEach((localidade) => {
        const chartSection = h("div", "sheet-chart-section");
        chartSection.appendChild(
          h(
            "p",
            "sheet-section-label",
            `Chamados por Disciplina — ${localidade}`,
          ),
        );
        const chartWrap = h("div", "sheet-chart-canvas-wrap");
        const canvas = document.createElement("canvas");
        chartWrap.appendChild(canvas);
        chartSection.appendChild(chartWrap);
        chartsRow.appendChild(chartSection);
        chartCanvases[localidade] = canvas;
      });
    }

    // ---- Tabela de pendências ----
    const tableSection = h("div", "sheet-table-section");
    tableSection.appendChild(
      h("p", "sheet-section-label", "Chamados Pendentes"),
    );
    const noteEl = h("p", "sheet-table-note", "");
    tableSection.appendChild(noteEl);

    const table = document.createElement("table");
    table.className = "sheet-table";
    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    ["Ordem", "Texto Breve", "Disciplina", "Status"].forEach((label) => {
      headRow.appendChild(h("th", null, label));
    });
    thead.appendChild(headRow);
    const tbody = document.createElement("tbody");
    table.appendChild(thead);
    table.appendChild(tbody);
    tableSection.appendChild(table);

    root.appendChild(tableSection);

    // ---- Rodapé ----
    const footer = h("div", "sheet-footer");
    const now = new Date();
    const generatedLabel = h(
      "span",
      null,
      `Gerado em ${now.toLocaleDateString("pt-BR")} às ${now.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })}`,
    );
    const pageNumberEl = h("span", null, "");
    footer.appendChild(generatedLabel);
    footer.appendChild(pageNumberEl);
    root.appendChild(footer);

    return {
      root,
      tableSectionEl: tableSection,
      noteEl,
      tbodyEl: tbody,
      footerEl: footer,
      pageNumberEl,
      chartCanvases,
    };
  }

  function buildInfoItem(label, value) {
    const wrap = h("div", "sheet-info__item");
    wrap.appendChild(h("p", "sheet-info__label", label));
    wrap.appendChild(h("p", "sheet-info__value", value || "-"));
    return wrap;
  }

  function buildKpiCard(label, value, color) {
    const card = h("div", "kpi-card");
    card.style.setProperty("--kpi-color", color);
    card.appendChild(h("div", "kpi-card__accent"));

    const body = h("div", "kpi-card__body");
    body.appendChild(h("p", "kpi-card__label", label));
    body.appendChild(h("p", "kpi-card__value", String(value)));
    card.appendChild(body);

    return card;
  }

  /**
   * Insere linhas de `records` no tbody da folha até que o conteúdo esteja
   * prestes a ultrapassar o limite vertical de 1056px. Retorna as linhas que
   * couberam e as que precisam ser levadas para a próxima folha.
   */
  function fitRowsToPage(page, records, compact) {
    const sheetRect = page.root.getBoundingClientRect();
    const footerHeight = page.footerEl.offsetHeight;
    const maxBottom = sheetRect.top + CONFIG.SHEET_HEIGHT - footerHeight - 10;

    const fitted = [];

    for (let i = 0; i < records.length; i++) {
      appendTableRow(page.tbodyEl, records[i]);
      const bottom = page.tbodyEl.getBoundingClientRect().bottom;

      if (bottom > maxBottom) {
        page.tbodyEl.removeChild(page.tbodyEl.lastElementChild);
        return { fitted, remaining: records.slice(fitted.length) };
      }
      fitted.push(records[i]);
    }

    return { fitted, remaining: [] };
  }

  function appendTableRow(tbodyEl, record) {
    const tr = document.createElement("tr");

    const tdOms = h("td", "col-oms", record.ordem || "-");
    const tdDescricao = h("td", "col-descricao", record.textoBreve || "-");
    const tdTipo = h("td", "col-tipo", record.disciplina || "-");

    const tdStatus = document.createElement("td");
    tdStatus.className = "col-status";
    const pill = h("span", "status-pill", record.status || "Pendente");
    tdStatus.appendChild(pill);

    tr.appendChild(tdOms);
    tr.appendChild(tdDescricao);
    tr.appendChild(tdTipo);
    tr.appendChild(tdStatus);
    tbodyEl.appendChild(tr);
    return tr;
  }

  function updateTableNote(
    noteEl,
    shownCount,
    totalCount,
    hasMore,
    startIndex,
  ) {
    if (startIndex !== undefined) {
      const from = startIndex + 1;
      const to = startIndex + shownCount;
      noteEl.textContent =
        `Exibindo chamados pendentes ${from} a ${to} de ${totalCount}.` +
        (hasMore ? " (continua na próxima página)" : "");
    } else {
      noteEl.textContent =
        `Total de chamados pendentes: ${totalCount}.` +
        (hasMore
          ? ` Exibindo os ${shownCount} primeiros — continua na próxima página.`
          : "");
    }
  }

  // ==========================================================================
  // 11. GRÁFICO (Chart.js)
  // ==========================================================================

  function renderChartWhenReady(canvas, groupedData, maxTotal) {
    return new Promise((resolve) => {
      if (!canvas) {
        resolve();
        return;
      }

      const wrapWidth = canvas.parentElement.clientWidth;
      const wrapHeight = canvas.parentElement.clientHeight;
      canvas.width = wrapWidth;
      canvas.height = wrapHeight;

      // Cor fixa por disciplina (mesma cor nos dois gráficos).
      const disciplinas = Object.values(CONFIG.DISCIPLINAS);
      const labels = groupedData.map((g) => g.disciplina);
      const values = groupedData.map((g) => g.total);
      const colors = labels.map(
        (label) =>
          CONFIG.CHART_PALETTE[
            disciplinas.indexOf(label) % CONFIG.CHART_PALETTE.length
          ],
      );

      const chart = new Chart(canvas.getContext("2d"), {
        type: "bar",
        data: {
          labels,
          datasets: [
            {
              data: values,
              backgroundColor: colors,
              borderRadius: 4,
              barThickness: 16,
              maxBarThickness: 20,
            },
          ],
        },
        options: {
          indexAxis: "y",
          responsive: false,
          animation: false,
          devicePixelRatio: CONFIG.EXPORT_SCALE,
          layout: { padding: { right: 28 } },
          plugins: {
            legend: { display: false },
            tooltip: { enabled: false },
          },
          scales: {
            x: {
              beginAtZero: true,
              suggestedMax: maxTotal,
              ticks: { precision: 0, font: { size: 10 }, color: "#667370" },
              grid: { color: "#eef2f1" },
            },
            y: {
              ticks: { font: { size: 11, weight: "600" }, color: "#262b2a" },
              grid: { display: false },
            },
          },
        },
        plugins: [valueLabelPlugin],
      });
      state.chartInstances.push(chart);

      // animation:false já renderiza de forma síncrona; aguarda um frame
      // extra por segurança antes de liberar a captura da imagem.
      requestAnimationFrame(() => requestAnimationFrame(resolve));
    });
  }

  /** Plugin simples do Chart.js para escrever o valor ao final de cada barra. */
  const valueLabelPlugin = {
    id: "valueLabelPlugin",
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      chart.getDatasetMeta(0).data.forEach((bar, index) => {
        const value = chart.data.datasets[0].data[index];
        ctx.save();
        ctx.fillStyle = "#262b2a";
        ctx.font = '600 11px "Segoe UI", Arial, sans-serif';
        ctx.textBaseline = "middle";
        ctx.textAlign = "left";
        ctx.fillText(String(value), bar.x + 6, bar.y);
        ctx.restore();
      });
    },
  };

  // ==========================================================================
  // 12. EXPORTAÇÃO — IMAGEM (PNG) e PDF
  // ==========================================================================

  async function waitForRenderComplete() {
    if (document.fonts && document.fonts.ready) {
      try {
        await document.fonts.ready;
      } catch (e) {
        /* segue mesmo assim */
      }
    }
    await new Promise((resolve) =>
      requestAnimationFrame(() => requestAnimationFrame(resolve)),
    );
  }

  async function captureAllPages() {
    await waitForRenderComplete();

    const pageEls = Array.from(
      dom.reportContainer.querySelectorAll(".report-sheet"),
    );
    const canvases = [];

    for (const pageEl of pageEls) {
      // eslint-disable-next-line no-await-in-loop
      const canvas = await html2canvas(pageEl, {
        scale: CONFIG.EXPORT_SCALE,
        useCORS: true,
        backgroundColor: "#ffffff",
        width: CONFIG.SHEET_WIDTH,
        height: CONFIG.SHEET_HEIGHT,
        windowWidth: CONFIG.SHEET_WIDTH,
        windowHeight: CONFIG.SHEET_HEIGHT,
      });
      canvases.push(canvas);
    }

    return canvases;
  }

  /** Empilha verticalmente várias folhas (canvases) em uma única imagem, com um pequeno respiro entre elas. */
  function combineCanvasesVertically(canvases) {
    const gap = 24 * CONFIG.EXPORT_SCALE;
    const width = canvases[0].width;
    const height =
      canvases.reduce((sum, c) => sum + c.height, 0) +
      gap * (canvases.length - 1);

    const combined = document.createElement("canvas");
    combined.width = width;
    combined.height = height;

    const ctx = combined.getContext("2d");
    ctx.fillStyle = "#e9edec";
    ctx.fillRect(0, 0, width, height);

    let offsetY = 0;
    canvases.forEach((canvas) => {
      ctx.drawImage(canvas, 0, offsetY);
      offsetY += canvas.height + gap;
    });

    return combined;
  }

  async function handleGerarImagem() {
    if (!state.reportData) return;
    setButtonLoading(dom.btnGerarImagem, true, "Gerando...");

    try {
      const canvases = await captureAllPages();
      const baseName = buildBaseFileName(state.reportData);

      // O relatório é entregue como UMA única imagem: quando há mais de uma
      // folha (paginação de pendências), as folhas são empilhadas verticalmente
      // em um único PNG — isso também evita o bloqueio do navegador para
      // downloads múltiplos disparados automaticamente pela mesma página.
      const finalCanvas =
        canvases.length > 1 ? combineCanvasesVertically(canvases) : canvases[0];
      const dataUrl = finalCanvas.toDataURL("image/png", 1.0);
      triggerDownload(dataUrl, `${baseName}.png`);

      showMessage("Imagem do relatório gerada com sucesso.", "success");
    } catch (err) {
      console.error(err);
      showMessage("Não foi possível gerar a imagem do relatório.", "error");
    } finally {
      setButtonLoading(dom.btnGerarImagem, false, "Gerar Imagem (PNG)");
    }
  }

  async function handleGerarPDF() {
    if (!state.reportData) return;
    setButtonLoading(dom.btnGerarPDF, true, "Gerando...");

    try {
      const canvases = await captureAllPages();
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({
        orientation: "portrait",
        unit: "px",
        format: [CONFIG.SHEET_WIDTH, CONFIG.SHEET_HEIGHT],
        compress: true,
      });

      canvases.forEach((canvas, index) => {
        if (index > 0)
          doc.addPage([CONFIG.SHEET_WIDTH, CONFIG.SHEET_HEIGHT], "portrait");
        const imgData = canvas.toDataURL("image/jpeg", 0.95);
        doc.addImage(
          imgData,
          "JPEG",
          0,
          0,
          CONFIG.SHEET_WIDTH,
          CONFIG.SHEET_HEIGHT,
        );
      });

      doc.save(`${buildBaseFileName(state.reportData)}.pdf`);
      showMessage("PDF do relatório gerado com sucesso.", "success");
    } catch (err) {
      console.error(err);
      showMessage("Não foi possível gerar o PDF do relatório.", "error");
    } finally {
      setButtonLoading(dom.btnGerarPDF, false, "Gerar PDF");
    }
  }

  // ==========================================================================
  // 13. UTILITÁRIOS
  // ==========================================================================

  function formatDateBR(isoDate) {
    if (!isoDate) return "-";
    const [year, month, day] = isoDate.split("-");
    return `${day}/${month}/${year}`;
  }

  function sanitizeFileNamePart(str) {
    return String(str || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-zA-Z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
  }

  function buildBaseFileName(reportData) {
    const dateForFile = reportData.data
      ? reportData.data.split("-").reverse().join("-")
      : "sem-data";
    return `Relatorio_Turno_CCM_${dateForFile}_${sanitizeFileNamePart(reportData.turno)}`;
  }

  function triggerDownload(dataUrl, filename) {
    const link = document.createElement("a");
    link.href = dataUrl;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  // ==========================================================================
  // DADOS MOCK — apenas para testes manuais durante o desenvolvimento.
  // Ative CONFIG.MOCK_DATA_ENABLED = true para usar. Remova este bloco (e a
  // referência acima) quando não for mais necessário.
  // ==========================================================================
  function gerarDadosMockParaTeste() {
    const grupos = Object.keys(CONFIG.DISCIPLINAS);
    const centros = ["RGMANUT", "SLMANUT"];
    const mock = [];
    for (let i = 1; i <= 23; i++) {
      mock.push({
        Ordem: String(40000000 + i),
        "texto breve": `Texto breve de teste do chamado número ${i}, gerado apenas para validação visual do layout.`,
        "CenTrab respon.": centros[i % centros.length],
        "Status usuário": i % 3 === 0 ? "PENDENTE" : "CONCLUÍDO",
        "Criado por": "C0731117",
        "Grp.plnj.PM": grupos[i % grupos.length],
      });
    }
    return mock;
  }
})();
