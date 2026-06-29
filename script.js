// --- ESTRUTURA GLOBAL DE DADOS DA EMPRESA ---
let db = { motos: [], seminovos: [], expansao: [] };
let mesesDisponiveis = [];

function formatarMoeda(valor) {
  if (isNaN(valor) || valor === null) valor = 0;
  return (
    "R$ " +
    valor.toLocaleString("pt-BR", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    })
  );
}

// Configurações estruturais de exibição do Chart.js
const chartConfigs = {
  responsive: true,
  maintainAspectRatio: false,
  plugins: { legend: { display: false } },
  scales: {
    y: { display: false },
    x: { grid: { display: false }, ticks: { font: { size: 9 } } }
  }
};

// Inicialização das Instâncias Globais dos Gráficos (Vazios antes do upload)
const chartMotosFluxo = new Chart(
  document.getElementById("chart-motos-fluxo").getContext("2d"),
  {
    type: "line",
    data: {
      labels: [],
      datasets: [
        { data: [], borderColor: "#16a34a", tension: 0.3, fill: false }
      ]
    },
    options: chartConfigs
  }
);
const chartMotosTaxas = new Chart(
  document.getElementById("chart-motos-taxas").getContext("2d"),
  {
    type: "bar",
    data: { labels: [], datasets: [{ data: [], backgroundColor: "#1e3a8a" }] },
    options: chartConfigs
  }
);
const chartSemiVendas = new Chart(
  document.getElementById("chart-semi-vendas").getContext("2d"),
  {
    type: "line",
    data: {
      labels: [],
      datasets: [{ data: [], borderColor: "#16a34a", tension: 0.3 }]
    },
    options: chartConfigs
  }
);
const chartSemiCarteira = new Chart(
  document.getElementById("chart-semi-carteira").getContext("2d"),
  {
    type: "bar",
    data: { labels: [], datasets: [{ data: [], backgroundColor: "#2563eb" }] },
    options: chartConfigs
  }
);
const chartSemiRisco = new Chart(
  document.getElementById("chart-semi-risco").getContext("2d"),
  {
    type: "line",
    data: {
      labels: [],
      datasets: [{ data: [], borderColor: "#dc2626", tension: 0.1 }]
    },
    options: chartConfigs
  }
);
const chartExpFranquias = new Chart(
  document.getElementById("chart-exp-franquias").getContext("2d"),
  {
    type: "bar",
    data: { labels: [], datasets: [{ data: [], backgroundColor: "#8b5cf6" }] },
    options: chartConfigs
  }
);
const chartExpMotos = new Chart(
  document.getElementById("chart-exp-motos").getContext("2d"),
  {
    type: "line",
    data: {
      labels: [],
      datasets: [{ data: [], borderColor: "#ec4899", tension: 0.3 }]
    },
    options: chartConfigs
  }
);

// --- CONVERSOR EXCLUSIVO DE PLANILHA EXCEL (.XLSX) ---
function importarExcelReal(input) {
  const file = input.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    // Reseta o banco temporário
    db.motos = [];
    db.seminovos = [];
    db.expansao = [];
    let mesesSet = new Set();
    let unidadesMotosSet = new Set();
    let unidadesSemiSet = new Set();
    let regExpSet = new Set();

    // Obtém a primeira aba de dados do arquivo Excel
    const nomeDaAba = workbook.SheetNames[0];
    const planilha = workbook.Sheets[nomeDaAba];

    // Extrai a tabela convertendo-a em matriz pura (linhas x colunas)
    const linhas = XLSX.utils.sheet_to_json(planilha, { header: 1 });

    if (linhas.length < 2) return alert("A planilha selecionada está vazia!");

    // Processa linha por linha pulando o cabeçalho (i=0)
    for (let i = 1; i < linhas.length; i++) {
      let col = linhas[i];
      if (!col || col.length === 0 || !col[0]) continue;

      let tipoAba = String(col[0]).toLowerCase().trim();
      let mes = String(col[1]).trim();
      if (!mes) continue;
      mesesSet.add(mes);

      if (tipoAba === "motos") {
        unidadesMotosSet.add(String(col[2]));
        db.motos.push({
          mes: mes,
          unidade: String(col[2]),
          qtdEmitidos: parseInt(col[3]) || 0,
          qtdRecebidos: parseInt(col[4]) || 0,
          emitidos: parseFloat(col[5]) || 0,
          recebidos: parseFloat(col[6]) || 0,
          royalties: parseFloat(col[7]) || 0,
          royaltiesRec: parseFloat(col[8]) || 0,
          splits: parseFloat(col[9]) || 0,
          splitsRec: parseFloat(col[10]) || 0,
          gestao: parseFloat(col[11]) || 0,
          gestaoRec: parseFloat(col[12]) || 0,
          mkt: parseFloat(col[13]) || 0,
          mktRec: parseFloat(col[14]) || 0,
          vendasQtd: parseInt(col[15]) || 0,
          status: String(col[16] || "ADIMPLENTE")
            .toUpperCase()
            .trim()
        });
      } else if (tipoAba === "seminovos") {
        unidadesSemiSet.add(String(col[2]));
        db.seminovos.push({
          mes: mes,
          unidade: String(col[2]),
          fatBruto: parseFloat(col[3]) || 0,
          boletosMes: parseFloat(col[4]) || 0,
          recebidos: parseFloat(col[5]) || 0,
          motosVendidas: parseInt(col[6]) || 0,
          futuras: parseFloat(col[7]) || 0,
          vencidos: parseFloat(col[8]) || 0
        });
      } else if (tipoAba === "expansao") {
        regExpSet.add(String(col[2]));
        db.expansao.push({
          mes: mes,
          regiao: String(col[2]),
          contract: String(col[3]),
          franBruto: parseFloat(col[4]) || 0,
          franPagas: parseFloat(col[5]) || 0,
          motosBruto: parseFloat(col[6]) || 0,
          motosPagas: parseFloat(col[7]) || 0,
          status: String(col[8] || "Regular").trim()
        });
      }
    }

    mesesDisponiveis = Array.from(mesesSet);

    // Alimenta as caixas de seleção dinamicamente conforme os dados do Excel
    atualizarDropdown(
      "global-period-select",
      mesesDisponiveis,
      "[Período Global]"
    );
    atualizarDropdown(
      "motos-periodo-select",
      mesesDisponiveis,
      "Data Início e Fim (Motos)"
    );
    atualizarDropdown(
      "semi-periodo-select",
      mesesDisponiveis,
      "Período de Venda (Mensal)"
    );
    atualizarDropdown(
      "exp-periodo-select",
      mesesDisponiveis,
      "Funil de Vendas Expansão"
    );

    atualizarDropdown(
      "motos-unidade-select",
      Array.from(unidadesMotosSet),
      "Filtrar por Unidade Master (Todas)",
      "Todas"
    );
    atualizarDropdown(
      "semi-unidade-select",
      Array.from(unidadesSemiSet),
      "Filtrar por Unidade Seminovos (Todas)",
      "Todas"
    );
    atualizarDropdown(
      "exp-regiao-select",
      Array.from(regExpSet),
      "Filtrar por Supervisor / Região (Todos)",
      "Todos"
    );

    // Configura o eixo do tempo nos gráficos de linhas/barras
    [
      chartMotosFluxo,
      chartMotosTaxas,
      chartSemiVendas,
      chartSemiCarteira,
      chartSemiRisco,
      chartExpFranquias,
      chartExpMotos
    ].forEach((c) => {
      c.data.labels = mesesDisponiveis;
    });

    renderDashboard();
  };
  reader.readAsArrayBuffer(file);
}

function atualizarDropdown(id, lista, textoPadrao, valorPadrao = "Todos") {
  const el = document.getElementById(id);
  el.innerHTML = `<option value="${valorPadrao}">${textoPadrao}</option>`;
  lista.forEach((item) => {
    if (item) el.innerHTML += `<option value="${item}">${item}</option>`;
  });
}

// --- RENDERIZADOR DINÂMICO DOS COMPONENTES E GRÁFICOS ---
function renderDashboard() {
  const globalMes = document.getElementById("global-period-select").value;

  const mUnidade = document.getElementById("motos-unidade-select").value;
  const mMes =
    document.getElementById("motos-periodo-select").value !== "Todos"
      ? document.getElementById("motos-periodo-select").value
      : globalMes;
  const taxaSelecionada = document.getElementById("motos-taxa-select").value;

  const sUnidade = document.getElementById("semi-unidade-select").value;
  const sMes =
    document.getElementById("semi-periodo-select").value !== "Todos"
      ? document.getElementById("semi-periodo-select").value
      : globalMes;

  const eRegiao = document.getElementById("exp-regiao-select").value;
  const eMes =
    document.getElementById("exp-periodo-select").value !== "Todos"
      ? document.getElementById("exp-periodo-select").value
      : globalMes;

  // --- CÁLCULOS DA ABA 1: MOTOS ---
  let fMotos = db.motos.filter(
    (i) =>
      (mUnidade === "Todas" || i.unidade === mUnidade) &&
      (mMes === "Todos" || i.mes === mMes)
  );
  let tEmitidos = 0,
    tRecebidos = 0,
    tQtdEmit = 0,
    tQtdRec = 0;
  let tRoyE = 0,
    tRoyR = 0,
    tSplE = 0,
    tSplR = 0,
    tGesE = 0,
    tGesR = 0,
    tMktE = 0,
    tMktR = 0;

  fMotos.forEach((i) => {
    tEmitidos += i.emitidos;
    tRecebidos += i.recebidos;
    tQtdEmit += i.qtdEmitidos;
    tQtdRec += i.qtdRecebidos;
    tRoyE += i.royalties;
    tRoyR += i.royaltiesRec;
    tSplE += i.splits;
    tSplR += i.splitsRec;
    tGesE += i.gestao;
    tGesR += i.gestaoRec;
    tMktE += i.mkt;
    tMktR += i.mktRec;
  });

  document.getElementById("motos-p-recebido").innerText =
    (tEmitidos > 0 ? ((tRecebidos / tEmitidos) * 100).toFixed(1) : "0") +
    "% recebido";
  document.getElementById("motos-rec-bruta").innerText = formatarMoeda(
    tRecebidos
  );
  document.getElementById("motos-rec-liquida").innerText = formatarMoeda(
    tRecebidos * 0.85
  );
  document.getElementById("motos-emitidos").innerText = formatarMoeda(
    tEmitidos
  );
  document.getElementById("motos-em-aberto").innerText = formatarMoeda(
    tEmitidos - tRecebidos
  );
  document.getElementById("motos-qtd-emitidos").innerText = tQtdEmit + " un";
  document.getElementById("motos-qtd-recebidos").innerText = tQtdRec + " un";

  document.getElementById("tax-roy-emit").innerText = formatarMoeda(tRoyE);
  document.getElementById("tax-roy-rec").innerText = formatarMoeda(tRoyR);
  document.getElementById("tax-roy-inad").innerText =
    (tRoyE > 0 ? (((tRoyE - tRoyR) / tRoyE) * 100).toFixed(1) : 0) + " %";
  document.getElementById("tax-split-emit").innerText = formatarMoeda(tSplE);
  document.getElementById("tax-split-rec").innerText = formatarMoeda(tSplR);
  document.getElementById("tax-split-inad").innerText =
    (tSplE > 0 ? (((tSplE - tSplR) / tSplE) * 100).toFixed(1) : 0) + " %";
  document.getElementById("tax-gestao-emit").innerText = formatarMoeda(tGesE);
  document.getElementById("tax-gestao-rec").innerText = formatarMoeda(tGesR);
  document.getElementById("tax-gestao-inad").innerText =
    (tGesE > 0 ? (((tGesE - tGesR) / tGesE) * 100).toFixed(1) : 0) + " %";
  document.getElementById("tax-mkt-emit").innerText = formatarMoeda(tMktE);
  document.getElementById("tax-mkt-rec").innerText = formatarMoeda(tMktR);
  document.getElementById("tax-mkt-inad").innerText =
    (tMktE > 0 ? (((tMktE - tMktR) / tMktE) * 100).toFixed(1) : 0) + " %";

  let tbodyMotos = document.querySelector("#table-motos tbody");
  tbodyMotos.innerHTML =
    fMotos.length === 0
      ? '<tr><td colspan="6" style="text-align:center; padding:15px;">Nenhum dado integrado</td></tr>'
      : "";
  let agrupadoMotos = {};
  fMotos.forEach((i) => {
    if (!agrupadoMotos[i.unidade])
      agrupadoMotos[i.unidade] = { vQtd: 0, emit: 0, rec: 0, status: i.status };
    agrupadoMotos[i.unidade].vQtd += i.vendasQtd;
    agrupadoMotos[i.unidade].emit += i.emitidos;
    agrupadoMotos[i.unidade].rec += i.recebidos;
  });
  for (let key in agrupadoMotos) {
    let item = agrupadoMotos[key];
    tbodyMotos.innerHTML += `<tr><td><strong>${key}</strong></td><td>${
      item.vQtd
    }</td><td>${formatarMoeda(
      item.emit
    )}</td><td style="color: #16a34a; font-weight: bold;">${formatarMoeda(
      item.rec
    )}</td><td style="color: #dc2626;">${formatarMoeda(
      item.emit - item.rec
    )}</td><td style="text-align: center;"><span class="badge ${item.status.toLowerCase()}">${
      item.status
    }</span></td></tr>`;
  }

  // --- CÁLCULOS DA ABA 2: SEMINOVOS ---
  let fSemi = db.seminovos.filter(
    (i) =>
      (sUnidade === "Todas" || i.unidade === sUnidade) &&
      (sMes === "Todos" || i.mes === sMes)
  );
  let sFatBruto = 0,
    sBoletos = 0,
    sRecebidos = 0,
    sMotosQtd = 0,
    sFuturas = 0,
    sVencidos = 0;
  fSemi.forEach((i) => {
    sFatBruto += i.fatBruto;
    sBoletos += i.boletosMes;
    sRecebidos += i.recebidos;
    sMotosQtd += i.motosVendidas;
    sFuturas += i.futuras;
    sVencidos += i.vencidos;
  });

  document.getElementById("semi-fat-bruto").innerText = formatarMoeda(
    sFatBruto
  );
  document.getElementById("semi-boletos-mes").innerText = formatarMoeda(
    sBoletos
  );
  document.getElementById("semi-recebidos").innerText = formatarMoeda(
    sRecebidos
  );
  document.getElementById("semi-qtd-motos").innerText = sMotosQtd + " un";
  document.getElementById("semi-futuras").innerText = formatarMoeda(sFuturas);
  document.getElementById("semi-vencidos").innerText = formatarMoeda(sVencidos);
  document.getElementById("semi-inad-pct").innerText =
    (sBoletos > 0 ? ((sVencidos / sBoletos) * 100).toFixed(1) : "0.0") + " %";

  let tbodySemi = document.querySelector("#table-seminovos tbody");
  tbodySemi.innerHTML =
    fSemi.length === 0
      ? '<tr><td colspan="7" style="text-align:center; padding:15px;">Nenhum dado integrado</td></tr>'
      : "";
  fSemi.forEach((i) => {
    let inadP =
      i.boletosMes > 0
        ? (((i.boletosMes - i.recebidos) / i.boletosMes) * 100).toFixed(1)
        : 0;
    tbodySemi.innerHTML += `<tr><td><strong>${i.unidade}</strong></td><td>${
      i.mes
    }</td><td>${i.motosVendidas} un</td><td>${formatarMoeda(
      i.fatBruto
    )}</td><td>${formatarMoeda(i.boletosMes)}</td><td>${formatarMoeda(
      i.futuras
    )}</td><td style="color:#dc2626; font-weight:bold;">${inadP} %</td></tr>`;
  });

  // --- CÁLCULOS DA ABA 3: EXPANSÃO ---
  let fExp = db.expansao.filter(
    (i) =>
      (eRegiao === "Todos" || i.regiao === eRegiao) &&
      (eMes === "Todos" || i.mes === eMes)
  );
  let eFranB = 0,
    eFranP = 0,
    eMotosB = 0,
    eMotosP = 0;
  fExp.forEach((i) => {
    eFranB += i.franBruto;
    eFranP += i.franPagas;
    eMotosB += i.motosBruto;
    eMotosP += i.motosPagas;
  });

  document.getElementById("exp-fran-bruto").innerText = formatarMoeda(eFranB);
  document.getElementById("exp-fran-pagas").innerText = formatarMoeda(eFranP);
  document.getElementById("exp-fran-aberto").innerText = formatarMoeda(
    eFranB - eFranP
  );
  document.getElementById("exp-fran-inad").innerText =
    (eFranB > 0 ? (((eFranB - eFranP) / eFranB) * 100).toFixed(1) : 0) + " %";
  document.getElementById("exp-motos-bruto").innerText = formatarMoeda(eMotosB);
  document.getElementById("exp-motos-pagas").innerText = formatarMoeda(eMotosP);
  document.getElementById("exp-motos-aberto").innerText = formatarMoeda(
    eMotosB - eMotosP
  );
  document.getElementById("exp-motos-inad").innerText =
    (eMotosB > 0 ? (((eMotosB - eMotosP) / eMotosB) * 100).toFixed(1) : 0) +
    " %";

  let tbodyExp = document.querySelector("#table-expansao tbody");
  tbodyExp.innerHTML =
    fExp.length === 0
      ? '<tr><td colspan="6" style="text-align:center; padding:15px;">Nenhum dado integrado</td></tr>'
      : "";
  fExp.forEach((i) => {
    let eInad =
      i.franBruto > 0
        ? (((i.franBruto - i.franPagas) / i.franBruto) * 100).toFixed(1)
        : 0;
    tbodyExp.innerHTML += `<tr><td><strong>${i.contract}</strong></td><td>${
      i.status
    }</td><td>${formatarMoeda(i.franBruto)}</td><td>${formatarMoeda(
      i.franPagas
    )}</td><td>${formatarMoeda(
      i.franPagas + i.motosPagas
    )}</td><td style="color:#dc2626; font-weight:bold;">${eInad} %</td></tr>`;
  });

  // --- ATUALIZAÇÃO REATIVA DOS GRÁFICOS NO TEMPO (EIXO X MENSAL) ---
  let evoMotosFluxo = [],
    evoMotosTaxas = [],
    evoSemiVendas = [],
    evoSemiCarteira = [],
    evoSemiRisco = [],
    evoExpFran = [],
    evoExpMotos = [];

  mesesDisponiveis.forEach((m) => {
    let dadosMesMotos = db.motos.filter(
      (i) => (mUnidade === "Todas" || i.unidade === mUnidade) && i.mes === m
    );
    let valorRecebidoTaxa = 0;
    let valorEmAbertoTaxa = 0;

    dadosMesMotos.forEach((i) => {
      if (taxaSelecionada === "royalties") {
        valorRecebidoTaxa += i.royaltiesRec;
        valorEmAbertoTaxa += i.royalties - i.royaltiesRec;
      } else if (taxaSelecionada === "splits") {
        valorRecebidoTaxa += i.splitsRec;
        valorEmAbertoTaxa += i.splits - i.splitsRec;
      } else if (taxaSelecionada === "gestao") {
        valorRecebidoTaxa += i.gestaoRec;
        valorEmAbertoTaxa += i.gestao - i.gestaoRec;
      } else if (taxaSelecionada === "mkt") {
        valorRecebidoTaxa += i.mktRec;
        valorEmAbertoTaxa += i.mkt - i.mktRec;
      } else {
        valorRecebidoTaxa += i.recebidos;
        valorEmAbertoTaxa += i.emitidos - i.recebidos;
      }
    });

    evoMotosFluxo.push(valorRecebidoTaxa);
    evoMotosTaxas.push(valorEmAbertoTaxa);

    evoSemiVendas.push(
      db.seminovos
        .filter(
          (i) => (sUnidade === "Todas" || i.unidade === sUnidade) && i.mes === m
        )
        .reduce((a, b) => a + b.fatBruto, 0)
    );
    evoSemiCarteira.push(
      db.seminovos
        .filter(
          (i) => (sUnidade === "Todas" || i.unidade === sUnidade) && i.mes === m
        )
        .reduce((a, b) => a + b.futuras, 0)
    );
    evoSemiRisco.push(
      db.seminovos
        .filter(
          (i) => (sUnidade === "Todas" || i.unidade === sUnidade) && i.mes === m
        )
        .reduce((a, b) => a + b.vencidos, 0)
    );
    evoExpFran.push(
      db.expansao
        .filter(
          (i) => (eRegiao === "Todos" || i.regiao === eRegiao) && i.mes === m
        )
        .reduce((a, b) => a + b.franBruto, 0)
    );
    evoExpMotos.push(
      db.expansao
        .filter(
          (i) => (eRegiao === "Todos" || i.regiao === eRegiao) && i.mes === m
        )
        .reduce((a, b) => a + b.motosBruto, 0)
    );
  });

  chartMotosFluxo.data.datasets[0].data = evoMotosFluxo;
  chartMotosFluxo.update();
  chartMotosTaxas.data.datasets[0].data = evoMotosTaxas;
  chartMotosTaxas.update();
  chartSemiVendas.data.datasets[0].data = evoSemiVendas;
  chartSemiVendas.update();
  chartSemiCarteira.data.datasets[0].data = evoSemiCarteira;
  chartSemiCarteira.update();
  chartSemiRisco.data.datasets[0].data = evoSemiRisco;
  chartSemiRisco.update();
  chartExpFranquias.data.datasets[0].data = evoExpFran;
  chartExpFranquias.update();
  chartExpMotos.data.datasets[0].data = evoExpMotos;
  chartExpMotos.update();

  // --- MÓDULO DO BANNER MACRO CONSOLIDADO ---
  let globalMotos = db.motos.filter(
    (i) => globalMes === "Todos" || i.mes === globalMes
  );
  let globalSemi = db.seminovos.filter(
    (i) => globalMes === "Todos" || i.mes === globalMes
  );
  let globalExp = db.expansao.filter(
    (i) => globalMes === "Todos" || i.mes === globalMes
  );

  let bEmitidos =
    globalMotos.reduce((a, b) => a + b.emitidos, 0) +
    globalSemi.reduce((a, b) => a + b.boletosMes, 0) +
    globalExp.reduce((a, b) => a + b.franBruto + b.motosBruto, 0);
  let bRecebidos =
    globalMotos.reduce((a, b) => a + b.recebidos, 0) +
    globalSemi.reduce((a, b) => a + b.recebidos, 0) +
    globalExp.reduce((a, b) => a + b.franPagas + b.motosPagas, 0);

  document.getElementById("kpi-fat-geral").innerText = formatarMoeda(
    bRecebidos
  );
  document.getElementById("kpi-bruto-conso").innerText = formatarMoeda(
    bEmitidos
  );
  document.getElementById("kpi-liq-conso").innerText = formatarMoeda(
    bRecebidos * 0.9
  );
  document.getElementById("kpi-subtext-liq").innerText =
    (bEmitidos > 0 ? ((bRecebidos / bEmitidos) * 100).toFixed(1) : 0) +
    "% do bruto";
  document.getElementById("kpi-inad-geral").innerText =
    (bEmitidos > 0
      ? (((bEmitidos - bRecebidos) / bEmitidos) * 100).toFixed(1)
      : 0) + " %";
}

// Ouvintes das chaves de filtros
document
  .getElementById("global-period-select")
  .addEventListener("change", () => {
    document.getElementById("motos-periodo-select").value = "Todos";
    document.getElementById("semi-periodo-select").value = "Todos";
    document.getElementById("exp-periodo-select").value = "Todos";
    renderDashboard();
  });
[
  "motos-unidade-select",
  "motos-periodo-select",
  "motos-taxa-select",
  "semi-unidade-select",
  "semi-periodo-select",
  "exp-regiao-select",
  "exp-periodo-select"
].forEach((id) => {
  document.getElementById(id).addEventListener("change", renderDashboard);
});

// Extrator Analítico das tabelas da interface para planilhas limpas .CSV
function exportTableToCSV(tableID, filename) {
  var csv = [];
  var rows = document.querySelectorAll("#" + tableID + " tr");
  for (var i = 0; i < rows.length; i++) {
    var row = [],
      cols = rows[i].querySelectorAll("td, th");
    for (var j = 0; j < cols.length; j++) {
      var data = cols[j].innerText
        .replace(/(\r\n|\n|\r)/gm, "")
        .replace(/^"|"$/g, "");
      row.push('"' + data + '"');
    }
    csv.push(row.join(";"));
  }
  var csvFile = new Blob(["\ufeff" + csv.join("\n")], {
    type: "text/csv;charset=utf-8;"
  });
  var downloadLink = document.createElement("a");
  downloadLink.download = filename + ".csv";
  downloadLink.href = window.URL.createObjectURL(csvFile);
  downloadLink.style.display = "none";
  document.body.appendChild(downloadLink);
  downloadLink.click();
  document.body.removeChild(downloadLink);
}