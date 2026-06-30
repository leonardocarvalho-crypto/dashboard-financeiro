// Variável global para armazenar os dados carregados na memória do Dashboard
let dadosDashboardGlobais = [];

// =========================================================================
// 1. EVENTO AO CARREGAR A PÁGINA (Garante a persistência dos dados salvos)
// =========================================================================
document.addEventListener("DOMContentLoaded", () => {
  // Carrega a biblioteca do Excel dinamicamente se ela não existir na página
  if (!window.XLSX) {
    const scriptXLSX = document.createElement("script");
    scriptXLSX.src =
      "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
    document.head.appendChild(scriptXLSX);
  }

  const dadosSalvos = localStorage.getItem("dadosDashboardMotos");
  if (dadosSalvos) {
    dadosDashboardGlobais = JSON.parse(dadosSalvos);
    configurarSeletorDatas(dadosDashboardGlobais);
    renderizarDashboard();
  }

  const inputUpload = document.getElementById("btn-importar");
  if (inputUpload) {
    inputUpload.addEventListener("change", lerPlanilhaExcel);
  }

  const filtroData = document.getElementById("filtro-data");
  if (filtroData) {
    filtroData.addEventListener("change", renderizarDashboard);
  }
});

// =========================================================================
// 2. IMPORTAÇÃO DO EXCEL REAL (.XLSX) COM SENHA E PERSISTÊNCIA
// =========================================================================
function lerPlanilhaExcel(evento) {
  const senhaCorreta = "12345678";
  const senhaDigitada = prompt(
    "Digite a senha de administrador para importar e salvar novos dados:"
  );

  if (senhaDigitada === null) {
    evento.target.value = "";
    return;
  }

  if (senhaDigitada !== senhaCorreta) {
    alert(
      "Senha incorreta! Você não tem permissão para alterar os dados deste Dashboard."
    );
    evento.target.value = "";
    return;
  }

  const arquivo = evento.target.files[0];
  if (!arquivo) return;

  const leitor = new FileReader();

  // 🔥 AQUI ESTÁ A MUDANÇA: Agora lê o arquivo como ArrayBuffer (formato do Excel bruto)
  leitor.readAsArrayBuffer(arquivo);

  leitor.onload = function (e) {
    try {
      const dadosBrutos = new Uint8Array(e.target.result);
      const livroExcel = XLSX.read(dadosBrutos, { type: "array" });

      // Pega a primeira aba da planilha do Excel
      const primeiraAbaNome = livroExcel.SheetNames[0];
      const abaDisponivel = livroExcel.Sheets[primeiraAbaNome];

      // Converte as linhas do Excel diretamente para JSON
      const dadosAgrupados = XLSX.utils.sheet_to_json(abaDisponivel);

      if (dadosAgrupados && dadosAgrupados.length > 0) {
        dadosDashboardGlobais = dadosAgrupados;

        // Grava permanentemente no navegador
        localStorage.setItem(
          "dadosDashboardMotos",
          JSON.stringify(dadosAgrupados)
        );

        alert(
          "Senha confirmada! Planilha do Excel (.xlsx) importada e salva com sucesso."
        );
        configurarSeletorDatas(dadosDashboardGlobais);
        renderizarDashboard();
      } else {
        alert("Erro: A planilha está vazia ou sem dados compatíveis.");
        evento.target.value = "";
      }
    } catch (erro) {
      console.error(erro);
      alert(
        "Erro ao processar o arquivo Excel. Verifique se o arquivo não está corrompido."
      );
      evento.target.value = "";
    }
  };
}

// =========================================================================
// 3. SELETOR DE DATAS CORPORATIVO (Padrão: ANO - mês 1,2,3,4,5...)
// =========================================================================
function configurarSeletorDatas(dados) {
  const seletor = document.getElementById("filtro-data");
  if (!seletor) return;

  seletor.innerHTML =
    '<option value="todas">Data Início e Fim (Motos)</option>';

  const mesesMap = {
    janeiro: 31,
    fevereiro: 28,
    março: 31,
    abril: 30,
    maio: 31,
    junho: 30,
    julho: 31,
    agosto: 31,
    setembro: 30,
    outubro: 31,
    novembro: 30,
    dezembro: 31
  };

  // Filtra os meses de forma única
  const periodosUnicos = [
    ...new Set(dados.map((item) => item.mes).filter(Boolean))
  ];

  periodosUnicos.forEach((periodo) => {
    const partes = periodo.toString().split("/");
    const mesNomeOriginal = partes[0] || "Janeiro";
    const mesNomeMinusculo = mesNomeOriginal.toLowerCase().trim();
    const anoNum = partes[1] || "2025";

    const limiteDias = mesesMap[mesNomeMinusculo] || 31;
    let arrayDias = [];
    for (let i = 1; i <= limiteDias; i++) {
      arrayDias.push(i);
    }
    const stringDias = arrayDias.join(",");

    const textoOpcaoFormatado = `${anoNum} - ${mesNomeMinusculo} ${stringDias}`;

    const opcao = document.createElement("option");
    opcao.value = periodo;
    opcao.textContent = textoOpcaoFormatado;
    seletor.appendChild(opcao);
  });
}

// =========================================================================
// 4. PROCESSAMENTO E CÁLCULO GERAL (Renderização na Tela)
// =========================================================================
function renderizarDashboard() {
  const seletor = document.getElementById("filtro-data");
  const periodoSelecionado = seletor ? seletor.value : "todas";

  let dadosFiltrados = dadosDashboardGlobais;
  if (periodoSelecionado !== "todas") {
    dadosFiltrados = dadosDashboardGlobais.filter(
      (item) => item.mes == periodoSelecionado
    );
  }

  let totalEmitido = 0;
  let totalRecebido = 0;
  let totalBoletosEmitidos = 0;
  let totalBoletosRecebidos = 0;
  let totalRoyaltiesEmitido = 0;
  let totalRoyaltiesRecebido = 0;
  let totalSplitsEmitido = 0;
  let totalSplitsRecebido = 0;
  let totalGestaoEmitido = 0;
  let totalGestaoRecebido = 0;
  let totalMktEmitido = 0;
  let totalMktRecebido = 0;

  dadosFiltrados.forEach((item) => {
    totalEmitido += parseFloat(item.emitidos || 0);
    totalRecebido += parseFloat(item.recebidos || 0);
    totalBoletosEmitidos += parseInt(item.qtdEmitidos || 0);
    totalBoletosRecebidos += parseInt(item.qtdRecebidos || 0);

    totalRoyaltiesEmitido += parseFloat(item.royalties || 0);
    totalRoyaltiesRecebido += parseFloat(item.royaltiesRec || 0);
    totalSplitsEmitido += parseFloat(item.splits || 0);
    totalSplitsRecebido += parseFloat(item.splitsRec || 0);
    totalGestaoEmitido += parseFloat(item.gestao || 0);
    totalGestaoRecebido += parseFloat(item.gestaoRec || 0);
    totalMktEmitido += parseFloat(item.mkt || 0);
    totalMktRecebido += parseFloat(item.mktRec || 0);
  });

  if (document.getElementById("motos-rec-bruta"))
    document.getElementById("motos-rec-bruta").textContent =
      "R$ " +
      totalRecebido.toLocaleString("pt-BR", { minimumFractionDigits: 2 });
  if (document.getElementById("motos-rec-liquida"))
    document.getElementById("motos-rec-liquida").textContent =
      "R$ " +
      totalRecebido.toLocaleString("pt-BR", { minimumFractionDigits: 2 });
  if (document.getElementById("motos-emitidos"))
    document.getElementById("motos-emitidos").textContent =
      "R$ " +
      totalEmitido.toLocaleString("pt-BR", { minimumFractionDigits: 2 });

  let totalEmAberto = totalEmitido - totalRecebido;
  if (document.getElementById("motos-em-aberto"))
    document.getElementById("motos-em-aberto").textContent =
      "R$ " +
      (totalEmAberto > 0 ? totalEmAberto : 0).toLocaleString("pt-BR", {
        minimumFractionDigits: 2
      });

  if (document.getElementById("motos-qtd-emitidos"))
    document.getElementById("motos-qtd-emitidos").textContent =
      totalBoletosEmitidos.toLocaleString("pt-BR") + " un";
  if (document.getElementById("motos-qtd-recebidos"))
    document.getElementById("motos-qtd-recebidos").textContent =
      totalBoletosRecebidos.toLocaleString("pt-BR") + " un";

  atualizarLinhaTabela(
    "motos-t1",
    totalRoyaltiesEmitido,
    totalRoyaltiesRecebido
  );
  atualizarLinhaTabela("motos-t2", totalSplitsEmitido, totalSplitsRecebido);
  atualizarLinhaTabela("motos-t3", totalGestaoEmitido, totalGestaoRecebido);
  atualizarLinhaTabela("motos-t4", totalMktEmitido, totalMktRecebido);
}

function atualizarLinhaTabela(idPrefixo, emitido, recebido) {
  const campoEmitido = document.getElementById(`${idPrefixo}-emitido`);
  const campoRecebido = document.getElementById(`${idPrefixo}-recebido`);
  const campoInad = document.getElementById(`${idPrefixo}-inad`);

  if (campoEmitido)
    campoEmitido.textContent =
      "R$ " + emitido.toLocaleString("pt-BR", { minimumFractionDigits: 2 });
  if (campoRecebido)
    campoRecebido.textContent =
      "R$ " + recebido.toLocaleString("pt-BR", { minimumFractionDigits: 2 });

  if (campoInad) {
    let inadPorcentagem = 0;
    if (emitido > 0) {
      inadPorcentagem = ((emitido - recebido) / emitido) * 100;
    }
    campoInad.textContent =
      (inadPorcentagem > 0 ? inadPorcentagem.toFixed(1) : "0.0") + " %";
  }
}