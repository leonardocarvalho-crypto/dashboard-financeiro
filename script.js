JavaScript;
// Variável global para armazenar os dados carregados na memória do Dashboard
let dadosDashboardGlobais = [];

// =========================================================================
// 1. EVENTO AO CARREGAR A PÁGINA (Garante a persistência dos dados salvos)
// =========================================================================
document.addEventListener("DOMContentLoaded", () => {
  // Garante o carregamento da biblioteca do Excel (SheetJS)
  if (!window.XLSX) {
    const scriptXLSX = document.createElement("script");
    scriptXLSX.src =
      "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
    document.head.appendChild(scriptXLSX);
  }

  // Tenta recuperar dados guardados anteriormente no navegador
  const dadosSalvos = localStorage.getItem("dadosDashboardMotos");
  if (dadosSalvos) {
    dadosDashboardGlobais = JSON.parse(dadosSalvos);
    console.log("Dados recuperados com sucesso!");
    configurarSeletorDatas(dadosDashboardGlobais);
    renderizarDashboard();
  }

  // Vincula o botão de importar planilha
  const inputUpload = document.getElementById("btn-importar");
  if (inputUpload) {
    inputUpload.addEventListener("change", lerPlanilhaExcel);
  }

  // Vincula a mudança de filtro de data
  const filtroData = document.getElementById("filtro-data");
  if (filtroData) {
    filtroData.addEventListener("change", renderizarDashboard);
  }
});

// =========================================================================
// 2. IMPORTAÇÃO DO EXCEL REAL (.XLSX) COM VALIDAÇÃO DE SENHA
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
  leitor.readAsArrayBuffer(arquivo);

  leitor.onload = function (e) {
    try {
      const dadosBrutos = new Uint8Array(e.target.result);
      const livroExcel = XLSX.read(dadosBrutos, { type: "array" });

      const primeiraAbaNome = livroExcel.SheetNames[0];
      const abaDisponivel = livroExcel.Sheets[primeiraAbaNome];

      // Converte as linhas da planilha histórica em dados utilizáveis
      const dadosConvertidos = XLSX.utils.sheet_to_json(abaDisponivel);

      if (dadosConvertidos && dadosConvertidos.length > 0) {
        // Filtra apenas as linhas referentes a 'motos' para segurança
        dadosDashboardGlobais = dadosConvertidos.filter(
          (item) =>
            item.operacao && item.operacao.toString().toLowerCase() === "motos"
        );

        // Grava permanentemente na memória do navegador do usuário
        localStorage.setItem(
          "dadosDashboardMotos",
          JSON.stringify(dadosDashboardGlobais)
        );

        alert(
          "Senha confirmada! Histórico completo carregado e salvo com sucesso."
        );
        configurarSeletorDatas(dadosDashboardGlobais);
        renderizarDashboard();
      } else {
        alert("Erro: A planilha está vazia ou os cabeçalhos estão incorretos.");
        evento.target.value = "";
      }
    } catch (erro) {
      console.error(erro);
      alert(
        "Erro ao ler o arquivo Excel. Verifique se o formato está correto."
      );
      evento.target.value = "";
    }
  };
}

// =========================================================================
// 3. SELETOR DE DATAS CORPORATIVO (Padrão solicitado: ANO - mês 1,2,3,4...)
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

  // Extrai os períodos (ex: "Maio/2025") sem repetição
  const periodosUnicos = [
    ...new Set(dados.map((item) => item.mes).filter(Boolean))
  ];

  // Ordenação básica para os meses ficarem organizados
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

    // Formato final idêntico ao modelo exigido: "2025 - maio 1,2,3,4..."
    const textoOpcaoFormatado = `${anoNum} - ${mesNomeMinusculo} ${stringDias}`;

    const opcao = document.createElement("option");
    opcao.value = periodo;
    opcao.textContent = textoOpcaoFormatado;
    seletor.appendChild(opcao);
  });
}

// =========================================================================
// 4. PROCESSAMENTO, FILTRAGEM E SOMA AUTOMÁTICA DAS CIDADES
// =========================================================================
function renderizarDashboard() {
  const seletor = document.getElementById("filtro-data");
  const periodoSelecionado = seletor ? seletor.value : "todas";

  // Aplica o filtro de período selecionado
  let dadosFiltrados = dadosDashboardGlobais;
  if (periodoSelecionado !== "todas") {
    dadosFiltrados = dadosDashboardGlobais.filter(
      (item) => item.mes == periodoSelecionado
    );
  }

  // Variáveis somadoras
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

  // Percorre todas as linhas somando os valores de todas as cidades agrupadas
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

  // Atualiza os Cards Visuais Principais
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

  // Atualiza as tabelas com os cálculos de inadimplência automáticos
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