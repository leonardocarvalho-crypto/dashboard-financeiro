// Variável global para armazenar os dados carregados na memória do Dashboard
let dadosDashboardGlobais = [];

// =========================================================================
// 1. EVENTO AO CARREGAR A PÁGINA (Garante a persistência dos dados salvos)
// =========================================================================
document.addEventListener("DOMContentLoaded", () => {
  const dadosSalvos = localStorage.getItem("dadosDashboardMotos");
  if (dadosSalvos) {
    try {
      dadosDashboardGlobais = JSON.parse(dadosSalvos);
      configurarSeletorDatas(dadosDashboardGlobais);
      renderizarDashboard();
    } catch (e) {
      console.error("Erro ao ler dados salvos:", e);
    }
  }

  const inputUpload = document.getElementById("btn-importar");
  if (inputUpload) {
    inputUpload.addEventListener("change", lerPlanilhaExcelReal);
  }

  const filtroData = document.getElementById("filtro-data");
  if (filtroData) {
    filtroData.addEventListener("change", renderizarDashboard);
  }
});

// =========================================================================
// 2. LEITOR DE EXCEL REAL (.XLSX) - VERSÃO ULTRA TOLERANTE
// =========================================================================
function lerPlanilhaExcelReal(evento) {
  if (typeof XLSX === "undefined") {
    alert(
      "Erro: A biblioteca do Excel ainda está sendo carregada pelo navegador. Aguarde 3 segundos e tente novamente."
    );
    evento.target.value = "";
    return;
  }

  const senhaCorreta = "12345678";
  const senhaDigitada = prompt(
    "Digite a senha de administrador para importar os dados do Excel:"
  );

  if (senhaDigitada === null) {
    evento.target.value = "";
    return;
  }

  if (senhaDigitada !== senhaCorreta) {
    alert("Senha incorreta!");
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

      // Converte a aba do Excel diretamente em JSON
      const dadosConvertidos = XLSX.utils.sheet_to_json(abaDisponivel);

      if (dadosConvertidos && dadosConvertidos.length > 0) {
        // 🔥 NORMALIZAÇÃO ULTRA INTELIGENTE: Remove espaços, acentos e maiúsculas das colunas para não dar erro
        const dadosNormalizados = dadosConvertidos.map((linha) => {
          const novaLinha = {};
          Object.keys(linha).forEach((chave) => {
            const chaveLimpa = chave
              .toLowerCase()
              .trim()
              .normalize("NFD")
              .replace(/[\u0300-\u036f]/g, ""); // Remove acentos (Ex: mês -> mes)
            novaLinha[chaveLimpa] = linha[chave];
          });
          return novaLinha;
        });

        // Filtra as linhas da operação 'motos'
        dadosDashboardGlobais = dadosNormalizados.filter(
          (item) =>
            item.operacao && item.operacao.toString().toLowerCase() === "motos"
        );

        if (dadosDashboardGlobais.length > 0) {
          // Salva na memória do navegador
          localStorage.setItem(
            "dadosDashboardMotos",
            JSON.stringify(dadosDashboardGlobais)
          );

          alert(
            "Sucesso! Planilha Excel importada e sincronizada perfeitamente."
          );
          configurarSeletorDatas(dadosDashboardGlobais);
          renderizarDashboard();
        } else {
          alert(
            "A planilha foi lida, mas nenhuma linha com a coluna 'operacao' contendo 'motos' foi encontrada. Verifique se escreveu certo no Excel."
          );
        }
      } else {
        alert("Erro: Planilha vazia.");
      }
    } catch (erro) {
      console.error(erro);
      alert(
        "Erro ao ler o arquivo Excel. Certifique-se de que é um arquivo .xlsx válido."
      );
    }
    evento.target.value = "";
  };
}

// =========================================================================
// 3. SELETOR DE DATAS CORPORATIVO
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
// 4. RENDERIZAÇÃO E SOMA DOS CARDS
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
    totalBoletosEmitidos += parseInt(item.qtdemitidos || 0);
    totalBoletosRecebidos += parseInt(item.qtdrecebidos || 0);

    totalRoyaltiesEmitido += parseFloat(item.royalties || 0);
    totalRoyaltiesRecebido += parseFloat(item.royaltiesrec || 0);
    totalSplitsEmitido += parseFloat(item.splits || 0);
    totalSplitsRecebido += parseFloat(item.splitsrec || 0);
    totalGestaoEmitido += parseFloat(item.gestao || 0);
    totalGestaoRecebido += parseFloat(item.gestaorec || 0);
    totalMktEmitido += parseFloat(item.mkt || 0);
    totalMktRecebido += parseFloat(item.mktrec || 0);
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

// ... Mantém a função atualizarLinhaTabela igual lá embaixo ...
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