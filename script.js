let dadosDashboardGlobais = [];

document.addEventListener("DOMContentLoaded", () => {
  // Tenta carregar o que estiver salvo
  try {
    const salvos = localStorage.getItem("dadosDashboardMotos");
    if (salvos) {
      dadosDashboardGlobais = JSON.parse(salvos);
      configurarSeletorDatas(dadosDashboardGlobais);
      renderizarDashboard();
    }
  } catch (e) {}

  const btn = document.getElementById("btn-importar");
  if (btn) btn.addEventListener("change", lerExcelSemErros);

  const filtro = document.getElementById("filtro-data");
  if (filtro) filtro.addEventListener("change", renderizarDashboard);
});

function lerExcelSemErros(e) {
  const senha = prompt("Digite a senha:");
  if (senha !== "12345678") {
    alert("Senha incorreta!");
    return;
  }

  const arq = e.target.files[0];
  if (!arq) return;

  const leitor = new FileReader();
  leitor.readAsArrayBuffer(arq);
  leitor.onload = function (evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const aba = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(aba, { defval: "" });

      // Força a conversão limpando nomes de colunas
      const limpos = json.map((linha) => {
        const n = {};
        Object.keys(linha).forEach((k) => {
          const chaveLimpa = k
            .toLowerCase()
            .trim()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "");
          n[chaveLimpa] = linha[k];
        });
        return n;
      });

      // SALVA DIRETO SEM FAZER FILTROS CHATOS
      dadosDashboardGlobais = limpos;
      localStorage.setItem(
        "dadosDashboardMotos",
        JSON.stringify(dadosDashboardGlobais)
      );

      // Atualiza a tela na marra
      configurarSeletorDatas(dadosDashboardGlobais);
      renderizarDashboard();

      // Recarrega a página sozinho para forçar o navegador a acordar
      setTimeout(() => {
        location.reload();
      }, 500);
    } catch (err) {
      alert("Erro ao ler o arquivo.");
    }
  };
}

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

  const listaMeses = [...new Set(dados.map((i) => i.mes).filter(Boolean))];
  listaMeses.forEach((m) => {
    const partes = m.toString().split("/");
    const mesNome = (partes[0] || "Maio").toLowerCase().trim();
    const ano = partes[1] || "2025";
    const dias = mesesMap[mesNome] || 31;

    let arrDias = [];
    for (let i = 1; i <= dias; i++) arrDias.push(i);

    const opt = document.createElement("option");
    opt.value = m;
    opt.textContent = `${ano} - ${mesNome} ${arrDias.join(",")}`;
    seletor.appendChild(opt);
  });
}

function renderizarDashboard() {
  const seletor = document.getElementById("filtro-data");
  const sel = seletor ? seletor.value : "todas";

  let filtrados = dadosDashboardGlobais;
  if (sel !== "todas") {
    filtrados = dadosDashboardGlobais.filter((i) => i.mes == sel);
  }

  let tEmitido = 0;
  let tRecebido = 0;
  let qEmit = 0;
  let qRec = 0;
  let rEmit = 0;
  let rRec = 0;
  let sEmit = 0;
  let sRec = 0;
  let gEmit = 0;
  let gRec = 0;
  let mEmit = 0;
  let mRec = 0;

  filtrados.forEach((i) => {
    tEmitido += parseFloat(i.emitidos || 0);
    tRecebido += parseFloat(i.recebidos || 0);

    // 🔥 CORREÇÃO AQUI: Aceita tanto tudo minúsculo quanto com letra maiúscula do seu Excel
    qEmit += parseInt(i.qtdemitidos || i.qtdEmitidos || 0);
    qRec += parseInt(i.qtdrecebidos || i.qtdRecebidos || 0);

    rEmit += parseFloat(i.royalties || 0);
    rRec += parseFloat(i.royaltiesrec || i.royaltiesRec || 0);
    sEmit += parseFloat(i.splits || 0);
    sRec += parseFloat(i.splitsrec || i.splitsRec || 0);
    gEmit += parseFloat(i.gestao || 0);
    gRec += parseFloat(i.gestaorec || i.gestaoRec || 0);
    mEmit += parseFloat(i.mkt || 0);
    mRec += parseFloat(i.mktrec || i.mktRec || 0);
  });

  const setTxt = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.textContent = val;
  };

  setTxt(
    "motos-rec-bruta",
    "R$ " + tRecebido.toLocaleString("pt-BR", { minimumFractionDigits: 2 })
  );
  setTxt(
    "motos-rec-liquida",
    "R$ " + tRecebido.toLocaleString("pt-BR", { minimumFractionDigits: 2 })
  );
  setTxt(
    "motos-emitidos",
    "R$ " + tEmitido.toLocaleString("pt-BR", { minimumFractionDigits: 2 })
  );

  let aberto = tEmitido - tRecebido;
  setTxt(
    "motos-em-aberto",
    "R$ " +
      (aberto > 0 ? aberto : 0).toLocaleString("pt-BR", {
        minimumFractionDigits: 2
      })
  );
  setTxt("motos-qtd-emitidos", qEmit.toLocaleString("pt-BR") + " un");
  setTxt("motos-qtd-recebidos", qRec.toLocaleString("pt-BR") + " un");

  const setLinha = (id, em, rec) => {
    setTxt(
      `${id}-emitido`,
      "R$ " + em.toLocaleString("pt-BR", { minimumFractionDigits: 2 })
    );
    setTxt(
      `${id}-recebido`,
      "R$ " + rec.toLocaleString("pt-BR", { minimumFractionDigits: 2 })
    );
    let inad = em > 0 ? ((em - rec) / em) * 100 : 0;
    setTxt(`${id}-inad`, (inad > 0 ? inad.toFixed(1) : "0.0") + " %");
  };

  setLinha("motos-t1", rEmit, rRec);
  setLinha("motos-t2", sEmit, sRec);
  setLinha("motos-t3", gEmit, gRec);
  setLinha("motos-t4", mEmit, mRec);
}