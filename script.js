// FUNÇÃO PARA LER O EXCEL E SALVAR NO NAVEGADOR (COM PROTEÇÃO POR SENHA)
function lerPlanilhaExcel(evento) {
  // 🔒 DEFINA A SUA SENHA AQUI (Troque '1234' pela senha que você quiser)
  const senhaCorreta = "1234";

  // Abre a caixinha na tela pedindo a senha
  const senhaDigitada = prompt(
    "Digite a senha de administrador para importar novos dados:"
  );

  // Se o usuário cancelar ou digitar errado, o sistema bloqueia
  if (senhaDigitada === null) {
    evento.target.value = ""; // Limpa o botão de upload
    return; // Cancela a operação
  }

  if (senhaDigitada !== senhaCorreta) {
    alert("Senha incorreta! Você não tem permissão para alterar o Dashboard.");
    evento.target.value = ""; // Limpa o botão de upload
    return; // Cancela a operação e não lê o arquivo
  }

  // 🔓 SE A SENHA ESTIVER CORRETA, O CÓDIGO CONTINUA NORMALMENTE:
  const arquivo = evento.target.files[0];
  if (!arquivo) return;

  const leitor = new FileReader();
  leitor.onload = function (e) {
    // ... (resto do seu código atual que processa o Excel e salva no localStorage) ...
    const dadosAgrupados = processarDadosExcel(e.target.result);
    if (dadosAgrupados && dadosAgrupados.length > 0) {
      dadosDashboardGlobais = dadosAgrupados;
      localStorage.setItem(
        "dadosDashboardMotos",
        JSON.stringify(dadosAgrupados)
      );
      alert("Senha confirmada. Planilha importada e salva com sucesso!");
      configurarSeletorDatas(dadosDashboardGlobais);
      renderizarDashboard();
    }
  };
  leitor.readAsText(arquivo);
}