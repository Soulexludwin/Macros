function onEdit(e) {
  // Evita erros 
  if (!e || !e.range) return;
  
  // encontra aba de destino com o nome de preferencia
  const nomeAbaHistorico = "XXXXXXX";
  const planilha = e.source;
  const abaAtual = planilha.getActiveSheet();
  const nomeAbaAtual = abaAtual.getName();
  
  // Impede que o script registre alterações feitas na própria aba de histórico
  if (nomeAbaAtual === nomeAbaHistorico) return;
  
  // encontrar ou 
  let abaHistorico = planilha.getSheetByName(nomeAbaHistorico);
  if (!abaHistorico) {
    //criar a aba de histórico
    abaHistorico = planilha.insertSheet(nomeAbaHistorico);
  }
  
  // modificação dos dados adicione uma nova const com um novo requerimento se precisar
  const dataHora = new Date();
  const celula = e.range.getA1Notation();
  const valorAntigo = e.oldValue !== undefined ? e.oldValue : "";
  const novoValor = e.value !== undefined ? e.value : "";
  const usuario = Session.getActiveUser().getEmail() || "Não identificado";
  const horas = Utilities.formatDate(new Date(), "GMT-3", "HH:mm:ss");
  
  // Monta os dados na coluna J (altere o "J" por qualquer coluna que você bem entender)
  const dadosColunaJ = [
    [dataHora],      // J4
    [nomeAbaAtual],  // J5
    [celula],        // J6
    [valorAntigo],   // J7
    [novoValor],     // J8
    [usuario],       // J9
    [horas]             // J10
    // Você que decide o tamanho ou proporção.
  ];
  
  // Atualiza diretamente o intervalo sobrescrevendo as informações antigas (J4:J10)
  abaHistorico.getRange("J4:J10").setValues(dadosColunaJ);
}