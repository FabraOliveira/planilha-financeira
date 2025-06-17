// ▪️ Funções por mês
function DistribuirJaneiro() { DistribuirParaMes("Janeiro", 0); }
function DistribuirFevereiro() { DistribuirParaMes("Fevereiro", 1); }
function DistribuirMarco() { DistribuirParaMes("Março", 2); }
function DistribuirAbril() { DistribuirParaMes("Abril", 3); }
function DistribuirMaio() { DistribuirParaMes("Maio", 4); }
function DistribuirJunho() { DistribuirParaMes("Junho", 5); }
function DistribuirJulho() { DistribuirParaMes("Julho", 6); }
function DistribuirAgosto() { DistribuirParaMes("Agosto", 7); }
function DistribuirSetembro() { DistribuirParaMes("Setembro", 8); }
function DistribuirOutubro() { DistribuirParaMes("Outubro", 9); }
function DistribuirNovembro() { DistribuirParaMes("Novembro", 10); }
function DistribuirDezembro() { DistribuirParaMes("Dezembro", 11); }

function DistribuirParaMes(nomeMes, indiceMes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOrigem = ss.getSheetByName(nomeMes);
  if (!sheetOrigem) return;

  const meses = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
  ];

  const linhaInicial = 60;
  const ultimaLinha = sheetOrigem.getRange("E:E").getLastRow();
  const totalLinhas = Math.max(ultimaLinha, linhaInicial) - linhaInicial + 1;
  const dadosOrigem = sheetOrigem.getRange(linhaInicial, 3, totalLinhas, 7).getValues(); // C60:I

  const buffersDestino = {}; // { nomeMes: { linhas: [], sheet: Sheet } }

  for (let i = 0; i < dadosOrigem.length; i++) {
    const linha = dadosOrigem[i];
    const produto = linha[2];        // Coluna E
    const parcelaAtual = linha[5];   // Coluna H
    const totalParcelas = linha[6];  // Coluna I

    if (!produto || !Number.isInteger(parcelaAtual) || !Number.isInteger(totalParcelas)) continue;
    if (totalParcelas <= 1 || parcelaAtual >= totalParcelas) continue;

    for (let p = 1; p <= totalParcelas - parcelaAtual; p++) {
      const destinoIndex = indiceMes + p;
      if (destinoIndex >= meses.length) break;

      const nomeDestino = meses[destinoIndex];
      const sheetDestino = ss.getSheetByName(nomeDestino);
      if (!sheetDestino) continue;

      if (!buffersDestino[nomeDestino]) {
        const dadosDestino = sheetDestino.getRange("C60:I" + sheetDestino.getLastRow()).getValues();
        const existentes = new Set(
          dadosDestino
            .filter(l => l[2] && Number.isInteger(l[5]))
            .map(l => `${l[2]}|${l[5]}`)
        );
        buffersDestino[nomeDestino] = {
          sheet: sheetDestino,
          existentes: existentes,
          linhasParaAdicionar: []
        };
      }

      const novaParcela = parcelaAtual + p;
      const chave = `${produto}|${novaParcela}`;
      const buffer = buffersDestino[nomeDestino];

      if (buffer.existentes.has(chave)) continue;

      const novaLinha = [...linha];
      novaLinha[5] = novaParcela; // Atualiza a parcela atual
      buffer.linhasParaAdicionar.push(novaLinha);
      buffer.existentes.add(chave);
    }
  }

  // ✅ Escreve todas as linhas acumuladas por aba (com verificação de segurança)
  for (let nome in buffersDestino) {
    const buffer = buffersDestino[nome];
    const sheet = buffer.sheet;
    const destino = sheet.getRange("E60:E1059").getValues();

    let base = 60;
    for (let i = 0; i < destino.length; i++) {
      if (!destino[i][0]) {
        base = 60 + i;
        break;
      }
    }

    if (buffer.linhasParaAdicionar.length > 0) {
      const linhas = buffer.linhasParaAdicionar.map(l => l.slice(0, 7));
      sheet.getRange(base, 3, linhas.length, 7).setValues(linhas);
    }
  }
}
