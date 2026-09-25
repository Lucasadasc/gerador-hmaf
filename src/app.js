const DIAS_SEMANA = ['DOMINGO', 'SEGUNDA-FEIRA', 'TERÇA-FEIRA', 'QUARTA-FEIRA', 'QUINTA-FEIRA', 'SEXTA-FEIRA', 'SÁBADO'];
const arquivo = document.querySelector('#arquivo');
const status = document.querySelector('#status');
const botoes = [...document.querySelectorAll('button')];

function nomeCurto(valor) {
  const partes = String(valor).trim().split(/\s+/);
  if (partes.length <= 1) return partes[0] || '';
  return `${partes[0]} ${partes[1]}${partes[2] && !['de', 'do', 'da', 'dos', 'das'].includes(partes[2].toLowerCase()) ? ` ${partes[2]}` : ''}`;
}

function identificarMesAno(dados) {
  for (const linha of dados.slice(0, 6)) for (const celula of linha || []) {
    const resultado = String(celula || '').match(/(\d{1,2})[/-](\d{4})/);
    if (resultado) return { mes: Number(resultado[1]), ano: Number(resultado[2]) };
  }
  const hoje = new Date();
  return { mes: hoje.getMonth() + 1, ano: hoje.getFullYear() };
}

function habilitarDownloads(habilitado) {
  botoes.forEach(botao => botao.disabled = !habilitado);
}

async function validarArquivo() {
  const selecionado = arquivo.files[0];
  habilitarDownloads(false);
  status.className = '';
  if (!selecionado) { status.textContent = ''; return; }
  if (!/\.xlsx?$/i.test(selecionado.name)) {
    status.className = 'erro'; status.textContent = 'Escolha um arquivo Excel (.xlsx ou .xls).'; return;
  }
  try {
    const livro = XLSX.read(await selecionado.arrayBuffer(), { type: 'array' });
    if (!livro.Sheets.ESCALA) throw new Error();
    habilitarDownloads(true); status.className = 'ok'; status.textContent = 'Planilha válida. Escolha o formato para baixar.';
  } catch (_) {
    status.className = 'erro'; status.textContent = 'A planilha precisa ter uma aba chamada "ESCALA".';
  }
}

async function lerEscala() {
  const selecionado = arquivo.files[0];
  if (!selecionado) throw new Error('Selecione uma planilha antes de continuar.');
  const livro = XLSX.read(await selecionado.arrayBuffer(), { type: 'array' });
  const aba = livro.Sheets.ESCALA;
  if (!aba) throw new Error('Não encontrei a aba "ESCALA" na planilha selecionada.');
  const dados = XLSX.utils.sheet_to_json(aba, { header: 1, defval: null });
  const { mes, ano } = identificarMesAno(dados);
  const escala = Array.from({ length: new Date(ano, mes, 0).getDate() }, () => ({ manha: [], tarde: [], noite: [] }));
  const codigos = {
    M: [[ 'manha', false ]], D: [[ 'manha', false ], [ 'tarde', false ]], P: [[ 'manha', false ], [ 'tarde', false ], [ 'noite', false ]],
    N: [[ 'noite', false ]], J: [[ 'manha', true ], [ 'tarde', true ]], K: [[ 'noite', true ]],
    'D/K': [[ 'manha', false ], [ 'tarde', false ], [ 'noite', true ]], 'J/N': [[ 'manha', true ], [ 'tarde', true ], [ 'noite', false ]]
  };
  for (const linha of dados.slice(5)) {
    if (!linha[0]) break;
    const nome = nomeCurto(linha[0]); if (nome === 'JUSTIS') continue;
    escala.forEach((dia, indice) => (codigos[String(linha[6 + indice] || '').trim().toUpperCase()] || []).forEach(([turno, eventual]) => dia[turno].push(eventual ? `${nome} (EVE)` : nome)));
  }
  return { escala, mes, ano };
}

function baixarExcel({ escala, mes, ano }) {
  const aba = XLSX.utils.aoa_to_sheet([]); aba['!cols'] = [{ wch: 10 }, { wch: 40 }, { wch: 10 }, { wch: 40 }, { wch: 10 }, { wch: 40 }]; aba['!merges'] = [];
  const escrever = (linha, coluna, valor, estilo = {}) => aba[XLSX.utils.encode_cell({ r: linha - 1, c: coluna - 1 })] = { t: 's', v: valor, s: estilo };
  let linha = 1, linhaInicial = 1;
  escala.forEach((turnos, indice) => {
    const data = new Date(ano, mes - 1, indice + 1), titulo = `${DIAS_SEMANA[data.getDay()]} - ${String(indice + 1).padStart(2, '0')}/${String(mes).padStart(2, '0')}/${ano}`;
    escrever(linha, 1, titulo, { font: { bold: true, color: { rgb: 'FF0000' } }, alignment: { horizontal: 'center' } }); aba['!merges'].push({ s: { r: linha - 1, c: 0 }, e: { r: linha - 1, c: 5 } }); linha++;
    [['Manhã', 1], ['Tarde', 3], ['Noite', 5]].forEach(([tituloTurno, coluna]) => { escrever(linha, coluna, tituloTurno, { font: { bold: true }, alignment: { horizontal: 'center' } }); aba['!merges'].push({ s: { r: linha - 1, c: coluna - 1 }, e: { r: linha - 1, c: coluna } }); }); linha++;
    const maior = Math.max(turnos.manha.length, turnos.tarde.length, turnos.noite.length);
    for (let posicao = 0; posicao < maior; posicao++) { [['manha', 2], ['tarde', 4], ['noite', 6]].forEach(([turno, coluna]) => { const pessoa = turnos[turno][posicao]; if (pessoa) escrever(linha, coluna, pessoa.replace(' (EVE)', ''), { font: { italic: pessoa.endsWith(' (EVE)') } }); }); linha += 2; }
    // O Python reserva 16 linhas por dia; preservamos esse espaçamento no Excel.
    linha = linhaInicial + 16; linhaInicial = linha;
  });
  aba['!ref'] = `A1:F${linha}`; const livro = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(livro, aba, 'Escala formatada'); XLSX.writeFile(livro, `escala_formatada_${ano}-${String(mes).padStart(2, '0')}.xlsx`);
}

// Medidas em twips: mesma proporção 10/40 usada pelo gerador Python.
const LARGURAS_COLUNAS = [695, 2780, 695, 2780, 695, 2780];
const BORDA = { style: docx.BorderStyle.SINGLE, size: 4, color: '000000' };
const BORDAS_TABELA = { top: BORDA, bottom: BORDA, left: BORDA, right: BORDA, insideHorizontal: BORDA, insideVertical: BORDA };

function paragrafo(texto, opcoes = {}) {
  return new docx.Paragraph({ alignment: opcoes.centro ? docx.AlignmentType.CENTER : undefined, spacing: { before: 0, after: 0, line: 240 }, children: [new docx.TextRun({ text: texto || '', bold: opcoes.negrito, italics: opcoes.italico, size: 20 })] });
}

function celula(texto = '', opcoes = {}) {
  return new docx.TableCell({
    columnSpan: opcoes.span,
    width: { size: opcoes.largura || 695, type: docx.WidthType.DXA },
    verticalAlign: docx.VerticalAlign.CENTER,
    margins: { top: 40, bottom: 40, left: 70, right: 70 },
    children: [paragrafo(texto, opcoes)]
  });
}

function linhaWord(celulas, altura) {
  return new docx.TableRow({ cantSplit: true, height: { value: altura, rule: docx.HeightRule.EXACTLY }, children: celulas });
}

async function baixarWord({ escala, mes, ano }) {
  const conteudo = [new docx.Paragraph({ text: 'Escala formatada', heading: docx.HeadingLevel.HEADING_1, spacing: { after: 240 } })];
  escala.forEach((turnos, indice) => {
    const data = new Date(ano, mes - 1, indice + 1), titulo = `${DIAS_SEMANA[data.getDay()]} - ${String(indice + 1).padStart(2, '0')}/${String(mes).padStart(2, '0')}/${ano}`;
    const linhas = [
      linhaWord([celula(titulo, { span: 6, largura: 10425, negrito: true, centro: true })], 397),
      linhaWord(['Manhã', 'Tarde', 'Noite'].map(tituloTurno => celula(tituloTurno, { span: 2, largura: 3475, negrito: true, centro: true })), 340)
    ];
    const maior = Math.max(turnos.manha.length, turnos.tarde.length, turnos.noite.length);
    for (let posicao = 0; posicao < maior; posicao++) {
      const celulas = []; ['manha', 'tarde', 'noite'].forEach((turno, coluna) => { const pessoa = turnos[turno][posicao] || '', inicio = coluna * 2; celulas.push(celula('', { largura: LARGURAS_COLUNAS[inicio] }), celula(pessoa.replace(' (EVE)', ''), { largura: LARGURAS_COLUNAS[inicio + 1], italico: pessoa.endsWith(' (EVE)') })); });
      linhas.push(linhaWord(celulas, 312), linhaWord(LARGURAS_COLUNAS.map(largura => celula('', { largura })), 312));
    }
    conteudo.push(new docx.Table({ rows: linhas, width: { size: 10425, type: docx.WidthType.DXA }, layout: docx.TableLayoutType.FIXED, borders: BORDAS_TABELA }), new docx.Paragraph({ text: '', spacing: { after: 120 } }));
  });
  const documento = new docx.Document({ sections: [{ properties: { page: { margin: { top: 850, right: 740, bottom: 850, left: 740 } } }, children: conteudo }] });
  const blob = await docx.Packer.toBlob(documento), url = URL.createObjectURL(blob), link = Object.assign(document.createElement('a'), { href: url, download: `escala_formatada_${ano}-${String(mes).padStart(2, '0')}.docx` }); link.click(); setTimeout(() => URL.revokeObjectURL(url), 500);
}

async function processar(formato) {
  try { botoes.forEach(botao => botao.disabled = true); status.className = ''; status.textContent = 'Formatando…'; const dados = await lerEscala(); if (formato === 'excel') baixarExcel(dados); else await baixarWord(dados); status.className = 'ok'; status.textContent = `Download de ${formato === 'excel' ? 'Excel' : 'Word'} iniciado.`; }
  catch (erro) { status.className = 'erro'; status.textContent = erro.message || 'Não foi possível gerar o arquivo.'; }
  finally { botoes.forEach(botao => botao.disabled = false); }
}
document.querySelector('#gerarExcel').addEventListener('click', () => processar('excel'));
document.querySelector('#gerarDocx').addEventListener('click', () => processar('word'));
arquivo.addEventListener('change', validarArquivo);
