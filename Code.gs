/**
 * Apps Script vinculado à planilha "Iscas Travessia - captura".
 * Uma planilha só tem UM script vinculado — este arquivo é o ÚNICO Code.gs
 * publicado, e atende Quiz, Checklist, Calculadora e Sorteio ao mesmo tempo,
 * roteando pelo campo "formulario" que cada página envia.
 *
 * DEPLOY (atualizar o que já existe, não criar um novo):
 * 1. Abra a planilha "Iscas Travessia - captura" → Extensões → Apps Script.
 * 2. Apague o conteúdo de Code.gs e cole este arquivo inteiro.
 * 3. Deploy → Gerenciar implantações → ícone de lápis na implantação ativa
 *    → Versão: "Nova versão" → Implantar.
 *    (Isso mantém a MESMA URL que quiz, checklist, calculadora e sorteio já usam —
 *    não precisa mexer no GOOGLE_SCRIPT_URL de nenhuma das páginas.)
 *
 * Esse mesmo arquivo está duplicado nas pastas Quiz Diagnóstico/, Checklist/,
 * Calculadora/ e Sorteio Live/ só pra ficar fácil de achar. Se editar, atualize os 4.
 *
 * UTM: cada página captura utm_source/medium/campaign/content/term da URL
 * (query string) e manda junto no POST. As 5 colunas de UTM são adicionadas
 * no fim de cada aba automaticamente, inclusive em abas que já tinham dados
 * antes (o cabeçalho se atualiza sozinho a cada envio).
 */

const UTM_HEADERS = ['UTM Source', 'UTM Medium', 'UTM Campaign', 'UTM Content', 'UTM Term'];

function doPost(e) {
  const data = e.parameter;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const utm = [
    data.utm_source || '', data.utm_medium || '', data.utm_campaign || '',
    data.utm_content || '', data.utm_term || ''
  ];

  if (data.formulario === 'checklist') {
    appendRow(ss, 'Checklist',
      ['Data/Hora', 'Nome', 'E-mail', 'Telefone', 'Faixa de renda'].concat(UTM_HEADERS),
      [new Date(), data.nome || '', data.email || '', data.telefone || '', data.renda || ''].concat(utm)
    );
  } else if (data.formulario === 'calculadora') {
    appendRow(ss, 'Calculadora',
      [
        'Data/Hora', 'Nome', 'WhatsApp', 'E-mail',
        'Idade atual', 'Idade quer aposentar', 'Renda mensal', 'Patrimônio atual',
        'Patrimônio necessário (nominal)', 'Patrimônio na idade-alvo (nominal)',
        'Idade que bate a meta', 'Diferença de anos (positivo = antecipa, negativo = atrasa)'
      ].concat(UTM_HEADERS),
      [
        new Date(), data.nome || '', data.whatsapp || '', data.email || '',
        data.idadeAtual || '', data.idadeAposentar || '', data.rendaMensal || '', data.patrimonioAtual || '',
        data.patrimonioNecessario || '', data.patrimonioNaIdadeAlvo || '',
        data.idadeBateMeta || '', data.diferencaAnos || ''
      ].concat(utm)
    );
  } else if (data.formulario === 'sorteio') {
    appendRow(ss, 'Sorteio',
      [
        'Data/Hora', 'Nome', 'WhatsApp', 'E-mail', 'Faixa de renda',
        'Maior dor hoje', 'Situação atual', 'Situação atual (Outro)',
        'O que já tentou / por que não funcionou', 'Como se sente',
        'Objetivo principal', 'Objetivo principal (Outro)',
        'Sabe quando vai alcançar sem contas', 'O que impede de organizar',
        'O que faria considerar entrar', 'Teria interesse',
        'Frustração daqui 1 ano', 'Pergunta sobre dinheiro'
      ].concat(UTM_HEADERS),
      [
        new Date(), data.nome || '', data.whatsapp || '', data.email || '', data.renda || '',
        data.maiorDor || '', data.situacao || '', data.situacaoOutro || '',
        data.jaTentou || '', data.comoSente || '',
        data.objetivo || '', data.objetivoOutro || '',
        data.sabeQuando || '', data.oQueImpede || '',
        data.oQueFaria || '', data.teriaInteresse || '',
        data.frustracao1Ano || '', data.perguntaDinheiro || ''
      ].concat(utm)
    );
  } else {
    // Comportamento original do Quiz "Descubra por que seu dinheiro não vira patrimônio"
    appendRow(ss, 'Quiz diagnóstico',
      ['Data/Hora', 'Nome', 'WhatsApp', 'E-mail', 'Renda', 'Perfil de vazamento'].concat(UTM_HEADERS),
      [new Date(), data.nome || '', data.whatsapp || '', data.email || '', data.renda || '', data.perfil || ''].concat(utm)
    );
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function appendRow(ss, sheetName, headers, row) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  // Reescreve a linha de cabeçalho sempre — idempotente, e "cura" sozinho
  // abas antigas que ainda não tinham as colunas de UTM.
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.appendRow(row);
}
