// =============================================
// BERNHOEFT — Gerador de Importação
// Ferramenta interna: XLS → JSON/CSV
// =============================================

// ── Estado ───────────────────────────────────
var STATE = {
  tipo: null,
  dados: null,
  outputStr: null,
  outputFmt: null,
};

// ── Tipos disponíveis ─────────────────────────
var TIPO_LABELS = {
  estagiario: '🎓 Estagiário',
  clt:        '💼 CLT',
  autonomo:   '🤝 Autônomo',
};

// =============================================
// 1. SELEÇÃO DE TIPO
// =============================================

function selecionarTipo(tipo) {
  STATE.tipo = tipo;

  // Visual: marcar botão selecionado
  ['estagiario', 'clt', 'autonomo'].forEach(function(t) {
    var btn = document.getElementById('btn-' + t);
    var chk = document.getElementById('check-' + t);
    if (t === tipo) {
      btn.classList.add('selected');
    } else {
      btn.classList.remove('selected');
    }
  });

  // Mostrar badge no card de upload
  var badge = document.getElementById('tipo-selected-badge');
  badge.textContent = TIPO_LABELS[tipo] || tipo;
  badge.classList.remove('hidden');

  // Atualizar hint de upload
  document.querySelector('.upload-hint').style.display = 'none';

  toast('Tipo selecionado: ' + (TIPO_LABELS[tipo] || tipo), 'ok');
}

// =============================================
// 2. UPLOAD / DRAG-AND-DROP
// =============================================

function triggerFileInput() {
  if (!STATE.tipo) {
    toast('⚠️ Selecione o tipo de contrato primeiro!', 'warn');
    document.getElementById('card-tipo').scrollIntoView({ behavior: 'smooth', block: 'center' });
    return;
  }
  document.getElementById('file-input').click();
}

function onDragOver(e) {
  e.preventDefault();
  document.getElementById('dropzone').classList.add('hover');
}

function onDragLeave(e) {
  document.getElementById('dropzone').classList.remove('hover');
}

function onDrop(e) {
  e.preventDefault();
  document.getElementById('dropzone').classList.remove('hover');
  var file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
  if (file) processarArquivo(file);
}

function onFileSelect(e) {
  var file = e.target.files && e.target.files[0];
  if (file) processarArquivo(file);
}

function processarArquivo(file) {
  if (!STATE.tipo) {
    toast('⚠️ Selecione o tipo de contrato primeiro!', 'warn');
    return;
  }

  var ext = file.name.split('.').pop().toLowerCase();
  if (ext !== 'xls' && ext !== 'xlsx') {
    toast('❌ Formato inválido. Use .xls ou .xlsx', 'err');
    return;
  }

  var dz = document.getElementById('dropzone');
  dz.querySelector('.dz-title').textContent = 'Lendo arquivo...';
  dz.querySelector('.dz-icon').textContent = '⏳';

  var reader = new FileReader();
  reader.onload = function(ev) {
    try {
      var dados = lerPlanilha(ev.target.result);
      if (!dados) {
        toast('❌ Não foi possível ler a planilha. Verifique o formato.', 'err');
        resetDropzone();
        return;
      }

      for (var k_d in dados) {
        if (typeof dados[k_d] === 'string' && k_d.toLowerCase().indexOf('email') === -1) {
          dados[k_d] = dados[k_d].toUpperCase();
        }
      }

      STATE.dados = dados;

      // Atualizar dropzone com sucesso
      dz.classList.add('loaded');
      dz.querySelector('.dz-icon').textContent = '✅';
      dz.querySelector('.dz-title').textContent = file.name;
      dz.querySelector('.dz-sub').textContent = (file.size / 1024).toFixed(1) + ' KB · Importado com sucesso!';
      var dzState = document.getElementById('dz-state');
      dzState.textContent = TIPO_LABELS[STATE.tipo];
      dzState.classList.remove('hidden');

      mostrarPreview(dados);

    } catch (err) {
      console.error('[processarArquivo]', err);
      toast('❌ Erro ao processar: ' + err.message, 'err');
      resetDropzone();
    }
  };
  reader.onerror = function() {
    toast('❌ Erro ao ler o arquivo.', 'err');
    resetDropzone();
  };
  reader.readAsArrayBuffer(file);
}

function resetDropzone() {
  var dz = document.getElementById('dropzone');
  dz.classList.remove('loaded', 'hover');
  dz.querySelector('.dz-icon').textContent = '📂';
  dz.querySelector('.dz-title').textContent = 'Arraste a planilha aqui';
  dz.querySelector('.dz-sub').textContent = 'ou clique para selecionar · .xls / .xlsx';
  document.getElementById('dz-state').classList.add('hidden');
  document.getElementById('file-input').value = '';
}

// =============================================
// 3. LEITURA DA PLANILHA (SheetJS)
// =============================================

// Mapeamento flexível: campo interno → possíveis cabeçalhos no XLS
var COLMAP = {
  empresa:          ['empresa', 'razão social', 'razao social', 'nome da empresa'],
  contato:          ['contato', 'pessoa de contato', 'responsável', 'responsavel'],
  telefone_emp:     ['telefone empresa', 'fone empresa', 'tel empresa', 'tel. empresa'],
  email_emp:        ['e-mail empresa', 'email empresa', 'email da empresa'],
  nome:             ['nome completo', 'nome', 'colaborador'],
  nascimento:       ['data de nascimento', 'nascimento', 'dt nascimento', 'dt. nascimento'],
  naturalidade:     ['naturalidade', 'cidade/uf', 'cidade uf'],
  raca:             ['raça', 'raca', 'cor/raça', 'etnia'],
  estado_civil:     ['estado civil', 'est civil', 'est. civil'],
  pai:              ['filiação pai', 'pai', 'nome do pai', 'nome pai'],
  mae:              ['filiação mãe', 'mae', 'mãe', 'nome da mãe', 'nome mae'],
  grau:             ['grau de instrução', 'instrução', 'escolaridade', 'grau instrucao'],
  deficiencia:      ['deficiência', 'deficiencia', 'portador deficiência', 'portador de deficiência'],
  deficiencia_qual: ['qual deficiência', 'qual deficiencia', 'tipo deficiência'],
  habilitacao:      ['cnh', 'habilitação', 'habilitacao'],
  seguro:           ['seguro desemprego', 'recebe seguro'],
  primeiro_emp:     ['primeiro emprego'],
  rg:               ['rg', 'identidade'],
  rg_expedicao:     ['emissão rg', 'emissao rg', 'data rg', 'data emissão rg', 'data emissao rg'],
  rg_orgao:         ['órgão expedidor', 'orgao expedidor', 'expedidor'],
  cpf:              ['cpf'],
  ctps_num:         ['ctps número', 'ctps num', 'ctps nº', 'ctps numero', 'ctps'],
  ctps_serie:       ['ctps série', 'ctps serie'],
  ctps_data:        ['ctps data', 'data ctps', 'expedição ctps', 'expedicao ctps'],
  ctps_local:       ['ctps local', 'ctps uf', 'local ctps'],
  titulo:           ['título eleitor', 'titulo eleitor', 'título de eleitor'],
  pis:              ['pis', 'pis/nit', 'nit'],
  reservista:       ['reservista', 'certificado reservista'],
  certidao_num:     ['certidão nº', 'certidao num', 'nº certidão', 'certidao numero'],
  certidao_livro:   ['certidão livro', 'certidao livro', 'livro'],
  certidao_folha:   ['certidão folha', 'certidao folha', 'folha'],
  certidao_data:    ['certidão data', 'certidao data', 'data certidão'],
  certidao_cartorio:['certidão cartório', 'certidao cartorio', 'cartório', 'cartorio'],
  endereco:         ['endereço', 'endereco', 'endereço completo', 'endereco completo'],
  bairro:           ['bairro'],
  cidade:           ['cidade'],
  cep:              ['cep'],
  fone:             ['telefone', 'fone', 'telefone pessoal', 'tel pessoal'],
  celular:          ['celular', 'telefone celular', 'cel', 'cel.', 'whatsapp'],
  email:            ['e-mail', 'email', 'e-mail pessoal', 'email pessoal'],
  conjuge:          ['cônjuge', 'conjuge', 'nome cônjuge', 'nome conjuge'],
  conjuge_nasc:     ['cônjuge nascimento', 'nasc cônjuge', 'nasc conjuge'],
  conjuge_cpf:      ['cônjuge cpf', 'cpf cônjuge', 'cpf conjuge'],
  conjuge_ir:       ['cônjuge ir', 'dep ir cônjuge'],
  admissao:         ['data admissão', 'admissão', 'data de admissão', 'dt admissão', 'dt. admissão', 'data admissao', 'admissao'],
  salario:          ['salário', 'salario', 'remuneração', 'remuneracao', 'bolsa'],
  funcao:           ['função', 'funcao', 'cargo', 'curso'],
  departamento:     ['departamento', 'setor'],
  filial:           ['filial', 'unidade'],
  contrato_exp:     ['contrato experiência', 'experiência', 'contrato experiencia'],
  dias_exp:         ['dias experiência', 'dias exp', 'dias experiencia'],
  horario:          ['horário', 'horario', 'jornada', 'horário trabalho', 'horario trabalho'],
  folgas:           ['folgas', 'dias folga'],
  exame:            ['exame admissional', 'exame'],
  adiantamento:     ['adiantamento'],
  sind:             ['contribuição sindical', 'sind', 'contribuicao sindical'],
  insalubridade:    ['insalubridade'],
  periculosidade:   ['periculosidade'],
  ticket:           ['ticket refeição', 'vale refeição', 'vr', 'ticket refeicao'],
  pensao:           ['pensão alimentícia', 'pensão', 'pensao'],
  obs:              ['observações', 'observacoes', 'obs'],
};

function lerPlanilha(buffer) {
  var wb = XLSX.read(buffer, { type: 'arraybuffer', cellDates: true });
  if (!wb || !wb.SheetNames || !wb.SheetNames.length) return null;

  var wsName = wb.SheetNames[0];
  var ws = wb.Sheets[wsName];
  var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });

  if (!rows || rows.length < 1) return null;

  // Encontrar linha de cabeçalho (primeira com conteúdo)
  var headerIdx = -1;
  for (var i = 0; i < Math.min(rows.length, 8); i++) {
    var row = rows[i];
    if (row && row.some(function(c) { return c && String(c).trim() !== ''; })) {
      headerIdx = i;
      break;
    }
  }

  if (headerIdx === -1) return null;

  // Encontrar linha de dados (próxima linha com conteúdo)
  var dataIdx = -1;
  for (var j = headerIdx + 1; j < Math.min(rows.length, 20); j++) {
    var r = rows[j];
    if (r && r.some(function(c) { return c && String(c).trim() !== ''; })) {
      dataIdx = j;
      break;
    }
  }

  if (dataIdx === -1) return null;

  var headers = rows[headerIdx].map(function(h) {
    return String(h || '').toLowerCase().trim().replace(/\s+/g, ' ');
  });
  var dataRow = rows[dataIdx];

  // Mapear campos via COLMAP
  var result = {};
  var campo, alternativas, val, k, alt, idx;
  for (campo in COLMAP) {
    alternativas = COLMAP[campo];
    val = '';
    for (k = 0; k < alternativas.length; k++) {
      alt = alternativas[k].toLowerCase().trim();
      idx = -1;
      for (var h = 0; h < headers.length; h++) {
        if (headers[h] === alt || headers[h].indexOf(alt) !== -1) {
          idx = h;
          break;
        }
      }
      if (idx !== -1 && dataRow[idx] !== undefined && dataRow[idx] !== null && String(dataRow[idx]).trim() !== '') {
        val = formatarValor(dataRow[idx]);
        break;
      }
    }
    result[campo] = val;
  }

  // Tentar ler filhos (até 5 linhas abaixo da linha de dados)
  var filhos = [];
  for (var fi = 1; fi <= 5; fi++) {
    var nk = 'filho ' + fi + ' nome';
    var nIdx = -1;
    for (var hh = 0; hh < headers.length; hh++) {
      if (headers[hh].indexOf('filho') !== -1 && headers[hh].indexOf('nome') !== -1) {
        nIdx = hh;
        break;
      }
    }
    if (nIdx !== -1 && dataRow[nIdx]) {
      filhos.push({
        nome: formatarValor(dataRow[nIdx]),
        nasc: '',
        cpf: '',
        ir: '',
      });
    }
  }
  result.filhos = filhos;

  return result;
}

function formatarValor(val) {
  if (val === null || val === undefined) return '';
  if (val instanceof Date) {
    var d = String(val.getDate()).padStart(2, '0');
    var m = String(val.getMonth() + 1).padStart(2, '0');
    var y = val.getFullYear();
    return d + '/' + m + '/' + y;
  }
  return String(val).trim();
}

// =============================================
// 4. PREVIEW DOS DADOS
// =============================================

var PREVIEW_SECTIONS = [
  {
    icon: '🏢', title: 'Dados da Empresa',
    campos: ['empresa', 'contato', 'telefone_emp', 'email_emp'],
    labels: { empresa: 'Empresa', contato: 'Contato', telefone_emp: 'Telefone', email_emp: 'E-mail' },
  },
  {
    icon: '👤', title: 'Dados Pessoais',
    campos: ['nome', 'nascimento', 'naturalidade', 'estado_civil', 'raca', 'etniaId', 'grau', 'pai', 'mae'],
    labels: { nome: 'Nome Completo', nascimento: 'Nascimento', naturalidade: 'Naturalidade', estado_civil: 'Estado Civil', raca: 'Raça/Cor', etniaId: 'Etnia ID', grau: 'Grau de Instrução', pai: 'Pai', mae: 'Mãe' },
  },
  {
    icon: '📄', title: 'Documentos',
    campos: ['rg', 'rg_expedicao', 'rg_orgao', 'cpf', 'pis', 'ctps_num', 'ctps_serie', 'titulo'],
    labels: { rg: 'RG', rg_expedicao: 'Emissão RG', rg_orgao: 'Órgão Expedidor', cpf: 'CPF', pis: 'PIS/NIT', ctps_num: 'CTPS Nº', ctps_serie: 'CTPS Série', titulo: 'Título Eleitor' },
  },
  {
    icon: '🏠', title: 'Endereço',
    campos: ['enderecoLogradouroId', 'endereco', 'enderecoNumero', 'bairro', 'cidade', 'enderecoCidadeId', 'cep', 'fone', 'celular', 'email'],
    labels: { enderecoLogradouroId: 'Logradouro ID', endereco: 'Endereço', enderecoNumero: 'Número', bairro: 'Bairro', cidade: 'Cidade', enderecoCidadeId: 'Cidade ID', cep: 'CEP', fone: 'Telefone', celular: 'Celular', email: 'E-mail Pessoal' },
  },
  {
    icon: '📋', title: 'Admissão',
    campos: ['admissao', 'contratoTipoId', 'categoriaeSocialId', 'classeId', 'percentualAdiantamento', 'salario', 'funcao', 'departamento', 'horario', 'filial'],
    labels: { admissao: 'Data Admissão', contratoTipoId: 'Tipo de Contrato ID', categoriaeSocialId: 'Categoria eSocial', classeId: 'Classe ID', percentualAdiantamento: 'Perc. Adiantamento', salario: 'Salário', funcao: 'Função/Cargo', departamento: 'Departamento', horario: 'Horário', filial: 'Filial' },
  },
];

function mostrarPreview(dados) {
  // Badge do tipo
  var badgeWrap = document.getElementById('preview-tipo-badge');
  badgeWrap.innerHTML = '<div class="preview-badge">' + (TIPO_LABELS[STATE.tipo] || STATE.tipo) + '</div>';

  // Gerar seções de preview
  var container = document.getElementById('preview-container');
  container.innerHTML = '';

  PREVIEW_SECTIONS.forEach(function(sec) {
    var div = document.createElement('div');
    div.className = 'preview-section';

    var head = document.createElement('div');
    head.className = 'preview-sec-head';
    head.innerHTML = '<span>' + sec.icon + '</span>' + sec.title;
    div.appendChild(head);

    var fields = document.createElement('div');
    fields.className = 'preview-fields';

    sec.campos.forEach(function(campo) {
      var val = dados[campo] || '';
      var f = document.createElement('div');
      f.className = 'preview-field';
      f.innerHTML =
        '<label>' + (sec.labels[campo] || campo) + '</label>' +
        '<p class="' + (val ? '' : 'empty') + '">' + (val || '—') + '</p>';
      fields.appendChild(f);
    });

    div.appendChild(fields);
    container.appendChild(div);
  });

  // Mostrar card de preview
  var cardPreview = document.getElementById('card-preview');
  cardPreview.classList.remove('hidden');

  setTimeout(function() {
    cardPreview.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }, 150);

  toast('✅ Planilha importada! Verifique os dados.', 'ok');
}

// =============================================
// 5. GERAR OUTPUT (JSON ou CSV)
// =============================================

function gerarOutput() {
  if (!STATE.dados || !STATE.tipo) {
    toast('❌ Nenhum dado disponível.', 'err');
    return;
  }

  var btn = document.getElementById('btn-gerar');
  btn.disabled = true;
  btn.textContent = '⏳ Gerando...';

  setTimeout(function() {
    try {
      var d = STATE.dados;

      if (STATE.tipo === 'autonomo') {
        STATE.outputStr = gerarCSVAutonomo(d);
        STATE.outputFmt = 'csv';
      } else {
        var json = gerarJSON(d, STATE.tipo);
        STATE.outputStr = JSON.stringify(json, null, 2);
        STATE.outputFmt = 'json';
      }

      // Exibir
      document.getElementById('output-pre').textContent = STATE.outputStr;

      var badge = document.getElementById('output-badge');
      if (STATE.outputFmt === 'json') {
        badge.className = 'output-fmt-badge json';
        badge.textContent = '✅ JSON — Pronto para importação';
        document.getElementById('output-subtitle').textContent = 'JSON gerado. Copie ou baixe para importar no sistema de RH.';
      } else {
        badge.className = 'output-fmt-badge csv';
        badge.textContent = '✅ CSV — Pronto para importação (Autônomo)';
        document.getElementById('output-subtitle').textContent = 'CSV gerado. Copie ou baixe para importar no sistema.';
      }

      var cardOutput = document.getElementById('card-output');
      cardOutput.classList.remove('hidden');
      setTimeout(function() {
        cardOutput.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }, 100);

      toast('✅ Arquivo de importação gerado!', 'ok');
    } catch (err) {
      console.error('[gerarOutput]', err);
      toast('❌ Erro ao gerar: ' + err.message, 'err');
    } finally {
      btn.disabled = false;
      btn.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/></svg> Gerar Arquivo de Importação';
    }
  }, 80);
}

// ─── Helpers ──────────────────────────────────
function limpar(val) {
  return String(val || '').replace(/\D/g, '');
}

function ddd(fone) {
  var n = limpar(fone);
  return n.length >= 10 ? n.slice(0, 2) : '';
}

function numero(fone) {
  var n = limpar(fone);
  return n.length >= 10 ? n.slice(2) : n;
}

function parseSal(val) {
  if (!val) return 0;
  return parseFloat(String(val).replace(/\./g, '').replace(',', '.')) || 0;
}

function fmtDate(val) {
  if (!val) return '';
  // Já no formato DD/MM/YYYY
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(val)) return val;
  // Formato YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}/.test(val)) {
    var p = val.split('-');
    return p[2] + '/' + p[1] + '/' + p[0];
  }
  return val;
}

// ─── Gerar JSON (CLT / Estagiário) ───────────
function gerarJSON(d, tipo) {
  var fone = d.fone || d.telefone_emp || '';
  var json = {
    tipo: tipo,
    empresaId: 0,
    funcionarioContribuinteId: 0,
    vFuncionarioContribuinteId: 0,
    nome: d.nome || '',
    email: d.email || '',
    enderecoCep: (function(c) { var n = limpar(c); return n.length === 8 ? n.slice(0,5) + '-' + n.slice(5) : n; })(d.cep),
    enderecoLogradouroId: d.enderecoLogradouroId || 0,
    endereco: d.endereco || '',
    enderecoNumero: d.enderecoNumero || '',
    enderecoComplemento: '',
    enderecoBairro: d.bairro || '',
    enderecoCidadeId: d.enderecoCidadeId || 0,
    dddTelefone: ddd(fone),
    telefone: numero(fone),
    dddCelular: ddd(d.celular),
    celular: numero(d.celular),
    estrangeiroEnderecoPaisId: 0,
    estrangeiroEndereco: '',
    estrangeiroEnderecoNumero: '',
    estrangeiroComplemento: '',
    estrangeiroEnderecoBairro: '',
    estrangeiroEnderecoCidade: '',
    estrangeiroEnderecoCodPostal: '',
    tipoSanguineoId: 0,
    etniaId: d.etniaId || 0,
    cabeloCorId: 0,
    olhoCorId: 0,
    sexo: 0,
    ehDeficiente: d.deficiencia === 'Sim' ? 1 : 0,
    deficienciaCota: 0,
    deficienciaFisica: 0,
    deficienciaAuditiva: 0,
    deficienciaVisual: 0,
    deficienciaIntelectual: 0,
    deficienciaMental: 0,
    deficienciaReabilitado: 0,
    deficienciaObservacao: d.deficiencia_qual || '',
    altura: 0,
    peso: 0,
    sinaisCorpo: 0,
    fotoNome: '',
    foto: '',
    paisNascimentoId: 0,
    nascimentoCidadeId: 0,
    estadoCivilId: 0,
    nascimentoData: d.nascimento || '',
    grauInstrucaoId: 0,
    paisNacionalidadeId: 0,
    estrangeiroDataChegadaBrasil: '',
    estrangeiroCasadoComBrasileiro: 0,
    estrangeiroComFilhoBrasileiro: 0,
    rneNumero: '',
    rneOrgaoEmissor: '',
    rneEmissao: '',
    estrangeiroCondicaoIngresso: 0,
    estrangeiroTipoResidencia: 0,
    admissaoData: d.admissao || '',
    entradaData: d.admissao || '',
    cadastroData: '',
    admissaoTipoId: 0,
    contratoTipoId: d.contratoTipoId || 0,
    ocupacaoNaturezaId: 0,
    cnpjEmpresaAnterior: '',
    transferenciaOnus: 0,
    transferenciaData: '',
    adicionalTempoServicoInicio: '',
    aposentadoriaData: '',
    desligamentoData: '',
    baixaData: '',
    matriculaAnterior: '',
    dataReintegracao: '',
    cpf: limpar(d.cpf),
    nomeImagemCPF: '',
    imagemCPF: '',
    rg: d.rg || '',
    rgOrgaoEmissor: d.rg_orgao || '',
    rgEmissao: d.rg_expedicao || '',
    rgUf: '',
    nomeImagemRG: '',
    imagemRG: '',
    tituloEleitor: d.titulo || '',
    tituloEleitorZona: '',
    tituloEleitorSecao: '',
    nomeImagemTituloEleitor: '',
    imagemTituloEleitor: '',
    pisNumero: d.pis || '',
    pisEmissao: '',
    nomeImagemPis: '',
    imagemPis: '',
    certificadoReservista: d.reservista || '',
    registroCivilId: 0,
    registroCivilTermoMatricula: d.certidao_num || '',
    registroCivilCartorio: d.certidao_cartorio || '',
    registroCivilLivro: d.certidao_livro || '',
    registroCivilFolha: d.certidao_folha || '',
    registroCivilCidadeId: 0,
    registroCivilEmissao: d.certidao_data || '',
    carteiraTrabalho: d.ctps_num || '',
    carteiraTrabalhoSerie: d.ctps_serie || '',
    carteiraTrabalhoSerieDigito: '',
    carteiraTrabalhoEmissao: d.ctps_data || '',
    carteiraTrabalhoUf: d.ctps_local || '',
    nomeImagemCTPS: '',
    imagemCTPS: '',
    cnh: d.cnh || d.habilitacao || '',
    cnhUf: '',
    cnhCategoria: '',
    cnhEmissao: '',
    cnhVencimento: '',
    cnhPrimeiraHabilitacao: '',
    nomeImagemCNH: '',
    imagemCNH: '',
    ricNumero: '',
    ricOrgaoEmissor: '',
    ricEmissao: '',
    maeNome: d.mae || '',
    paiNome: d.pai || '',
    conjugeNome: d.conjuge || '',
    conjugeNascimentoCidadeId: 0,
    conjugeNascimentoData: d.conjuge_nasc || '',
    nomeSocial: '',
    categoriaeSocialId: d.categoriaeSocialId || 0,
    fgtsOcorrenciaId: 0,
    fgtsConta: '',
    regimePrevidenciario: 0,
    sindicatoId: 0,
    sindicalizado: d.sind === 'Sim' ? 1 : 0,
    classeId: d.classeId || { funcionario: 0, contribuinte: 0 },
    centroCustoId: '',
    departamentoId: d.departamento || '',
    cartaoPonto: '',
    fichaRegistro: '',
    livro: '',
    folha: '',
    regimeJornadaTrabalhoId: 0,
    tipoJornada: 0,
    tipoJornadaDescricao: d.horario || '',
    horarioNoturno: 0,
    tipoEscalaId: 0,
    descansoSemanalId: 0,
    quadroHorarioId: 0,
    cargoId: 0,
    ocNumero: '',
    ocOrgaoEmissor: '',
    ocEmissao: '',
    ocValidade: '',
    formaPagamento: '',
    funcionarioTipoId: 0,
    salarioInicial: parseSal(d.salario),
    remuneracao: parseSal(d.salario),
    percentualComissao: 0,
    horaMensal: 0,
    horaSemanal: 0,
    horaDiaria: 0,
    insalubridadeAdicional: parseFloat(d.insalubridade) || 0,
    insalubridadeIncidenciaId: 0,
    periculosidadeAdicional: parseFloat(d.periculosidade) || 0,
    periculosidadeIncidenciaId: 0,
    noturnoAdicional: 0,
    noturnoIncidenciaId: 0,
    valorPrevidenciaPrivada: 0,
    valorPrevidenciaPrivada13: 0,
    prazoExperiencia: parseInt(d.dias_exp) || 0,
    prazoExperienciaFim: '',
    prazoExperienciaProrrogacao: 0,
    prazoExperienciaProrrogacaoFim: '',
    bancoId: 0,
    bancoContaAgencia: '',
    bancoConta: '',
    bancoContaDigito: '',
    bancoContaTipoId: 0,
    bancoModoPagamento: 0,
    cartaoSalario: '',
    recebeValeRefeicao: 0,
    cartaoVR: '',
    recebeValeAlimentacao: 0,
    cartaoVA: '',
    recebeValeTransporte: 0,
    cartaoVT: '',
    percentualAdiantamento: d.percentualAdiantamento || 0,
    contribuicaoSindical: 0,
    recebeAdiantamento: d.adiantamento === 'Sim' ? 1 : 0,
    regimeTempoParcial: 0,
    beneficioDesemprego: d.seguro === 'Sim' ? 1 : 0,
    descSimpIRRF: 0,
    observacao: d.obs || '',
    numeroRecibo: '',
    dataIntegracao: '',
    qualificacaoStatus: '',
    qualificacaoMensagem: '',
    qualificacaoOrientacao: '',
  };

  return [json];
}

// ─── Gerar CSV para Autônomo ──────────────────
function gerarCSVAutonomo(d) {
  var row = new Array(42).fill('');
  var tel = limpar(d.fone || d.telefone_emp || '');

  row[0]  = limpar(d.cpf);
  row[1]  = d.nome || '';
  row[2]  = limpar(d.cep);
  row[3]  = 'Rua';
  row[4]  = d.endereco || '';
  row[5]  = 'S/N';
  row[6]  = '';
  row[7]  = d.bairro || '';
  row[8]  = '';
  row[9]  = '';
  if (tel.length >= 10) {
    row[10] = tel.slice(0, 2);
    row[11] = tel.slice(2);
  }
  row[14] = d.email || '';
  row[15] = fmtDate(d.nascimento);
  row[16] = '';
  row[17] = '';
  row[18] = d.mae || '';
  row[19] = '07';  // Grau instrução padrão
  row[20] = '';
  row[21] = '0';   // Estado civil padrão
  row[22] = d.rg || '';
  row[23] = d.rg_orgao || '';
  row[24] = fmtDate(d.rg_expedicao);
  row[25] = '';
  row[26] = d.reservista || '';
  row[27] = d.titulo || '';
  row[31] = d.ctps_num || '';
  row[32] = d.ctps_serie || '';
  row[34] = fmtDate(d.ctps_data);
  row[36] = d.pis || '';
  row[37] = '3532-30';  // CBO
  row[38] = '13';       // SEFIP
  row[39] = '712';      // eSocial
  row[40] = '20';       // INSS
  row[41] = '20';       // IR

  return row.join(';') + ';';
}

// =============================================
// 6. AÇÕES DE OUTPUT
// =============================================

function copiar() {
  if (!STATE.outputStr) return;

  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(STATE.outputStr).then(function() {
      flashCopy();
    }).catch(function() {
      copiarFallback();
    });
  } else {
    copiarFallback();
  }
}

function copiarFallback() {
  var ta = document.createElement('textarea');
  ta.value = STATE.outputStr;
  ta.style.position = 'fixed';
  ta.style.opacity = '0';
  document.body.appendChild(ta);
  ta.select();
  try {
    document.execCommand('copy');
    flashCopy();
  } catch (e) {
    toast('❌ Não foi possível copiar. Selecione e copie manualmente.', 'err');
  }
  document.body.removeChild(ta);
}

function flashCopy() {
  var btn = document.getElementById('btn-copy');
  btn.classList.add('ok');
  btn.innerHTML = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg> Copiado!';
  toast('✅ Copiado para a área de transferência!', 'ok');
  setTimeout(function() {
    btn.classList.remove('ok');
    btn.innerHTML = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg> Copiar';
  }, 2500);
}

function baixar() {
  if (!STATE.outputStr) return;

  var nome = (STATE.dados && STATE.dados.nome ? STATE.dados.nome : 'Colaborador');
  nome = nome.replace(/[^a-zA-Z0-9À-ÿ ]/g, '').trim().replace(/\s+/g, '_');
  var ts = new Date().toISOString().slice(0, 16).replace(/[:.]/g, '-');
  var ext = STATE.outputFmt === 'csv' ? 'csv' : 'json';
  var mime = STATE.outputFmt === 'csv' ? 'text/csv;charset=utf-8;' : 'application/json';
  var fileName = 'Importacao_' + nome + '_' + ts + '.' + ext;

  var blob = new Blob([STATE.outputStr], { type: mime });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  setTimeout(function() {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 100);

  toast('⬇️ "' + fileName + '" baixado!', 'ok');
}

// =============================================
// 7. REINICIAR
// =============================================

function reiniciar() {
  STATE.tipo = null;
  STATE.dados = null;
  STATE.outputStr = null;
  STATE.outputFmt = null;

  // Reset botões de tipo
  ['estagiario', 'clt', 'autonomo'].forEach(function(t) {
    document.getElementById('btn-' + t).classList.remove('selected');
  });

  // Reset badge
  var badge = document.getElementById('tipo-selected-badge');
  badge.classList.add('hidden');
  badge.textContent = '';

  // Reset dropzone
  resetDropzone();

  // Mostrar hint
  document.querySelector('.upload-hint').style.display = '';

  // Ocultar cards
  document.getElementById('card-preview').classList.add('hidden');
  document.getElementById('card-output').classList.add('hidden');
  document.getElementById('preview-container').innerHTML = '';
  document.getElementById('output-pre').textContent = '';

  window.scrollTo({ top: 0, behavior: 'smooth' });
  toast('↩ Pronto para nova importação.', '');
}

// =============================================
// 8. TOAST
// =============================================

var _toastTimer = null;

function toast(msg, tipo) {
  var el = document.getElementById('toast');
  el.textContent = msg;
  el.className = 'toast show' + (tipo ? ' ' + tipo : '');
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(function() {
    el.className = 'toast';
  }, 3500);
}
