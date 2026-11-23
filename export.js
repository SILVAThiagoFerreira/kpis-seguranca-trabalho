(function () {
  'use strict';

  // Exportação em PowerPoint (.pptx) no padrão das lâminas Enaex: fundo branco,
  // título vermelho em caixa alta, logo à direita e rodapé "Strictly private".
  // O conteúdo é lido da própria página já renderizada: o arquivo traz exatamente
  // o que está na tela (filtros, área escolhida e tabelas filtradas por busca).
  // Gráficos são nativos do PowerPoint e continuam editáveis.

  var PPTX_URL = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
  var LOGO_URL = 'assets/enaex-referencia.png';
  var C = {
    branco: 'FFFFFF', vermelho: 'B5121B', escuro: '3C3F44', texto: '4A4A4A',
    medio: '8A8F96', linha: 'C4C7CC', fundoCard: 'F4F5F6', verde: '167C62'
  };
  var X0 = 0.5, W = 12.333;

  // ---------- leitura da página ----------

  function limpo(valor) {
    return String(valor == null ? '' : valor).replace(/\s+/g, ' ').trim();
  }

  function textoDe(elemento) {
    return elemento ? limpo(elemento.textContent) : '';
  }

  function visivel(elemento) {
    return !elemento.closest('[hidden]') && elemento.getClientRects().length > 0;
  }

  // "4,2%" -> 4.2 · "3068.13" -> 3068.13 · "1.234,5" -> 1234.5
  function numero(texto) {
    var t = limpo(texto).replace(/[%\s]/g, '');
    if (t.indexOf(',') >= 0) t = t.replace(/\./g, '').replace(',', '.');
    var n = parseFloat(t);
    return isFinite(n) ? n : 0;
  }

  // Gráfico agrupado desenhado com retângulos simples: cada barra traz <title> "Mês: valor série".
  function coletarGraficoAgrupado(svg) {
    var categorias = [], series = [], percentual = false;
    [].slice.call(svg.querySelectorAll('rect > title')).forEach(function (titulo) {
      var m = limpo(titulo.textContent).match(/^(.*?):\s*(-?[\d.,]+)\s*(%?)\s*(.*)$/);
      if (!m) return;
      var categoria = m[1], nomeSerie = limpo(m[4]) || 'Valor';
      if (categorias.indexOf(categoria) < 0) categorias.push(categoria);
      var serie = series.filter(function (s) { return s.chave === nomeSerie; })[0];
      if (!serie) {
        var cor = (titulo.parentNode.getAttribute('fill') || '').replace('#', '').toUpperCase();
        serie = { chave: nomeSerie, nome: nomeSerie.charAt(0).toUpperCase() + nomeSerie.slice(1), cor: /^[0-9A-F]{6}$/.test(cor) ? cor : null, registros: {} };
        series.push(serie);
      }
      serie.registros[categoria] = numero(m[2]);
      if (m[3]) percentual = true;
    });
    series.forEach(function (s) {
      s.valores = categorias.map(function (c) { return s.registros[c] || 0; });
      delete s.registros;
      delete s.chave;
    });
    return { tipo: 'barra', categorias: categorias, series: series, percentual: percentual };
  }

  function coletarGrafico(svg) {
    if (!svg.querySelector('rect.bar, circle.dot')) return coletarGraficoAgrupado(svg);
    var linha = svg.querySelector('polyline') !== null;
    var pontos = [].slice.call(svg.querySelectorAll(linha ? 'circle.dot' : 'rect.bar'));
    var rotulos = [].slice.call(svg.querySelectorAll('text.axis:not(.axis-y)'));
    var valores = [].slice.call(svg.querySelectorAll('text.bar-value'));
    var categorias = [], valoresNumericos = [], percentual = false;

    pontos.forEach(function (ponto) {
      var x = linha
        ? Number(ponto.getAttribute('cx'))
        : Number(ponto.getAttribute('x')) + Number(ponto.getAttribute('width')) / 2;
      var mesmoX = function (el) { return Math.abs(Number(el.getAttribute('x')) - x) < 1; };

      // Rótulos podem ocupar mais de uma linha: juntamos os trechos com o mesmo x.
      var nome = rotulos.filter(mesmoX).map(function (el) { return limpo(el.textContent); }).join(' ');

      var exibido = valores.filter(mesmoX).map(function (el) { return limpo(el.textContent); })[0];
      if (exibido === undefined) {
        // Sem rótulo de valor na barra: usa a dica "Nome: valor | detalhe".
        var dica = ponto.getAttribute('data-tooltip') || '';
        exibido = dica.slice(dica.indexOf(': ') + 2).split(' |')[0];
      }
      if (/%/.test(exibido)) percentual = true;
      categorias.push(nome);
      valoresNumericos.push(numero(exibido));
    });

    return { tipo: linha ? 'linha' : 'barra', categorias: categorias, series: [{ nome: '', valores: valoresNumericos }], percentual: percentual };
  }

  // Célula com mais de um bloco (ex.: data + "Faltam 200 dias") vira "data · Faltam 200 dias".
  function textoCelula(celula) {
    return [].slice.call(celula.childNodes).map(function (no) { return limpo(no.textContent); }).filter(Boolean).join(' · ');
  }

  function coletarTabela(tabela) {
    var colunas = [].slice.call(tabela.querySelectorAll('thead th')).map(textoDe);
    var linhas = [].slice.call(tabela.querySelectorAll('tbody tr'))
      .filter(function (tr) {
        return !tr.hidden && !tr.classList.contains('empty-row') && !tr.classList.contains('empty-data-row');
      })
      .map(function (tr) { return [].slice.call(tr.cells).map(textoCelula); });
    return { colunas: colunas, linhas: linhas };
  }

  function coletarPainel(painel) {
    // A legenda traz o nome da série (1º trecho) e, em seguida, complementos como "Total acumulado: 66".
    var trechosLegenda = [].slice.call(painel.querySelectorAll('.legend span')).map(textoDe).filter(Boolean);
    var item = {
      kicker: textoDe(painel.querySelector('.panel-kicker')),
      titulo: textoDe(painel.querySelector('h3')),
      textos: (trechosLegenda.length > 1 ? [trechosLegenda.slice(1).join(' · ')] : [])
        .concat([].slice.call(painel.querySelectorAll('p.footnote')).map(textoDe))
        .filter(Boolean),
      contagem: '',
      legenda: trechosLegenda[0] || ''
    };

    var resumo = painel.querySelector('.control-summary');
    if (resumo) item.contagem = [].slice.call(resumo.querySelectorAll('span')).map(textoDe).filter(Boolean).join(' · ');
    var registros = painel.querySelector('.table-tools small');
    if (!item.contagem && registros) item.contagem = textoDe(registros);

    var svg = painel.querySelector('.chart svg');
    var tabela = painel.querySelector('table');
    var lista = painel.querySelector('.list');

    if (svg) {
      item.tipo = 'grafico';
      item.grafico = coletarGrafico(svg);
      item.grafico.series.forEach(function (serie) { serie.nome = serie.nome || item.legenda || item.titulo; });
      // Nomes de séries já aparecem no gráfico: não os repetimos como nota.
      item.textos = item.textos.filter(function (texto) {
        return !item.grafico.series.some(function (serie) { return serie.nome === texto; });
      });
    } else if (tabela) {
      item.tipo = 'tabela';
      var dados = coletarTabela(tabela);
      item.colunas = dados.colunas;
      item.linhas = dados.linhas;
    } else if (lista) {
      item.tipo = 'lista';
      item.colunas = ['Indicador', 'Valor'];
      item.linhas = [].slice.call(lista.querySelectorAll('.list-row')).map(function (row) {
        return [textoDe(row.querySelector('strong')), textoDe(row.querySelector('.badge'))];
      });
    } else {
      item.tipo = 'texto';
    }
    return item;
  }

  function coletarSecao(secao) {
    return {
      numero: textoDe(secao.querySelector('.section-title span')),
      titulo: textoDe(secao.querySelector('.section-title h2')),
      descricao: textoDe(secao.querySelector('.section-description')),
      indicadores: [].slice.call(secao.querySelectorAll('.metric')).filter(visivel).map(function (metrica) {
        return {
          rotulo: textoDe(metrica.querySelector('.metric-label')),
          valor: textoDe(metrica.querySelector('.metric-value')),
          nota: textoDe(metrica.querySelector('.metric-note')),
          tipo: metrica.classList.contains('alert') ? 'alerta' : metrica.classList.contains('good') ? 'bom' : ''
        };
      }),
      paineis: [].slice.call(secao.querySelectorAll('article.panel')).filter(visivel).map(coletarPainel)
    };
  }

  function descreverFiltros() {
    var campos = [['filter-area', 'Área'], ['filter-year', 'Ano'], ['filter-month', 'Mês'], ['filter-status', 'Status']];
    return campos.map(function (campo) {
      var select = document.getElementById(campo[0]);
      var opcao = select && select.options[select.selectedIndex];
      return select ? campo[1] + ': ' + limpo(opcao ? opcao.textContent : '') : '';
    }).filter(Boolean).join(' · ');
  }

  function coletarModelo() {
    var secoes = [].slice.call(document.querySelectorAll('#dashboard section.section'))
      .filter(visivel)
      .map(coletarSecao);
    return {
      capa: {
        titulo: 'Saúde & Segurança do Trabalho',
        subtitulo: 'Pessoas e Processos',
        filtros: descreverFiltros(),
        gerado: new Date().toLocaleDateString('pt-BR')
      },
      secoes: secoes
    };
  }

  // ---------- montagem do .pptx ----------

  function construirApresentacao(modelo, PptxGen, logo) {
    var pptx = new PptxGen();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.title = 'HSEC · Pessoas e Processos';
    pptx.author = 'US Vale Verde';
    pptx.company = 'Enaex Brasil';

    var slides = [];

    // O logo vai no modelo de slide (uma única cópia no arquivo), não em cada lâmina.
    if (logo) {
      pptx.defineSlideMaster({
        title: 'HSEC',
        background: { color: C.branco },
        objects: [{ image: { data: logo, x: X0 + W - 1.57, y: 0.4, w: 1.57, h: 0.5 } }]
      });
    }

    function texto(slide, conteudo, opcoes) {
      var o = { fontFace: 'Arial', fontSize: 10.5, color: C.texto, margin: 0, valign: 'top' };
      for (var chave in opcoes) o[chave] = opcoes[chave];
      slide.addText(conteudo, o);
    }

    function novoSlide() {
      var slide = pptx.addSlide(logo ? { masterName: 'HSEC' } : undefined);
      slide.background = { color: C.branco };
      slides.push(slide);
      return slide;
    }

    function cabecalho(slide, kicker, titulo, subtitulo) {
      if (kicker) texto(slide, kicker.toUpperCase(), { x: X0, y: 0.42, w: 9.5, h: 0.25, fontSize: 10, color: C.medio, charSpacing: 1 });
      texto(slide, titulo.toUpperCase(), { x: X0, y: 0.66, w: 10.6, h: 0.5, fontSize: 20, bold: true, color: C.vermelho, valign: 'middle' });
      slide.addShape('line', { x: X0, y: 1.3, w: W, h: 0, line: { color: C.linha, width: 0.75 } });
      if (subtitulo) texto(slide, subtitulo, { x: X0, y: 1.38, w: W, h: 0.28, fontSize: 10.5, color: C.medio });
    }

    function rodape(slide) {
      slide.addShape('line', { x: X0, y: 6.98, w: W, h: 0, line: { color: C.escuro, width: 1 } });
      texto(slide, 'US Vale Verde · Saúde & Segurança do Trabalho · Pessoas e Processos', { x: X0, y: 7.04, w: 8, h: 0.22, fontSize: 8, color: C.medio });
      texto(slide, 'STRICTLY PRIVATE AND CONFIDENTIAL   ·   ' + slides.length, { x: X0 + W - 5, y: 7.04, w: 5, h: 0.22, fontSize: 8, color: C.escuro, align: 'right' });
    }

    function notas(slide, textos, y) {
      if (!textos || !textos.length) return;
      slide.addText(textos.map(function (t, i) {
        return { text: t, options: { breakLine: i < textos.length - 1 } };
      }), { x: X0, y: y, w: W, h: 0.5, fontFace: 'Arial', fontSize: 10.5, color: C.texto, margin: 0, valign: 'top' });
    }

    function capa() {
      var slide = novoSlide();
      slide.addShape('rect', { x: X0, y: 2.3, w: 1.1, h: 0.06, fill: { color: C.vermelho }, line: { type: 'none' } });
      texto(slide, modelo.capa.titulo.toUpperCase(), { x: X0, y: 2.5, w: W, h: 0.8, fontSize: 34, bold: true, color: C.vermelho });
      texto(slide, modelo.capa.subtitulo, { x: X0, y: 3.35, w: W, h: 0.5, fontSize: 20, color: C.escuro });
      texto(slide, 'Filtros aplicados: ' + (modelo.capa.filtros || '—'), { x: X0, y: 4.1, w: W, h: 0.3, fontSize: 12, color: C.texto });
      texto(slide, 'Gerado em ' + modelo.capa.gerado, { x: X0, y: 4.42, w: W, h: 0.3, fontSize: 12, color: C.texto });
      rodape(slide);
    }

    function indicadores(secao) {
      var slide = novoSlide();
      cabecalho(slide, secao.numero + ' · Indicadores', secao.titulo, secao.descricao);
      var itens = secao.indicadores;
      if (!itens.length) texto(slide, 'Sem indicadores para os filtros atuais.', { x: X0, y: 2.1, w: W, h: 0.4, fontSize: 12 });
      var colunas = Math.min(4, itens.length) || 1;
      var espaco = 0.3;
      var largura = (W - espaco * (colunas - 1)) / colunas;
      var altura = 2.0;
      itens.forEach(function (item, i) {
        var coluna = i % colunas, linhaCard = Math.floor(i / colunas);
        var x = X0 + coluna * (largura + espaco), y = 2.1 + linhaCard * (altura + espaco);
        var cor = item.tipo === 'alerta' ? C.vermelho : item.tipo === 'bom' ? C.verde : C.escuro;
        slide.addShape('rect', { x: x, y: y, w: largura, h: altura, fill: { color: C.fundoCard }, line: { type: 'none' } });
        slide.addShape('rect', { x: x, y: y, w: largura, h: 0.06, fill: { color: cor }, line: { type: 'none' } });
        texto(slide, item.rotulo.toUpperCase(), { x: x + 0.25, y: y + 0.25, w: largura - 0.5, h: 0.5, fontSize: 10, color: C.medio, charSpacing: 0.5 });
        texto(slide, item.valor, { x: x + 0.25, y: y + 0.8, w: largura - 0.5, h: 0.7, fontSize: 34, bold: true, color: cor });
        texto(slide, item.nota, { x: x + 0.25, y: y + 1.5, w: largura - 0.5, h: 0.35, fontSize: 10, color: C.texto });
      });
      rodape(slide);
    }

    function grafico(painel) {
      var slide = novoSlide();
      cabecalho(slide, painel.kicker, painel.titulo, painel.contagem);
      var g = painel.grafico;
      var series = g.series.map(function (s) {
        return { name: s.nome, labels: g.categorias, values: s.valores };
      });
      var opcoes = {
        x: X0, y: 1.8, w: W, h: 4.4,
        chartColors: g.series.map(function (s) { return s.cor || C.vermelho; }),
        showValue: true,
        dataLabelFontFace: 'Arial', dataLabelFontSize: 10, dataLabelColor: C.escuro,
        dataLabelFormatCode: g.percentual ? '0.0"%"' : 'General',
        catAxisLabelFontFace: 'Arial', catAxisLabelFontSize: 10, catAxisLabelColor: C.texto,
        valAxisHidden: true, valAxisMinVal: 0, valGridLine: { style: 'none' },
        showLegend: true, legendPos: 'b', legendFontFace: 'Arial', legendFontSize: 10, legendColor: C.texto
      };
      if (g.tipo === 'linha') {
        opcoes.lineSize = 2;
        opcoes.lineDataSymbol = 'circle';
        opcoes.lineDataSymbolSize = 7;
        opcoes.dataLabelPosition = 't';
        slide.addChart(pptx.charts.LINE, series, opcoes);
      } else {
        opcoes.barDir = 'col';
        opcoes.barGapWidthPct = 60;
        opcoes.dataLabelPosition = 'outEnd';
        slide.addChart(pptx.charts.BAR, series, opcoes);
      }
      notas(slide, painel.textos, 6.3);
      rodape(slide);
    }

    function tabelas(painel) {
      var colunas = painel.colunas || [], linhas = painel.linhas || [];
      var larguras = colunas.map(function (nome, i) {
        var maior = Math.max(String(nome).length, 6);
        linhas.forEach(function (l) { maior = Math.max(maior, String(l[i] || '').length); });
        return Math.min(maior, 60);
      });
      var somaLarguras = larguras.reduce(function (s, v) { return s + v; }, 0) || 1;
      var colW = larguras.map(function (v) { return W * v / somaLarguras; });
      // Altura estimada de cada linha: texto que quebra em mais de uma linha ocupa mais espaço.
      var alturaLinha = function (linha) {
        var linhasTexto = colunas.reduce(function (maximo, _, i) {
          return Math.max(maximo, Math.ceil(String(linha[i] || '—').length / Math.max(1, colW[i] * 13)));
        }, 1);
        return Math.max(0.34, 0.12 + linhasTexto * 0.19);
      };
      // Quebra as linhas em páginas que cabem na área útil (cabeçalho de 0,36 pol. já descontado).
      var paginas = [], atual = [], alturaAtual = 0;
      linhas.forEach(function (linha) {
        var altura = alturaLinha(linha);
        if (atual.length && alturaAtual + altura > 4.1) { paginas.push(atual); atual = []; alturaAtual = 0; }
        atual.push(linha);
        alturaAtual += altura;
      });
      if (atual.length || !paginas.length) paginas.push(atual);
      var contagem = painel.contagem || (linhas.length + ' registros');
      for (var p = 0; p < paginas.length; p++) {
        var slide = novoSlide();
        var subtitulo = paginas.length > 1 ? contagem + '  ·  página ' + (p + 1) + ' de ' + paginas.length : contagem;
        cabecalho(slide, painel.kicker, painel.titulo, subtitulo);
        if (!colunas.length || !linhas.length) {
          texto(slide, 'Base atualizada sem registros compatíveis.', { x: X0, y: 1.9, w: W, h: 0.4, fontSize: 12 });
        } else {
          var fatia = paginas[p];
          var cabecalhoTabela = colunas.map(function (coluna) {
            return { text: coluna.toUpperCase(), options: { bold: true, color: C.branco, fill: { color: C.escuro }, fontSize: 9.5 } };
          });
          var corpo = fatia.map(function (linha) {
            return colunas.map(function (_, i) {
              return { text: linha[i] || '—', options: { color: C.texto, fontSize: 10 } };
            });
          });
          slide.addTable([cabecalhoTabela].concat(corpo), {
            x: X0, y: 1.8, w: W,
            colW: colW,
            rowH: [0.36].concat(fatia.map(alturaLinha)),
            fontFace: 'Arial',
            valign: 'middle',
            border: { type: 'solid', pt: 0.5, color: C.linha }
          });
        }
        if (p === paginas.length - 1) notas(slide, painel.textos, 6.3);
        rodape(slide);
      }
    }

    function textoLivre(painel) {
      var slide = novoSlide();
      cabecalho(slide, painel.kicker, painel.titulo, painel.contagem);
      notas(slide, painel.textos, 1.9);
      rodape(slide);
    }

    capa();
    modelo.secoes.forEach(function (secao) {
      indicadores(secao);
      secao.paineis.forEach(function (painel) {
        if (painel.tipo === 'grafico') grafico(painel);
        else if (painel.tipo === 'tabela' || painel.tipo === 'lista') tabelas(painel);
        else textoLivre(painel);
      });
    });
    return pptx;
  }

  // ---------- navegador: carga, botão e download ----------

  function carregarPptx() {
    if (window.PptxGenJS) return Promise.resolve(window.PptxGenJS);
    return new Promise(function (ok, erro) {
      var script = document.createElement('script');
      script.src = PPTX_URL;
      script.onload = function () {
        if (window.PptxGenJS) ok(window.PptxGenJS);
        else erro(new Error('o gerador de PowerPoint não carregou.'));
      };
      script.onerror = function () { erro(new Error('não foi possível carregar o gerador de PowerPoint. Verifique a conexão.')); };
      document.head.appendChild(script);
    });
  }

  function carregarLogo() {
    return fetch(LOGO_URL).then(function (r) { return r.blob(); }).then(function (blob) {
      return new Promise(function (ok, erro) {
        var leitor = new FileReader();
        leitor.onload = function () { ok(leitor.result); };
        leitor.onerror = erro;
        leitor.readAsDataURL(blob);
      });
    }).catch(function () { return null; });
  }

  function nomeArquivo() {
    var d = new Date();
    var data = [d.getFullYear(), String(d.getMonth() + 1).padStart(2, '0'), String(d.getDate()).padStart(2, '0')].join('-');
    return 'HSEC-Pessoas-e-Processos_' + data + '.pptx';
  }

  function baixarArquivo(blob, nome) {
    var url = URL.createObjectURL(blob);
    var link = document.createElement('a');
    link.href = url;
    link.download = nome;
    document.body.appendChild(link);
    link.click();
    setTimeout(function () { URL.revokeObjectURL(url); link.remove(); }, 60000);
  }

  function ligarBotao() {
    var botao = document.getElementById('btn-pptx');
    var status = document.getElementById('export-status');
    if (!botao) return;
    var rotulo = botao.textContent;

    var atualizar = function () {
      if (botao.dataset.working) return;
      botao.disabled = !document.querySelector('#dashboard section.section');
    };

    botao.addEventListener('click', function () {
      botao.dataset.working = '1';
      botao.disabled = true;
      botao.textContent = 'Gerando PowerPoint…';
      if (status) status.textContent = '';
      Promise.all([carregarPptx(), carregarLogo()])
        .then(function (carga) {
          var modelo = coletarModelo();
          if (!modelo.secoes.length) throw new Error('não há indicadores visíveis para exportar.');
          return construirApresentacao(modelo, carga[0], carga[1]).write({ outputType: 'blob' });
        })
        .then(function (blob) { baixarArquivo(blob, nomeArquivo()); })
        .catch(function (erro) {
          if (status) status.textContent = 'Não foi possível gerar o PowerPoint: ' + erro.message;
        })
        .finally(function () {
          delete botao.dataset.working;
          botao.textContent = rotulo;
          atualizar();
        });
    });

    var painel = document.getElementById('dashboard');
    if (painel) new MutationObserver(atualizar).observe(painel, { childList: true, subtree: true });
    atualizar();
  }

  window.HSEC_EXPORT = { coletarModelo: coletarModelo, construirApresentacao: construirApresentacao };

  if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', ligarBotao);
    else ligarBotao();
  }
})();
