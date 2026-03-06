// ==UserScript==
// @name         Diario GED para acompanhamento de lançamentos e FICAI
// @namespace    http://tampermonkey.net/
// @version      v1.0
// @description  Automação para emitir diários, extrair atestados e baixar HTML com Alertas Específicos por Aluno
// @author       Lucas Monteiro
// @match        http://sigeduca.seduc.mt.gov.br/ged/hwgedemitediarioclasse.aspx?*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    let executandoLoop = false;
    const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    // Objetos globais de extração
    let alunosDados = {};
    let metadadosTurma = { ip: "", fp: "", turma: "", turno: "" };
    let disciplinasLidasNoCabecalho = new Set();

    // --- FUNÇÕES DE APOIO ---
    function recortarString(strPrincipal, strInicio, strFinal) {
        var indiceInicio = strPrincipal.indexOf(strInicio);
        var indiceFinal = strPrincipal.indexOf(strFinal);
        if (indiceInicio === -1 || indiceFinal === -1) return "";
        indiceInicio += strInicio.length;
        return strPrincipal.substring(indiceInicio, indiceFinal).trim();
    }

    function ultimoElementoTDPrimeiroTR(tabela) {
        var primeiroTR = tabela.querySelector('tr');
        if (!primeiroTR) return null;
        var tds = primeiroTR.querySelectorAll('td');
        if (tds.length === 0) return null;
        return tds[tds.length - 1];
    }

    function formatarData(dataString) {
        var partes = dataString.trim().split(' ');
        var dia = parseInt(partes[0]);
        var mes = parseInt(partes[1]);
        if (isNaN(dia) || isNaN(mes)) return dataString;
        var diaFormatado = dia < 10 ? '0' + dia : dia.toString();
        var mesFormatado = mes < 10 ? '0' + mes : mes.toString();
        return diaFormatado + '/' + mesFormatado;
    }

    function isNotificationHidden(docObject) {
        var notification = docObject.getElementById('gx_ajax_notification');
        if (notification) {
            var displayStyle = docObject.defaultView.getComputedStyle(notification).getPropertyValue('display');
            return displayStyle === 'none';
        }
        return true;
    }

    // --- 1. Função de Extração de Dados (Fase 1 - Diários) ---
    function extrairDadosIframe() {
        const iframe = document.getElementById('iframeImpressao');
        if (!iframe || !iframe.contentWindow || !iframe.contentWindow.document) return false;

        const iframeDoc = iframe.contentWindow.document;
        const tabelas = iframeDoc.getElementsByTagName("table");
        let extraiuAlgo = false;
        let disciplinaAtual = "Desconhecida";

        for (var n = 0; n < tabelas.length; n++) {
            var minhaTabela = tabelas[n];
            var ultimoTD = ultimoElementoTDPrimeiroTR(minhaTabela);

            if (ultimoTD) {
                var spanHeader = ultimoTD.getElementsByTagName("span")[0];
                if (!spanHeader) continue;

                var headerText = spanHeader.textContent.trim();

                if (headerText === "Disciplina:") {
                    try {
                        var linhasDd = minhaTabela.getElementsByTagName("tr")[0].getElementsByTagName("tr");
                        var tempDisc = linhasDd[3].getElementsByTagName("td")[1].getElementsByTagName("span")[2].textContent.trim();
                        if (tempDisc) {
                            disciplinaAtual = tempDisc;
                            disciplinasLidasNoCabecalho.add(disciplinaAtual);
                        }

                        if (!metadadosTurma.turma) {
                            var tempCabecalho = linhasDd[1].getElementsByTagName("td")[1].getElementsByTagName("span")[2].textContent.trim();
                            metadadosTurma.ip = recortarString(tempCabecalho, '', 'FP:').replace('IP:', '').trim();
                            metadadosTurma.fp = recortarString(tempCabecalho, 'FP:', 'Turma:');
                            metadadosTurma.turma = recortarString(tempCabecalho, 'Turma:', 'Turno:');
                            metadadosTurma.turno = tempCabecalho.split('Turno:')[1].trim();
                        }
                    } catch (e) { console.error("Erro ao ler cabeçalho:", e); }
                }
                else if (headerText === "TF" || headerText === "Situação") {
                    var linhas = minhaTabela.getElementsByTagName("tr");
                    var datas = [];

                    for (var i = 0; i < linhas.length; i++) {
                        if (i == 0) continue;
                        if (i == 1) {
                            var dias = linhas[i].getElementsByTagName("td");
                            for (var j = 0; j < dias.length; j++) {
                                var spanData = dias[j].getElementsByTagName("span")[0];
                                if (spanData) datas.push(formatarData(spanData.textContent.trim()));
                            }
                            continue;
                        }

                        var tdCodigo = linhas[i].getElementsByTagName("td")[0];
                        if (!tdCodigo) continue;
                        var spanCodigo = tdCodigo.getElementsByTagName("span")[0];
                        if (!spanCodigo) continue;

                        var codigoEstudante = spanCodigo.textContent.trim();
                        if (codigoEstudante === '') continue;

                        var nomeEstudante = linhas[i].getElementsByTagName("td")[2].getElementsByTagName("span")[0].textContent.trim();

                        if (!alunosDados[codigoEstudante]) {
                            alunosDados[codigoEstudante] = { nome: nomeEstudante, calendario: {}, atestados: [] };
                        }

                        for (var k = 0; k < datas.length; k++) {
                            var tdData = linhas[i].getElementsByTagName("td")[3 + k];
                            if (tdData) {
                                var spanPresenca = tdData.getElementsByTagName("span")[0];
                                var status = spanPresenca ? spanPresenca.textContent.trim() : "";
                                if (!alunosDados[codigoEstudante].calendario[datas[k]]) alunosDados[codigoEstudante].calendario[datas[k]] = {};
                                if (status !== "") alunosDados[codigoEstudante].calendario[datas[k]][disciplinaAtual] = status;
                            }
                        }
                    }
                    extraiuAlgo = true;
                }
            }
        }
        return extraiuAlgo;
    }

    // --- 2. Função de Extração de Atestados (Fase 2) ---
    async function extrairAtestadosIframe(listaAlunosIds) {
        const iframe = document.getElementById('iframeImpressao');
        const status = document.getElementById('statusProgresso');
        const barra = document.getElementById('barraProgresso');

        status.innerText = "Carregando tela de Atestados...";
        barra.style.width = `0%`;
        iframe.src = 'http://sigeduca.seduc.mt.gov.br/ged/hwmgedatestado.aspx';

        let iframeDoc = iframe.contentWindow.document;
        let tentativasLoad = 0;
        while (!iframeDoc.getElementById('vGEDALUCOD') && tentativasLoad < 30) {
            await delay(1000);
            iframeDoc = iframe.contentWindow.document;
            tentativasLoad++;
        }

        if (tentativasLoad >= 30) {
            alert("Aviso: A tela de Atestados demorou muito para carregar. Os atestados não foram lidos.");
            return;
        }

        await delay(1500);

        for (let i = 0; i < listaAlunosIds.length; i++) {
            if (!executandoLoop) break;

            let alunoCodigo = listaAlunosIds[i];
            status.innerText = `Lendo Atestados: Aluno ${i + 1} de ${listaAlunosIds.length}`;
            barra.style.width = `${Math.round(((i + 1) / listaAlunosIds.length) * 100)}%`;

            let inputAluno = iframeDoc.getElementById('vGEDALUCOD');
            if (inputAluno) {
                inputAluno.value = alunoCodigo;
                if ("createEvent" in iframeDoc) {
                    var evt = iframeDoc.createEvent("HTMLEvents");
                    evt.initEvent("change", false, true);
                    inputAluno.dispatchEvent(evt);
                } else {
                    inputAluno.fireEvent("onchange");
                }
            } else { continue; }

            await delay(300);

            let btnConsultar = iframeDoc.getElementsByName('BCONSULTAR')[0] || iframeDoc.querySelector('.btnConsultar');
            if (btnConsultar) btnConsultar.click();

            await delay(300);
            while (!isNotificationHidden(iframeDoc)) { await delay(300); }

            let docTabela = iframeDoc;
            if (iframe.contentWindow.frames.length > 0 && iframe.contentWindow.frames[0].document.getElementById('GriddetalhesContainerTbl')) {
                docTabela = iframe.contentWindow.frames[0].document;
            }

            let selectPag = docTabela.getElementById('vPAG');
            let totalPaginas = selectPag ? selectPag.options.length : 1;

            for (let p = 1; p <= totalPaginas; p++) {
                if (!executandoLoop) break;

                if (p > 1 && selectPag) {
                    status.innerText = `Lendo Atestados: Aluno ${i + 1} (Pág. ${p}/${totalPaginas})`;
                    selectPag.value = p.toString();

                    if ("createEvent" in docTabela) {
                        var evtPag = docTabela.createEvent("HTMLEvents");
                        evtPag.initEvent("change", false, true);
                        selectPag.dispatchEvent(evtPag);
                    } else {
                        selectPag.fireEvent("onchange");
                    }

                    try { docTabela.defaultView.gx.evt.execEvt('EVPAG.CLICK.', selectPag); } catch(e){}

                    await delay(300);
                    while (!isNotificationHidden(docTabela)) { await delay(300); }
                }

                let tabelaDetalhes = docTabela.getElementById('GriddetalhesContainerTbl');
                if (tabelaDetalhes) {
                    let totalLinhas = tabelaDetalhes.rows.length;
                    if (totalLinhas > 1) {
                        for (let n = 1; n < totalLinhas; n++) {
                            let numStr = ("0000" + n).slice(-4);
                            try {
                                let dataIni = docTabela.getElementById('span_vGEDATEPERINI_' + numStr)?.textContent.trim() || '';
                                let dataFim = docTabela.getElementById('span_vGEDATEPERFIN_' + numStr)?.textContent.trim() || '';
                                let tipoJust = docTabela.getElementById('span_vGEDATETIPO_' + numStr)?.textContent.trim() || '';

                                if (dataIni) {
                                    alunosDados[alunoCodigo].atestados.push({ dataIni, dataFim, tipoJust });
                                }
                            } catch (err) { console.log(err); }
                        }
                    }
                }
            }
            await delay(300);
        }
    }

    function verificaAtestadoNoDia(dataCalendario, atestadosArray) {
        if (!atestadosArray || atestadosArray.length === 0) return null;
        let [diaC, mesC] = dataCalendario.split('/').map(Number);
        let numDiaAtual = mesC * 100 + diaC;

        for (let at of atestadosArray) {
            if (!at.dataIni) continue;
            let [diaI, mesI] = at.dataIni.split('/').map(Number);
            let [diaF, mesF] = (at.dataFim || at.dataIni).split('/').map(Number);
            let numDiaIni = mesI * 100 + diaI;
            let numDiaFim = mesF * 100 + diaF;

            if (numDiaAtual >= numDiaIni && numDiaAtual <= numDiaFim) {
                return at.tipoJust;
            }
        }
        return null;
    }

    // --- 3. Função para Gerar e Baixar HTML ---
    function gerarEBaixarRelatorioHTML() {
        let todasAsDatas = new Set();
        Object.values(alunosDados).forEach(aluno => {
            Object.keys(aluno.calendario).forEach(data => todasAsDatas.add(data));
        });

        let datasOrdenadas = Array.from(todasAsDatas).sort((a, b) => {
            let [da, ma] = a.split('/');
            let [db, mb] = b.split('/');
            return (ma + da).localeCompare(mb + db);
        });

        let alunosArray = Object.keys(alunosDados).map(cod => {
            let aluno = { codigo: cod, ...alunosDados[cod], totalFaltas: 0, totalPresencas: 0, faltasJustificadas: 0 };

            Object.keys(aluno.calendario).forEach(data => {
                let registrosDoDia = aluno.calendario[data];
                let temJustificativa = verificaAtestadoNoDia(data, aluno.atestados);

                Object.values(registrosDoDia).forEach(status => {
                    if (status.toUpperCase() === 'F') {
                        aluno.totalFaltas++;
                        if (temJustificativa) aluno.faltasJustificadas++;
                    }
                    else aluno.totalPresencas++;
                });
            });

            let totalLancs = aluno.totalFaltas + aluno.totalPresencas;
            aluno.porcentagemBruta = totalLancs > 0 ? ((aluno.totalFaltas / totalLancs) * 100).toFixed(1) + '%' : '0%';

            let faltasReais = aluno.totalFaltas - aluno.faltasJustificadas;
            aluno.porcentagemLiquida = totalLancs > 0 ? ((faltasReais / totalLancs) * 100).toFixed(1) + '%' : '0%';

            return aluno;
        }).sort((a, b) => a.nome.localeCompare(b.nome));

        // --- CÁLCULO DOS ALERTAS ---
        let disciplinasComRegistro = new Set();

        alunosArray.forEach(aluno => {
            Object.values(aluno.calendario).forEach(dia => {
                Object.keys(dia).forEach(disc => disciplinasComRegistro.add(disc));
            });
        });

        let disciplinasZeradas = Array.from(disciplinasLidasNoCabecalho).filter(d => !disciplinasComRegistro.has(d));
        let alunosZerados = alunosArray.filter(a => a.totalFaltas === 0 && a.totalPresencas === 0 && !a.nome.toUpperCase().includes("(TRANSFER"));

        // NOVO: Verifica alunos que têm lançamentos, mas faltam matérias específicas
        let alunosComDisciplinasFaltando = [];

        alunosArray.forEach(aluno => {
            // Se for transferido ou se estiver 100% zerado (já aparece no alerta 2), ignora.
            if (aluno.nome.toUpperCase().includes("(TRANSFER") || (aluno.totalFaltas === 0 && aluno.totalPresencas === 0)) return;

            let disciplinasDoAluno = new Set();
            Object.values(aluno.calendario).forEach(dia => {
                Object.keys(dia).forEach(disc => disciplinasDoAluno.add(disc));
            });

            // Matérias lidas que o aluno NÃO possui lançamentos
            let faltamNoAluno = Array.from(disciplinasLidasNoCabecalho).filter(d => !disciplinasDoAluno.has(d));

            // Remove as disciplinas que já estão 100% zeradas para a turma (para não poluir a lista)
            faltamNoAluno = faltamNoAluno.filter(d => !disciplinasZeradas.includes(d));

            if (faltamNoAluno.length > 0) {
                alunosComDisciplinasFaltando.push({ nome: aluno.nome, faltam: faltamNoAluno });
            }
        });

        // --- MONTA A CAIXA DE ALERTAS ---
        let htmlAlertas = '';
        if (alunosZerados.length > 0 || disciplinasZeradas.length > 0 || alunosComDisciplinasFaltando.length > 0) {
            htmlAlertas += `<div class="alerta-box">`;
            htmlAlertas += `<h3>⚠️ Alertas de Lançamento Pendente</h3>`;

            if (disciplinasZeradas.length > 0) {
                htmlAlertas += `<p><strong>Disciplinas SEM nenhum lançamento no período (em branco na turma):</strong></p><ul>`;
                disciplinasZeradas.forEach(d => htmlAlertas += `<li>${d}</li>`);
                htmlAlertas += `</ul>`;
            }

            if (alunosZerados.length > 0) {
                htmlAlertas += `<p><strong>Alunos SEM nenhuma presença ou falta lançada geral (Ignorando Transferidos):</strong></p><ul>`;
                alunosZerados.forEach(a => htmlAlertas += `<li>${a.nome}</li>`);
                htmlAlertas += `</ul>`;
            }

            // Renderiza o novo alerta individual
            if (alunosComDisciplinasFaltando.length > 0) {
                htmlAlertas += `<p><strong>Alunos com disciplinas específicas EM BRANCO (Furo parcial):</strong></p><ul>`;
                alunosComDisciplinasFaltando.forEach(item => {
                    htmlAlertas += `<li>${item.nome}: <span style="color:#d93025; font-weight:bold;">${item.faltam.join(', ')}</span></li>`;
                });
                htmlAlertas += `</ul>`;
            }

            htmlAlertas += `</div>`;
        }

        let html = `
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Calendário de Faltas - ${metadadosTurma.turma || 'Relatório'}</title>
                <style>
                    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f4f6f9; color: #333; padding: 20px; margin: 0; }
                    .header-top { text-align: center; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); margin-bottom: 20px; border-top: 5px solid #0056b3; }
                    .header-top h1 { color: #0056b3; margin: 0 0 10px 0; font-size: 24px; }

                    .container { max-width: 1200px; margin: 0 auto; }

                    .alerta-box { background: #fff3cd; border-left: 5px solid #ffeeba; padding: 15px 20px; margin-bottom: 20px; border-radius: 4px; color: #856404; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
                    .alerta-box h3 { margin-top: 0; color: #856404; }
                    .alerta-box ul { margin-top: 5px; margin-bottom: 15px; }

                    .aluno-card { background: #fff; border-radius: 8px; padding: 15px; margin-bottom: 20px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
                    .aluno-nome { font-size: 18px; font-weight: bold; margin-bottom: 10px; border-bottom: 2px solid #eee; padding-bottom: 5px; color: #2c3e50; }
                    .calendario-grid { display: flex; flex-wrap: wrap; gap: 6px; }

                    .dia-bloco { width: 45px; height: 48px; border-radius: 6px; display: flex; flex-direction: column; align-items: center; justify-content: center; border: 1px solid #ddd; font-weight: bold; cursor: pointer; position: relative; }
                    .data-label { font-size: 10px; opacity: 0.8; margin-bottom: 2px; }
                    .falta-count { font-size: 15px; }

                    .dia-bloco.atestado { background-color: #e8f0fe; border-color: #d2e3fc; color: #1967d2; }
                    .dia-bloco.falta-total { background-color: #d93025; border-color: #b31412; color: #ffffff; }
                    .dia-bloco.faltou { background-color: #ffeaea; border-color: #ffc2c2; color: #d93025; }
                    .dia-bloco.presente { background-color: #e6f4ea; border-color: #ceead6; color: #1e8e3e; }
                    .dia-bloco.sem-aula { background-color: #f1f3f4; border-color: #dadce0; color: #80868b; cursor: default; }

                    .tooltip-custom { visibility: hidden; background-color: rgba(30, 41, 59, 0.98); color: #fff; text-align: left; border-radius: 6px; padding: 10px 12px; position: absolute; z-index: 999; bottom: 110%; left: 50%; transform: translateX(-50%); font-size: 12px; font-weight: normal; box-shadow: 0 4px 10px rgba(0,0,0,0.3); opacity: 0; transition: opacity 0.2s, bottom 0.2s; pointer-events: none; white-space: nowrap; }
                    .tooltip-custom::after { content: ""; position: absolute; top: 100%; left: 50%; margin-left: -5px; border-width: 5px; border-style: solid; border-color: rgba(30, 41, 59, 0.98) transparent transparent transparent; }
                    .dia-bloco:hover .tooltip-custom { visibility: visible; opacity: 1; bottom: 120%; }

                    .resumo-table { width: 100%; border-collapse: collapse; background: #fff; box-shadow: 0 2px 5px rgba(0,0,0,0.1); margin-bottom: 40px; border-radius: 8px; overflow: hidden; }
                    .resumo-table th { background-color: #0056b3; color: white; padding: 12px; border: 1px solid #004494; font-size: 13px; }
                    .resumo-table td { padding: 10px; border: 1px solid #ddd; text-align: center; font-size: 13px; }
                    .resumo-table tr:nth-child(even) { background-color: #f9f9f9; }

                    .box-justificativas { background: #fff; border-left: 5px solid #1967d2; padding: 20px; border-radius: 8px; margin-bottom: 40px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
                    .box-justificativas h2 { margin-top: 0; color: #1967d2; font-size: 18px; }
                    .box-justificativas ul { margin-bottom: 0; }
                    .box-justificativas li { margin-bottom: 5px; font-size: 14px; }

                    @media print {
                        body { background: #fff; padding: 0; }
                        .header-top, .alerta-box, .box-justificativas { box-shadow: none; border: 1px solid #ccc; }
                        .aluno-card, .resumo-table { box-shadow: none; border: 1px solid #ccc; page-break-inside: avoid; }
                        .tooltip-custom { display: none !important; }
                        .dia-bloco { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                    }
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="header-top">
                        <h1>Calendário Consolidado de Faltas e Atestados</h1>
                        <p><strong>Turma:</strong> ${metadadosTurma.turma || '-'} &nbsp;&nbsp;|&nbsp;&nbsp; <strong>Turno:</strong> ${metadadosTurma.turno || '-'}</p>
                        <p><strong>Período:</strong> ${metadadosTurma.ip || '-'} a ${metadadosTurma.fp || '-'}</p>
                    </div>
                    ${htmlAlertas}
        `;

        alunosArray.forEach(aluno => {
            html += `<div class="aluno-card">`;
            html += `<div class="aluno-nome">${aluno.nome} <span style="font-size:12px; color:#7f8c8d; font-weight:normal;">(Cód: ${aluno.codigo})</span></div>`;
            html += `<div class="calendario-grid">`;

            datasOrdenadas.forEach(data => {
                let registrosDia = aluno.calendario[data] || {};
                let materiasComFalta = [];
                let materiasComPresenca = [];
                let materiasLancadas = Object.keys(registrosDia);

                materiasLancadas.forEach(disc => {
                    if (registrosDia[disc].toUpperCase() === 'F') materiasComFalta.push(disc);
                    else materiasComPresenca.push(disc);
                });

                let qtdFaltas = materiasComFalta.length;
                let temLancamento = materiasLancadas.length > 0;
                let tipoJustificativa = verificaAtestadoNoDia(data, aluno.atestados);

                let classeCss = ''; let textoExibicao = ''; let htmlTooltip = ''; let textoAlerta = '';

                if (tipoJustificativa) {
                    classeCss = 'atestado';
                    textoExibicao = 'A';
                    let det = "";
                    if (materiasComFalta.length > 0) det += `<br><em>Havia falta lançada em: ${materiasComFalta.join(', ')}</em>`;
                    htmlTooltip = `<div style="text-align:left;"><strong>Justificativa/Atestado:</strong><br><span style="color:#a8c7fa;">${tipoJustificativa}</span>${det}</div>`;
                    textoAlerta = `Atestado no dia:\n${tipoJustificativa}`;
                } else if (!temLancamento) {
                    classeCss = 'sem-aula';
                    textoExibicao = '-';
                    htmlTooltip = 'Nenhum lançamento';
                    textoAlerta = 'Nenhum lançamento neste dia.';
                } else {
                    if (qtdFaltas > 0) {
                        classeCss = (materiasComPresenca.length === 0) ? 'falta-total' : 'faltou';
                        textoExibicao = `${qtdFaltas}F`;
                    } else {
                        classeCss = 'presente';
                        textoExibicao = 'OK';
                    }
                    let divFaltasHtml = materiasComFalta.length > 0 ? `<div style="margin-top:6px; color:#ffb3b3;"><strong>Faltou em:</strong><br>- ${materiasComFalta.join('<br>- ')}</div>` : '';
                    let divPresencasHtml = materiasComPresenca.length > 0 ? `<div style="margin-top:6px; color:#85e085;"><strong>Presente em:</strong><br>- ${materiasComPresenca.join('<br>- ')}</div>` : '';
                    htmlTooltip = `<div style="text-align:left;"><strong>Resumo de Lançamentos</strong>${divFaltasHtml}${divPresencasHtml}</div>`;

                    let txtFaltas = materiasComFalta.length > 0 ? `\n\n[ FALTAS ]\n- ${materiasComFalta.join('\n- ')}` : '';
                    let txtPresencas = materiasComPresenca.length > 0 ? `\n\n[ PRESENÇAS ]\n- ${materiasComPresenca.join('\n- ')}` : '';
                    textoAlerta = `Resumo de Lançamentos do dia ${data}${txtFaltas}${txtPresencas}`;
                }

                html += `
                    <div class="dia-bloco ${classeCss}" onclick="if('${classeCss}' !== 'sem-aula') alert('${textoAlerta}')">
                        <span class="tooltip-custom">${htmlTooltip}</span>
                        <div class="data-label">${data}</div>
                        <div class="falta-count">${textoExibicao}</div>
                    </div>
                `;
            });

            html += `</div></div>`;
        });

        // --- Tabela de Resumo Final ---
        html += `
                    <h2 style="text-align:center; color:#0056b3; margin-top:40px;">Resumo da Turma</h2>
                    <table class="resumo-table">
                        <thead>
                            <tr>
                                <th>Turma</th>
                                <th>Turno</th>
                                <th>Código</th>
                                <th style="text-align: left;">Nome</th>
                                <th>Qtd. Presenças</th>
                                <th>Qtd. Faltas (Bruto)</th>
                                <th>% Faltas (Bruto)</th>
                                <th>Qtd. Faltas Justificadas</th>
                                <th style="background-color:#1967d2;">% Faltas c/ Desconto</th>
                            </tr>
                        </thead>
                        <tbody>
        `;

        alunosArray.forEach(aluno => {
            html += `
                            <tr>
                                <td>${metadadosTurma.turma || '-'}</td>
                                <td>${metadadosTurma.turno || '-'}</td>
                                <td>${aluno.codigo}</td>
                                <td style="text-align: left; font-weight: bold; color: #333;">${aluno.nome}</td>
                                <td style="color: #1e8e3e; font-weight: bold;">${aluno.totalPresencas}</td>
                                <td style="color: #d93025; font-weight: bold;">${aluno.totalFaltas}</td>
                                <td>${aluno.porcentagemBruta}</td>
                                <td style="color: #1967d2; font-weight: bold;">${aluno.faltasJustificadas}</td>
                                <td style="font-weight: bold;">${aluno.porcentagemLiquida}</td>
                            </tr>
            `;
        });
        html += `</tbody></table>`;

        // Seção da Lista de Justificativas / Atestados
        let alunosComAtestado = alunosArray.filter(a => a.atestados && a.atestados.length > 0);

        if (alunosComAtestado.length > 0) {
            html += `
                    <div class="box-justificativas">
                        <h2>Lista de Atestados e Justificativas no Período</h2>
                        <ul>
            `;
            alunosComAtestado.forEach(aluno => {
                aluno.atestados.forEach(at => {
                    let periodoStr = (at.dataIni === at.dataFim || !at.dataFim) ? at.dataIni : `${at.dataIni} a ${at.dataFim}`;
                    html += `<li><strong>${aluno.nome}</strong> (${aluno.codigo}): ${periodoStr} - <span style="color:#555;">${at.tipoJust}</span></li>`;
                });
            });
            html += `</ul></div>`;
        }

        html += `
                </div>
            </body>
            </html>
        `;

        const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const linkDeDownload = document.createElement('a');
        let nomeArquivoLimpo = (metadadosTurma.turma || "Turma_Desconhecida").replace(/[^a-zA-Z0-9]/g, '_');
        linkDeDownload.href = url;
        linkDeDownload.download = `Relatorio_Faltas_${nomeArquivoLimpo}.html`;

        document.body.appendChild(linkDeDownload);
        linkDeDownload.click();
        document.body.removeChild(linkDeDownload);
        URL.revokeObjectURL(url);
    }

    // --- FUNÇÕES DE SETUP DO SISTEMA (IFRAME E PAINEL) ---
    function prepararIframe() {
        let container = document.getElementById('containerIframeImpressao');
        if (!container) {
            container = document.createElement('div');
            container.id = 'containerIframeImpressao';
            container.style.cssText = 'position: absolute; width: 0; height: 0; overflow: hidden; visibility: hidden; opacity: 0; border: none;';

            let iframe = document.createElement('iframe');
            iframe.id = 'iframeImpressao';
            iframe.name = 'iframeImpressao';
            iframe.style.cssText = 'width: 100%; height: 100%; border: none;';

            container.appendChild(iframe);
            document.body.appendChild(container);
        }
        return container;
    }

    function criarPainelLateral() {
        if (document.getElementById('painelAutomacaoSeduc')) return;

        const painel = document.createElement('div');
        painel.id = 'painelAutomacaoSeduc';
        painel.style.cssText = 'position: fixed; top: 50px; left: 10px; width: 220px; background: #f9f9f9; border: 2px solid #0056b3; border-radius: 8px; padding: 15px; z-index: 10000; box-shadow: 2px 2px 10px rgba(0,0,0,0.3); font-family: Arial, sans-serif;';

        let htmlPainel = `
            <h3 style="margin: 0 0 10px 0; font-size: 14px; color: #0056b3; text-align: center;">Automação de Diários</h3>
            <div id="containerBotoesBimestre" style="display: flex; flex-direction: column; gap: 8px; margin-bottom: 15px;"></div>
            <div style="margin-bottom: 5px; font-size: 12px; font-weight: bold; color: #333;" id="statusProgresso">Aguardando...</div>
            <div style="width: 100%; height: 15px; background: #ddd; border-radius: 5px; overflow: hidden; margin-bottom: 10px;">
                <div id="barraProgresso" style="width: 0%; height: 100%; background: #28a745; transition: width 0.3s ease;"></div>
            </div>
            <button id="btnPararAutomacao" style="width: 100%; padding: 8px; background: #dc3545; color: white; border: none; border-radius: 4px; cursor: pointer; display: none; margin-bottom: 5px; font-weight: bold;">Parar Loop</button>
            <button id="btnVerRelatorio" style="width: 100%; padding: 8px; background: #28a745; color: white; border: none; border-radius: 4px; cursor: pointer; display: none; font-weight: bold;">Baixar Relatório HTML</button>
        `;

        painel.innerHTML = htmlPainel;
        document.body.appendChild(painel);

        const selectBimestre = document.getElementById('vGEDPERCOD');
        const containerBotoes = document.getElementById('containerBotoesBimestre');

        if (selectBimestre && containerBotoes) {
            Array.from(selectBimestre.options).forEach(opcao => {
                if (opcao.value !== "0") {
                    let btn = document.createElement('button');
                    btn.innerText = `Processar ${opcao.text}`;
                    btn.style.cssText = 'padding: 8px; background: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px;';
                    btn.onclick = (e) => { e.preventDefault(); iniciarLoop(opcao.value, opcao.text); };
                    containerBotoes.appendChild(btn);
                }
            });
        }

        document.getElementById('btnPararAutomacao').onclick = () => {
            executandoLoop = false;
            document.getElementById('statusProgresso').innerText = "Interrompido!";
            document.getElementById('statusProgresso').style.color = "#dc3545";
            document.getElementById('btnPararAutomacao').style.display = 'none';
        };

        document.getElementById('btnVerRelatorio').onclick = () => { gerarEBaixarRelatorioHTML(); };
    }

    // --- 4. Lógica do Loop Principal ---
    async function iniciarLoop(valorBimestre, textoBimestre) {
        if (executandoLoop) return;
        executandoLoop = true;

        alunosDados = {};
        metadadosTurma = { ip: "", fp: "", turma: "", turno: "" };
        disciplinasLidasNoCabecalho.clear();

        const selectBimestre = document.getElementById('vGEDPERCOD');
        const btnParar = document.getElementById('btnPararAutomacao');
        const btnVerRelatorio = document.getElementById('btnVerRelatorio');
        const barra = document.getElementById('barraProgresso');
        const status = document.getElementById('statusProgresso');

        btnParar.style.display = 'block';
        btnVerRelatorio.style.display = 'none';
        barra.style.width = '0%';
        status.style.color = "#333";

        status.innerText = "Selecionando opção 'Preenchido'...";
        const radioPreenchido = document.querySelector('input[name="vOPCAOREL"][value="1"]');
        if (radioPreenchido && !radioPreenchido.checked) {
            radioPreenchido.checked = true;
            radioPreenchido.dispatchEvent(new Event('focus', { bubbles: true }));
            radioPreenchido.click();
            radioPreenchido.dispatchEvent(new Event('change', { bubbles: true }));
            radioPreenchido.dispatchEvent(new Event('blur', { bubbles: true }));
            await delay(1500);
        }

        selectBimestre.value = valorBimestre;
        selectBimestre.dispatchEvent(new Event('focus'));
        selectBimestre.dispatchEvent(new Event('change'));
        selectBimestre.dispatchEvent(new Event('blur'));

        status.innerText = `Carregando ${textoBimestre}...`;
        await delay(2000);

        const disciplinasSelectAtualizado = document.getElementById('vDISCIPLINAAREACOD');
        const disciplinas = Array.from(disciplinasSelectAtualizado.options).filter(opt => opt.value !== "0");
        const total = disciplinas.length;

        prepararIframe();
        const iframeImpressao = document.getElementById('iframeImpressao');

        // ======= FASE 1: LER DIÁRIOS =======
        for (let i = 0; i < total; i++) {
            if (!executandoLoop) break;
            const disc = disciplinas[i];
            barra.style.width = `${Math.round(((i + 1) / total) * 100)}%`;
            status.innerText = `Lendo Diário: ${disc.text.substring(0, 15)}...`;

            const discSelect = document.getElementById('vDISCIPLINAAREACOD');
            discSelect.value = disc.value;
            discSelect.dispatchEvent(new Event('focus'));
            discSelect.dispatchEvent(new Event('change'));
            discSelect.dispatchEvent(new Event('blur'));

            await delay(500);
            while (!isNotificationHidden(document)) { await delay(300); }

            // Tratamento Específico para Período Único
            const selectBimestreLoop = document.getElementById('vGEDPERCOD');
            if (selectBimestreLoop) {
                let optionExists = Array.from(selectBimestreLoop.options).some(opt => opt.value === valorBimestre);

                if (!optionExists) {
                    let option21Exists = Array.from(selectBimestreLoop.options).some(opt => opt.value === "21");
                    if (option21Exists) {
                        selectBimestreLoop.value = "21";
                        selectBimestreLoop.dispatchEvent(new Event('focus'));
                        selectBimestreLoop.dispatchEvent(new Event('change'));
                        selectBimestreLoop.dispatchEvent(new Event('blur'));
                        await delay(500);
                        while (!isNotificationHidden(document)) { await delay(300); }
                    }
                } else if (selectBimestreLoop.value !== valorBimestre) {
                    selectBimestreLoop.value = valorBimestre;
                    selectBimestreLoop.dispatchEvent(new Event('focus'));
                    selectBimestreLoop.dispatchEvent(new Event('change'));
                    selectBimestreLoop.dispatchEvent(new Event('blur'));
                    await delay(500);
                    while (!isNotificationHidden(document)) { await delay(300); }
                }
            }

            const btnImprimir = document.querySelector('input[name="BIMPRIMIR"]');
            if (btnImprimir) {
                const form = btnImprimir.closest('form') || document.forms[0];
                let targetOriginal = form ? form.getAttribute('target') : null;
                if (form) form.setAttribute('target', 'iframeImpressao');

                btnImprimir.value = "1";

                let timeoutId;
                let promiseLoadIframe = new Promise((resolve, reject) => {
                    const aoCarregar = () => {
                        iframeImpressao.removeEventListener('load', aoCarregar);
                        clearTimeout(timeoutId);
                        resolve();
                    };
                    iframeImpressao.addEventListener('load', aoCarregar);

                    timeoutId = setTimeout(() => {
                        iframeImpressao.removeEventListener('load', aoCarregar);
                        reject(new Error("Timeout_Sigeduca"));
                    }, 40000);
                });

                btnImprimir.click();

                if (form) {
                    setTimeout(() => {
                        if (targetOriginal === null) form.removeAttribute('target');
                        else form.setAttribute('target', targetOriginal);
                    }, 1000);
                }

                try {
                    await promiseLoadIframe;
                    await delay(500);
                } catch (err) {
                    executandoLoop = false;
                    status.innerText = "Erro de Conexão!";
                    status.style.color = "#dc3545";
                    btnParar.style.display = 'none';
                    alert("Aviso: O diário demorou mais de 40 segundos para carregar.\n\nPossível problema de conexão com o SIGEDUCA. A extração foi abortada. Por favor, recarregue a página e tente novamente.");
                    return;
                }
            }
            extrairDadosIframe();
        }

        // ======= FASE 2: LER ATESTADOS =======
        if (executandoLoop) {
            let listaCodigos = Object.keys(alunosDados);
            if(listaCodigos.length > 0) {
                await extrairAtestadosIframe(listaCodigos);
            }
        }

        // ======= FINALIZAR =======
        if (executandoLoop) {
            status.innerText = "Extração Concluída!";
            status.style.color = "#28a745";
            btnVerRelatorio.style.display = 'block';

            gerarEBaixarRelatorioHTML();
            alert(`Processo 100% concluído!\n\nDiários e Atestados da turma '${metadadosTurma.turma || "Desconhecida"}' foram analisados.\nO HTML deve ter sido baixado.`);
        }

        executandoLoop = false;
        btnParar.style.display = 'none';
    }

    const windowOpenOriginal = window.open;
    window.open = function(url) {
        if (executandoLoop) {
            prepararIframe();
            document.getElementById('iframeImpressao').src = url;
            return null;
        }
        return windowOpenOriginal.apply(this, arguments);
    };

    function inicializar() {
        criarPainelLateral();
    }

    const observer = new MutationObserver(inicializar);
    observer.observe(document.body, { childList: true, subtree: true });
    inicializar();

})();
