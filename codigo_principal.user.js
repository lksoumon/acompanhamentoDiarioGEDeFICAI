// ==UserScript==
// @name         Diario GED para acompanhamento de lançamentos e FICAI (Completo)
// @namespace    http://tampermonkey.net/
// @version      v2.5-FaltasReais-LegendaAuditoria
// @description  Atestados e Justificativas em Azul. Detecta aulas geminadas. Exporta Faltas REAIS (sem justif) em JSON. Legenda na auditoria.
// @author       Lucas Monteiro
// @match        http://sigeduca.seduc.mt.gov.br/ged/hwgedemitediarioclasse.aspx?*
// @grant        none
// @updateURL    https://github.com/lksoumon/acompanhamentoDiarioGEDeFICAI/raw/refs/heads/main/codigo_principal.user.js
// @downloadURL  https://github.com/lksoumon/acompanhamentoDiarioGEDeFICAI/raw/refs/heads/main/codigo_principal.user.js
// ==/UserScript==

(function() {
    'use strict';

    let executandoLoop = false;
    const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    let alunosDados = {};
    let metadadosTurma = { ano: "", ip: "", fp: "", turma: "", turno: "", matrizDesc: "" };
    let disciplinasLidasNoCabecalho = new Set();
    let diasLecionadosGlobais = {};
    let diasLetivosGlobais = [];
    let infoServidores = { lista: [], chDisciplina: {}, chTotal: 0 };

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
        if (!dataString) return "";
        var strLimpa = dataString.replace(/[\n\r]/g, ' ').replace(/\s+/g, ' ').trim();
        var partes = strLimpa.split(' ');
        if (partes.length < 2) {
             if (strLimpa.includes('/')) partes = strLimpa.split('/');
             else return dataString;
        }
        var dia = parseInt(partes[0]); var mes = parseInt(partes[1]);
        if (isNaN(dia) || isNaN(mes)) return dataString;
        return (dia < 10 ? '0' + dia : dia) + '/' + (mes < 10 ? '0' + mes : mes);
    }

    function isNotificationHidden(docObject) {
        var notification = docObject.getElementById('gx_ajax_notification');
        if (notification) return docObject.defaultView.getComputedStyle(notification).getPropertyValue('display') === 'none';
        return true;
    }

    function parseDataBR(dataStr) {
        let partes = dataStr.trim().split('/');
        let dia = parseInt(partes[0]); let mes = parseInt(partes[1]);
        let ano = partes.length === 3 ? parseInt(partes[2]) : (parseInt(metadadosTurma.ano) || new Date().getFullYear());
        return new Date(ano, mes - 1, dia);
    }

    function getInicioSemana(date) {
        let d = new Date(date); let day = d.getDay();
        let diff = d.getDate() - day + (day === 0 ? -6 : 1);
        return new Date(d.setDate(diff));
    }

    function qtdeDiasLetivosNaSemana(strSemana) {
        let [inicioStr, fimStr] = strSemana.split(' a ');
        let dIni = parseDataBR(inicioStr); let dFim = parseDataBR(fimStr);
        let cont = 0;
        diasLetivosGlobais.forEach(dl => {
            let d = parseDataBR(dl);
            if (d >= dIni && d <= dFim) cont++;
        });
        return cont;
    }

    function verificaAtestadoNoDia(dataCalendario, atestadosArray) {
        if (!atestadosArray || atestadosArray.length === 0) return null;
        let [diaC, mesC] = dataCalendario.split('/').map(Number);
        let numDiaAtual = mesC * 100 + diaC;
        for (let at of atestadosArray) {
            if (!at.dataIni) continue;
            let [diaI, mesI] = at.dataIni.split('/').map(Number);
            let [diaF, mesF] = (at.dataFim || at.dataIni).split('/').map(Number);
            let numDiaIni = mesI * 100 + diaI; let numDiaFim = mesF * 100 + diaF;
            if (numDiaAtual >= numDiaIni && numDiaAtual <= numDiaFim) return at.tipoJust;
        }
        return null;
    }

    function prepararIframe() {
        let container = document.getElementById('containerIframeImpressao');
        if (!container) {
            container = document.createElement('div');
            container.id = 'containerIframeImpressao';
            container.style.cssText = 'position: absolute; width: 0; height: 0; overflow: hidden; visibility: hidden; opacity: 0; border: none;';
            let iframe = document.createElement('iframe');
            iframe.id = 'iframeImpressao'; iframe.name = 'iframeImpressao';
            iframe.style.cssText = 'width: 100%; height: 100%; border: none;';
            container.appendChild(iframe); document.body.appendChild(container);
        }
        return container;
    }

    function criarPainelLateral() {
        if (document.getElementById('painelAutomacaoSeduc')) return;
        const painel = document.createElement('div');
        painel.id = 'painelAutomacaoSeduc';
        painel.style.cssText = 'position: fixed; top: 50px; left: 10px; width: 280px; background: #f9f9f9; border: 2px solid #0056b3; border-radius: 8px; padding: 15px; z-index: 10000; box-shadow: 2px 2px 10px rgba(0,0,0,0.3); font-family: Arial, sans-serif;';

        painel.innerHTML = `
            <h3 style="margin: 0 0 10px 0; font-size: 14px; color: #0056b3; text-align: center;">Automação de Diários</h3>
            <div style="background: #e8f0fe; padding: 8px; border-radius: 4px; margin-bottom: 15px; font-size: 11px;">
                <strong>Extração Local:</strong><br>
                <label style="display:block; margin-top:4px; cursor:pointer;"><input type="checkbox" id="chkOpProfessores" checked> Quadro Servidores</label>
                <label style="display:block; margin-top:4px; cursor:pointer;"><input type="checkbox" id="chkOpAuditoria" checked> Auditoria Semanal</label>
                <label style="display:block; margin-top:4px; cursor:pointer;"><input type="checkbox" id="chkOpAtestados" checked> Atestados/Justificativas</label>
            </div>
            <div id="containerBotoesBimestre" style="display: flex; flex-direction: column; gap: 8px; margin-bottom: 15px;"></div>
            <div style="margin-bottom: 5px; font-size: 12px; font-weight: bold; color: #333;" id="statusProgresso">Aguardando...</div>
            <div style="width: 100%; height: 15px; background: #ddd; border-radius: 5px; overflow: hidden; margin-bottom: 10px;">
                <div id="barraProgresso" style="width: 0%; height: 100%; background: #28a745; transition: width 0.3s ease;"></div>
            </div>
            <button id="btnPararAutomacao" style="width: 100%; padding: 8px; background: #dc3545; color: white; border: none; border-radius: 4px; cursor: pointer; display: none; margin-bottom: 5px; font-weight: bold;">Parar Loop</button>
        `;
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
            executandoLoop = false; document.getElementById('statusProgresso').innerText = "Interrompido!";
            document.getElementById('statusProgresso').style.color = "#dc3545"; document.getElementById('btnPararAutomacao').style.display = 'none';
        };
    }

    function extrairDadosServidores(iframeDoc) {
        let tabela = iframeDoc.getElementById('Grid1ContainerTbl');
        if (!tabela) return false;
        for (let i = 1; i < tabela.rows.length; i++) {
            let cols = tabela.rows[i].cells; if (cols.length < 18) continue;
            let servidor = cols[11].innerText.trim(); let inicio = cols[4].innerText.trim();
            let fim = cols[5].innerText.trim(); let substituicao = cols[14].innerText.trim();
            let disciplina = cols[16].innerText.trim(); let chAula = parseInt(cols[17].innerText.trim(), 10);
            infoServidores.lista.push({ servidor, inicio, fim, substituicao, disciplina, chAula });
            if (!isNaN(chAula) && disciplina) {
                if (!infoServidores.chDisciplina[disciplina] || chAula > infoServidores.chDisciplina[disciplina]) infoServidores.chDisciplina[disciplina] = chAula;
            }
        }
        infoServidores.chTotal = Object.values(infoServidores.chDisciplina).reduce((a, b) => a + b, 0);
        return true;
    }

    function extrairDadosIframe() {
        const iframeDoc = document.getElementById('iframeImpressao')?.contentWindow?.document;
        if (!iframeDoc) return false;

        const tabelas = iframeDoc.getElementsByTagName("table");
        let extraiuAlgo = false;
        let disciplinaAtual = "Desconhecida";
        let cabecalhosLidosNesteIframe = new Set();

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
                            if (!diasLecionadosGlobais[disciplinaAtual]) diasLecionadosGlobais[disciplinaAtual] = [];
                        }

                        if (!metadadosTurma.turma) {
                            var tempCabecalho = linhasDd[1].getElementsByTagName("td")[1].getElementsByTagName("span")[2].textContent.trim();
                            metadadosTurma.ip = recortarString(tempCabecalho, '', 'FP:').replace('IP:', '').trim();
                            metadadosTurma.fp = recortarString(tempCabecalho, 'FP:', 'Turma:');
                            metadadosTurma.turma = recortarString(tempCabecalho, 'Turma:', 'Turno:');
                            metadadosTurma.turno = tempCabecalho.split('Turno:')[1].trim();
                        }
                    } catch (e) { }
                }
                else if (headerText === "TF" || headerText === "Situação") {
                    var linhas = minhaTabela.getElementsByTagName("tr");
                    var datas = [];

                    for (var i = 0; i < linhas.length; i++) {
                        if (i == 0) continue;
                        if (i == 1) {
                            var dias = linhas[i].getElementsByTagName("td");
                            for (var j = 0; j < dias.length; j++) {
                                var textoData = dias[j].textContent || dias[j].innerText;
                                let dFormatada = formatarData(textoData);

                                if (/^\d{2}\/\d{2}$/.test(dFormatada)) {
                                    datas.push(dFormatada);
                                    if (!cabecalhosLidosNesteIframe.has(disciplinaAtual)) {
                                        diasLecionadosGlobais[disciplinaAtual].push(dFormatada);
                                    }
                                }
                            }
                            cabecalhosLidosNesteIframe.add(disciplinaAtual);
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

                                if (!alunosDados[codigoEstudante].calendario[datas[k]]) {
                                    alunosDados[codigoEstudante].calendario[datas[k]] = {};
                                }
                                if (!alunosDados[codigoEstudante].calendario[datas[k]][disciplinaAtual]) {
                                    alunosDados[codigoEstudante].calendario[datas[k]][disciplinaAtual] = [];
                                }

                                if (status !== "") {
                                    alunosDados[codigoEstudante].calendario[datas[k]][disciplinaAtual].push(status);
                                }
                            }
                        }
                    }
                    extraiuAlgo = true;
                }
            }
        }
        return extraiuAlgo;
    }

    async function extrairAtestadosIframe(listaAlunosIds) {
        const iframe = document.getElementById('iframeImpressao');
        const status = document.getElementById('statusProgresso'); const barra = document.getElementById('barraProgresso');
        status.innerText = "Carregando tela de Atestados..."; barra.style.width = `0%`;
        iframe.src = 'http://sigeduca.seduc.mt.gov.br/ged/hwmgedatestado.aspx';

        let iframeDoc = iframe.contentWindow.document; let tentativasLoad = 0;
        while (!iframeDoc.getElementById('vGEDALUCOD') && tentativasLoad < 30) { await delay(1000); iframeDoc = iframe.contentWindow.document; tentativasLoad++; }
        if (tentativasLoad >= 30) return; await delay(1500);

        for (let i = 0; i < listaAlunosIds.length; i++) {
            if (!executandoLoop) break;
            let alunoCodigo = listaAlunosIds[i];
            status.innerText = `Lendo Atestados: Aluno ${i + 1} de ${listaAlunosIds.length}`; barra.style.width = `${Math.round(((i + 1) / listaAlunosIds.length) * 100)}%`;

            let inputAluno = iframeDoc.getElementById('vGEDALUCOD');
            if (inputAluno) {
                inputAluno.value = alunoCodigo;
                if ("createEvent" in iframeDoc) { var evt = iframeDoc.createEvent("HTMLEvents"); evt.initEvent("change", false, true); inputAluno.dispatchEvent(evt); }
                else { inputAluno.fireEvent("onchange"); }
            } else { continue; }

            await delay(300);
            let btnConsultar = iframeDoc.getElementsByName('BCONSULTAR')[0] || iframeDoc.querySelector('.btnConsultar');
            if (btnConsultar) btnConsultar.click();
            await delay(300); while (!isNotificationHidden(iframeDoc)) { await delay(300); }

            let docTabela = iframeDoc;
            if (iframe.contentWindow.frames.length > 0 && iframe.contentWindow.frames[0].document.getElementById('GriddetalhesContainerTbl')) {
                docTabela = iframe.contentWindow.frames[0].document;
            }

            let selectPag = docTabela.getElementById('vPAG'); let totalPaginas = selectPag ? selectPag.options.length : 1;
            for (let p = 1; p <= totalPaginas; p++) {
                if (!executandoLoop) break;
                if (p > 1 && selectPag) {
                    selectPag.value = p.toString();
                    if ("createEvent" in docTabela) { var evtPag = docTabela.createEvent("HTMLEvents"); evtPag.initEvent("change", false, true); selectPag.dispatchEvent(evtPag); }
                    else { selectPag.fireEvent("onchange"); }
                    try { docTabela.defaultView.gx.evt.execEvt('EVPAG.CLICK.', selectPag); } catch(e){}
                    await delay(300); while (!isNotificationHidden(docTabela)) { await delay(300); }
                }

                let tabelaDetalhes = docTabela.getElementById('GriddetalhesContainerTbl');
                if (tabelaDetalhes && tabelaDetalhes.rows.length > 1) {
                    for (let n = 1; n < tabelaDetalhes.rows.length; n++) {
                        let numStr = ("0000" + n).slice(-4);
                        try {
                            let dataIni = docTabela.getElementById('span_vGEDATEPERINI_' + numStr)?.textContent.trim() || '';
                            let dataFim = docTabela.getElementById('span_vGEDATEPERFIN_' + numStr)?.textContent.trim() || '';
                            let tipoJust = docTabela.getElementById('span_vGEDATETIPO_' + numStr)?.textContent.trim() || '';
                            if (dataIni) alunosDados[alunoCodigo].atestados.push({ dataIni, dataFim, tipoJust });
                        } catch (err) {}
                    }
                }
            }
            await delay(300);
        }
    }

    // --- HTML INDIVIDUAL COM SOMA E EXPORTAÇÃO JSON ---
    function gerarEBaixarRelatorioHTML(opts) {
        let anoDaTurma = parseInt(metadadosTurma.ano) || new Date().getFullYear();
        let todasAsDatas = new Set();
        Object.values(alunosDados).forEach(aluno => {
            Object.keys(aluno.calendario).forEach(data => { if (/^\d{2}\/\d{2}$/.test(data.trim())) todasAsDatas.add(data.trim()); });
        });

        let datasOrdenadas = Array.from(todasAsDatas).sort((a, b) => {
            let [da, ma] = a.split('/'); let [db, mb] = b.split('/'); return (ma + da).localeCompare(mb + db);
        });

        let semanas = {};
        datasOrdenadas.forEach(data => {
            let dObj = parseDataBR(data); let seg = getInicioSemana(dObj);
            let sab = new Date(seg); sab.setDate(seg.getDate() + 5);
            let idSemana = `${seg.getDate().toString().padStart(2,'0')}/${(seg.getMonth()+1).toString().padStart(2,'0')} a ${sab.getDate().toString().padStart(2,'0')}/${(sab.getMonth()+1).toString().padStart(2,'0')}`;
            if(!semanas[idSemana]) semanas[idSemana] = { datas: [] };
            semanas[idSemana].datas.push(data);
        });

        let semanasKeys = Object.keys(semanas).sort((a, b) => {
            let [diaA, mesA] = a.split(' a ')[0].split('/'); let [diaB, mesB] = b.split(' a ')[0].split('/');
            return (mesA+diaA).localeCompare(mesB+diaB);
        });

        let datasPorMes = {};
        datasOrdenadas.forEach(data => {
            let partes = data.split('/'); let mes = partes[1];
            let key = `${anoDaTurma}-${mes}`;
            if (!datasPorMes[key]) datasPorMes[key] = { mesStr: mes, anoNum: anoDaTurma, datas: [] };
            datasPorMes[key].datas.push(data);
        });
        const nomeMeses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

        // Prepara dados dos alunos e estrutura de Faltas (JSON) EXCLUINDO justificadas
        let alunosArray = Object.keys(alunosDados).map(cod => {
            let aluno = { codigo: cod, ...alunosDados[cod], totalFaltas: 0, totalPresencas: 0, faltasJustificadas: 0, relatorioFaltasJson: {} };

            Object.keys(aluno.calendario).forEach(data => {
                if (!/^\d{2}\/\d{2}$/.test(data.trim())) return;
                let registrosDoDia = aluno.calendario[data];
                let temJustificativa = verificaAtestadoNoDia(data, aluno.atestados);
                let matComFaltaNoDia = [];

                Object.keys(registrosDoDia).forEach(disc => {
                    let statusArr = registrosDoDia[disc];
                    if (!Array.isArray(statusArr)) statusArr = [statusArr];

                    statusArr.forEach(status => {
                        if (status.toUpperCase() === 'F') {
                            aluno.totalFaltas++;
                            if (temJustificativa) {
                                aluno.faltasJustificadas++;
                            } else {
                                // MÁGICA: Só vai para o JSON se a falta NÃO for justificada
                                matComFaltaNoDia.push(disc);
                            }
                        } else {
                            aluno.totalPresencas++;
                        }
                    });
                });

                if (matComFaltaNoDia.length > 0) {
                    aluno.relatorioFaltasJson[data] = matComFaltaNoDia;
                }
            });

            let totalLancs = aluno.totalFaltas + aluno.totalPresencas;
            aluno.porcentagemBruta = totalLancs > 0 ? ((aluno.totalFaltas / totalLancs) * 100).toFixed(1) + '%' : '0%';
            let faltasReais = aluno.totalFaltas - aluno.faltasJustificadas;
            aluno.porcentagemLiquida = totalLancs > 0 ? ((faltasReais / totalLancs) * 100).toFixed(1) + '%' : '0%';
            return aluno;
        }).sort((a, b) => a.nome.localeCompare(b.nome));

        let html = `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Auditoria GED - ${metadadosTurma.matrizDesc || 'Matriz'} - ${metadadosTurma.turma || 'Turma'}</title>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/tablesort/5.2.1/tablesort.min.js"></script>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/tablesort/5.2.1/sorts/tablesort.number.min.js"></script>
            <style>
                body { font-family: 'Segoe UI', sans-serif; background: #f4f6f9; color: #333; padding: 20px; margin: 0; }
                .header-top { text-align: center; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); margin-bottom: 20px; border-top: 5px solid #0056b3; }
                .header-top h1 { color: #0056b3; margin: 0 0 5px 0; font-size: 24px; }
                .header-top h3 { color: #666; margin: 0; font-size: 14px; font-weight: normal; }
                .container { max-width: 1300px; margin: 0 auto; }
                table { width: 100%; border-collapse: collapse; background: #fff; box-shadow: 0 2px 5px rgba(0,0,0,0.1); margin-bottom: 40px; border-radius: 8px; overflow: hidden; }
                th { background-color: #0056b3; color: white; padding: 10px; border: 1px solid #004494; font-size: 12px; cursor: pointer; }
                td { padding: 8px; border: 1px solid #ddd; text-align: center; font-size: 12px; }
                tr:nth-child(even) { background-color: #f9f9f9; }
                .alerta-celula { background-color: #ffcccc !important; color: #b30000 !important; font-weight: bold; }
                .alerta-celula-amarelo { background-color: #fff3cd !important; color: #856404 !important; font-weight: bold; }
                .aluno-card { background: #fff; border-radius: 8px; padding: 15px; margin-bottom: 20px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
                .aluno-nome { font-size: 16px; font-weight: bold; margin-bottom: 10px; border-bottom: 2px solid #eee; padding-bottom: 5px; color: #2c3e50; display: flex; justify-content: space-between; align-items: center;}
                .btn-copy { background-color: #1967d2; color: #fff; border: none; padding: 4px 10px; border-radius: 4px; font-size: 11px; cursor: pointer; transition: 0.2s; font-weight: bold;}
                .btn-copy:hover { background-color: #11468f; }
                .calendario-wrapper { display: flex; flex-wrap: wrap; gap: 20px; }
                .mes-container { border: 1px solid #e0e0e0; padding: 10px; border-radius: 6px; background: #fafafa; min-width: 250px; }
                .calendario-grid-mes { display: grid; grid-template-columns: repeat(7, 34px); gap: 4px; justify-content: center; text-align:center; }
                .dia-semana-cabecalho { font-size: 10px; font-weight: bold; color: #7f8c8d; }
                .dia-bloco { width: 34px; height: 38px; border-radius: 4px; display: flex; flex-direction: column; align-items: center; justify-content: center; border: 1px solid #ddd; font-weight: bold; cursor: pointer; position: relative; background: #fff; }
                .data-label { font-size: 8px; opacity: 0.8; margin-bottom: 2px; }
                .falta-count { font-size: 12px; }
                .dia-bloco.vazio { background: transparent; border: 1px dashed #ddd; color: #ccc; cursor: default; }
                .dia-bloco.atestado { background-color: #e8f0fe; border-color: #d2e3fc; color: #1967d2; }
                .dia-bloco.falta-total { background-color: #d93025; border-color: #b31412; color: #ffffff; }
                .dia-bloco.faltou { background-color: #ffeaea; border-color: #ffc2c2; color: #d93025; }
                .dia-bloco.presente { background-color: #e6f4ea; border-color: #ceead6; color: #1e8e3e; }
                .tooltip-custom { visibility: hidden; background-color: rgba(30, 41, 59, 0.98); color: #fff; text-align: left; border-radius: 6px; padding: 10px; position: absolute; z-index: 999; bottom: 110%; left: 50%; transform: translateX(-50%); font-size: 12px; font-weight: normal; opacity: 0; pointer-events: none; white-space: nowrap; }
                .dia-bloco:hover .tooltip-custom { visibility: visible; opacity: 1; bottom: 120%; }
                .legenda-auditoria { background: #fff3cd; border-left: 5px solid #ffeeba; padding: 12px; margin-bottom: 20px; border-radius: 4px; color: #856404; font-size: 13px; }
            </style>
            </head><body><div class="container">
                <div class="header-top">
                    <h1>Auditoria Integrada - Turma: ${metadadosTurma.turma || '-'}</h1>
                    <h3>${metadadosTurma.matrizDesc || 'Matriz não identificada'}</h3>
                </div>`;

        if (opts.professores && infoServidores.lista.length > 0) {
            html += `<h2>1. Quadro de Professores Vinculados</h2><table><thead><tr><th style="text-align: left;">Servidor</th><th style="text-align: left;">Componente Curricular</th><th>C.H. Semanal</th><th>Início</th><th>Fim</th><th>Substituto?</th></tr></thead><tbody>`;
            infoServidores.lista.forEach(s => { html += `<tr><td style="text-align: left;">${s.servidor}</td><td style="text-align: left;">${s.disciplina}</td><td>${s.chAula}</td><td>${s.inicio}</td><td>${s.fim}</td><td>${s.substituicao}</td></tr>`; });
            html += `</tbody></table>`;
        }

        if (opts.auditoria) {
            let disciplinasAuditadas = Array.from(disciplinasLidasNoCabecalho).sort();

            html += `<h2>2. Auditoria Semanal de Lançamentos</h2>
                     <div class="legenda-auditoria">
                        <strong>⚠️ Entendendo as cores:</strong><br>
                        - <strong style="color: #b30000;">Vermelho:</strong> O professor lançou aulas a <strong>MENOS</strong> em uma semana normal (5 dias letivos), ou lançou <strong>A MAIS</strong> que a base.<br>
                        - <strong style="color: #856404;">Amarelo:</strong> Lançou aulas a <strong>MENOS</strong>, mas a semana tem <strong>menos de 5 dias letivos</strong> (feriados, etc). Precisa de conferência manual. (Obs: Se ele atingiu a carga base mesmo no feriado, fica sem cor).
                     </div>
                     <div style="overflow-x: auto;"><table style="white-space: nowrap;">
                    <thead><tr><th style="text-align: left;">Componente</th><th>C.H. Base</th>`;

            semanasKeys.forEach(sem => {
                let numDiasLetivos = qtdeDiasLetivosNaSemana(sem);
                html += `<th>${sem.replace(' a ', '<br>a ')}<br><span style="font-size:10px; font-weight:normal;">(${numDiasLetivos} DL)</span></th>`;
            });
            html += `<th>Total Lançado</th></tr></thead><tbody>`;

            let totaisPorSemana = new Array(semanasKeys.length).fill(0);
            let totalGeralLancadoEscola = 0;

            disciplinasAuditadas.forEach(d => {
                let nomeDisciplinaFormatado = d.length > 32 ? d.substring(0, 32) + '...' : d;
                let chEsperada = infoServidores.chDisciplina[d] || '?';
                let rowHtml = `<tr><td style="text-align: left; font-weight: bold;">${nomeDisciplinaFormatado}</td><td>${chEsperada}</td>`;
                let totalLancadoMateria = 0;

                semanasKeys.forEach((idSemana, idxSemana) => {
                    let datasDaSemana = semanas[idSemana]?.datas || [];
                    let numDiasLetivos = qtdeDiasLetivosNaSemana(idSemana);
                    let qtde = 0;

                    if (diasLecionadosGlobais[d]) {
                        diasLecionadosGlobais[d].forEach(dataLancada => {
                            if (datasDaSemana.includes(dataLancada)) qtde++;
                        });
                    }

                    totalLancadoMateria += qtde;
                    totaisPorSemana[idxSemana] += qtde;

                    let classe = '';
                    if (chEsperada !== '?') {
                        if (qtde > chEsperada) {
                            classe = 'alerta-celula'; // Sempre erro vermelho se passou da Base
                        } else if (qtde < chEsperada) {
                            classe = (numDiasLetivos < 5) ? 'alerta-celula-amarelo' : 'alerta-celula';
                        }
                    }
                    rowHtml += `<td class="${classe}">${qtde}</td>`;
                });
                totalGeralLancadoEscola += totalLancadoMateria;
                rowHtml += `<td style="font-weight:bold;">${totalLancadoMateria}</td></tr>`;
                html += rowHtml;
            });

            // LINHA DE SOMA TOTAL
            html += `<tr style="background-color: #e8f0fe;"><td style="text-align: left; font-weight: bold; color: #1967d2;">TOTAL LANÇADO NA SEMANA</td><td style="font-weight: bold; color: #1967d2;">-</td>`;
            totaisPorSemana.forEach(tot => {
                html += `<td style="font-weight:bold; color: #1967d2;">${tot}</td>`;
            });
            html += `<td style="font-weight:bold; color: #1967d2; font-size: 14px;">${totalGeralLancadoEscola}</td></tr>`;

            html += `</tbody></table></div>`;
        }

        html += `<h2>3. Calendário Detalhado por Aluno</h2>`;
        alunosArray.forEach(aluno => {

            // Objeto JSON pronto
            let objetoJsonFaltas = {
                codigo: aluno.codigo,
                nome: aluno.nome,
                turma: metadadosTurma.turma || '',
                faltas: aluno.relatorioFaltasJson // Aqui já estão sem as justificadas!
            };
            let jsonStringSafe = JSON.stringify(objetoJsonFaltas).replace(/'/g, "&apos;");

            html += `<div class="aluno-card">
                        <div class="aluno-nome">
                            <span>${aluno.nome} <span style="font-size:12px; color:#7f8c8d; font-weight:normal;">(Cód: ${aluno.codigo})</span></span>
                            <button class="btn-copy" data-json='${jsonStringSafe}' onclick="copiarJson(this)">📋 Copiar Faltas JSON</button>
                        </div>
                        <div class="calendario-wrapper">`;

            Object.keys(datasPorMes).sort().forEach(key => {
                let info = datasPorMes[key]; let mIdx = parseInt(info.mesStr, 10) - 1; let anoNum = parseInt(info.anoNum, 10);

                html += `<div class="mes-container"><h4 class="mes-titulo">${nomeMeses[mIdx]} ${anoNum}</h4><div class="calendario-grid-mes">
                    <div class="dia-semana-cabecalho">D</div><div class="dia-semana-cabecalho">S</div><div class="dia-semana-cabecalho">T</div><div class="dia-semana-cabecalho">Q</div><div class="dia-semana-cabecalho">Q</div><div class="dia-semana-cabecalho">S</div><div class="dia-semana-cabecalho">S</div>`;

                let primeiroDia = new Date(anoNum, mIdx, 1).getDay(); let totalDiasNoMes = new Date(anoNum, mIdx + 1, 0).getDate();
                for (let b = 0; b < primeiroDia; b++) html += `<div class="dia-bloco vazio" style="border:none;"></div>`;

                for (let d = 1; d <= totalDiasNoMes; d++) {
                    let dataStr = `${d.toString().padStart(2,'0')}/${info.mesStr}`;
                    if (info.datas.includes(dataStr)) {
                        let registrosDia = aluno.calendario[dataStr] || {};
                        let materiasComFalta = []; let materiasComPresenca = [];
                        let materiasLancadas = Object.keys(registrosDia);

                        materiasLancadas.forEach(disc => {
                            let stArr = registrosDia[disc];
                            if(!Array.isArray(stArr)) stArr = [stArr];
                            stArr.forEach(st => {
                                if (st.toUpperCase() === 'F') materiasComFalta.push(disc);
                                else materiasComPresenca.push(disc);
                            });
                        });

                        let qtdFaltas = materiasComFalta.length; let temLancamento = materiasLancadas.length > 0;
                        let tipoJustificativa = verificaAtestadoNoDia(dataStr, aluno.atestados);
                        let classeCss = ''; let textoExibicao = ''; let htmlTooltip = '';

                        if (tipoJustificativa) {
                            let strNormalizada = tipoJustificativa.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
                            classeCss = 'atestado'; textoExibicao = strNormalizada.includes('medico') ? 'A' : 'J';
                            let det = materiasComFalta.length > 0 ? `<br><em>Falta em: ${materiasComFalta.join(', ')}</em>` : "";
                            htmlTooltip = `<div style="text-align:left;"><strong>Justificativa:</strong><br><span style="color:#a8c7fa;">${tipoJustificativa}</span>${det}</div>`;
                        } else if (!temLancamento) {
                            classeCss = 'sem-aula'; textoExibicao = '-'; htmlTooltip = 'Nenhum lançamento';
                        } else {
                            if (qtdFaltas > 0) { classeCss = (materiasComPresenca.length === 0) ? 'falta-total' : 'faltou'; textoExibicao = `${qtdFaltas}F`; }
                            else { classeCss = 'presente'; textoExibicao = 'OK'; }
                            let divFaltasHtml = materiasComFalta.length > 0 ? `<div style="margin-top:6px; color:#ffb3b3;"><strong>Faltou:</strong><br>- ${materiasComFalta.join('<br>- ')}</div>` : '';
                            let divPresencasHtml = materiasComPresenca.length > 0 ? `<div style="margin-top:6px; color:#85e085;"><strong>Presente:</strong><br>- ${materiasComPresenca.join('<br>- ')}</div>` : '';
                            htmlTooltip = `<div style="text-align:left;"><strong>Resumo</strong>${divFaltasHtml}${divPresencasHtml}</div>`;
                        }
                        html += `<div class="dia-bloco ${classeCss}"><span class="tooltip-custom">${htmlTooltip}</span><div class="data-label">${d.toString().padStart(2,'0')}</div><div class="falta-count">${textoExibicao}</div></div>`;
                    } else {
                        html += `<div class="dia-bloco vazio"><div class="data-label">${d.toString().padStart(2,'0')}</div></div>`;
                    }
                }
                html += `</div></div>`;
            });
            html += `</div></div>`;
        });

        html += `<table style="margin-top: 30px;"><thead><tr><th data-sort-method="number">Código</th><th style="text-align: left;">Nome do Estudante</th><th data-sort-method="number">Qtd. Presenças</th><th data-sort-method="number">Qtd. Faltas (Bruto)</th><th>% Faltas (Bruto)</th><th data-sort-method="number">Faltas Justificadas</th><th>% Faltas c/ Desconto</th></tr></thead><tbody>`;
        alunosArray.forEach(aluno => {
            html += `<tr><td>${aluno.codigo}</td><td style="text-align: left; font-weight: bold;">${aluno.nome}</td><td style="color: #1e8e3e; font-weight: bold;">${aluno.totalPresencas}</td><td style="color: #d93025; font-weight: bold;">${aluno.totalFaltas}</td><td>${aluno.porcentagemBruta}</td><td style="color: #1967d2; font-weight: bold;">${aluno.faltasJustificadas}</td><td style="font-weight: bold;">${aluno.porcentagemLiquida}</td></tr>`;
        });
        html += `</tbody></table>
            <script>
                document.addEventListener('DOMContentLoaded', function() {
                    document.querySelectorAll('table').forEach(function(table) { new Tablesort(table); });
                });
                function copiarJson(btn) {
                    navigator.clipboard.writeText(btn.getAttribute('data-json')).then(() => {
                        let originalText = btn.innerText;
                        btn.innerText = '✅ Copiado!';
                        btn.style.backgroundColor = '#28a745';
                        setTimeout(() => {
                            btn.innerText = originalText;
                            btn.style.backgroundColor = '#1967d2';
                        }, 2000);
                    });
                }
            </script>
            </div></body></html>`;

        const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const linkDeDownload = document.createElement('a');

        let nomeArquivoLimpo = (metadadosTurma.turma || "Auditoria").replace(/[^a-zA-Z0-9]/g, '_').replace(/_+/g, '_');
        let matrizLimpa = (metadadosTurma.matrizDesc || "Matriz").replace(/[^a-zA-Z0-9]/g, '_').replace(/_+/g, '_');

        linkDeDownload.href = url;
        linkDeDownload.download = `Auditoria_${matrizLimpa}_${nomeArquivoLimpo}.html`;
        document.body.appendChild(linkDeDownload); linkDeDownload.click(); document.body.removeChild(linkDeDownload); URL.revokeObjectURL(url);
    }

    async function iniciarLoop(valorBimestre, textoBimestre) {
        if (executandoLoop) return; executandoLoop = true;
        alunosDados = {}; metadadosTurma = { ano: "", ip: "", fp: "", turma: "", turno: "", matrizDesc: "" };
        disciplinasLidasNoCabecalho.clear(); diasLecionadosGlobais = {}; diasLetivosGlobais = [];
        infoServidores = { lista: [], chDisciplina: {}, chTotal: 0 };

        const btnParar = document.getElementById('btnPararAutomacao'); const barra = document.getElementById('barraProgresso'); const status = document.getElementById('statusProgresso');
        const optProfessores = document.getElementById('chkOpProfessores').checked; const optAuditoria = document.getElementById('chkOpAuditoria').checked; const optAtestados = document.getElementById('chkOpAtestados').checked;

        btnParar.style.display = 'block'; barra.style.width = '0%'; status.style.color = "#333";

        try {
            let actionCompleto = document.getElementById("MAINFORM").action; let parametrosAction = actionCompleto.split('?')[1].split(',');
            let cidade = parametrosAction[0]; metadadosTurma.ano = parametrosAction[1]; let escola = parametrosAction[2];
            let sala = parametrosAction[3]; let turnoNum = parametrosAction[4]; let chaveDesc1 = parametrosAction[5]; let matriz = parametrosAction[6];
            let turnoTexto = document.getElementById("span_vGERTRNCOD") ? document.getElementById("span_vGERTRNCOD").innerText.trim() : "VESPERTINO";
            let matrizTexto = document.getElementById("span_vGERDESCMAT") ? document.getElementById("span_vGERDESCMAT").innerText.trim() : "";
            metadadosTurma.matrizDesc = matrizTexto.replace(/[^a-zA-Z0-9\s]/g, '').trim().replace(/\s+/g, '_');

            prepararIframe(); const iframeImpressao = document.getElementById('iframeImpressao');

            if (optAuditoria) {
                status.innerText = "Coletando Calendário Escolar...";
                iframeImpressao.src = `http://sigeduca.seduc.mt.gov.br/grh/hwmgrhcalendarioimp.aspx?${metadadosTurma.ano},${escola}`;
                await new Promise((resolve) => {
                    const aoCarregar = () => { iframeImpressao.removeEventListener('load', aoCarregar); resolve(); };
                    iframeImpressao.addEventListener('load', aoCarregar);
                    setTimeout(() => { iframeImpressao.removeEventListener('load', aoCarregar); resolve(); }, 20000);
                });
                await delay(1500);

                try {
                    let targetDoc = iframeImpressao.contentWindow.document;
                    if (iframeImpressao.contentWindow.frames.length > 0) { try { targetDoc = iframeImpressao.contentWindow.frames[0].document; } catch(e){} }

                    targetDoc.querySelectorAll('[id*="vDATAINDICE_"]').forEach(span => {
                        let match = span.id.match(/span_(W\d{4}\d{4})vDATAINDICE_(\d{4})/);
                        if (match) {
                            let legendaEl = targetDoc.getElementById(`${match[1]}TLEGENDA_${match[2]}`);
                            if (legendaEl) {
                                let discrica = legendaEl.innerText.trim();
                                if(discrica === "L" || discrica.includes("- L") || discrica.includes("L -")) {
                                    const partesData = span.innerText.trim().split('/');
                                    if (partesData.length >= 2) {
                                        let ano = partesData[2] || metadadosTurma.ano; if (ano.length === 2) ano = '20' + ano;
                                        diasLetivosGlobais.push(`${partesData[0]}/${partesData[1]}/${ano}`);
                                    }
                                }
                            }
                        }
                    });
                    diasLetivosGlobais = [...new Set(diasLetivosGlobais)];
                } catch (e) { console.error("Erro calendário:", e); }
            }

            if (optProfessores || optAuditoria) {
                status.innerText = "Coletando Quadro de Professores...";
                iframeImpressao.src = `http://sigeduca.seduc.mt.gov.br/ged/hwmgrhturmaservidor.aspx?${metadadosTurma.ano},${escola},${cidade},${sala},,${turnoNum},${chaveDesc1},${turnoTexto},HWMGrhLotTurma.aspx%3f0%2c0%2c0%2c0,${matriz},,,`;
                await new Promise((resolve) => {
                    const aoCarregar = () => { iframeImpressao.removeEventListener('load', aoCarregar); resolve(); };
                    iframeImpressao.addEventListener('load', aoCarregar); setTimeout(() => { iframeImpressao.removeEventListener('load', aoCarregar); resolve(); }, 20000);
                });
                await delay(1500);
                if (iframeImpressao.contentWindow) extrairDadosServidores(iframeImpressao.contentWindow.document);
            }

        } catch (e) {}

        status.innerText = "Selecionando opção 'Preenchido'...";
        const radioPreenchido = document.querySelector('input[name="vOPCAOREL"][value="1"]');
        if (radioPreenchido && !radioPreenchido.checked) {
            radioPreenchido.checked = true; radioPreenchido.dispatchEvent(new Event('focus', { bubbles: true }));
            radioPreenchido.click(); radioPreenchido.dispatchEvent(new Event('change', { bubbles: true }));
            radioPreenchido.dispatchEvent(new Event('blur', { bubbles: true })); await delay(1500);
        }

        let selectBimestreInicial = document.getElementById('vGEDPERCOD');
        if (selectBimestreInicial) {
            let optReqExiste = Array.from(selectBimestreInicial.options).some(opt => opt.value === valorBimestre);
            let opt21Existe = Array.from(selectBimestreInicial.options).some(opt => opt.value === "21");
            if (!optReqExiste && opt21Existe) { valorBimestre = "21"; textoBimestre = "Período Único"; }
        }

        const selectBimestre = document.getElementById('vGEDPERCOD');
        selectBimestre.value = valorBimestre; selectBimestre.dispatchEvent(new Event('focus')); selectBimestre.dispatchEvent(new Event('change')); selectBimestre.dispatchEvent(new Event('blur'));

        status.innerText = `Carregando ${textoBimestre}...`; await delay(2000);
        const disciplinasSelectAtualizado = document.getElementById('vDISCIPLINAAREACOD');
        const disciplinas = Array.from(disciplinasSelectAtualizado.options).filter(opt => opt.value !== "0");
        const total = disciplinas.length; const iframeImpressao = document.getElementById('iframeImpressao');

        for (let i = 0; i < total; i++) {
            if (!executandoLoop) break;
            const disc = disciplinas[i]; barra.style.width = `${Math.round(((i + 1) / total) * 100)}%`; status.innerText = `Lendo Diário: ${disc.text.substring(0, 15)}...`;

            const discSelect = document.getElementById('vDISCIPLINAAREACOD');
            discSelect.value = disc.value; discSelect.dispatchEvent(new Event('focus')); discSelect.dispatchEvent(new Event('change')); discSelect.dispatchEvent(new Event('blur'));

            await delay(500); while (!isNotificationHidden(document)) { await delay(300); }

            const selectBimestreLoop = document.getElementById('vGEDPERCOD');
            if (selectBimestreLoop && selectBimestreLoop.value !== valorBimestre) {
                selectBimestreLoop.value = valorBimestre; selectBimestreLoop.dispatchEvent(new Event('focus')); selectBimestreLoop.dispatchEvent(new Event('change'));
                selectBimestreLoop.dispatchEvent(new Event('blur')); await delay(500); while (!isNotificationHidden(document)) { await delay(300); }
            }

            const btnImprimir = document.querySelector('input[name="BIMPRIMIR"]');
            if (btnImprimir) {
                const form = btnImprimir.closest('form') || document.forms[0];
                let targetOriginal = form ? form.getAttribute('target') : null;
                if (form) form.setAttribute('target', 'iframeImpressao');
                btnImprimir.value = "1";

                let timeoutId;
                let promiseLoadIframe = new Promise((resolve, reject) => {
                    const aoCarregar = () => { iframeImpressao.removeEventListener('load', aoCarregar); clearTimeout(timeoutId); resolve(); };
                    iframeImpressao.addEventListener('load', aoCarregar);
                    timeoutId = setTimeout(() => { iframeImpressao.removeEventListener('load', aoCarregar); reject(new Error("Timeout_Sigeduca")); }, 40000);
                });

                btnImprimir.click();
                if (form) setTimeout(() => { if (targetOriginal === null) form.removeAttribute('target'); else form.setAttribute('target', targetOriginal); }, 1000);

                try { await promiseLoadIframe; await delay(500); }
                catch (err) { executandoLoop = false; status.innerText = "Erro!"; status.style.color = "#dc3545"; btnParar.style.display = 'none'; return; }
            }
            extrairDadosIframe();
        }

        if (executandoLoop && optAtestados) {
            let listaCodigos = Object.keys(alunosDados);
            if(listaCodigos.length > 0) await extrairAtestadosIframe(listaCodigos);
        }

        if (executandoLoop) {
            status.innerText = "Extração Concluída!"; status.style.color = "#28a745";

            const urlParams = new URLSearchParams(window.location.search);
            if (urlParams.get('autorun') === '1') {
                let dbExport = {
                    metadadosTurma: metadadosTurma,
                    infoServidores: infoServidores,
                    alunosDados: alunosDados,
                    diasLecionadosGlobais: diasLecionadosGlobais,
                    diasLetivosGlobais: diasLetivosGlobais
                };
                window.parent.postMessage({ type: 'GED_AUTO_DONE', payload: dbExport }, "*");
            } else {
                gerarEBaixarRelatorioHTML({
                    professores: optProfessores,
                    auditoria: optAuditoria,
                    atestados: optAtestados
                });
            }
        }
        executandoLoop = false; btnParar.style.display = 'none';
    }

    const windowOpenOriginal = window.open;
    window.open = function(url) {
        if (executandoLoop) { prepararIframe(); document.getElementById('iframeImpressao').src = url; return null; }
        return windowOpenOriginal.apply(this, arguments);
    };

    function inicializar() {
        criarPainelLateral();
        const urlParams = new URLSearchParams(window.location.search);
        if (urlParams.get('autorun') === '1') {
            document.getElementById('painelAutomacaoSeduc').style.display = 'none';
            let bim = urlParams.get('bimestre') || "1";
            setTimeout(() => { iniciarLoop(bim, "Bimestre " + bim); }, 1500);
        }
    }

    const observer = new MutationObserver(inicializar);
    observer.observe(document.body, { childList: true, subtree: true });
    inicializar();

})();
