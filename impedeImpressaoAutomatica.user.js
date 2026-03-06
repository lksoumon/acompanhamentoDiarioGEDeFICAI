// ==UserScript==
// @name         GED - Bloquear Impressão Automática do Relatório (Definitivo)
// @namespace    http://tampermonkey.net/
// @version      2.0
// @description  Fura o sandbox do Tampermonkey e injeta o bloqueio direto no HTML
// @author       Você
// @match        http://sigeduca.seduc.mt.gov.br/ged/hwgedteladocumento.aspx*
// @run-at       document-start
// @icon         https://www.google.com/s2/favicons?sz=64&domain=gov.br
// @grant        unsafeWindow
// @updateURL    https://github.com/lksoumon/acompanhamentoDiarioGEDeFICAI/raw/refs/heads/main/impedeImpressaoAutomatica.user.js
// @downloadURL  https://github.com/lksoumon/acompanhamentoDiarioGEDeFICAI/raw/refs/heads/main/impedeImpressaoAutomatica.user.js
// ==/UserScript==

(function() {
    'use strict';

    // 1. Tenta bloquear diretamente a janela real da página (ignorando o sandbox)
    try {
        unsafeWindow.print = function() { console.log("Print nativo bloqueado via unsafeWindow!"); };
    } catch(e) {}

    // 2. Cria um script físico para injetar no código-fonte da página
    const scriptDeBloqueio = document.createElement('script');
    scriptDeBloqueio.textContent = `
        // Sobrescreve e tranca a função print no ambiente real da página
        window.print = function() { console.log('Print bloqueado pelo script injetado!'); };
        try {
            Object.defineProperty(window, 'print', {
                value: function() { console.log('Print completamente desativado.'); },
                writable: false,
                configurable: false
            });
        } catch(e) {}
    `;

    // Injeta o script o mais rápido possível (antes do sistema conseguir rodar qualquer coisa)
    const observer = new MutationObserver(() => {
        if (document.head || document.documentElement) {
            (document.head || document.documentElement).appendChild(scriptDeBloqueio);
            observer.disconnect(); // Para de observar assim que injetar
        }
    });
    observer.observe(document, { childList: true, subtree: true });

    // 3. Continua caçando o infame <body onload="window.print()"> do GeneXus
    window.addEventListener('DOMContentLoaded', () => {
        if (document.body) {
            const onloadAttr = document.body.getAttribute('onload');
            if (onloadAttr && onloadAttr.toLowerCase().includes('print')) {
                document.body.removeAttribute('onload');
            }
            // Força a anulação se foi criado via JavaScript
            document.body.onload = null;
        }
    });

})();
