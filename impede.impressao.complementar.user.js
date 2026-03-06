// ==UserScript==
// @name         Não imprimir calendário
// @namespace    http://tampermonkey.net/
// @version      2026-02-27
// @description  Bloqueia a impressão automática
// @author       You
// @match        http://sigeduca.seduc.mt.gov.br/grh/hwmgrhcalendarioimp.aspx?*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=gov.br
// @run-at       document-start
// @updateURL    https://github.com/lksoumon/acompanhamentoDiarioGEDeFICAI/raw/refs/heads/main/impede.impressao.complementar.user.js
// @downloadURL  https://github.com/lksoumon/acompanhamentoDiarioGEDeFICAI/raw/refs/heads/main/impede.impressao.complementar.user.js
// ==/UserScript==

(function() {
    'use strict';

    // Sobrescreve a função nativa de impressão na janela principal antes que a página carregue
    unsafeWindow.print = function() {
        console.log("Impressão automática bloqueada pelo Tampermonkey.");
    };

})();
