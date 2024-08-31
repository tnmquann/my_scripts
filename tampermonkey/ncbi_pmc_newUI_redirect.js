// ==UserScript==
// @name         Redirect to NCBI PMC new UI (2024)
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  Redirect from ncbi.nlm.nih.gov/pmc/* to pmc.ncbi.nlm.nih.gov
// @author       https://github.com/tnmquann/
// @match        https://www.ncbi.nlm.nih.gov/pmc/*
// @grant        none
// @run-at       document-start
// ==/UserScript==

(function() {
    'use strict';

    var oldURL = window.location.href;
    var newURL = oldURL.replace('www.ncbi.nlm.nih.gov/pmc', 'pmc.ncbi.nlm.nih.gov');

    window.location.replace(newURL);
})();