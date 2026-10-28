export function inizializzaImport(onImportate) {
    const clientiSalvati = localStorage.getItem("anagraficaLogistica");
    const destinazioniSalvate = localStorage.getItem("anagraficaDestinazioniLogistica");
    const anagrafiche = {
        clienti: clientiSalvati ? JSON.parse(clientiSalvati) : null,
        destinazioni: destinazioniSalvate ? JSON.parse(destinazioniSalvate) : []
    };

    document.getElementById("importClienti").addEventListener("change", e => {
        importaFileExcel(e, "clienti", "Clienti", anagrafiche, onImportate);
    });
    document.getElementById("importDestinazioni").addEventListener("change", e => {
        importaFileExcel(e, "destinazioni", "ClientiZ", anagrafiche, onImportate);
    });

    if (anagrafiche.clienti || anagrafiche.destinazioni.length > 0) {
        mostraAnagrafiche(anagrafiche.clienti || [], anagrafiche.destinazioni);
    }
    return anagrafiche;
}

function importaFileExcel(evento, tipo, nomeAnagrafica, anagrafiche, onImportate) {
    const file = evento.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(eventoLettura) {
        try {
            const data = new Uint8Array(eventoLettura.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const foglio = workbook.Sheets[workbook.SheetNames[0]];
            const datiRaw = XLSX.utils.sheet_to_json(foglio);

            const anagrafica = datiRaw.map(riga => ({
                nome: riga["Ragione Sociale"] || riga["Rag. Soc."] || riga["Cliente"] || "",
                via: riga["Indirizzo"] || riga["Via"] || "",
                comune: riga["Città"] || riga["Citta"] || riga["Comune"] || ""
            }));

            const chiave = tipo === "clienti" ? "anagraficaLogistica" : "anagraficaDestinazioniLogistica";
            localStorage.setItem(chiave, JSON.stringify(anagrafica));
            localStorage.setItem("dataUltimoImport", new Date().toLocaleDateString());

            if (tipo === "clienti") {
                anagrafiche.clienti = anagrafica;
            } else {
                anagrafiche.destinazioni = anagrafica;
            }
            onImportate(anagrafiche.clienti, anagrafiche.destinazioni);
            mostraAnagrafiche(anagrafiche.clienti || [], anagrafiche.destinazioni);
            mostraStato(`${nomeAnagrafica} importati: ${anagrafica.length} righe.`);
        } catch (errore) {
            mostraStato(`Errore durante l’importazione: ${errore.message}`, true);
        }
    };
    reader.onerror = () => mostraStato("Errore durante la lettura del file Excel.", true);
    reader.readAsArrayBuffer(file);
}

function mostraAnagrafiche(clienti, destinazioni) {
    const data = localStorage.getItem("dataUltimoImport") || "recente";
    mostraStato(`Anagrafiche caricate: Clienti ${clienti.length}, ClientiZ ${destinazioni.length} (aggiornate al: ${data}).`);
}

function mostraStato(messaggio, errore = false) {
    const info = document.getElementById("infoImport");
    info.textContent = messaggio;
    info.style.color = errore ? "#a4262c" : "#107c10";
}
