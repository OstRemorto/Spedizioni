export function inizializzaImport(onImportate) {
    const fileInputClienti = document.getElementById("importClienti");
    if (fileInputClienti) {
        fileInputClienti.addEventListener("change", gestisciImportExcel("clienti", onImportate), false);
    }
    const fileInputDestinazioni = document.getElementById("importDestinazioni");
    if (fileInputDestinazioni) {
        fileInputDestinazioni.addEventListener("change", gestisciImportExcel("clientiZ", onImportate), false);
    }

    const scriptSheetJs = document.querySelector('script[src*="xlsx.full.min.js"]');
    if (scriptSheetJs) {
        scriptSheetJs.addEventListener("error", () => {
            mostraStatoImport("Errore: impossibile caricare la libreria Excel. Verifica la connessione e ricarica il pannello.", true);
        });
    }

    let clienti = null;
    let destinazioni = [];
    try {
        const clientiSalvati = localStorage.getItem("anagraficaLogistica");
        if (clientiSalvati) clienti = JSON.parse(clientiSalvati);
        const destinazioniSalvate = localStorage.getItem("anagraficaDestinazioniLogistica");
        if (destinazioniSalvate) destinazioni = JSON.parse(destinazioniSalvate);
        if (clienti || destinazioni.length > 0) {
            aggiornaInterfacciaImport(clienti || [], destinazioni);
        }
    } catch (errore) {
        mostraStatoImport(`Impossibile caricare le anagrafiche salvate: ${errore.message}`, true);
    }

    return { clienti, destinazioni };
}

function gestisciImportExcel(tipo, onImportate) {
    return function(e) {
        const file = e.target.files[0];
        if (!file) return;

        if (!/\.(xlsx|xls)$/i.test(file.name)) {
            mostraStatoImport("Seleziona un file Excel (.xlsx o .xls).", true);
            return;
        }
        const etichetta = tipo === "clienti" ? "Clienti" : "ClientiZ";
        mostraStatoImport(`Lettura di "${file.name}" (${etichetta}) in corso...`);
        const reader = new FileReader();

        reader.onload = function(evento) {
            try {
                if (typeof XLSX === "undefined") {
                    throw new Error("La libreria Excel non è disponibile.");
                }
                const workbook = XLSX.read(new Uint8Array(evento.target.result), { type: "array" });
                const nomeFoglio = workbook.SheetNames[0];
                if (!nomeFoglio) {
                    throw new Error("Il file Excel non contiene fogli.");
                }

                const righe = XLSX.utils.sheet_to_json(workbook.Sheets[nomeFoglio]);
                const anagrafica = normalizzaAnagrafica(righe);
                salvaAnagraficaImportata(tipo, anagrafica, onImportate);
            } catch (errore) {
                mostraStatoImport(`Importazione non riuscita: ${errore.message}`, true);
            }
        };
        reader.onerror = () => {
            mostraStatoImport("Errore durante la lettura del file Excel. Riprova selezionando il file.", true);
        };
        reader.onabort = () => {
            mostraStatoImport("Lettura del file Excel annullata.", true);
        };
        reader.readAsArrayBuffer(file);
    };
}

function normalizzaIntestazione(chiave) {
    return chiave.normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");
}

function normalizzaAnagrafica(righe) {
    return righe.map(riga => {
        const colonne = Object.keys(riga).reduce((mappa, chiave) => {
            mappa[normalizzaIntestazione(chiave)] = riga[chiave];
            return mappa;
        }, {});
        return {
            nome: String(colonne.ragionesociale || colonne.ragsoc || colonne.cliente || "").trim(),
            via: String(colonne.indirizzo || colonne.via || "").trim(),
            comune: String(colonne.citta || colonne.comune || "").trim()
        };
    }).filter(cliente => cliente.nome || cliente.via || cliente.comune);
}

function salvaAnagraficaImportata(tipo, anagrafica, onImportate) {
    if (anagrafica.length === 0) {
        throw new Error("Il primo foglio è vuoto o non contiene righe con Rag. Soc., Città e Indirizzo. Nessuna anagrafica è stata aggiornata.");
    }

    const chiave = tipo === "clienti" ? "anagraficaLogistica" : "anagraficaDestinazioniLogistica";
    try {
        localStorage.setItem(chiave, JSON.stringify(anagrafica));
        localStorage.setItem("dataUltimoImport", new Date().toLocaleDateString());
        const clientiSalvati = tipo === "clienti"
            ? anagrafica
            : JSON.parse(localStorage.getItem("anagraficaLogistica") || "null");
        const destinazioniSalvate = tipo === "clientiZ"
            ? anagrafica
            : JSON.parse(localStorage.getItem("anagraficaDestinazioniLogistica") || "[]");
        onImportate(clientiSalvati, destinazioniSalvate);
        aggiornaInterfacciaImport(clientiSalvati || [], destinazioniSalvate);
        mostraStatoImport(`${tipo === "clienti" ? "Clienti" : "ClientiZ"} importati correttamente: ${anagrafica.length} record.`);
    } catch (errore) {
        if (errore.name === "QuotaExceededError") {
            mostraStatoImport("Il browser non ha spazio sufficiente per salvare l’anagrafica. Il file è stato letto, ma i dati non sono stati salvati.", true);
            return;
        }
        throw errore;
    }
}

function mostraStatoImport(messaggio, errore = false) {
    const info = document.getElementById("infoImport");
    info.textContent = messaggio;
    info.style.color = errore ? "#a4262c" : "#605e5c";
}

function aggiornaInterfacciaImport(clienti, destinazioni) {
    const info = document.getElementById("infoImport");
    const data = localStorage.getItem("dataUltimoImport") || "recente";
    info.textContent = `Anagrafiche attive: Clienti ${clienti.length}, ClientiZ ${destinazioni.length} (aggiornate al: ${data})`;
    info.style.color = "#107c10";
}
