const statoSelettoreImportKey = "spedizioniStatoSelettoreImport";

export function inizializzaImport(onImportate) {
    const fileInput = document.getElementById("importExcel");
    if (fileInput) {
        fileInput.addEventListener("click", registraAperturaSelettore, false);
        fileInput.addEventListener("change", gestisciImportExcel(onImportate), false);
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

    mostraStatoSelettoreImport();
    return { clienti, destinazioni };
}

function registraAperturaSelettore() {
    try {
        localStorage.setItem(statoSelettoreImportKey, JSON.stringify({
            fase: "aperto",
            timestamp: Date.now()
        }));
        mostraStatoImport("Il clic sul selettore è arrivato all’add-in. Seleziona il file Excel.");
    } catch (errore) {
        mostraStatoImport(`Errore diagnostico del selettore: ${errore.message}`, true);
    }
}

function gestisciImportExcel(onImportate) {
    return function(e) {
        const file = e.target.files[0];
        if (!file) return;

        try {
            localStorage.setItem(statoSelettoreImportKey, JSON.stringify({
                fase: "selezionato",
                nomeFile: file.name,
                timestamp: Date.now()
            }));

        } catch (errore) {
            mostraStatoImport(`Impossibile registrare la selezione del file: ${errore.message}`, true);
            return;
        }

        if (!/\.(xlsx|xls)$/i.test(file.name)) {
            mostraStatoImport("Seleziona un file Excel (.xlsx o .xls).", true);
            localStorage.removeItem(statoSelettoreImportKey);
            return;
        }
        const tipoImport = document.getElementById("tipoAnagraficaImport").value;
        mostraStatoImport(`Lettura di "${file.name}" come ${tipoImport === "clienti" ? "anagrafica Clienti" : "anagrafica ClientiZ"} in corso...`);
        const reader = new FileReader();

        reader.onload = function(evento) {
            try {
                if (typeof XLSX === "undefined") {
                    throw new Error("La libreria Excel non è disponibile. Ricarica il pannello e riprova.");
                }
                const workbook = XLSX.read(new Uint8Array(evento.target.result), { type: "array" });
                const nomeFoglio = workbook.SheetNames[0];
                if (!nomeFoglio) {
                    throw new Error("Il file Excel non contiene fogli.");
                }

                const righe = XLSX.utils.sheet_to_json(workbook.Sheets[nomeFoglio]);
                const anagrafica = normalizzaAnagrafica(righe);
                salvaAnagraficaImportata(tipoImport, anagrafica, onImportate);
            } catch (errore) {
                localStorage.removeItem(statoSelettoreImportKey);
                mostraStatoImport(`Importazione non riuscita: ${errore.message}`, true);
            }
        };
        reader.onerror = () => {
            localStorage.removeItem(statoSelettoreImportKey);
            mostraStatoImport("Errore durante la lettura del file Excel. Riprova selezionando il file.", true);
        };
        reader.onabort = () => {
            localStorage.removeItem(statoSelettoreImportKey);
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
    localStorage.setItem(chiave, JSON.stringify(anagrafica));
    localStorage.setItem("dataUltimoImport", new Date().toLocaleDateString());
    localStorage.removeItem(statoSelettoreImportKey);
    const clientiSalvati = tipo === "clienti"
        ? anagrafica
        : JSON.parse(localStorage.getItem("anagraficaLogistica") || "null");
    const destinazioniSalvate = tipo === "clientiZ"
        ? anagrafica
        : JSON.parse(localStorage.getItem("anagraficaDestinazioniLogistica") || "[]");
    onImportate(clientiSalvati, destinazioniSalvate);
    aggiornaInterfacciaImport(clientiSalvati || [], destinazioniSalvate);
    mostraStatoImport(`${tipo === "clienti" ? "Clienti" : "ClientiZ"} importati correttamente: ${anagrafica.length} record.`);
}

function mostraStatoImport(messaggio, errore = false) {
    const info = document.getElementById("infoImport");
    info.textContent = messaggio;
    info.style.color = errore ? "#a4262c" : "#605e5c";
}

function mostraStatoSelettoreImport() {
    const valore = localStorage.getItem(statoSelettoreImportKey);
    if (!valore) return;

    try {
        const stato = JSON.parse(valore);
        if (Date.now() - stato.timestamp > 5 * 60 * 1000) {
            localStorage.removeItem(statoSelettoreImportKey);
            return;
        }

        if (stato.fase === "selezionato") {
            mostraStatoImport(`Outlook ha riaperto il pannello durante l’importazione di "${stato.nomeFile}". Il file è stato selezionato, ma la lettura non è arrivata a termine.`, true);
        } else {
            mostraStatoImport("Il clic sul selettore è stato ricevuto, ma non è arrivato l’evento di scelta del file. Potrebbe essere stata annullata la finestra oppure Outlook ha ricaricato il pannello.", true);
        }
    } catch (errore) {
        localStorage.removeItem(statoSelettoreImportKey);
        mostraStatoImport(`Stato diagnostico non leggibile: ${errore.message}`, true);
    }
}

function aggiornaInterfacciaImport(clienti, destinazioni) {
    const info = document.getElementById("infoImport");
    const data = localStorage.getItem("dataUltimoImport") || "recente";
    info.textContent = `Anagrafiche attive: Clienti ${clienti.length}, ClientiZ ${destinazioni.length} (aggiornate al: ${data})`;
    info.style.color = "#107c10";
}
