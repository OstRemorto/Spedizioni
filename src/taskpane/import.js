const statoSelettoreImportKey = "spedizioniStatoSelettoreImport";

export function inizializzaImport(onImportate) {
    const fileInput = document.getElementById("importExcel");
    if (fileInput) {
        fileInput.addEventListener("click", registraAperturaSelettore, false);
        fileInput.addEventListener("change", gestisciImportExcel(onImportate), false);
    }

    const btnImportaIncollati = document.getElementById("btnImportaIncollati");
    if (btnImportaIncollati) {
        btnImportaIncollati.addEventListener("click", importaDatiIncollati(onImportate), false);
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
            e.target.value = "";
        } catch (errore) {
            mostraStatoImport(`Impossibile registrare la selezione del file: ${errore.message}`, true);
            return;
        }

        if (!/\.(xlsx|xls)$/i.test(file.name)) {
            mostraStatoImport("Seleziona un file Excel (.xlsx o .xls) con i fogli CLIENTI e DESTINAZIONIZ.", true);
            localStorage.removeItem(statoSelettoreImportKey);
            return;
        }
        mostraStatoImport(`Lettura di "${file.name}" in corso...`);
        const reader = new FileReader();

        reader.onload = function(evento) {
            try {
                if (typeof XLSX === "undefined") {
                    throw new Error("La libreria Excel non è disponibile. Ricarica il pannello e riprova.");
                }
                const workbook = XLSX.read(new Uint8Array(evento.target.result), { type: "array" });
                const trovaFoglio = nome => {
                    const nomeFoglio = workbook.SheetNames.find(
                        nomeDisponibile => nomeDisponibile.trim().toUpperCase() === nome
                    );
                    return nomeFoglio ? workbook.Sheets[nomeFoglio] : null;
                };
                const foglioClienti = trovaFoglio("CLIENTI");
                const foglioDestinazioni = trovaFoglio("DESTINAZIONIZ");
                if (!foglioClienti || !foglioDestinazioni) {
                    const mancanti = [
                        !foglioClienti && "CLIENTI",
                        !foglioDestinazioni && "DESTINAZIONIZ"
                    ].filter(Boolean).join(" e ");
                    mostraStatoImport(`Nel file manca il foglio ${mancanti}. Fogli trovati: ${workbook.SheetNames.join(", ")}. Nessuna anagrafica è stata aggiornata.`, true);
                    localStorage.removeItem(statoSelettoreImportKey);
                    return;
                }

                const clienti = normalizzaAnagrafica(XLSX.utils.sheet_to_json(foglioClienti));
                const destinazioni = normalizzaAnagrafica(XLSX.utils.sheet_to_json(foglioDestinazioni));
                salvaAnagraficheImportate(clienti, destinazioni, onImportate);
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

function importaDatiIncollati(onImportate) {
    return function() {
        try {
            const clienti = normalizzaAnagrafica(parseTsv(document.getElementById("incollaClienti").value));
            const destinazioni = normalizzaAnagrafica(parseTsv(document.getElementById("incollaDestinazioni").value));
            salvaAnagraficheImportate(clienti, destinazioni, onImportate);
        } catch (errore) {
            mostraStatoImport(`Importazione dei dati incollati non riuscita: ${errore.message}`, true);
        }
    };
}

function parseTsv(testo) {
    const righe = [];
    let riga = [];
    let cella = "";
    let dentroVirgolette = false;
    const contenuto = testo.replace(/^\uFEFF/, "");

    for (let i = 0; i < contenuto.length; i++) {
        const carattere = contenuto[i];
        if (carattere === '"' && dentroVirgolette && contenuto[i + 1] === '"') {
            cella += '"';
            i++;
        } else if (carattere === '"') {
            dentroVirgolette = !dentroVirgolette;
        } else if (carattere === "\t" && !dentroVirgolette) {
            riga.push(cella);
            cella = "";
        } else if ((carattere === "\n" || carattere === "\r") && !dentroVirgolette) {
            riga.push(cella);
            righe.push(riga);
            riga = [];
            cella = "";
            if (carattere === "\r" && contenuto[i + 1] === "\n") i++;
        } else {
            cella += carattere;
        }
    }

    if (cella || riga.length > 0) {
        riga.push(cella);
        righe.push(riga);
    }
    if (dentroVirgolette) {
        throw new Error("Il testo incollato contiene virgolette non bilanciate.");
    }
    if (righe.length < 2) {
        throw new Error("Incolla le righe di Excel includendo la riga delle intestazioni.");
    }

    const intestazioni = righe[0].map(normalizzaIntestazione);
    return righe.slice(1).map(valori => {
        const record = {};
        intestazioni.forEach((intestazione, indice) => {
            record[intestazione] = valori[indice] || "";
        });
        return record;
    });
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

function salvaAnagraficheImportate(clienti, destinazioni, onImportate) {
    if (clienti.length === 0 || destinazioni.length === 0) {
        throw new Error("Entrambi i fogli devono contenere righe con Rag. Soc., Città e Indirizzo. Nessuna anagrafica è stata aggiornata.");
    }

    localStorage.setItem("anagraficaLogistica", JSON.stringify(clienti));
    localStorage.setItem("anagraficaDestinazioniLogistica", JSON.stringify(destinazioni));
    localStorage.setItem("dataUltimoImport", new Date().toLocaleDateString());
    localStorage.removeItem(statoSelettoreImportKey);
    onImportate(clienti, destinazioni);
    aggiornaInterfacciaImport(clienti, destinazioni);
    document.getElementById("incollaClienti").value = "";
    document.getElementById("incollaDestinazioni").value = "";
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
