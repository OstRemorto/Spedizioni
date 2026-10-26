/* taskpane.js - Versione DEFINITIVA con Fix Indici e Checkbox */
// 1. Cambia l'import (lasciamo quello di default come backup)
import { anagraficaClienti as databaseIniziale } from "./clienti.js";

let anagraficaClienti = [...databaseIniziale]; // Inizia con quelli del file JS
let anagraficaDestinazioni = [];

let tappeDelGiro = [];
let clienteCorrente = null;
let destinazioneCorrente = null;
let dragIdx = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Add-in pronto!");

        // --- GESTIONE IMPORT EXCEL ---
        const fileInput = document.getElementById("importExcel");
        if (fileInput) {
            fileInput.addEventListener("change", gestisciImportExcel, false);
        }

        const datiSalvati = localStorage.getItem("anagraficaLogistica");
        if (datiSalvati) {
            anagraficaClienti = JSON.parse(datiSalvati);
        }
        const destinazioniSalvate = localStorage.getItem("anagraficaDestinazioniLogistica");
        if (destinazioniSalvate) {
            anagraficaDestinazioni = JSON.parse(destinazioniSalvate);
        }
        if (datiSalvati || destinazioniSalvate) {
            aggiornaInterfacciaImport(true);
        }

        // --- 1. FIX CHECKBOX (Era sparito il collegamento) ---
        const checkDest = document.getElementById("checkDestinazione");
        const boxDest = document.getElementById("boxDestinazione");
        if (checkDest) {
            checkDest.onchange = function() {
                boxDest.style.display = this.checked ? "block" : "none";
            };
        }

        // --- 2. SWITCH RADIO (Manuale vs Ricerca) ---
        document.querySelectorAll('input[name="tipoDest"]').forEach(radio => {
            radio.onchange = function() {
                const usaRicerca = this.value !== "manuale";
                document.getElementById("uiRicercaDest").style.display = usaRicerca ? "block" : "none";
                document.getElementById("uiManualeDest").style.display = this.value === "manuale" ? "block" : "none";
                document.getElementById("labelRicercaDest").textContent = this.value === "clientiZ"
                    ? "Cerca destinazione ClientiZ"
                    : "Cerca Cliente";
                document.getElementById("inputDestinazione").value = "";
                document.getElementById("suggDestinazione").style.display = "none";
                destinazioneCorrente = null;
            };
        });

        setupRicerca("inputCliente", "suggCliente", false);
        setupRicerca("inputDestinazione", "suggDestinazione", true);

        document.getElementById("btnAggiungi").onclick = aggiungiTappa;
        document.getElementById("btnCompila").onclick = compilaAppuntamentoAperto;

        caricaDatiEsistenti();
    }
});

function setupRicerca(inputId, suggId, isDestinazione) {
    const input = document.getElementById(inputId);
    const box = document.getElementById(suggId);

    input.onkeyup = () => {
        const testo = input.value.toLowerCase();
        box.innerHTML = "";
        if (testo.length < 2) { box.style.display = "none"; return; }

        const listaRicerca = isDestinazione &&
            document.querySelector('input[name="tipoDest"]:checked').value === "clientiZ"
            ? anagraficaDestinazioni
            : anagraficaClienti;
        const risultati = listaRicerca.filter(c => 
            c.nome.toLowerCase().includes(testo) || c.comune.toLowerCase().includes(testo)
        );

        risultati.forEach(c => {
            const div = document.createElement("div");
            div.className = "sugg-item";
            div.innerHTML = `<b>${c.nome}</b> <span style="font-size:10px;">(${c.comune})</span>`;
            div.onclick = () => {
                input.value = c.nome;
                box.style.display = "none";
                if (isDestinazione) {
                    destinazioneCorrente = c;
                } else {
                    clienteCorrente = c;
                    if (!document.getElementById("checkDestinazione").checked) {
                        destinazioneCorrente = c;
                    }
                }
            };
            box.appendChild(div);
        });
        box.style.display = "block";
    };
}

function aggiungiTappa() {
    const tipoTappa = document.getElementById("tipoTappa").value;
    if (!tipoTappa) { alert("Seleziona Ritiro o Consegna!"); return; }
    if (!clienteCorrente) { alert("Seleziona prima il cliente!"); return; }

    const quantita = {};
    for (const [campo, etichetta] of [
        ["inputCommerciale", "Commerciale"],
        ["inputTravi", "Travi"],
        ["inputPiastre", "Piastre"]
    ]) {
        const valore = document.getElementById(campo).value;
        const numero = valore === "" ? 0 : Number(valore);
        if (!Number.isSafeInteger(numero) || numero < 0 || numero > 20) {
            alert(`Inserisci una quantità intera da 0 a 20 per ${etichetta}.`);
            return;
        }
        quantita[etichetta.toLowerCase()] = numero;
    }

    let destFinal = "";
    let indFinal = "";
    const isDiverso = document.getElementById("checkDestinazione").checked;
    const isManuale = document.querySelector('input[name="tipoDest"]:checked').value === "manuale";

    if (!isDiverso) {
        destFinal = clienteCorrente.nome;
        indFinal = `${clienteCorrente.via}, ${clienteCorrente.comune}`;
    } else if (isManuale) {
        destFinal = document.getElementById("nomeManuale").value;
        indFinal = document.getElementById("indManuale").value;
        if (!destFinal || !indFinal) { alert("Dati cantiere mancanti!"); return; }
    } else {
        if (!destinazioneCorrente) { alert("Scegli un destinatario!"); return; }
        destFinal = destinazioneCorrente.nome;
        indFinal = `${destinazioneCorrente.via}, ${destinazioneCorrente.comune}`;
    }

    tappeDelGiro.push({
        tipo: tipoTappa,
        intestatario: clienteCorrente.nome,
        destinazione: destFinal,
        indirizzo: indFinal,
        ...quantita,
        note: document.getElementById("inputNote").value || "-"
    });

    renderizzaLista();
    resetCampi();
}

function resetCampi() {
    document.getElementById("tipoTappa").value = "consegna";
    document.getElementById("inputCommerciale").value = "0";
    document.getElementById("inputTravi").value = "0";
    document.getElementById("inputPiastre").value = "0";
    document.getElementById("inputNote").value = "";
    document.getElementById("inputCliente").value = "";
    document.getElementById("inputDestinazione").value = "";
    document.getElementById("nomeManuale").value = "";
    document.getElementById("indManuale").value = "";
    document.getElementById("checkDestinazione").checked = false;
    document.getElementById("boxDestinazione").style.display = "none";
    clienteCorrente = null;
    destinazioneCorrente = null;
}

function renderizzaLista() {
    const consegne = document.getElementById("listaConsegne");
    const ritiri = document.getElementById("listaRitiri");
    document.getElementById("countTappe").innerText = tappeDelGiro.length;

    consegne.innerHTML = "";
    ritiri.innerHTML = "";
    const tappePerTipo = {
        consegna: tappeDelGiro.filter(t => t.tipo !== "ritiro"),
        ritiro: tappeDelGiro.filter(t => t.tipo === "ritiro")
    };

    if (tappePerTipo.consegna.length === 0) {
        consegne.innerHTML = '<em style="color:#666; font-size:10px;">Nessuna consegna aggiunta.</em>';
    }
    if (tappePerTipo.ritiro.length === 0) {
        ritiri.innerHTML = '<em style="color:#666; font-size:10px;">Nessun ritiro aggiunto.</em>';
    }

    Object.entries(tappePerTipo).forEach(([tipo, tappe]) => {
        tappe.forEach((t, posizione) => {
            const i = tappeDelGiro.indexOf(t);
            const div = document.createElement("div");
            div.className = "tappa-row";
            div.draggable = true;
            div.innerHTML = `
                <button class="btn-remove" onclick="event.stopPropagation(); window.rimuoviTappa(${i})">✕</button>
                <span style="color:#999; margin-right:5px;">☰</span>
                <b>${posizione+1}. ${t.destinazione}</b><br>
                <small>📍 ${t.indirizzo}</small><br>
                <small style="color:#0078d4;">📦 Comm. ${t.commerciale || 0} | Travi ${t.travi || 0} | Piastre ${t.piastre || 0} | 📝 ${t.note}</small>
            `;
            div.ondragstart = () => { div.classList.add('dragging'); dragIdx = i; };
            div.ondragend = () => div.classList.remove('dragging');
            div.ondragover = (e) => e.preventDefault();
            div.ondrop = () => {
                if (dragIdx === null || tappeDelGiro[dragIdx].tipo !== t.tipo) return;
                const item = tappeDelGiro.splice(dragIdx, 1)[0];
                const indiceDestinazione = tappeDelGiro.indexOf(t);
                tappeDelGiro.splice(indiceDestinazione, 0, item);
                dragIdx = null;
                renderizzaLista();
            };
            (tipo === "ritiro" ? ritiri : consegne).appendChild(div);
        });
    });
}

window.rimuoviTappa = (i) => { tappeDelGiro.splice(i, 1); renderizzaLista(); };

function caricaDatiEsistenti() {
    const item = Office.context.mailbox.item;
    item.subject.getAsync((res) => {
        if (res.value) {
            const m = res.value.match(/GIRO: (.*?) \(/);
            if (m) document.getElementById("vettore").value = m[1];
        }
    });

    item.body.getAsync(Office.CoercionType.Html, (res) => {
        if (res.value) {
            const doc = new DOMParser().parseFromString(res.value, "text/html");
            const tabelleTipizzate = doc.querySelectorAll("table[data-tipo]");
            const tabelle = tabelleTipizzate.length > 0
                ? Array.from(tabelleTipizzate)
                : Array.from(doc.querySelectorAll("table"));
            if (tabelle.some(tabella => tabella.querySelectorAll("tr").length > 1)) {
                tappeDelGiro = [];
                tabelle.forEach(tabella => {
                    const righe = tabella.querySelectorAll("tr");
                    const intestazione = tabella.previousElementSibling;
                    const tipo = tabella.dataset.tipo === "ritiro" ||
                        (intestazione && intestazione.textContent.trim().toLowerCase().includes("ritiri"))
                        ? "ritiro"
                        : "consegna";
                    for (let i = 1; i < righe.length; i++) {
                        const c = righe[i].querySelectorAll("td");
                        if (c.length >= 5) {
                            const destB = c[1].querySelector("b") ? c[1].querySelector("b").innerText : c[1].innerText;
                            const contoB = c[1].querySelector("i") ? c[1].querySelector("i").innerText.replace("(Fattura a: ", "").replace(")", "") : destB;
                            const tappa = {
                                tipo,
                                destinazione: destB,
                                intestatario: contoB,
                                indirizzo: c[2].innerText
                            };
                            if (c.length >= 7) {
                                tappa.commerciale = (c[3].innerText.match(/☐/g) || []).length;
                                tappa.travi = (c[4].innerText.match(/☐/g) || []).length;
                                tappa.piastre = (c[5].innerText.match(/☐/g) || []).length;
                                tappa.note = c[6].innerText;
                            } else {
                                tappa.commerciale = 0;
                                tappa.travi = 0;
                                tappa.piastre = 0;
                                tappa.note = c[4].innerText;
                            }
                            tappeDelGiro.push(tappa);
                        }
                    }
                });
                renderizzaLista();
            }
        }
    });
}

function compilaAppuntamentoAperto() {
    const item = Office.context.mailbox.item;
    const vet = document.getElementById("vettore").value || "Giro";
    if (tappeDelGiro.length === 0) { alert("Aggiungi almeno una tappa prima di compilare l'appuntamento."); return; }

    item.subject.setAsync(`GIRO: ${vet} (${tappeDelGiro.length} tappe)`);

    let html = `<div style="font-family:sans-serif;"><h2 style="color:#0078d4;">🚚 SCHEDA DI CARICO</h2><p><b>Vettore:</b> ${vet}</p>`;
    ["consegna", "ritiro"].forEach(tipo => {
        const tappe = tappeDelGiro.filter(t => (t.tipo || "consegna") === tipo);
        if (tappe.length === 0) return;

        const titolo = tipo === "consegna" ? "CONSEGNE" : "RITIRI";
        const colore = tipo === "consegna" ? "#0078d4" : "#107c10";
        html += `<h3 style="color:${colore};">${titolo}</h3>
            <table data-tipo="${tipo}" border="1" cellpadding="8" style="border-collapse:collapse; width:100%; font-size:11pt;">
            <tr style="background:#f3f2f1;"><th>#</th><th>Destinazione</th><th>Indirizzo</th><th>Commerciale</th><th>Travi</th><th>Piastre</th><th>Note</th></tr>`;

        tappe.forEach((t, i) => {
            const nota = t.intestatario !== t.destinazione ? `<br><i style="font-size:10px;">(Fattura a: ${t.intestatario})</i>` : "";
            const caselle = quantita => Array.from({ length: quantita || 0 }, () => "&#9744;").join(" ");
            html += `<tr><td></td><td><b>${t.destinazione}</b>${nota}</td><td>${t.indirizzo}</td>
                <td style="font-family:'Segoe UI Symbol',sans-serif;">${caselle(t.commerciale)}</td>
                <td style="font-family:'Segoe UI Symbol',sans-serif;">${caselle(t.travi)}</td>
                <td style="font-family:'Segoe UI Symbol',sans-serif;">${caselle(t.piastre)}</td>
                <td>${t.note}</td></tr>`;
        });
        html += "</table>";
    });

    html += `<p style="margin-top:20px; font-size:14pt;"><b>Si prega gentilmente di pesare il camion in entrata ed in uscita.</b></p></div>`;
    item.body.setAsync(html, { coercionType: Office.CoercionType.Html });
}

function gestisciImportExcel(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const foglioClienti = workbook.Sheets.CLIENTI;
            const foglioDestinazioni = workbook.Sheets.DESTINAZIONIZ;
            if (!foglioClienti || !foglioDestinazioni) {
                const mancanti = [
                    !foglioClienti && "CLIENTI",
                    !foglioDestinazioni && "DESTINAZIONIZ"
                ].filter(Boolean).join(" e ");
                alert(`Nel file Excel manca il foglio richiesto: ${mancanti}. Nessuna anagrafica è stata aggiornata.`);
                return;
            }

            const leggiAnagrafica = foglio => XLSX.utils.sheet_to_json(foglio)
                .map(riga => {
                    const colonne = Object.keys(riga).reduce((mappa, chiave) => {
                        const normalizzata = chiave.normalize("NFD")
                            .replace(/[\u0300-\u036f]/g, "")
                            .toLowerCase()
                            .replace(/[^a-z0-9]/g, "");
                        mappa[normalizzata] = riga[chiave];
                        return mappa;
                    }, {});
                    return {
                        nome: String(colonne.ragionesociale || colonne.ragsoc || colonne.cliente || "").trim(),
                        via: String(colonne.indirizzo || colonne.via || "").trim(),
                        comune: String(colonne.citta || colonne.comune || "").trim()
                    };
                })
                .filter(cliente => cliente.nome || cliente.via || cliente.comune);

            const clientiImportati = leggiAnagrafica(foglioClienti);
            const destinazioniImportate = leggiAnagrafica(foglioDestinazioni);
            if (clientiImportati.length === 0 || destinazioniImportate.length === 0) {
                alert("Uno dei fogli CLIENTI o DESTINAZIONIZ non contiene righe utilizzabili. Nessuna anagrafica è stata aggiornata.");
                return;
            }

            anagraficaClienti = clientiImportati;
            anagraficaDestinazioni = destinazioniImportate;
            localStorage.setItem("anagraficaLogistica", JSON.stringify(anagraficaClienti));
            localStorage.setItem("anagraficaDestinazioniLogistica", JSON.stringify(anagraficaDestinazioni));
            localStorage.setItem("dataUltimoImport", new Date().toLocaleDateString());
            aggiornaInterfacciaImport(true);
        } catch (errore) {
            alert(`Impossibile leggere il file Excel: ${errore.message}`);
        }
    };
    reader.onerror = () => alert("Errore durante la lettura del file Excel.");
    reader.readAsArrayBuffer(file);
}

function aggiornaInterfacciaImport(caricato) {
    const info = document.getElementById("infoImport");
    if (caricato) {
        const data = localStorage.getItem("dataUltimoImport") || "recente";
        info.innerHTML = `<b style="color: #107c10;">✅ Anagrafiche attive</b> Clienti: ${anagraficaClienti.length}, ClientiZ: ${anagraficaDestinazioni.length} (aggiornate al: ${data})`;
    }
}