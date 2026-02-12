/* taskpane.js - Versione DEFINITIVA con Fix Indici e Checkbox */
// 1. Cambia l'import (lasciamo quello di default come backup)
import { anagraficaClienti as databaseIniziale } from "./clienti.js";

let anagraficaClienti = [...databaseIniziale]; // Inizia con quelli del file JS

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
            aggiornaInterfacciaImport(true); // Diciamo all'utente che è tutto pronto
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
                document.getElementById("uiRicercaDest").style.display = this.value === "ricerca" ? "block" : "none";
                document.getElementById("uiManualeDest").style.display = this.value === "manuale" ? "block" : "none";
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

        const risultati = anagraficaClienti.filter(c => 
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
    if (!clienteCorrente) { alert("Seleziona prima il cliente!"); return; }

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
        intestatario: clienteCorrente.nome,
        destinazione: destFinal,
        indirizzo: indFinal,
        ordine: document.getElementById("inputOrdine").value || "-",
        note: document.getElementById("inputNote").value || "-"
    });

    renderizzaLista();
    resetCampi();
}

function resetCampi() {
    document.getElementById("inputOrdine").value = "";
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
    const container = document.getElementById("listaContainer");
    document.getElementById("countTappe").innerText = tappeDelGiro.length;
    container.innerHTML = tappeDelGiro.length === 0 ? '<em style="color:#666; font-size:12px;">Nessuna tappa.</em>' : "";

    tappeDelGiro.forEach((t, i) => {
        const div = document.createElement("div");
        div.className = "tappa-row";
        div.draggable = true;
        div.innerHTML = `
            <button class="btn-remove" onclick="event.stopPropagation(); window.rimuoviTappa(${i})">✕</button>
            <span style="color:#999; margin-right:5px;">☰</span>
            <b>${i+1}. ${t.destinazione}</b><br>
            <small>📍 ${t.indirizzo}</small><br>
            <small style="color:#0078d4;">📦 ${t.ordine} | 📝 ${t.note}</small>
        `;
        div.ondragstart = () => { div.classList.add('dragging'); dragIdx = i; };
        div.ondragend = () => div.classList.remove('dragging');
        div.ondragover = (e) => e.preventDefault();
        div.ondrop = () => {
            const item = tappeDelGiro.splice(dragIdx, 1)[0];
            tappeDelGiro.splice(i, 0, item);
            renderizzaLista();
        };
        container.appendChild(div);
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
            const righe = doc.querySelectorAll("table tr");
            if (righe.length > 1) {
                tappeDelGiro = [];
                for (let i = 1; i < righe.length; i++) {
                    const c = righe[i].querySelectorAll("td");
                    if (c.length >= 5) {
                        // --- FIX INDICI: 3 è Ordine, 4 è Note ---
                        const destB = c[1].querySelector("b") ? c[1].querySelector("b").innerText : c[1].innerText;
                        const contoB = c[1].querySelector("i") ? c[1].querySelector("i").innerText.replace("(Fattura a: ", "").replace(")", "") : destB;
                        tappeDelGiro.push({
                            destinazione: destB,
                            intestatario: contoB,
                            indirizzo: c[2].innerText,
                            ordine: c[3].innerText,
                            note: c[4].innerText
                        });
                    }
                }
                renderizzaLista();
            }
        }
    });
}

function compilaAppuntamentoAperto() {
    const item = Office.context.mailbox.item;
    const vet = document.getElementById("vettore").value || "Giro";
    if (tappeDelGiro.length === 0) return;

    item.subject.setAsync(`GIRO: ${vet} (${tappeDelGiro.length} tappe)`);

    let html = `<div style="font-family:sans-serif;"><h2 style="color:#0078d4;">🚚 SCHEDA DI CARICO</h2><p><b>Vettore:</b> ${vet}</p>
        <table border="1" cellpadding="8" style="border-collapse:collapse; width:100%;">
        <tr style="background:#f3f2f1;"><th>#</th><th>Destinazione</th><th>Indirizzo</th><th>Rif. Ordine</th><th>Note</th></tr>`;

    tappeDelGiro.forEach((t, i) => {
        const nota = t.intestatario !== t.destinazione ? `<br><i style="font-size:11px;">(Fattura a: ${t.intestatario})</i>` : "";
        html += `<tr><td>${i+1}</td><td><b>${t.destinazione}</b>${nota}</td><td>${t.indirizzo}</td><td>${t.ordine}</td><td>${t.note}</td></tr>`;
    });

    html += `</table></div>`;
    item.body.setAsync(html, { coercionType: Office.CoercionType.Html });
}

function gestisciImportExcel(e) {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const foglio = workbook.Sheets[workbook.SheetNames[0]];
        const datiRaw = XLSX.utils.sheet_to_json(foglio);

        // Mappatura (personalizzala con i nomi delle tue colonne Excel)
        anagraficaClienti = datiRaw.map(riga => ({
            nome: riga["Ragione Sociale"] || riga["Cliente"] || "",
            via: riga["Indirizzo"] || riga["Via"] || "",
            comune: riga["Città"] || riga["Comune"] || ""
        }));

        // --- IL TRUCCO ---
        localStorage.setItem("anagraficaLogistica", JSON.stringify(anagraficaClienti));
        localStorage.setItem("dataUltimoImport", new Date().toLocaleDateString());

        aggiornaInterfacciaImport(true);
    };
    reader.readAsArrayBuffer(file);
}

function aggiornaInterfacciaImport(caricato) {
    const info = document.getElementById("infoImport");
    if (caricato) {
        const data = localStorage.getItem("dataUltimoImport") || "recente";
        info.innerHTML = `<b style="color: #107c10;">✅ Database attivo</b> (Aggiornato al: ${data})`;
    }
}