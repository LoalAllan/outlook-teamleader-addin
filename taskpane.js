/* ============================================
   TEAMLEADER OFFERTE ADD-IN - JAVASCRIPT LOGICA
   ============================================ */

// ============================================
// CONFIGURATIE - PAS DIT AAN!
// ============================================

/**
 * BELANGRIJK: Vervang deze URL met jouw eigen webhook URL
 * Dit is waar de email data naartoe gestuurd wordt
 */
const WEBHOOK_URL = "https://JOUW-WEBHOOK-URL.com/api/offerte";

// ============================================
// INITIALISATIE
// ============================================

/**
 * Office.onReady wordt aangeroepen zodra Office.js geladen is
 * Dit is het startpunt van onze add-in
 */
Office.onReady(function(info) {
    // Controleer of we in Outlook draaien
    if (info.host === Office.HostType.Outlook) {
        console.log("Add-in geladen in Outlook");
        
        // Laad de email informatie
        laadEmailInfo();
        
        // Stel de event listeners in
        setupEventListeners();
    } else {
        console.log("Add-in draait niet in Outlook");
        toonStatus("error", "Deze add-in werkt alleen in Outlook");
    }
});

// ============================================
// EMAIL INFORMATIE LADEN
// ============================================

/**
 * Haalt informatie op over de huidige geselecteerde email
 * en toont dit in de interface
 */
function laadEmailInfo() {
    try {
        // Haal het huidige mail item op
        const item = Office.context.mailbox.item;
        
        if (item) {
            // Toon het onderwerp
            const onderwerpElement = document.getElementById("email-subject");
            onderwerpElement.textContent = item.subject || "(Geen onderwerp)";
            
            // Toon de afzender
            const afzenderElement = document.getElementById("email-sender");
            if (item.from) {
                // item.from kan een object zijn met displayName en emailAddress
                const afzenderNaam = item.from.displayName || item.from.emailAddress || "Onbekend";
                afzenderElement.textContent = afzenderNaam;
            } else {
                afzenderElement.textContent = "Onbekend";
            }
            
            console.log("Email info geladen:", {
                onderwerp: item.subject,
                afzender: item.from
            });
        } else {
            console.log("Geen email item gevonden");
            document.getElementById("email-subject").textContent = "Geen email geselecteerd";
            document.getElementById("email-sender").textContent = "-";
        }
    } catch (error) {
        console.error("Fout bij laden email info:", error);
        document.getElementById("email-subject").textContent = "Fout bij laden";
        document.getElementById("email-sender").textContent = "-";
    }
}

// ============================================
// EVENT LISTENERS
// ============================================

/**
 * Stelt alle event listeners in
 */
function setupEventListeners() {
    // Voeg click event toe aan de upload knop
    const uploadKnop = document.getElementById("upload-btn");
    uploadKnop.addEventListener("click", handleUpload);
}

// ============================================
// UPLOAD FUNCTIE (HOOFDLOGICA)
// ============================================

/**
 * Wordt aangeroepen wanneer de gebruiker op "Opladen naar TeamLeader" klikt
 * Dit is de hoofdfunctie die de data naar de webhook stuurt
 */
async function handleUpload() {
    const uploadKnop = document.getElementById("upload-btn");
    
    // Disable de knop en toon loading status
    uploadKnop.disabled = true;
    uploadKnop.innerHTML = '<span class="spinner"></span> <span class="btn-text">Bezig met uploaden...</span>';
    verbergStatus();
    
    try {
        // Haal het huidige mail item op
        const item = Office.context.mailbox.item;
        
        if (!item) {
            throw new Error("Geen email geselecteerd");
        }
        
        // Haal het geselecteerde bedrijf op (ABAssociates of AE+)
        const geselecteerdBedrijf = document.querySelector('input[name="company"]:checked').value;
        
        // Verzamel alle data die we willen versturen
        const payload = {
            // Email identificatie
            messageId: item.itemId,                    // Outlook's interne ID
            conversationId: item.conversationId,       // Conversatie ID (voor email threads)
            
            // Email inhoud
            onderwerp: item.subject,
            
            // Afzender informatie
            afzender: {
                naam: item.from ? item.from.displayName : null,
                email: item.from ? item.from.emailAddress : null
            },
            
            // Ontvanger informatie (optioneel)
            ontvanger: {
                naam: item.to && item.to[0] ? item.to[0].displayName : null,
                email: item.to && item.to[0] ? item.to[0].emailAddress : null
            },
            
            // Jouw selectie
            bedrijf: geselecteerdBedrijf,
            
            // Metadata
            timestamp: new Date().toISOString(),
            bron: "Outlook Add-in"
        };
        
        console.log("Versturen naar webhook:", payload);
        
        // Verstuur naar de webhook
        const response = await fetch(WEBHOOK_URL, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                // Voeg hier eventueel extra headers toe zoals API keys
                // "Authorization": "Bearer JOUW_API_KEY"
            },
            body: JSON.stringify(payload)
        });
        
        // Controleer of het gelukt is
        if (response.ok) {
            const resultaat = await response.text();
            console.log("Webhook response:", resultaat);
            toonStatus("success", "✅ Succesvol geüpload naar TeamLeader!");
        } else {
            // Server gaf een foutcode terug
            throw new Error("Server fout: " + response.status + " " + response.statusText);
        }
        
    } catch (error) {
        // Er ging iets mis
        console.error("Upload fout:", error);
        toonStatus("error", "❌ Fout: " + error.message);
    } finally {
        // Reset de knop (altijd, of het nu lukte of niet)
        uploadKnop.disabled = false;
        uploadKnop.innerHTML = '<span class="btn-icon">📤</span> <span class="btn-text">Opladen naar TeamLeader</span>';
    }
}

// ============================================
// HULPFUNCTIES
// ============================================

/**
 * Toont een status bericht aan de gebruiker
 * @param {string} type - "success", "error", of "loading"
 * @param {string} bericht - Het bericht om te tonen
 */
function toonStatus(type, bericht) {
    const statusElement = document.getElementById("status-message");
    statusElement.textContent = bericht;
    statusElement.className = "status-message " + type;
}

/**
 * Verbergt het status bericht
 */
function verbergStatus() {
    const statusElement = document.getElementById("status-message");
    statusElement.className = "status-message hidden";
}

// ============================================
// EXTRA: INTERNET MESSAGE ID OPHALEN (OPTIONEEL)
// ============================================

/**
 * Haalt de Internet Message ID op via de REST API
 * Dit is een meer universele ID die ook werkt buiten Outlook
 * 
 * LET OP: Dit vereist extra permissies en configuratie
 * Gebruik dit alleen als je dit echt nodig hebt
 */
async function haalInternetMessageId() {
    return new Promise(function(resolve) {
        try {
            const item = Office.context.mailbox.item;
            
            // Probeer via de REST API
            Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function(result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const token = result.value;
                    const itemId = Office.context.mailbox.convertToRestId(
                        item.itemId,
                        Office.MailboxEnums.RestVersion.v2_0
                    );
                    
                    // Hier zou je een REST call kunnen doen om meer details op te halen
                    resolve(itemId);
                } else {
                    resolve(null);
                }
            });
        } catch (error) {
            console.error("Kon Internet Message ID niet ophalen:", error);
            resolve(null);
        }
    });
}

