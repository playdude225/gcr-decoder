Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("load-btn").onclick = loadEmailBody;
        document.getElementById("decode-btn").onclick = runDecoder;
    }
});

function loadEmailBody() {
    // Načítanie textu z aktuálne vybraného mailu
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            document.getElementById("input").value = result.value;
            runDecoder(); // Po načítaní automaticky dekóduj
        }
    });
}

function runDecoder() {
    let text = document.getElementById("input").value;
    let lines = text.split("\n");
    let decodedHTML = "";
    let replyText = "GCR\n/REG\nLZIB\n";

    lines.forEach(line => {
        let trimmedLine = line.trim();
        if(!/^[NCDR]/i.test(trimmedLine) || trimmedLine.length < 10) return;

        let actionLetter = trimmedLine[0]; 
        let isOdlet = trimmedLine[1] === " "; 
        let typPohybu = isOdlet ? "Odlet" : "Prílet";
        let triedaPohybu = isOdlet ? "odlet" : "prilet";

        let parts = trimmedLine.split(/\s+/);
        let flight = isOdlet ? parts[1] : parts[0].substring(1);
        let date = parts.find(p => /\d{2}[A-Z]{3}/.test(p)) || "-";
        let aircraft = parts.find(p => /^\d{3}[A-Z0-9]{4}$/.test(p)) || "-";
        let routeMatch = trimmedLine.match(/[A-Z]{4}\d{4}|\d{4}[A-Z]{4}/);
        let route = routeMatch ? routeMatch[0] : "-";
        
        let remainder = "";
        if (routeMatch) {
            let routeIndex = trimmedLine.indexOf(route);
            remainder = trimmedLine.substring(routeIndex + route.length).trim();
        }

        decodedHTML += `
            <div class="data-row">
                <span class="label">Pohyb: <span class="status-tag ${triedaPohybu}">${typPohybu}</span></span>
                <span class="value">${actionLetter} ${flight} | ${date} | ${route}</span>
                <div style="font-size: 0.8em; color: #94a3b8; margin-top: 4px;">${remainder}</div>
            </div>`;

        let separator = isOdlet ? " " : "";
        let prefix = (actionLetter.toUpperCase() === "N" || actionLetter.toUpperCase() === "R") ? "K" : "X";
        replyText += `${prefix}${separator}${flight} ${date} ${aircraft} ${route} ${remainder}`.trim().replace(/\s+/g, " ") + "\n";
    });

    replyText += "GI AUTO RESPONSE";
    document.getElementById("decoded").innerHTML = decodedHTML;
    document.getElementById("reply").innerText = replyText;
}