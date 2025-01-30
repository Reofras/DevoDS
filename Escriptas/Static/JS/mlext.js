const xmlData = `<?xml version="1.0" encoding="UTF-8"?>
<nfeProc xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00">
    <NFe xmlns="http://www.portalfiscal.inf.br/nfe">
        <infNFe Id="NFe21250109396414000156550010000249501502942965" versao="4.00">
            <ide>
                <nNF>24950</nNF>
            </ide>
            <infAdic>
                <infCpl>@DEV ;@REF 457029324; REF A NOTA FISCAL DE ENTRADA No 996101 PECA USADA. | |</infCpl>
            </infAdic>
            <det>
                <prod>
                    <cProd>3JK6G</cProd>
                    <qCom>3</qCom>
                </prod>
            </det>
            <det>
                <prod>
                    <cProd>0K1OP</cProd>
                    <qCom>5</qCom>
                </prod>
            </det>
        </infNFe>
    </NFe>
</nfeProc>`;

function processarNfe(xmlData) {
    try {
        // Analisando o XML
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlData, "text/xml");

        // Verifica se o XML é válido
        const errorNode = xmlDoc.querySelector("parsererror");
        if (errorNode) {
            throw new Error("Erro ao analisar o XML: " + errorNode.textContent);
        }

        // Definindo os namespaces
        const namespaces = { nfe: "http://www.portalfiscal.inf.br/nfe" };

        // Coletando dados da NFe
        const nNF = xmlDoc.querySelector("nfeProc NFe infNFe ide nNF")?.textContent || "N/A";
        const infCpl = xmlDoc.querySelector("nfeProc NFe infNFe infAdic infCpl")?.textContent || "";
        const callNumber = infCpl.includes('@REF') ? infCpl.split('@REF ')[1].split(';')[0] : "N/A";

        let dadosNFe = {
            "Número da NFe": nNF,
            "Call Number": callNumber,
            "Volume Total": 0,
            "Part Numbers": [],
            "Condição das Peças": []
        };

        // Iterando sobre os itens (det) no XML
        const dets = xmlDoc.querySelectorAll("nfeProc NFe infNFe det");
        dets.forEach(det => {
            const partNumber = det.querySelector("prod cProd")?.textContent || "N/A";
            const volume = parseInt(det.querySelector("prod qCom")?.textContent || 0, 10);
            const condicaoPeca = infCpl.includes("USADA") ? "USADA" : "NOVA";

            dadosNFe["Part Numbers"].push(partNumber);
            dadosNFe["Volume Total"] += volume;
            dadosNFe["Condição das Peças"].push(condicaoPeca);
        });

        return dadosNFe;
    } catch (error) {
        console.error("Erro ao processar a NF-e:", error);
        return null;
    }
}

// Executando o processamento do XML
const resultado = processarNfe(xmlData);
if (resultado) {
    console.log(resultado);
} else {
    console.log("Falha ao processar o XML.");
}