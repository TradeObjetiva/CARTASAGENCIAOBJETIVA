function drawWrappedText(page, text, options) {
    const { font, boldFont, size, x, y, maxWidth, lineHeight, firstLineIndent = 0, color = rgb01(0, 0, 0) } = options;

    const parts = String(text || "").split(/(<b>|<\/b>|\s+)/g).filter(Boolean);

    let isBold = false;
    let lines = [];
    let currentLine = [];
    let currentLineWidth = 0;

    for (const part of parts) {
        if (part === "<b>") { isBold = true; continue; }
        if (part === "</b>") { isBold = false; continue; }

        let cFont = (isBold && boldFont) ? boldFont : font;
        let cWidth = cFont.widthOfTextAtSize(part, size);
        let actualMaxWidth = lines.length === 0 ? maxWidth - firstLineIndent : maxWidth;

        if (part.trim() === "") {
            if (currentLineWidth === 0) continue;
            currentLine.push({ text: part, font: cFont, width: cWidth });
            currentLineWidth += cWidth;
            continue;
        }

        if (currentLineWidth + cWidth > actualMaxWidth && currentLineWidth > 0) {
            lines.push(currentLine);
            currentLine = [{ text: part, font: cFont, width: cWidth }];
            currentLineWidth = cWidth;
        } else {
            currentLine.push({ text: part, font: cFont, width: cWidth });
            currentLineWidth += cWidth;
        }
    }
    if (currentLine.length > 0) lines.push(currentLine);

    let currentY = y;
    for (let i = 0; i < lines.length; i++) {
        let line = lines[i];
        let currentX = x + (i === 0 ? firstLineIndent : 0);
        let grouped = [];
        for (const seg of line) {
            if (grouped.length > 0 && grouped[grouped.length - 1].font === seg.font) {
                grouped[grouped.length - 1].text += seg.text;
            } else {
                grouped.push({ text: seg.text, font: seg.font });
            }
        }

        for (const g of grouped) {
            page.drawText(g.text, { x: currentX, y: currentY, size, font: g.font, color });
            currentX += g.font.widthOfTextAtSize(g.text, size);
        }
        currentY -= lineHeight;
    }

    return currentY;
}

function sZ(text) {
    return String(text || "")
        .replace(/[\u2018\u2019]/g, "'") // smart single quotes
        .replace(/[\u201C\u201D]/g, '"') // smart double quotes
        .replace(/[\u2013\u2014]/g, '-') // en/em dashes
        .replace(/[\u2026]/g, '...') // ellipsis
        .replace(/[^\x00-\xFF]/g, ''); // strip any other non-WinAnsi/Latin1 char that breaks pdf-lib
}

async function buildPdfBytes(targetGroups) {
    if (!state.templateBytes) {
        throw new Error("Envie o PDF modelo primeiro.");
    }
    if (!targetGroups || !targetGroups.length) {
        throw new Error("Carregue a planilha primeiro.");
    }

    const cidadeData = document.getElementById("cidadeData");
    const cargo = document.getElementById("cargo");
    const docTipo = document.getElementById("docTipo");
    const docNumero = document.getElementById("docNumero");
    const empresaRemetente = document.getElementById("empresaRemetente");
    const textoVinculoModo = document.getElementById("textoVinculoModo");
    const zipName = document.getElementById("zipName");

    const { PDFDocument, StandardFonts, rgb } = PDFLib;

    const outputPdf = await PDFDocument.create();
    const templatePdfDoc = await PDFDocument.load(state.templateBytes);

    const helvetica = await outputPdf.embedFont(StandardFonts.Helvetica);
    const helveticaBold = await outputPdf.embedFont(StandardFonts.HelveticaBold);

    const templatePage = templatePdfDoc.getPage(0);

    for (const grupo of targetGroups) {
        const [copiedPage] = await outputPdf.copyPages(templatePdfDoc, [0]);
        const page = outputPdf.addPage(copiedPage);

        const marcas = grupo.marcas.join(", ");

        let modoVinculo = textoVinculoModo ? textoVinculoModo.value : "local";
        let destinatario = sZ(grupo.local) || "-";
        if (modoVinculo === "razao" && grupo.razaoSocial) destinatario = sZ(grupo.razaoSocial);
        if (modoVinculo === "rede" && grupo.rede) destinatario = sZ(grupo.rede);

        const responsabilidade = getTextoResponsabilidade(destinatario);

        page.drawText(sZ(cidadeData.value) || "Rio de Janeiro, 24 de março de 2026.", {
            x: 95,
            y: 750,
            size: 11.5,
            font: helveticaBold,
            color: rgb(...rgb01(20, 20, 20)),
        });

        let empresaText = "O+ MARKETING E RECURSOS HUMANOS LTDA";
        if (empresaRemetente && empresaRemetente.options && empresaRemetente.options.length > 0) {
            empresaText = empresaRemetente.options[empresaRemetente.selectedIndex].text;
        }

        drawWrappedText(page, `<b>DE:</b> ${sZ(empresaText)}`, {
            font: helvetica,
            boldFont: helveticaBold,
            size: 11.5,
            x: 95,
            y: 690,
            maxWidth: 440,
            lineHeight: 15,
            color: rgb(...rgb01(20, 20, 20)),
        });

        drawWrappedText(page, `<b>PARA:</b> ${sZ(grupo.local) || "-"}`, {
            font: helvetica,
            boldFont: helveticaBold,
            size: 11.5,
            x: 95,
            y: 670,
            maxWidth: 440,
            lineHeight: 15,
            color: rgb(...rgb01(20, 20, 20)),
        });

        // Some models hide address; we keep it if parsed, aligned like PARA.
        let cy = 600; // start text much lower
        if (sZ(grupo.endereco)) {
            let addrY = drawWrappedText(page, `<b>Endereço:</b> ${sZ(grupo.endereco)}`, {
                font: helvetica,
                boldFont: helveticaBold,
                size: 11.2,
                x: 95,
                y: 650,
                maxWidth: 440,
                lineHeight: 15,
                color: rgb(...rgb01(20, 20, 20)),
            });
            // if we drew address, cy can be relatively placed, but the gap is still huge
            cy = addrY - 40;
        }

        let valDoc = sZ(docNumero?.value || "-");
        let tipoD = sZ(docTipo?.value || "CPF");
        let cg = sZ(cargo?.value || "Promotor de Merchandising");

        const paragrafo1 =
            `Viemos por meio desta, apresentar o(a) nosso(a) funcionário(a) Sr(a) ` +
            `<b>${sZ(grupo.agente) || "-"}</b>, ${tipoD}: <b>${valDoc}</b>, ` +
            `que atuará na função de ${cg} ` +
            `dos produtos da empresa <b>${sZ(marcas) || "-"}</b>, no <b>${sZ(grupo.local) || "-"}</b>.`;

        cy = drawWrappedText(page, paragrafo1, {
            font: helvetica,
            boldFont: helveticaBold,
            size: 11.4,
            x: 95,
            y: cy,
            maxWidth: 450,
            lineHeight: 18,
            firstLineIndent: 30,
            color: rgb(...rgb01(20, 20, 20)),
        });

        cy -= 25; // Large gap between paragraphs

        cy = drawWrappedText(page, responsabilidade, {
            font: helvetica,
            boldFont: helveticaBold,
            size: 11.4,
            x: 95,
            y: cy,
            maxWidth: 450,
            lineHeight: 18,
            firstLineIndent: 30,
            color: rgb(...rgb01(20, 20, 20)),
        });

        page.drawText("Atenciosamente,", {
            x: 95,
            y: cy - 50, // Huge gap for Atenciosamente
            size: 11.5,
            font: helveticaBold, // Expected image has Atenciosamente in bold
            color: rgb(...rgb01(20, 20, 20)),
        });
    }

    return await outputPdf.save();
}

async function gerarPDF() {
    try {
        const pdfBytes = await buildPdfBytes(state.grouped);
        const zipName = document.getElementById("zipName");
        let dt = new Date().toISOString().slice(0, 10);
        let name = zipName && zipName.value ? zipName.value : `cartas_lojas_${dt}.pdf`;
        if (!name.toLowerCase().endsWith('.pdf')) name += '.pdf';

        baixarArquivo(pdfBytes, name, "application/pdf");
    } catch (e) {
        alert(e.message);
    }
}

async function gerarPreviaPDF(indice) {
    try {
        const grupo = state.grouped[indice];
        if (!grupo) return;

        const pdfBytes = await buildPdfBytes([grupo]);
        const blob = new Blob([pdfBytes], { type: "application/pdf" });
        const url = URL.createObjectURL(blob);

        const iframe = document.getElementById("previewIframe");
        if (iframe) {
            iframe.src = url;
            iframe.style.display = "block";
        }
    } catch (e) {
        console.error(e);
        alert("Não foi possível carregar a prévia: " + e.message);
    }
}
