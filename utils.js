function normalizeText(value) {
    return String(value ?? "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ")
        .trim();
}

function getValue(row, key) {
    const target = normalizeText(key).toUpperCase();
    for (const currentKey of Object.keys(row || {})) {
        if (normalizeText(currentKey).toUpperCase() === target) {
            return row[currentKey];
        }
    }
    return "";
}

function limparMarca(texto) {
    let valor = String(texto || "").trim();

    valor = valor.replace(/^PESQUISA\s+/i, "");
    valor = valor.replace(/^DE\s+AVARIA\s*-\s*/i, "");
    valor = valor.replace(/^AVARIA\s*-\s*/i, "");

    return valor.trim();
}

function uniqueNormalized(arr) {
    const used = new Set();
    return arr.filter((item) => {
        const key = normalizeText(item).toUpperCase();
        if (!key || used.has(key)) return false;
        used.add(key);
        return true;
    });
}

function montarEndereco(row) {
    const logradouro = String(getValue(row, "LOGRADOURO") || "").trim();
    const bairro = String(getValue(row, "BAIRRO") || "").trim();
    const municipio = String(getValue(row, "MUNICÍPIO") || getValue(row, "MUNICIPIO") || "").trim();

    const partes = [logradouro, bairro, municipio].filter(Boolean);

    if (partes.length === 0) return "";
    if (partes.length === 1) return partes[0];
    if (partes.length === 2) return `${partes[0]} - ${partes[1]}`;

    return `${partes[0]}, ${partes[1]} - ${partes[2]}`;
}

function getTextoResponsabilidade(destinatario) {
    const texto = "Declaramos para os devidos fins que o(a) referido(a) funcionário(a) não possui nenhum vínculo empregatício com o <b>[DESTINATARIO]</b>, sendo de nossa inteira responsabilidade todo e qualquer ônus trabalhistas, previdenciário e outros.";
    return texto.replaceAll("[DESTINATARIO]", destinatario || "DESTINATÁRIO");
}

function rgb01(r, g, b) {
    return [r / 255, g / 255, b / 255];
}

function baixarArquivo(bytes, fileName, mimeType) {
    const blob = new Blob([bytes], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
}
