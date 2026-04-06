function agruparLojas(rows) {
    const map = new Map();

    rows.forEach((row) => {
        const local = String(getValue(row, "LOCAL") || "").trim();
        if (!local) return;

        if (!map.has(local)) {
            map.set(local, {
                local,
                agente: String(getValue(row, "AGENTE") || "").trim(),
                endereco: montarEndereco(row),
                rede: String(getValue(row, "REDE") || "").trim(),
                razaoSocial: String(getValue(row, "RAZÃO SOCIAL") || getValue(row, "RAZAO SOCIAL") || "").trim(),
                marcas: [],
            });
        }

        const grupo = map.get(local);
        const marca = limparMarca(getValue(row, "FORM"));

        if (marca) grupo.marcas.push(marca);
        if (!grupo.agente) grupo.agente = String(getValue(row, "AGENTE") || "").trim();
        if (!grupo.endereco) grupo.endereco = montarEndereco(row);
        if (!grupo.rede) grupo.rede = String(getValue(row, "REDE") || "").trim();
        if (!grupo.razaoSocial) {
            grupo.razaoSocial = String(getValue(row, "RAZÃO SOCIAL") || getValue(row, "RAZAO SOCIAL") || "").trim();
        }
    });

    return [...map.values()].map((grupo) => ({
        ...grupo,
        marcas: uniqueNormalized(grupo.marcas),
    }));
}

async function carregarExcel(file) {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    state.rows = json;
    state.grouped = agruparLojas(json);
    if(window.atualizarListaPreview) window.atualizarListaPreview();
}
