document.addEventListener("DOMContentLoaded", () => {
    const excelFile = document.getElementById("excelFile");
    const templatePdf = document.getElementById("templatePdf");
    const generateBtn = document.getElementById("generateBtn");
    const previewSelect = document.getElementById("previewSelect");
    const previewBtn = document.getElementById("previewBtn");

    window.atualizarListaPreview = function() {
        if (!previewSelect) return;
        previewSelect.innerHTML = "";

        state.grouped.forEach((grupo, index) => {
            const opt = document.createElement("option");
            opt.value = index;
            opt.textContent = grupo.local;
            previewSelect.appendChild(opt);
        });
    };

    excelFile?.addEventListener("change", async (e) => {
        const file = e.target.files?.[0];
        if (!file) return;

        try {
            await carregarExcel(file);
            const numLojas = state.grouped.length;
            const statusBox = document.getElementById("statusBox");
            if (statusBox) statusBox.textContent = `${numLojas} lojas identificadas.`;
        } catch (error) {
            console.error(error);
            alert("Erro ao ler a planilha.");
        }
    });

    templatePdf?.addEventListener("change", async (e) => {
        const file = e.target.files?.[0];
        if (!file) return;

        try {
            state.templateBytes = await file.arrayBuffer();
            alert("PDF modelo carregado com sucesso.");
        } catch (error) {
            console.error(error);
            alert("Erro ao carregar o PDF modelo.");
        }
    });

    generateBtn?.addEventListener("click", async () => {
        try {
            generateBtn.textContent = "Gerando...";
            generateBtn.classList.add("loading");
            await gerarPDF();
        } catch (error) {
            console.error(error);
            alert("Erro ao gerar o PDF. Verifique o console.");
        } finally {
            generateBtn.textContent = "Gerar PDF";
            generateBtn.classList.remove("loading");
        }
    });

    previewBtn?.addEventListener("click", () => {
        const idx = parseInt(previewSelect.value, 10);
        if (!isNaN(idx)) {
            const previewPage = document.getElementById("previewPage");
            if (previewPage) previewPage.style.display = "none";
            
            previewBtn.textContent = "Atualizando...";
            // Pequeno delay para a UI respirar e o layout do iframe redesenhar
            setTimeout(async () => {
                await gerarPreviaPDF(idx);
                previewBtn.textContent = "Atualizar prévia";
            }, 100);
        } else {
            alert("A planilha não está carregada ou está vazia.");
        }
    });

    previewSelect?.addEventListener("change", () => {
        // Option to auto-trigger when dropdown changes? Let's just require clicking the button to avoid constant re-rendering.
    });
});
