let tagsEncontradas = [];

// Garante a conexão do Suplemento com a API nativa do Word
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("btnMapear").onclick = mapearTags;
        document.getElementById("btnPreencher").onclick = preencherDocumento;
    }
});

// 1. Lê o texto do documento e mapeia as tags <<TAG>> via Regex Nativo do JS
async function mapearTags() {
    exibirStatus("Mapeando tags no documento...", "info");

    try {
        await Word.run(async (contexto) => {
            // Obtém o texto completo do corpo do documento
            const corpo = contexto.document.body;
            corpo.load("text");
            await contexto.sync();

            const textoCompleto = corpo.text || "";

            // Expressão Regular para capturar qualquer valor entre << e >>
            const regexTags = /<<([^>]+)>>/g;
            const conjunto = new Set();
            let correspondencia;

            while ((correspondencia = regexTags.exec(textoCompleto)) !== null) {
                if (correspondencia[1]) {
                    const tagLimpa = correspondencia[1].trim();
                    if (tagLimpa) {
                        conjunto.add(tagLimpa);
                    }
                }
            }

            tagsEncontradas = Array.from(conjunto);

            if (tagsEncontradas.length === 0) {
                exibirStatus("Nenhuma tag <<TAG>> encontrada no documento.", "erro");
                document.getElementById("areaFormulario").style.display = "none";
                return;
            }

            gerarFormulario(tagsEncontradas);
            exibirStatus(`Encontrada(s) ${tagsEncontradas.length} tag(s).`, "sucesso");
        });
    } catch (erro) {
        exibirStatus("Erro ao acessar o documento: " + erro.message, "erro");
    }
}

// 2. Preenche os valores informados pelo usuário substituindo as marcas
async function preencherDocumento() {
    exibirStatus("Substituindo tags no documento...", "info");

    const valores = {};
    document.querySelectorAll("#meuFormulario input").forEach(input => {
        valores[input.getAttribute("data-tag")] = input.value || "";
    });

    try {
        await Word.run(async (contexto) => {
            for (const [tag, valor] of Object.entries(valores)) {
                // Busca exatamente o texto da tag específica <<TAG>> sem usar wildcards
                const termoBusca = `<<${tag}>>`;
                const busca = contexto.document.body.search(termoBusca, { matchCase: false });
                busca.load("items");
                await contexto.sync();

                busca.items.forEach(item => {
                    item.insertText(valor, Word.InsertLocation.replace);
                });
            }
            await contexto.sync();
        });

        exibirStatus("Documento preenchido e salvo automaticamente no OneDrive!", "sucesso");
    } catch (erro) {
        exibirStatus("Erro ao preencher o documento: " + erro.message, "erro");
    }
}

function gerarFormulario(tags) {
    const form = document.getElementById("meuFormulario");
    form.innerHTML = "";

    tags.forEach((tag, idx) => {
        const idCampo = `campo_${idx}`;
        const rotulo = tag.charAt(0).toUpperCase() + tag.slice(1);

        const grupo = document.createElement("div");
        grupo.className = "grupo-campo";
        grupo.innerHTML = `
            <label for="${idCampo}">${rotulo}:</label>
            <input type="text" id="${idCampo}" data-tag="${tag}" placeholder="Digite ${rotulo}..." />
        `;
        form.appendChild(grupo);
    });

    document.getElementById("areaFormulario").style.display = "block";
}

function exibirStatus(mensagem, tipo) {
    const status = document.getElementById("status");
    status.textContent = mensagem;
    status.className = `status ${tipo}`;
}