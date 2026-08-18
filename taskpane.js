let tagsEncontradas = [];

// Garante a conexão do Suplemento com a API nativa do Word
Office.onReady((info) => {
    console.log("[LOG COMMENT] Garante a conexão do Suplemento com a API nativa do Word");
    console.log("=== Office.onReady Disparado ===");
    console.log("Host:", info.host);

    if (info.host === Office.HostType.Word) {
        console.log("Ambiente Word confirmado. Registrando eventos dos botões...");
        document.getElementById("btnMapear").onclick = mapearTags;
        document.getElementById("btnPreencher").onclick = preencherDocumento;
    }
});

// 1. Mapeia as tags <<TAG>> lendo parágrafo por parágrafo (compatível com Word Online)
async function mapearTags() {
    console.log("[LOG COMMENT] 1. Mapeia as tags <<TAG>> lendo parágrafo por parágrafo (compatível com Word Online)");
    console.log("[PASSAGEM 1] Entrou na função mapearTags()");
    exibirStatus("Mapeando tags no documento...", "info");

    try {
        await Word.run(async (contexto) => {
            console.log("[PASSAGEM 2] Entrou no bloco Word.run()");

            // Carrega os parágrafos do documento
            console.log("[LOG COMMENT] Carrega os parágrafos do documento");
            const paragrafos = contexto.document.body.paragraphs;
            paragrafos.load("text");
            
            // Sincroniza a leitura com o Word Online
            console.log("[LOG COMMENT] Sincroniza a leitura com o Word Online");
            await contexto.sync();
            console.log("[PASSAGEM 3] contexto.sync() dos parágrafos executado com sucesso!");

            const conjunto = new Set();
            const regexTags = /<<([^>]+)>>/g;

            // Percorre parágrafo por parágrafo buscando as tags
            console.log("[LOG COMMENT] Percorre parágrafo por parágrafo buscando as tags");
            console.log("Total de parágrafos encontrados:", paragrafos.items.length);

            for (let i = 0; i < paragrafos.items.length; i++) {
                const textoParagrafo = paragrafos.items[i].text || "";
                let correspondencia;

                while ((correspondencia = regexTags.exec(textoParagrafo)) !== null) {
                    if (correspondencia[1]) {
                        const tagLimpa = correspondencia[1].trim();
                        if (tagLimpa) {
                            console.log(` -> Tag capturada no parágrafo ${i + 1}:`, tagLimpa);
                            conjunto.add(tagLimpa);
                        }
                    }
                }
            }

            tagsEncontradas = Array.from(conjunto);
            console.log("Array final de tags encontradas:", tagsEncontradas);

            if (tagsEncontradas.length === 0) {
                console.warn("Nenhuma tag <<TAG>> encontrada no documento.");
                exibirStatus("Nenhuma tag <<TAG>> encontrada no documento.", "erro");
                document.getElementById("areaFormulario").style.display = "none";
                return;
            }

            gerarFormulario(tagsEncontradas);
            exibirStatus(`Encontrada(s) ${tagsEncontradas.length} tag(s).`, "sucesso");
            console.log("[PASSAGEM 4] Mapeamento concluído com sucesso!");
        });
    } catch (erro) {
        console.error("=========================================");
        console.error("ERRO CAPTURADO EM mapearTags():", erro);
        console.error("Nome do Erro:", erro.name);
        console.error("Mensagem do Erro:", erro.message);
        if (erro.debugInfo) {
            console.error("DebugInfo completo da API:", JSON.stringify(erro.debugInfo));
        }
        console.error("=========================================");
        exibirStatus("Erro ao acessar o documento: " + erro.message, "erro");
    }
}

// 2. Preenche os valores informados pelo usuário substituindo as marcas
async function preencherDocumento() {
    console.log("[LOG COMMENT] 2. Preenche os valores informados pelo usuário substituindo as marcas");
    console.log("[PASSAGEM 1] Entrou na função preencherDocumento()");
    exibirStatus("Substituindo tags no documento...", "info");

    const valores = {};
    document.querySelectorAll("#meuFormulario input").forEach(input => {
        valores[input.getAttribute("data-tag")] = input.value || "";
    });

    try {
        await Word.run(async (contexto) => {
            console.log("[PASSAGEM 2] Entrou no Word.run() de substituição");

            for (const [tag, valor] of Object.entries(valores)) {
                const termoBusca = `<<${tag}>>`;
                console.log(`Buscando no documento pela tag: "${termoBusca}"...`);

                // Busca o texto exato da tag
                console.log("[LOG COMMENT] Busca o texto exato da tag");
                const busca = contexto.document.body.search(termoBusca, { matchCase: false });
                busca.load("items");
                
                await contexto.sync();
                console.log(`Encontradas ${busca.items.length} ocorrência(s) de: "${termoBusca}"`);

                // Substitui cada ocorrência encontrada
                console.log("[LOG COMMENT] Substitui cada ocorrência encontrada");
                busca.items.forEach(item => {
                    item.insertText(valor, Word.InsertLocation.replace);
                });
            }

            await contexto.sync();
            console.log("[PASSAGEM 3] Substituição gravada no Word com sucesso!");
        });

        exibirStatus("Documento preenchido e salvo automaticamente!", "sucesso");
    } catch (erro) {
        console.error("=========================================");
        console.error("ERRO CAPTURADO EM preencherDocumento():", erro);
        console.error("Mensagem do Erro:", erro.message);
        console.error("=========================================");
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