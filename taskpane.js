let tagsEncontradas = [];

Office.onReady((info) => {
    console.log("=== Office.onReady disparado ===");
    console.log("Host detectado:", info.host);
    console.log("Plataforma:", info.platform);

    if (info.host === Office.HostType.Word) {
        console.log("Ambiente do Word confirmado. Registrando eventos dos botões...");
        document.getElementById("btnMapear").onclick = mapearTags;
        document.getElementById("btnPreencher").onclick = preencherDocumento;
    } else {
        console.warn("Atenção: O suplemento não está rodando dentro do Microsoft Word.");
    }
});

// 1. Lê o texto do documento e mapeia as tags <<TAG>>
async function mapearTags() {
    console.log("----------------------------------------");
    console.log("[INÍCIO] Função mapearTags() iniciada.");
    exibirStatus("Mapeando tags no documento...", "info");

    try {
        console.log("[Etapa 1] Chamando Word.run()...");
        await Word.run(async (contexto) => {
            console.log("[Etapa 2] Contexto carregado. Carregando corpo do documento...");
            
            const corpo = contexto.document.body;
            corpo.load("text");

            console.log("[Etapa 3] Sincronizando com o Word (contexto.sync)...");
            await contexto.sync();
            console.log("[Etapa 4] Sincronização concluída com sucesso!");

            const textoCompleto = corpo.text || "";
            console.log("[Etapa 5] Tamanho do texto lido:", textoCompleto.length, "caracteres.");
            console.log("[Texto Extraído Preview]:", textoCompleto.substring(0, 150) + "...");

            // Expressão Regular para capturar valores entre << e >>
            const regexTags = /<<([^>]+)>>/g;
            const conjunto = new Set();
            let correspondencia;

            console.log("[Etapa 6] Executando Regex no texto...");
            while ((correspondencia = regexTags.exec(textoCompleto)) !== null) {
                if (correspondencia[1]) {
                    const tagLimpa = correspondencia[1].trim();
                    if (tagLimpa) {
                        console.log(" -> Tag encontrada pelo Regex:", tagLimpa);
                        conjunto.add(tagLimpa);
                    }
                }
            }

            tagsEncontradas = Array.from(conjunto);
            console.log("[Etapa 7] Lista final de tags mapeadas:", tagsEncontradas);

            if (tagsEncontradas.length === 0) {
                console.warn("[Resultado] Nenhuma tag encontrada no texto lido.");
                exibirStatus("Nenhuma tag <<TAG>> encontrada no documento.", "erro");
                document.getElementById("areaFormulario").style.display = "none";
                return;
            }

            console.log("[Etapa 8] Gerando campos de formulário na tela...");
            gerarFormulario(tagsEncontradas);
            exibirStatus(`Encontrada(s) ${tagsEncontradas.length} tag(s).`, "sucesso");
            console.log("[SUCESSO] Mapeamento concluído.");
        });
    } catch (erro) {
        console.error("========================================");
        console.error("[ERRO em mapearTags]:", erro);
        console.error("Nome do Erro:", erro.name);
        console.error("Mensagem de Erro:", erro.message);
        if (erro.debugInfo) {
            console.error("Debug Info da API Office:", JSON.stringify(erro.debugInfo));
        }
        console.error("========================================");
        
        exibirStatus("Erro ao acessar o documento: " + erro.message, "erro");
    }
}

// 2. Preenche os valores informados pelo usuário substituindo as marcas
async function preencherDocumento() {
    console.log("----------------------------------------");
    console.log("[INÍCIO] Função preencherDocumento() iniciada.");
    exibirStatus("Substituindo tags no documento...", "info");

    const valores = {};
    document.querySelectorAll("#meuFormulario input").forEach(input => {
        const tag = input.getAttribute("data-tag");
        const val = input.value || "";
        valores[tag] = val;
        console.log(` -> Campo [${tag}]: "${val}"`);
    });

    try {
        console.log("[Etapa 1] Chamando Word.run() para substituição...");
        await Word.run(async (contexto) => {
            for (const [tag, valor] of Object.entries(valores)) {
                const termoBusca = `<<${tag}>>`;
                console.log(`[Substituindo] Procurando ocorrências de: "${termoBusca}"...`);

                const busca = contexto.document.body.search(termoBusca, { matchCase: false });
                busca.load("items");
                
                await contexto.sync();
                console.log(` -> Encontradas ${busca.items.length} ocorrência(s) da tag ${tag}.`);

                busca.items.forEach((item, index) => {
                    console.log(`   --> Substituindo ocorrência #${index + 1} por: "${valor}"`);
                    item.insertText(valor, Word.InsertLocation.replace);
                });
            }

            console.log("[Etapa 2] Sincronizando alterações de texto no Word...");
            await contexto.sync();
            console.log("[Etapa 3] Alterações gravadas com sucesso no documento!");
        });

        exibirStatus("Documento preenchido e salvo automaticamente!", "sucesso");
        console.log("[SUCESSO] Processo de preenchimento concluído.");
    } catch (erro) {
        console.error("========================================");
        console.error("[ERRO em preencherDocumento]:", erro);
        console.error("Mensagem de Erro:", erro.message);
        console.error("========================================");
        
        exibirStatus("Erro ao preencher o documento: " + erro.message, "erro");
    }
}

function gerarFormulario(tags) {
    console.log("Desenhando elementos HTML do formulário...");
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