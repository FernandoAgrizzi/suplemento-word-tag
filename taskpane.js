let tagsEncontradas = [];

// Conecta o suplemento com a API oficial do Office
Office.onReady((info) => {
    console.log("=== Office.onReady Disparado ===");
    if (info.host === Office.HostType.Word) {
        document.getElementById("btnMapear").onclick = mapearTags;
        document.getElementById("btnPreencher").onclick = preencherDocumento;
        document.getElementById("btnGerarPdf").onclick = gerarEBaixarPDF;
        
        // Botão Dev para recarregar o script sem remover o suplemento
        document.getElementById("btnReload").onclick = () => {
            console.log("Recarregando o painel do suplemento...");
            window.location.reload(true);
        };
    }
});

// 1. Mapeia as tags <<TAG>> no Word
async function mapearTags() {
    exibirStatus("Mapeando tags no documento...", "info");

    try {
        await Word.run(async (contexto) => {
            const corpo = contexto.document.body;
            corpo.load("text");
            await contexto.sync();

            const textoCompleto = corpo.text || "";
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
        console.error("Erro no mapearTags:", erro);
        exibirStatus("Erro ao acessar o documento: " + erro.message, "erro");
    }
}

// 2. Preenche os valores alterando o documento original do Word
async function preencherDocumento() {
    exibirStatus("Substituindo tags no documento...", "info");

    const valores = {};
    document.querySelectorAll("#meuFormulario input").forEach(input => {
        valores[input.getAttribute("data-tag")] = input.value || "";
    });

    try {
        await Word.run(async (contexto) => {
            for (const [tag, valor] of Object.entries(valores)) {
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
        console.error("Erro ao preencher:", erro);
        exibirStatus("Erro ao preencher o documento: " + erro.message, "erro");
    }
}

// 3. Processa e baixa um novo arquivo em PDF sem alterar a matriz original no Word
async function gerarEBaixarPDF() {
    exibirStatus("Gerando novo documento para download em PDF...", "info");

    const valores = {};
    document.querySelectorAll("#meuFormulario input").forEach(input => {
        valores[input.getAttribute("data-tag")] = input.value || "";
    });

    try {
        await Word.run(async (contexto) => {
            const paragrafos = contexto.document.body.paragraphs;
            paragrafos.load("text");
            await contexto.sync();

            let conteudoHTML = "<html><head><title>Documento Preenchido</title><style>body{font-family:Arial,sans-serif;padding:30px;line-height:1.6;color:#333;}</style></head><body>";

            for (let i = 0; i < paragrafos.items.length; i++) {
                let texto = paragrafos.items[i].text || "";
                
                // Aplica a substituição dos campos apenas na memória de exportação
                for (const [tag, valor] of Object.entries(valores)) {
                    const regex = new RegExp(`<<${tag}>>`, 'gi');
                    texto = texto.replace(regex, valor);
                }
                
                conteudoHTML += `<p>${texto}</p>`;
            }

            conteudoHTML += "</body></html>";

            // Abre a janela temporária de impressão/PDF sem afetar o Word Online
            const janelaDownload = window.open("", "_blank");
            janelaDownload.document.write(conteudoHTML);
            janelaDownload.document.close();
            
            setTimeout(() => {
                janelaDownload.print();
                exibirStatus("Caixa de diálogo aberta. Selecione 'Salvar como PDF' para baixar.", "sucesso");
            }, 500);
        });
    } catch (erro) {
        console.error("Erro ao processar PDF:", erro);
        exibirStatus("Erro ao processar PDF: " + erro.message, "erro");
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