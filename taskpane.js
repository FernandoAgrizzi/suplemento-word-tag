let tagsEncontradas = [];

// Conecta o suplemento com a API oficial do Office
Office.onReady((info) => {
    console.log("=== Office.onReady Disparado 1.1.1 ===");
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

        exibirStatus("Documento preenchido e salvo no OneDrive!", "sucesso");
    } catch (erro) {
        console.error("Erro ao preencher:", erro);
        exibirStatus("Erro ao preencher o documento: " + erro.message, "erro");
    }
}

// 3. Gera e baixa o PDF oficial usando a API Nativa da Microsoft, sem danificar a matriz
async function gerarEBaixarPDF() {
    exibirStatus("Preenchendo documento para exportação...", "info");

    const valores = {};
    document.querySelectorAll("#meuFormulario input").forEach(input => {
        valores[input.getAttribute("data-tag")] = input.value || "";
    });

    try {
        // Passo A: Preenche o documento no Word
        await Word.run(async (contexto) => {
            for (const [tag, valor] of Object.entries(valores)) {
                if (valor) {
                    const termoBusca = `<<${tag}>>`;
                    const busca = contexto.document.body.search(termoBusca, { matchCase: false });
                    busca.load("items");
                    await contexto.sync();

                    busca.items.forEach(item => {
                        item.insertText(valor, Word.InsertLocation.replace);
                    });
                }
            }
            await contexto.sync();
        });

        exibirStatus("Gerando PDF oficial via motor da Microsoft...", "info");

        // Passo B: Pega o arquivo PDF compilado nativamente
        const pdfBlob = await obterPdfBlobNativo();

        // Passo C: Dispara o download automático no computador
        const url = URL.createObjectURL(pdfBlob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "Contrato_Preenchido.pdf";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        exibirStatus("Restaurando modelo original no Word...", "info");

        // Passo D: Desfaz as alterações trazendo as tags <<TAG>> de volta
        await Word.run(async (contexto) => {
            for (const [tag, valor] of Object.entries(valores)) {
                if (valor) {
                    const busca = contexto.document.body.search(valor, { matchCase: false });
                    busca.load("items");
                    await contexto.sync();

                    busca.items.forEach(item => {
                        item.insertText(`<<${tag}>>`, Word.InsertLocation.replace);
                    });
                }
            }
            await contexto.sync();
        });

        exibirStatus("PDF baixado com sucesso! Modelo original preservado.", "sucesso");

    } catch (erro) {
        console.error("Erro ao gerar PDF:", erro);
        exibirStatus("Erro ao processar PDF nativo: " + erro.message, "erro");
    }
}

// Função utilitária para extrair os blocos binários do PDF nativo do Office
function obterPdfBlobNativo() {
    return new Promise((resolve, reject) => {
        Office.context.document.getFileAsync(
            Office.FileType.Pdf,
            { sliceSize: 65536 },
            (resultado) => {
                if (resultado.status === Office.AsyncResultStatus.Failed) {
                    reject(new Error(resultado.error.message));
                    return;
                }

                const arquivo = resultado.value;
                const contagemSlices = arquivo.sliceCount;
                const slicesCarregados = [];
                let slicesLidos = 0;

                function lerSlice(index) {
                    arquivo.getSliceAsync(index, (resultadoSlice) => {
                        if (resultadoSlice.status === Office.AsyncResultStatus.Failed) {
                            arquivo.closeAsync();
                            reject(new Error(resultadoSlice.error.message));
                            return;
                        }

                        slicesCarregados[index] = new Uint8Array(resultadoSlice.value.data);
                        slicesLidos++;

                        if (slicesLidos === contagemSlices) {
                            arquivo.closeAsync();
                            const blob = new Blob(slicesCarregados, { type: "application/pdf" });
                            resolve(blob);
                        } else {
                            lerSlice(slicesLidos);
                        }
                    });
                }

                lerSlice(0);
            }
        );
    });
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