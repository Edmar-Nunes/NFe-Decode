// Variáveis globais
let nfeData = null;
let xmlDoc = null;
let produtosExibidos = []; // Array para armazenar os produtos exibidos na tabela

// Inicialização quando a página carrega
window.onload = function() {
    // Configurar event listeners
    document.getElementById("xmlFile").addEventListener("change", handleFileSelect);
    document.getElementById("excelBtn").addEventListener("click", exportToExcel);
    document.getElementById("pdfBtn").addEventListener("click", exportToPDF);
    
    // Inicialmente desabilitar botões de exportação
    document.getElementById("excelBtn").disabled = true;
    document.getElementById("pdfBtn").disabled = true;
};

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const parser = new DOMParser();
            xmlDoc = parser.parseFromString(e.target.result, "text/xml");
            
            // Verifica se é um XML válido de NF-e
            const nfe = xmlDoc.querySelector("NFe, nfeProc");
            if (!nfe) {
                throw new Error("XML não é uma NF-e válida");
            }

            parseXML(xmlDoc);
            
            // Mostra os dados e habilita botões de exportação
            document.getElementById("invoiceInfo").style.display = "block";
            document.getElementById("noData").style.display = "none";
            document.getElementById("excelBtn").disabled = false;
            document.getElementById("pdfBtn").disabled = false;
            
        } catch (error) {
            alert("Erro ao processar XML: " + error.message);
            console.error(error);
        }
    };
    
    reader.readAsText(file);
}

// ============================================
// FUNÇÕES PRINCIPAIS
// ============================================

function parseXML(xml) {
    // Função auxiliar para buscar tags
    const getTag = (tag, parent = xml) => {
        const el = parent.getElementsByTagName(tag)[0];
        return el ? el.textContent : "";
    };

    // Busca o elemento infNFe
    let infNFe = xml.querySelector("infNFe");
    if (!infNFe) {
        const nfeProc = xml.querySelector("nfeProc");
        if (nfeProc) infNFe = nfeProc.querySelector("infNFe");
    }
    
    if (!infNFe) {
        throw new Error("Estrutura da NF-e não encontrada");
    }

    // Emitente
    const emit = infNFe.querySelector("emit");
    document.getElementById("emitente").innerHTML = `
        <div class="info-item"><span class="info-label">Nome/Razão Social:</span><span class="info-value">${getTag("xNome", emit)}</span></div>
        <div class="info-item"><span class="info-label">CNPJ:</span><span class="info-value">${formatCNPJ(getTag("CNPJ", emit) || getTag("CPF", emit))}</span></div>
        <div class="info-item"><span class="info-label">IE:</span><span class="info-value">${getTag("IE", emit)}</span></div>
        <div class="info-item"><span class="info-label">Endereço:</span><span class="info-value">${getTag("xLgr", emit)}, ${getTag("nro", emit)}</span></div>
        <div class="info-item"><span class="info-label">Bairro:</span><span class="info-value">${getTag("xBairro", emit)}</span></div>
        <div class="info-item"><span class="info-label">Cidade/UF:</span><span class="info-value">${getTag("xMun", emit)}/${getTag("UF", emit)}</span></div>
        <div class="info-item"><span class="info-label">CEP:</span><span class="info-value">${formatCEP(getTag("CEP", emit))}</span></div>
    `;

    // Destinatário
    const dest = infNFe.querySelector("dest");
    document.getElementById("destinatario").innerHTML = `
        <div class="info-item"><span class="info-label">Nome/Razão Social:</span><span class="info-value">${getTag("xNome", dest)}</span></div>
        <div class="info-item"><span class="info-label">CNPJ/CPF:</span><span class="info-value">${formatCNPJ(getTag("CNPJ", dest) || getTag("CPF", dest))}</span></div>
        <div class="info-item"><span class="info-label">IE:</span><span class="info-value">${getTag("IE", dest) || "N/A"}</span></div>
        <div class="info-item"><span class="info-label">Cidade/UF:</span><span class="info-value">${getTag("xMun", dest)}/${getTag("UF", dest)}</span></div>
        <div class="info-item"><span class="info-label">Email:</span><span class="info-value">${getTag("email", dest) || "N/A"}</span></div>
    `;

    // Informações da NF-e
    const ide = infNFe.querySelector("ide");
    document.getElementById("nfeInfo").innerHTML = `
        <div class="info-item"><span class="info-label">Número:</span><span class="info-value">${getTag("nNF", ide)}</span></div>
        <div class="info-item"><span class="info-label">Série:</span><span class="info-value">${getTag("serie", ide)}</span></div>
        <div class="info-item"><span class="info-label">Chave de Acesso:</span><span class="info-value">${formatChaveAcesso(getTag("chNFe", xml) || infNFe.getAttribute("Id")?.replace('NFe', ''))}</span></div>
        <div class="info-item"><span class="info-label">Data Emissão:</span><span class="info-value">${formatDate(getTag("dhEmi", ide) || getTag("dEmi", ide))}</span></div>
        <div class="info-item"><span class="info-label">Modelo:</span><span class="info-value">${getTag("mod", ide)}</span></div>
        <div class="info-item"><span class="info-label">Tipo de Operação:</span><span class="info-value">${getTipoOperacao(getTag("tpNF", ide))}</span></div>
    `;

    // Totais
    const total = infNFe.querySelector("total");
    const icmsTot = total ? total.querySelector("ICMSTot") : null;
    document.getElementById("totais").innerHTML = `
        <div class="info-item"><span class="info-label">Valor Produtos:</span><span class="info-value currency">R$ ${formatCurrency(getTag("vProd", icmsTot))}</span></div>
        <div class="info-item"><span class="info-label">Valor NF:</span><span class="info-value currency">R$ ${formatCurrency(getTag("vNF", icmsTot))}</span></div>
        <div class="info-item"><span class="info-label">Valor ICMS:</span><span class="info-value currency">R$ ${formatCurrency(getTag("vICMS", icmsTot))}</span></div>
        <div class="info-item"><span class="info-label">Valor Descontos:</span><span class="info-value currency">R$ ${formatCurrency(getTag("vDesc", icmsTot))}</span></div>
        <div class="info-item"><span class="info-label">Valor Frete:</span><span class="info-value currency">R$ ${formatCurrency(getTag("vFrete", icmsTot))}</span></div>
        <div class="info-item"><span class="info-label">Valor Seguro:</span><span class="info-value currency">R$ ${formatCurrency(getTag("vSeg", icmsTot))}</span></div>
    `;

    // Informações adicionais
    const infAdic = infNFe.querySelector("infAdic");
    const infCpl = infAdic ? infAdic.querySelector("infCpl") : null;
    document.getElementById("infAdicionais").innerHTML = infCpl ? 
        `<div class="info-text">${infCpl.textContent}</div>` : 
        '<div class="no-data" style="margin:0;">Nenhuma informação adicional</div>';

    // Produtos
    const produtos = infNFe.querySelectorAll("det");
    let produtosHTML = '';
    
    // Limpa o array de produtos exibidos
    produtosExibidos = [];
    
    if (produtos.length > 0) {
        produtosHTML = `
            <table id="productsTableData">
                <thead>
                    <tr>
                        <th>Código</th>
                        <th>Descrição</th>
                        <th>Lote/Validade/Fabricação</th>
                        <th>EAN</th>
                        <th>NCM</th>
                        <th>Qtd</th>
                        <th>Unid.</th>
                        <th>Valor Unit.</th>
                        <th>Valor Total</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        produtos.forEach((det, index) => {
            const prod = det.querySelector("prod");
            const xProd = getTag("xProd", prod);
            
            // BUSCA CORRETA DE LOTE E VALIDADE
            const loteInfo = extrairLoteValidadeFabricacao(det);
            
            // Formatar as informações de lote para exibição
            const loteDisplay = formatarLoteDisplay(loteInfo);
            
            produtosHTML += `
                <tr>
                    <td>${getTag("cProd", prod)}</td>
                    <td>${xProd}</td>
                    <td class="lote-info">${loteDisplay}</td>
                    <td>${getTag("cEAN", prod)}</td>
                    <td>${getTag("NCM", prod)}</td>
                    <td>${formatNumber(getTag("qCom", prod))}</td>
                    <td>${getTag("uCom", prod)}</td>
                    <td class="currency">R$ ${formatCurrency(getTag("vUnCom", prod))}</td>
                    <td class="currency">R$ ${formatCurrency(getTag("vProd", prod))}</td>
                </tr>
            `;
            
            // Adiciona ao array de produtos exibidos
            produtosExibidos.push({
                codigo: getTag("cProd", prod),
                descricao: xProd,
                loteDisplay: loteDisplay,
                loteInfo: loteInfo,
                ean: getTag("cEAN", prod),
                ncm: getTag("NCM", prod),
                quantidade: getTag("qCom", prod),
                unidade: getTag("uCom", prod),
                valorUnitario: getTag("vUnCom", prod),
                valorTotal: getTag("vProd", prod)
            });
        });
        
        produtosHTML += `
                </tbody>
            </table>
        `;
        
        // Armazena os dados para exportação
        nfeData = {
            emitente: {
                nome: getTag("xNome", emit),
                cnpj: getTag("CNPJ", emit) || getTag("CPF", emit),
                endereco: `${getTag("xLgr", emit)}, ${getTag("nro", emit)}`,
                cidade: `${getTag("xMun", emit)}/${getTag("UF", emit)}`
            },
            destinatario: {
                nome: getTag("xNome", dest),
                cnpj: getTag("CNPJ", dest) || getTag("CPF", dest),
                cidade: `${getTag("xMun", dest)}/${getTag("UF", dest)}`
            },
            nfeInfo: {
                numero: getTag("nNF", ide),
                serie: getTag("serie", ide),
                dataEmissao: getTag("dhEmi", ide) || getTag("dEmi", ide),
                valorTotal: getTag("vNF", icmsTot),
                chaveAcesso: getTag("chNFe", xml) || infNFe.getAttribute("Id")?.replace('NFe', '')
            },
            produtos: produtosExibidos,
            informacoesAdicionais: infCpl ? infCpl.textContent : "Nenhuma informação adicional"
        };
        
    } else {
        produtosHTML = '<div class="no-data">Nenhum produto encontrado</div>';
        nfeData = null;
    }
    
    document.getElementById("productsTable").innerHTML = produtosHTML;
}

// ============================================
// FUNÇÕES PARA BUSCA DE LOTE/VALIDADE/FABRICAÇÃO
// ============================================

/**
 * FUNÇÃO CORRIGIDA: Extrai lote, validade e fabricação do elemento det
 */
function extrairLoteValidadeFabricacao(detElement) {
    // Procura a tag infAdProd DENTRO do elemento det
    const infAdProdElement = detElement.querySelector("infAdProd");
    
    if (!infAdProdElement || !infAdProdElement.textContent) {
        return {
            lote: "N/A",
            validade: "N/A",
            fabricacao: "N/A",
            textoOriginal: ""
        };
    }
    
    const texto = infAdProdElement.textContent.trim();
    
    // Extrai todas as ocorrências de VALIDADE e LOTE
    const lotes = extrairTodasOcorrencias(texto, "LOTE:");
    const validades = extrairTodasOcorrencias(texto, "VALIDADE:");
    const fabricacoes = extrairTodasOcorrencias(texto, "FABRICAÇÃO:");
    
    return {
        lote: lotes.length > 0 ? lotes.join(', ') : "N/A",
        validade: validades.length > 0 ? validades.join(', ') : "N/A",
        fabricacao: fabricacoes.length > 0 ? fabricacoes.join(', ') : "N/A",
        textoOriginal: texto
    };
}

/**
 * Extrai todas as ocorrências de um prefixo no texto
 */
function extrairTodasOcorrencias(texto, prefixo) {
    const resultados = [];
    const regex = new RegExp(`${prefixo}\\s*([^\\s\\-]+)`, 'gi');
    
    let match;
    while ((match = regex.exec(texto)) !== null) {
        // Para evitar loops infinitos
        if (match.index === regex.lastIndex) {
            regex.lastIndex++;
        }
        
        if (match[1]) {
            resultados.push(match[1].trim());
        }
    }
    
    return resultados;
}

/**
 * Formata a exibição das informações de lote para a tabela
 */
function formatarLoteDisplay(loteInfo) {
    if (loteInfo.textoOriginal) {
        // Formata o texto original com quebras de linha
        return loteInfo.textoOriginal
            .replace(/\s+-\s+/g, '\n')  // Substitui " - " por quebra de linha
            .replace(/VALIDADE:/g, '\nVALIDADE:')
            .replace(/LOTE:/g, '\nLOTE:')
            .replace(/FABRICAÇÃO:/g, '\nFABRICAÇÃO:')
            .trim()
            .replace(/^\n+/, ''); // Remove linha vazia no início
    }
    
    let display = '';
    if (loteInfo.lote !== "N/A") display += `LOTE: ${loteInfo.lote}\n`;
    if (loteInfo.validade !== "N/A") display += `VALIDADE: ${loteInfo.validade}\n`;
    if (loteInfo.fabricacao !== "N/A") display += `FABRICAÇÃO: ${loteInfo.fabricacao}`;
    
    return display.trim() || 'N/A';
}

/**
 * Formata as informações de lote para exportação (sem quebras de linha)
 */
function formatarLoteParaExportacao(loteInfo) {
    if (loteInfo.textoOriginal) {
        // Para exportação, mantém o texto original mas substitui quebras
        return loteInfo.textoOriginal
            .replace(/\n/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
    }
    
    let exportacao = '';
    if (loteInfo.lote !== "N/A") exportacao += `LOTE: ${loteInfo.lote} | `;
    if (loteInfo.validade !== "N/A") exportacao += `VALIDADE: ${loteInfo.validade} | `;
    if (loteInfo.fabricacao !== "N/A") exportacao += `FABRICAÇÃO: ${loteInfo.fabricacao} | `;
    
    exportacao = exportacao.replace(/\s\|\s$/, '');
    
    return exportacao || 'N/A';
}

/**
 * Formata lote/validade para PDF (mais compacto)
 */
function formatarLoteParaPDF(loteInfo) {
    if (loteInfo.textoOriginal) {
        // Formato compacto: "V:02/2028 L:MAS202501-02/28"
        let resultado = '';
        
        if (loteInfo.validade !== "N/A") {
            // Pega apenas a primeira validade
            const primeiraValidade = loteInfo.validade.split(',')[0];
            resultado += `V:${primeiraValidade.trim()} `;
        }
        
        if (loteInfo.lote !== "N/A") {
            // Pega apenas o primeiro lote
            const primeiroLote = loteInfo.lote.split(',')[0];
            resultado += `L:${primeiroLote.trim()}`;
        }
        
        return resultado.trim();
    }
    
    // Fallback
    let resultado = '';
    if (loteInfo.validade !== "N/A") {
        const primeiraValidade = loteInfo.validade.split(',')[0];
        resultado += `V:${primeiraValidade.trim()} `;
    }
    if (loteInfo.lote !== "N/A") {
        const primeiroLote = loteInfo.lote.split(',')[0];
        resultado += `L:${primeiroLote.trim()}`;
    }
    
    return resultado.trim() || 'N/A';
}

// ============================================
// FUNÇÕES DE EXPORTAÇÃO (Excel e PDF)
// ============================================

function exportToExcel() {
    if (!nfeData || !produtosExibidos.length) {
        alert('Nenhum dado disponível para exportação. Carregue um arquivo XML primeiro.');
        return;
    }
    
    try {
        // Cria uma nova workbook
        const wb = XLSX.utils.book_new();
        
        // Planilha de informações da NF-e
        const infoSheetData = [
            ["RELATÓRIO DE NF-e", "", "", "", ""],
            ["Data de exportação:", new Date().toLocaleDateString('pt-BR'), "", "", ""],
            ["", "", "", "", ""],
            ["EMITENTE", "", "", "DESTINATÁRIO", ""],
            ["Nome/Razão Social", nfeData.emitente.nome, "", "Nome/Razão Social", nfeData.destinatario.nome],
            ["CNPJ", formatCNPJ(nfeData.emitente.cnpj), "", "CNPJ/CPF", formatCNPJ(nfeData.destinatario.cnpj)],
            ["Endereço", nfeData.emitente.endereco, "", "Cidade/UF", nfeData.destinatario.cidade],
            ["Cidade/UF", nfeData.emitente.cidade, "", "", ""],
            ["", "", "", "", ""],
            ["INFORMAÇÕES DA NF-e", "", "", "", ""],
            ["Número", nfeData.nfeInfo.numero, "", "Série", nfeData.nfeInfo.serie],
            ["Data Emissão", formatDate(nfeData.nfeInfo.dataEmissao), "", "Valor Total", `R$ ${formatCurrency(nfeData.nfeInfo.valorTotal)}`],
            ["Chave de Acesso", nfeData.nfeInfo.chaveAcesso, "", "", ""],
            ["", "", "", "", ""],
            ["INFORMAÇÕES ADICIONAIS", "", "", "", ""]
        ];
        
        // Adiciona informações adicionais
        if (nfeData.informacoesAdicionais && nfeData.informacoesAdicionais !== "Nenhuma informação adicional") {
            const infLines = nfeData.informacoesAdicionais.split('\n');
            infLines.forEach(line => {
                infoSheetData.push([line, "", "", "", ""]);
            });
        } else {
            infoSheetData.push(["Nenhuma informação adicional", "", "", "", ""]);
        }
        
        const infoSheet = XLSX.utils.aoa_to_sheet(infoSheetData);
        
        // Estilização (largura de colunas)
        const infoColWidths = [
            {wch: 25}, {wch: 40}, {wch: 5}, {wch: 25}, {wch: 40}
        ];
        infoSheet['!cols'] = infoColWidths;
        
        XLSX.utils.book_append_sheet(wb, infoSheet, "Informações NF-e");
        
        // Planilha de produtos
        const produtosSheetData = [
            ["PRODUTOS/SERVIÇOS", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", ""],
            ["Código", "Descrição", "Lote/Validade/Fabricação", "EAN", "NCM", "Quantidade", "Unidade", "Valor Unitário", "Valor Total"]
        ];
        
        produtosExibidos.forEach(prod => {
            // Formata as informações de lote para exportação
            const loteExport = formatarLoteParaExportacao(prod.loteInfo);
            
            produtosSheetData.push([
                prod.codigo,
                prod.descricao,
                loteExport,
                prod.ean,
                prod.ncm,
                formatNumber(prod.quantidade),
                prod.unidade,
                `R$ ${formatCurrency(prod.valorUnitario)}`,
                `R$ ${formatCurrency(prod.valorTotal)}`
            ]);
        });
        
        // Adiciona totais
        const totalProdutos = produtosExibidos.reduce((sum, prod) => sum + parseFloat(prod.valorTotal || 0), 0);
        produtosSheetData.push(
            ["", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "TOTAL:", `R$ ${formatCurrency(totalProdutos.toString())}`]
        );
        
        const produtosSheet = XLSX.utils.aoa_to_sheet(produtosSheetData);
        
        // Estilização produtos
        const prodColWidths = [
            {wch: 15}, {wch: 40}, {wch: 50}, {wch: 15}, {wch: 10}, {wch: 10}, {wch: 8}, {wch: 15}, {wch: 15}
        ];
        produtosSheet['!cols'] = prodColWidths;
        
        XLSX.utils.book_append_sheet(wb, produtosSheet, "Produtos");
        
        // Planilha detalhada de lotes (OPCIONAL)
        const lotesSheetData = [
            ["DETALHES DE LOTE E VALIDADE", "", "", "", ""],
            ["", "", "", "", ""],
            ["Item", "Descrição", "Lote", "Validade", "Texto Original"]
        ];
        
        produtosExibidos.forEach((prod, index) => {
            lotesSheetData.push([
                index + 1,
                prod.descricao,
                prod.loteInfo.lote !== "N/A" ? prod.loteInfo.lote : "N/A",
                prod.loteInfo.validade !== "N/A" ? prod.loteInfo.validade : "N/A",
                prod.loteInfo.textoOriginal || "N/A"
            ]);
        });
        
        const lotesSheet = XLSX.utils.aoa_to_sheet(lotesSheetData);
        const lotesColWidths = [
            {wch: 8}, {wch: 40}, {wch: 30}, {wch: 15}, {wch: 50}
        ];
        lotesSheet['!cols'] = lotesColWidths;
        
        XLSX.utils.book_append_sheet(wb, lotesSheet, "Detalhes Lotes");
        
        // Gera o arquivo Excel
        const fileName = `NF-e_${nfeData.nfeInfo.numero || 'export'}_${new Date().toISOString().slice(0,10)}.xlsx`;
        XLSX.writeFile(wb, fileName);
        
    } catch (error) {
        alert("Erro ao exportar para Excel: " + error.message);
        console.error(error);
    }
}

function exportToPDF() {
    if (!nfeData || !produtosExibidos.length) {
        alert('Nenhum dado disponível para exportação. Carregue um arquivo XML primeiro.');
        return;
    }
    
    try {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('l', 'mm', 'a4'); // 'l' para landscape (paisagem)
        
        // Configurações
        const pageWidth = doc.internal.pageSize.getWidth();
        const pageHeight = doc.internal.pageSize.getHeight();
        const margin = 10;
        let yPos = margin;
        
        // Cabeçalho
        doc.setFontSize(14);
        doc.setTextColor(41, 128, 185);
        doc.text("RELATÓRIO DE NOTA FISCAL ELETRÔNICA", pageWidth / 2, yPos, { align: 'center' });
        yPos += 7;
        
        doc.setFontSize(9);
        doc.setTextColor(100, 100, 100);
        doc.text(`Data de exportação: ${new Date().toLocaleDateString('pt-BR')} ${new Date().toLocaleTimeString('pt-BR')}`, pageWidth / 2, yPos, { align: 'center' });
        yPos += 12;
        
        // Informações básicas (mais compactas)
        doc.setFontSize(10);
        doc.setTextColor(44, 62, 80);
        doc.setFont(undefined, 'bold');
        doc.text("INFORMAÇÕES DA NF-e", margin, yPos);
        yPos += 6;
        
        doc.setFontSize(8);
        doc.setFont(undefined, 'normal');
        doc.text(`Número: ${nfeData.nfeInfo.numero || 'N/A'}`, margin, yPos);
        doc.text(`Série: ${nfeData.nfeInfo.serie || 'N/A'}`, margin + 60, yPos);
        doc.text(`Data: ${formatDate(nfeData.nfeInfo.dataEmissao)}`, margin + 120, yPos);
        doc.text(`Valor: R$ ${formatCurrency(nfeData.nfeInfo.valorTotal)}`, margin + 180, yPos);
        yPos += 5;
        doc.text(`Chave: ${nfeData.nfeInfo.chaveAcesso ? nfeData.nfeInfo.chaveAcesso.substring(0, 50) + '...' : 'N/A'}`, margin, yPos);
        yPos += 10;
        
        // Emitente e Destinatário lado a lado
        const colWidth = (pageWidth - 3 * margin) / 2;
        
        // Emitente
        doc.setFontSize(9);
        doc.setFont(undefined, 'bold');
        doc.text("EMITENTE", margin, yPos);
        yPos += 5;
        
        doc.setFontSize(8);
        doc.setFont(undefined, 'normal');
        const emitenteText = [
            `Nome: ${nfeData.emitente.nome.substring(0, 40) + (nfeData.emitente.nome.length > 40 ? '...' : '')}`,
            `CNPJ: ${formatCNPJ(nfeData.emitente.cnpj)}`,
            `Cidade: ${nfeData.emitente.cidade}`
        ];
        
        emitenteText.forEach(line => {
            if (yPos < pageHeight - 20) {
                doc.text(line, margin, yPos);
                yPos += 4;
            }
        });
        
        // Destinatário (coluna direita)
        yPos = margin + 40; // Reset Y para mesma linha
        doc.setFontSize(9);
        doc.setFont(undefined, 'bold');
        doc.text("DESTINATÁRIO", margin + colWidth + margin, yPos);
        yPos += 5;
        
        doc.setFontSize(8);
        doc.setFont(undefined, 'normal');
        const destinatarioText = [
            `Nome: ${nfeData.destinatario.nome.substring(0, 40) + (nfeData.destinatario.nome.length > 40 ? '...' : '')}`,
            `CNPJ: ${formatCNPJ(nfeData.destinatario.cnpj)}`,
            `Cidade: ${nfeData.destinatario.cidade}`
        ];
        
        destinatarioText.forEach(line => {
            if (yPos < pageHeight - 20) {
                doc.text(line, margin + colWidth + margin, yPos);
                yPos += 4;
            }
        });
        
        // Posição para tabela (abaixo das informações)
        yPos = margin + 60;
        
        // Tabela de produtos com larguras ajustadas
        doc.setFontSize(10);
        doc.setFont(undefined, 'bold');
        doc.text("PRODUTOS/SERVIÇOS", margin, yPos);
        yPos += 7;
        
        // Configuração da tabela com colunas otimizadas
        const headers = [
            ["Item", "Código", "Descrição", "Lote/Validade", "Qtd", "Valor R$"]
        ];
        
        const data = produtosExibidos.map((prod, index) => {
            // Descrição reduzida
            const descricaoReduzida = prod.descricao.length > 25 ? 
                prod.descricao.substring(0, 25) + '...' : prod.descricao;
            
            // Lote/Validade compacto para PDF
            const loteCompacto = formatarLoteParaPDF(prod.loteInfo);
            
            return [
                (index + 1).toString(),
                prod.codigo,
                descricaoReduzida,
                loteCompacto,
                formatNumber(prod.quantidade),
                formatCurrency(prod.valorTotal)
            ];
        });
        
        // Configuração das colunas (larguras em mm)
        const columnStyles = {
            0: { cellWidth: 12, halign: 'center' }, // Item
            1: { cellWidth: 25 }, // Código
            2: { cellWidth: 50 }, // Descrição
            3: { cellWidth: 45 }, // Lote/Validade
            4: { cellWidth: 18, halign: 'right' }, // Qtd
            5: { cellWidth: 25, halign: 'right' } // Valor
        };
        
        doc.autoTable({
            startY: yPos,
            head: headers,
            body: data,
            margin: { left: margin, right: margin },
            styles: { 
                fontSize: 7, // Fonte menor
                cellPadding: 1.5, // Menos padding
                overflow: 'linebreak', // Quebra de linha
                cellWidth: 'wrap' // Ajusta largura
            },
            headStyles: { 
                fillColor: [41, 128, 185], 
                textColor: 255,
                fontSize: 8,
                fontStyle: 'bold'
            },
            alternateRowStyles: { fillColor: [245, 245, 245] },
            columnStyles: columnStyles,
            tableWidth: 'auto',
            didParseCell: function(data) {
                // Garante que texto muito longo seja truncado
                if (data.cell.raw && data.cell.raw.length > 60) {
                    data.cell.text = data.cell.raw.substring(0, 60) + '...';
                }
            },
            didDrawPage: function(data) {
                // Adiciona número da página
                doc.setFontSize(8);
                doc.text(`Página ${data.pageNumber}`, pageWidth / 2, pageHeight - 5, { align: 'center' });
            }
        });
        
        // Última posição Y após a tabela
        let finalY = doc.lastAutoTable.finalY || yPos + 100;
        
        // Informações adicionais em nova página se necessário
        if (finalY > pageHeight - 40) {
            doc.addPage();
            finalY = margin;
        }
        
        // Informações adicionais (se couber)
        if (nfeData.informacoesAdicionais && nfeData.informacoesAdicionais !== "Nenhuma informação adicional") {
            doc.setFontSize(10);
            doc.setFont(undefined, 'bold');
            doc.text("INFORMAÇÕES ADICIONAIS", margin, finalY + 10);
            
            doc.setFontSize(8);
            doc.setFont(undefined, 'normal');
            const infText = nfeData.informacoesAdicionais;
            if (infText) {
                const lines = doc.splitTextToSize(infText, pageWidth - 2 * margin);
                doc.text(lines, margin, finalY + 18);
            }
        }
        
        // Gera o PDF
        const fileName = `NF-e_${nfeData.nfeInfo.numero || 'export'}_${new Date().toISOString().slice(0,10)}.pdf`;
        doc.save(fileName);
        
    } catch (error) {
        alert("Erro ao exportar para PDF: " + error.message);
        console.error(error);
    }
}

// ============================================
// FUNÇÕES AUXILIARES DE FORMATAÇÃO
// ============================================

function formatCNPJ(cnpj) {
    if (!cnpj) return "N/A";
    const clean = cnpj.replace(/\D/g, '');
    if (clean.length === 11) {
        return clean.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
    } else if (clean.length === 14) {
        return clean.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5');
    }
    return cnpj;
}

function formatCEP(cep) {
    if (!cep) return "N/A";
    const clean = cep.replace(/\D/g, '');
    if (clean.length === 8) {
        return clean.replace(/(\d{5})(\d{3})/, '$1-$2');
    }
    return cep;
}

function formatChaveAcesso(chave) {
    if (!chave) return "N/A";
    const clean = chave.replace(/\D/g, '');
    if (clean.length === 44) {
        return clean.match(/.{1,4}/g).join(' ');
    }
    return chave;
}

function formatDate(dateString) {
    if (!dateString || dateString === "N/A") return "N/A";
    try {
        // Remove horário se existir
        const datePart = dateString.split('T')[0];
        const [year, month, day] = datePart.split('-');
        if (year && month && day) {
            return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${year}`;
        }
        
        // Tenta formato brasileiro
        const [day2, month2, year2] = dateString.split('/');
        if (day2 && month2 && year2) {
            return `${day2.padStart(2, '0')}/${month2.padStart(2, '0')}/${year2}`;
        }
        
        return dateString;
    } catch {
        return dateString;
    }
}

function formatCurrency(value) {
    if (!value || value === "N/A") return "0,00";
    const number = parseFloat(value);
    if (isNaN(number)) return "0,00";
    return number.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatNumber(value) {
    if (!value) return "0";
    const number = parseFloat(value);
    if (isNaN(number)) return "0";
    return number.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 4 });
}

function getTipoOperacao(tipo) {
    switch(tipo) {
        case "0": return "Entrada";
        case "1": return "Saída";
        default: return "N/A";
    }
}
