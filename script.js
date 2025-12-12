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
            lotes: [],
            validades: [],
            fabricacoes: [],
            textoOriginal: ""
        };
    }
    
    const texto = infAdProdElement.textContent.trim();
    
    // Extrai TODAS as ocorrências (para múltiplos lotes/validades)
    const lotes = extrairTodosPadroes(texto, /LOTE:\s*([^\s-]+)/gi);
    const validades = extrairTodosPadroes(texto, /VALIDADE:\s*([^\s-]+)/gi);
    const fabricacoes = extrairTodosPadroes(texto, /FABRICAÇÃO:\s*([^\s-]+)/gi);
    
    return {
        lotes: lotes,
        validades: validades,
        fabricacoes: fabricacoes,
        textoOriginal: texto
    };
}

/**
 * Extrai todos os padrões do texto (incluindo múltiplos)
 */
function extrairTodosPadroes(texto, regex) {
    const resultados = [];
    let match;
    
    while ((match = regex.exec(texto)) !== null) {
        if (match[1]) {
            resultados.push(match[1].trim());
        }
    }
    
    return resultados;
}

/**
 * Formata a exibição das informações de lote para a tabela HTML
 */
function formatarLoteDisplay(loteInfo) {
    if (loteInfo.textoOriginal) {
        return loteInfo.textoOriginal
            .replace(/\s+-\s+/g, '\n')
            .replace(/VALIDADE:/g, '\nVALIDADE:')
            .replace(/LOTE:/g, '\nLOTE:')
            .replace(/FABRICAÇÃO:/g, '\nFABRICAÇÃO:')
            .trim()
            .replace(/^\n+/, '');
    }
    
    let display = '';
    
    // Exibe todas as validades
    if (loteInfo.validades.length > 0) {
        loteInfo.validades.forEach((val, i) => {
            display += `VALIDADE: ${val}\n`;
        });
    }
    
    // Exibe todos os lotes
    if (loteInfo.lotes.length > 0) {
        loteInfo.lotes.forEach((lote, i) => {
            display += `LOTE: ${lote}\n`;
        });
    }
    
    // Exibe todas as fabricações
    if (loteInfo.fabricacoes.length > 0) {
        loteInfo.fabricacoes.forEach((fab, i) => {
            display += `FABRICAÇÃO: ${fab}\n`;
        });
    }
    
    return display.trim() || 'N/A';
}

/**
 * Formata lote/validade para Excel (inclui todos os múltiplos)
 */
function formatarLoteParaExcel(loteInfo) {
    let resultado = [];
    
    // Para Excel, formata cada par validade-lote
    const maxItems = Math.max(loteInfo.validades.length, loteInfo.lotes.length);
    
    for (let i = 0; i < maxItems; i++) {
        let item = '';
        if (loteInfo.validades[i]) {
            item += `V:${loteInfo.validades[i]}`;
        }
        if (loteInfo.lotes[i]) {
            if (item) item += ' ';
            item += `L:${loteInfo.lotes[i]}`;
        }
        if (item) {
            resultado.push(item);
        }
    }
    
    if (resultado.length > 0) {
        return resultado.join(' | ');
    }
    
    // Fallback para texto original
    if (loteInfo.textoOriginal) {
        return loteInfo.textoOriginal.replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
    }
    
    return 'N/A';
}

/**
 * Formata lote/validade para PDF (compacto, mas mostra múltiplos)
 */
function formatarLoteParaPDF(loteInfo) {
    let resultado = [];
    
    // Para PDF, mostra até 2 pares validade-lote
    const maxItems = Math.min(Math.max(loteInfo.validades.length, loteInfo.lotes.length), 2);
    
    for (let i = 0; i < maxItems; i++) {
        let item = '';
        if (loteInfo.validades[i]) {
            item += `V${i+1}:${loteInfo.validades[i]}`;
        }
        if (loteInfo.lotes[i]) {
            if (item) item += ' ';
            item += `L${i+1}:${loteInfo.lotes[i]}`;
        }
        if (item) {
            resultado.push(item);
        }
    }
    
    if (resultado.length > 0) {
        return resultado.join('; ');
    }
    
    // Se houver mais itens, indica com "..."
    if (loteInfo.validades.length > 2 || loteInfo.lotes.length > 2) {
        const count = Math.max(loteInfo.validades.length, loteInfo.lotes.length);
        return `V1:${loteInfo.validades[0]} L1:${loteInfo.lotes[0]}... (+${count-1})`;
    }
    
    return 'N/A';
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
        
        // Planilha de produtos com TODAS as informações
        const produtosSheetData = [
            ["PRODUTOS/SERVIÇOS", "", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "", ""],
            ["Item", "Código", "Descrição", "Validade(s)", "Lote(s)", "Qtd", "Unidade", "Valor Unitário", "Valor Total", "Informações Completas"]
        ];
        
        produtosExibidos.forEach((prod, index) => {
            // Formata separadamente validades e lotes
            const validadesStr = prod.loteInfo.validades.length > 0 ? 
                prod.loteInfo.validades.join(', ') : 'N/A';
            const lotesStr = prod.loteInfo.lotes.length > 0 ? 
                prod.loteInfo.lotes.join(', ') : 'N/A';
            
            produtosSheetData.push([
                (index + 1).toString(),
                prod.codigo,
                prod.descricao,
                validadesStr,
                lotesStr,
                formatNumber(prod.quantidade),
                prod.unidade,
                `R$ ${formatCurrency(prod.valorUnitario)}`,
                `R$ ${formatCurrency(prod.valorTotal)}`,
                prod.loteInfo.textoOriginal || 'N/A'
            ]);
        });
        
        // Adiciona totais
        const totalProdutos = produtosExibidos.reduce((sum, prod) => sum + parseFloat(prod.valorTotal || 0), 0);
        produtosSheetData.push(
            ["", "", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "TOTAL:", `R$ ${formatCurrency(totalProdutos.toString())}`, ""]
        );
        
        const produtosSheet = XLSX.utils.aoa_to_sheet(produtosSheetData);
        
        // Estilização produtos
        const prodColWidths = [
            {wch: 8},  // Item
            {wch: 15}, // Código
            {wch: 35}, // Descrição
            {wch: 20}, // Validade(s)
            {wch: 25}, // Lote(s)
            {wch: 10}, // Qtd
            {wch: 10}, // Unidade
            {wch: 15}, // Valor Unitário
            {wch: 15}, // Valor Total
            {wch: 40}  // Informações Completas
        ];
        produtosSheet['!cols'] = prodColWidths;
        
        XLSX.utils.book_append_sheet(wb, produtosSheet, "Produtos");
        
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
        const doc = new jsPDF('l', 'mm', 'a4'); // Landscape
        
        // Configurações da página
        const pageWidth = doc.internal.pageSize.getWidth();
        const pageHeight = doc.internal.pageSize.getHeight();
        const margin = 15;
        let yPos = margin;
        
        // Cabeçalho
        doc.setFontSize(16);
        doc.setTextColor(41, 128, 185);
        doc.setFont(undefined, 'bold');
        doc.text("RELATÓRIO DE NOTA FISCAL ELETRÔNICA", pageWidth / 2, yPos, { align: 'center' });
        yPos += 8;
        
        doc.setFontSize(10);
        doc.setTextColor(100, 100, 100);
        doc.setFont(undefined, 'normal');
        doc.text(`Data de exportação: ${new Date().toLocaleDateString('pt-BR')} ${new Date().toLocaleTimeString('pt-BR')}`, pageWidth / 2, yPos, { align: 'center' });
        yPos += 15;
        
        // Linha de informações principais
        doc.setFontSize(11);
        doc.setTextColor(44, 62, 80);
        doc.setFont(undefined, 'bold');
        doc.text("NF-e", margin, yPos);
        
        doc.setFontSize(10);
        doc.setFont(undefined, 'normal');
        doc.text(`Nº ${nfeData.nfeInfo.numero}`, margin + 20, yPos);
        doc.text(`Série ${nfeData.nfeInfo.serie}`, margin + 60, yPos);
        doc.text(`Emissão ${formatDate(nfeData.nfeInfo.dataEmissao)}`, margin + 100, yPos);
        doc.text(`Valor Total: R$ ${formatCurrency(nfeData.nfeInfo.valorTotal)}`, margin + 160, yPos);
        yPos += 7;
        
        // Emitente
        doc.setFontSize(11);
        doc.setFont(undefined, 'bold');
        doc.text("EMITENTE", margin, yPos);
        
        doc.setFontSize(10);
        doc.setFont(undefined, 'normal');
        doc.text(`${nfeData.emitente.nome}`, margin + 40, yPos);
        doc.text(`CNPJ: ${formatCNPJ(nfeData.emitente.cnpj)}`, margin + 140, yPos);
        yPos += 6;
        doc.text(`${nfeData.emitente.cidade}`, margin + 40, yPos);
        yPos += 10;
        
        // Destinatário
        doc.setFontSize(11);
        doc.setFont(undefined, 'bold');
        doc.text("DESTINATÁRIO", margin, yPos);
        
        doc.setFontSize(10);
        doc.setFont(undefined, 'normal');
        doc.text(`${nfeData.destinatario.nome}`, margin + 40, yPos);
        doc.text(`CNPJ: ${formatCNPJ(nfeData.destinatario.cnpj)}`, margin + 140, yPos);
        yPos += 6;
        doc.text(`${nfeData.destinatario.cidade}`, margin + 40, yPos);
        yPos += 15;
        
        // Tabela de produtos - OTIMIZADA para landscape
        doc.setFontSize(12);
        doc.setFont(undefined, 'bold');
        doc.text("PRODUTOS/SERVIÇOS", margin, yPos);
        yPos += 8;
        
        // Cabeçalhos da tabela
        const headers = [
            ["Item", "Código", "Descrição", "Validade/Lote", "Qtd", "Valor R$"]
        ];
        
        // Dados da tabela com formatação otimizada
        const data = produtosExibidos.map((prod, index) => {
            // Descrição reduzida mas legível
            const descricao = prod.descricao.length > 35 ? 
                prod.descricao.substring(0, 35) + '...' : prod.descricao;
            
            // Lote/Validade formatado para PDF (mostra múltiplos se houver)
            let loteValidade = formatarLoteParaPDF(prod.loteInfo);
            
            return [
                (index + 1).toString(),
                prod.codigo,
                descricao,
                loteValidade,
                formatNumber(prod.quantidade),
                `R$ ${formatCurrency(prod.valorTotal)}`
            ];
        });
        
        // Larguras das colunas (total 277mm em landscape A4)
        const columnStyles = {
            0: { cellWidth: 15, halign: 'center' },  // Item
            1: { cellWidth: 25 },                    // Código
            2: { cellWidth: 60 },                    // Descrição (mais larga)
            3: { cellWidth: 55 },                    // Validade/Lote (mais larga)
            4: { cellWidth: 20, halign: 'right' },   // Qtd
            5: { cellWidth: 30, halign: 'right' }    // Valor
        };
        
        doc.autoTable({
            startY: yPos,
            head: headers,
            body: data,
            margin: { left: margin, right: margin },
            styles: { 
                fontSize: 8,
                cellPadding: 2,
                overflow: 'linebreak',
                cellWidth: 'wrap',
                lineColor: [200, 200, 200],
                lineWidth: 0.1
            },
            headStyles: { 
                fillColor: [41, 128, 185], 
                textColor: 255,
                fontSize: 9,
                fontStyle: 'bold',
                halign: 'center'
            },
            alternateRowStyles: { fillColor: [250, 250, 250] },
            columnStyles: columnStyles,
            tableWidth: pageWidth - 2 * margin,
            didParseCell: function(data) {
                // Ajusta células com texto muito longo
                if (data.column.index === 2 && data.cell.raw && data.cell.raw.length > 45) {
                    data.cell.text = data.cell.raw.substring(0, 45) + '...';
                }
                if (data.column.index === 3 && data.cell.raw && data.cell.raw.length > 40) {
                    data.cell.text = data.cell.raw.substring(0, 40) + '...';
                }
            },
            willDrawCell: function(data) {
                // Adiciona bordas mais claras
                data.doc.setDrawColor(200, 200, 200);
            },
            didDrawPage: function(data) {
                // Número da página no rodapé
                doc.setFontSize(9);
                doc.setTextColor(100, 100, 100);
                doc.text(`Página ${data.pageNumber}`, pageWidth / 2, pageHeight - 10, { align: 'center' });
                
                // Linha separadora
                doc.setDrawColor(200, 200, 200);
                doc.line(margin, pageHeight - 15, pageWidth - margin, pageHeight - 15);
            }
        });
        
        // Pega a posição final após a tabela
        let finalY = doc.lastAutoTable.finalY || yPos + 100;
        
        // Informações adicionais em nova página se necessário
        if (nfeData.informacoesAdicionais && nfeData.informacoesAdicionais !== "Nenhuma informação adicional") {
            if (finalY > pageHeight - 30) {
                doc.addPage();
                finalY = margin;
            }
            
            doc.setFontSize(11);
            doc.setFont(undefined, 'bold');
            doc.setTextColor(44, 62, 80);
            doc.text("INFORMAÇÕES ADICIONAIS", margin, finalY + 10);
            
            doc.setFontSize(9);
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
        const datePart = dateString.split('T')[0];
        const [year, month, day] = datePart.split('-');
        if (year && month && day) {
            return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${year}`;
        }
        
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
