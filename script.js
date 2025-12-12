// Variáveis globais
let nfeData = null;
let xmlDoc = null;
let produtosExibidos = []; // Array para armazenar os produtos exibidos na tabela

// Inicialização
document.addEventListener('DOMContentLoaded', function() {
    document.getElementById("xmlFile").addEventListener("change", function(event) {
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
    });
});

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
            
            // BUSCA COMPLETA DE LOTE, VALIDADE E FABRICAÇÃO
            const loteInfo = buscarLoteValidadeFabricacao(det, xProd, xml);
            
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
            
            // Adiciona ao array de produtos exibidos (EXATAMENTE COMO NA TABELA)
            produtosExibidos.push({
                codigo: getTag("cProd", prod),
                descricao: xProd,
                loteDisplay: loteDisplay, // Mantém o formato exibido
                loteInfo: loteInfo, // Mantém também as informações originais para exportação detalhada
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
            produtos: produtosExibidos, // Usa os produtos exibidos
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
 * FUNÇÃO PRINCIPAL: Busca lote, validade e fabricação seguindo ordem de prioridade
 */
function buscarLoteValidadeFabricacao(detElement, xProd, xmlDoc) {
    const resultado = {
        lote: "N/A",
        validade: "N/A",
        fabricacao: "N/A",
        loteSource: "",
        validadeSource: "",
        fabricacaoSource: ""
    };

    // ETAPA 1: BUSCAR NOS CAMPOS OFICIAIS (PRIORIDADE A)
    // Caminho: det → prod → rastro → (nLote, dVal, dFab)
    const prod = detElement.querySelector("prod");
    if (prod) {
        const rastro = prod.querySelector("rastro");
        if (rastro) {
            // Lote no campo oficial nLote
            const nLoteElement = rastro.querySelector("nLote");
            if (nLoteElement && nLoteElement.textContent.trim()) {
                resultado.lote = nLoteElement.textContent.trim();
                resultado.loteSource = "campo";
            }
            
            // Validade no campo oficial dVal
            const dValElement = rastro.querySelector("dVal");
            if (dValElement && dValElement.textContent.trim()) {
                resultado.validade = formatarDataExport(dValElement.textContent.trim());
                resultado.validadeSource = "campo";
            }
            
            // Fabricação no campo oficial dFab
            const dFabElement = rastro.querySelector("dFab");
            if (dFabElement && dFabElement.textContent.trim()) {
                resultado.fabricacao = formatarDataExport(dFabElement.textContent.trim());
                resultado.fabricacaoSource = "campo";
            }
        }
        
        // ETAPA 2: BUSCAR NA DESCRIÇÃO (PRIORIDADE B)
        // Se algum campo estiver faltando, procura na descrição
        if (xProd) {
            // Busca lote
            if (resultado.lote === "N/A") {
                const loteEncontrado = buscarNaDescricao(xProd, ['LOTE', 'LT', 'L']);
                if (loteEncontrado) {
                    resultado.lote = loteEncontrado;
                    resultado.loteSource = "descricao";
                }
            }
            
            // Busca validade
            if (resultado.validade === "N/A") {
                const validadeEncontrada = buscarNaDescricao(xProd, ['VALIDADE', 'VAL', 'VENC', 'V']);
                if (validadeEncontrada) {
                    resultado.validade = formatarDataExport(validadeEncontrada);
                    resultado.validadeSource = "descricao";
                }
            }
            
            // Busca fabricação
            if (resultado.fabricacao === "N/A") {
                const fabricacaoEncontrada = buscarNaDescricao(xProd, ['FABRICAÇÃO', 'FAB', 'FABRIC', 'DFAB']);
                if (fabricacaoEncontrada) {
                    resultado.fabricacao = formatarDataExport(fabricacaoEncontrada);
                    resultado.fabricacaoSource = "descricao";
                }
            }
        }
        
        // ETAPA 3: BUSCAR EM INFADPROD (PRIORIDADE C)
        const infAdProdElement = prod.querySelector("infAdProd");
        if (infAdProdElement) {
            const infAdProd = infAdProdElement.textContent;
            
            if (resultado.lote === "N/A") {
                const loteEncontrado = buscarNaDescricao(infAdProd, ['LOTE', 'LT', 'L']);
                if (loteEncontrado) {
                    resultado.lote = loteEncontrado;
                    resultado.loteSource = "infAdProd";
                }
            }
            
            if (resultado.validade === "N/A") {
                const validadeEncontrada = buscarNaDescricao(infAdProd, ['VALIDADE', 'VAL', 'VENC', 'V']);
                if (validadeEncontrada) {
                    resultado.validade = formatarDataExport(validadeEncontrada);
                    resultado.validadeSource = "infAdProd";
                }
            }
            
            if (resultado.fabricacao === "N/A") {
                const fabricacaoEncontrada = buscarNaDescricao(infAdProd, ['FABRICAÇÃO', 'FAB', 'FABRIC', 'DFAB']);
                if (fabricacaoEncontrada) {
                    resultado.fabricacao = formatarDataExport(fabricacaoEncontrada);
                    resultado.fabricacaoSource = "infAdProd";
                }
            }
        }
    }
    
    // ETAPA 4: BUSCAR EM INFCPL (PRIORIDADE D)
    if (resultado.lote === "N/A" || resultado.validade === "N/A" || resultado.fabricacao === "N/A") {
        const infCplElement = xmlDoc.querySelector("infCpl");
        if (infCplElement) {
            const infCpl = infCplElement.textContent;
            
            if (resultado.lote === "N/A") {
                const loteEncontrado = buscarNaDescricao(infCpl, ['LOTE', 'LT', 'L']);
                if (loteEncontrado) {
                    resultado.lote = loteEncontrado;
                    resultado.loteSource = "infCpl";
                }
            }
            
            if (resultado.validade === "N/A") {
                const validadeEncontrada = buscarNaDescricao(infCpl, ['VALIDADE', 'VAL', 'VENC', 'V']);
                if (validadeEncontrada) {
                    resultado.validade = formatarDataExport(validadeEncontrada);
                    resultado.validadeSource = "infCpl";
                }
            }
            
            if (resultado.fabricacao === "N/A") {
                const fabricacaoEncontrada = buscarNaDescricao(infCpl, ['FABRICAÇÃO', 'FAB', 'FABRIC', 'DFAB']);
                if (fabricacaoEncontrada) {
                    resultado.fabricacao = formatarDataExport(fabricacaoEncontrada);
                    resultado.fabricacaoSource = "infCpl";
                }
            }
        }
    }
    
    return resultado;
}

/**
 * FUNÇÃO AUXILIAR: Busca informação na descrição baseada em termos-chave
 */
function buscarNaDescricao(texto, palavrasChave) {
    if (!texto) return null;
    
    for (const palavra of palavrasChave) {
        // Regex para buscar a palavra-chave seguida de valor
        const regex = new RegExp(`${palavra}[\\s\\:\\-]*([^\\s\\,\\;\\n\\r]+)`, 'i');
        const match = texto.match(regex);
        
        if (match && match[1]) {
            // Limpar o valor encontrado
            let valor = match[1].trim();
            valor = valor.replace(/[\\,\\.\\;\\:]+$/, '');
            
            // Remover parenteses e outros caracteres especiais
            valor = valor.replace(/[\(\)\[\]\{\}]/g, '');
            
            return valor;
        }
    }
    
    return null;
}

/**
 * FUNÇÃO AUXILIAR: Formata data para exportação (YYYY-MM-DD)
 */
function formatarDataExport(dataStr) {
    if (!dataStr) return dataStr;
    
    try {
        // Remove horário se existir
        const dataLimpa = dataStr.split('T')[0].split(' ')[0];
        
        let partes;
        if (dataLimpa.includes('-')) {
            partes = dataLimpa.split('-');
        } else if (dataLimpa.includes('/')) {
            partes = dataLimpa.split('/');
        } else if (dataLimpa.length === 8) {
            // Formato YYYYMMDD
            return `${dataLimpa.substring(0,4)}-${dataLimpa.substring(4,6)}-${dataLimpa.substring(6,8)}`;
        } else if (dataLimpa.length === 6) {
            // Formato DDMMYY
            const dia = dataLimpa.substring(0,2);
            const mes = dataLimpa.substring(2,4);
            const ano = dataLimpa.substring(4,6);
            return `20${ano}-${mes}-${dia}`;
        } else {
            return dataLimpa;
        }
        
        // Determinar formato
        if (partes[0].length === 4) {
            // Formato YYYY-MM-DD
            return `${partes[0]}-${partes[1].padStart(2, '0')}-${partes[2].padStart(2, '0')}`;
        } else if (partes[2] && partes[2].length === 4) {
            // Assume formato DD-MM-YYYY
            return `${partes[2]}-${partes[1].padStart(2, '0')}-${partes[0].padStart(2, '0')}`;
        } else if (partes[2] && partes[2].length === 2) {
            // Assume formato DD-MM-YY
            return `20${partes[2]}-${partes[1].padStart(2, '0')}-${partes[0].padStart(2, '0')}`;
        }
        
        return dataLimpa;
    } catch (error) {
        return dataStr;
    }
}

/**
 * FUNÇÃO: Formata a exibição das informações de lote (EXATAMENTE COMO NA TABELA)
 */
function formatarLoteDisplay(loteInfo) {
    let display = '';
    
    if (loteInfo.lote !== "N/A") {
        display += `LOTE: ${loteInfo.lote}`;
        if (loteInfo.loteSource) {
            display += ` [${getFonteLabel(loteInfo.loteSource)}]`;
        }
        display += '\n';
    }
    
    if (loteInfo.validade !== "N/A") {
        display += `VALIDADE: ${loteInfo.validade}`;
        if (loteInfo.validadeSource) {
            display += ` [${getFonteLabel(loteInfo.validadeSource)}]`;
        }
        display += '\n';
    }
    
    if (loteInfo.fabricacao !== "N/A") {
        display += `FABRICAÇÃO: ${loteInfo.fabricacao}`;
        if (loteInfo.fabricacaoSource) {
            display += ` [${getFonteLabel(loteInfo.fabricacaoSource)}]`;
        }
    }
    
    if (display === '') {
        display = 'Nenhuma informação';
    }
    
    return display.trim();
}

/**
 * FUNÇÃO: Formata as informações de lote para exportação (sem quebras de linha)
 */
function formatarLoteParaExportacao(loteInfo) {
    let exportacao = '';
    
    if (loteInfo.lote !== "N/A") {
        exportacao += `LOTE: ${loteInfo.lote}`;
        if (loteInfo.loteSource) {
            exportacao += ` (${getFonteLabel(loteInfo.loteSource)})`;
        }
        exportacao += ' | ';
    }
    
    if (loteInfo.validade !== "N/A") {
        exportacao += `VALIDADE: ${loteInfo.validade}`;
        if (loteInfo.validadeSource) {
            exportacao += ` (${getFonteLabel(loteInfo.validadeSource)})`;
        }
        exportacao += ' | ';
    }
    
    if (loteInfo.fabricacao !== "N/A") {
        exportacao += `FABRICAÇÃO: ${loteInfo.fabricacao}`;
        if (loteInfo.fabricacaoSource) {
            exportacao += ` (${getFonteLabel(loteInfo.fabricacaoSource)})`;
        }
    }
    
    // Remove o último separador se existir
    exportacao = exportacao.replace(/\s\|\s$/, '');
    
    if (exportacao === '') {
        exportacao = 'Nenhuma informação';
    }
    
    return exportacao;
}

/**
 * FUNÇÃO AUXILIAR: Retorna label amigável para a fonte
 */
function getFonteLabel(fonte) {
    const labels = {
        'campo': 'Campo',
        'descricao': 'Texto',
        'infAdProd': 'Inf.Prod',
        'infCpl': 'Inf.Cpl'
    };
    return labels[fonte] || fonte;
}

// ============================================
// FUNÇÕES DE EXPORTAÇÃO (Excel e PDF)
// ============================================

function exportToExcel() {
    if (!nfeData || !produtosExibidos.length) return;
    
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
        
        // Planilha de produtos (EXATAMENTE COMO NA TABELA)
        const produtosSheetData = [
            ["PRODUTOS/SERVIÇOS", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", ""],
            ["Código", "Descrição", "Lote/Validade/Fabricação", "EAN", "NCM", "Quantidade", "Unidade", "Valor Unitário", "Valor Total"]
        ];
        
        produtosExibidos.forEach(prod => {
            // Formata as informações de lote para exportação (em uma linha só)
            const loteExport = formatarLoteParaExportacao(prod.loteInfo);
            
            produtosSheetData.push([
                prod.codigo,
                prod.descricao,
                loteExport, // Usa a formatação para exportação
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
            ["DETALHES DE LOTE, VALIDADE E FABRICAÇÃO", "", "", "", ""],
            ["", "", "", "", ""],
            ["Item", "Descrição", "Lote", "Validade", "Fabricação", "Fonte"]
        ];
        
        produtosExibidos.forEach((prod, index) => {
            lotesSheetData.push([
                index + 1,
                prod.descricao,
                prod.loteInfo.lote !== "N/A" ? `${prod.loteInfo.lote} (${getFonteLabel(prod.loteInfo.loteSource)})` : "N/A",
                prod.loteInfo.validade !== "N/A" ? `${prod.loteInfo.validade} (${getFonteLabel(prod.loteInfo.validadeSource)})` : "N/A",
                prod.loteInfo.fabricacao !== "N/A" ? `${prod.loteInfo.fabricacao} (${getFonteLabel(prod.loteInfo.fabricacaoSource)})` : "N/A",
                `${getFonteLabel(prod.loteInfo.loteSource)}/${getFonteLabel(prod.loteInfo.validadeSource)}/${getFonteLabel(prod.loteInfo.fabricacaoSource)}`
            ]);
        });
        
        const lotesSheet = XLSX.utils.aoa_to_sheet(lotesSheetData);
        const lotesColWidths = [
            {wch: 8}, {wch: 40}, {wch: 20}, {wch: 15}, {wch: 15}, {wch: 20}
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
    if (!nfeData || !produtosExibidos.length) return;
    
    try {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('p', 'mm', 'a4');
        
        // Configurações
        const pageWidth = doc.internal.pageSize.getWidth();
        const margin = 15;
        let yPos = margin;
        
        // Cabeçalho
        doc.setFontSize(16);
        doc.setTextColor(41, 128, 185);
        doc.text("RELATÓRIO DE NOTA FISCAL ELETRÔNICA", pageWidth / 2, yPos, { align: 'center' });
        yPos += 8;
        
        doc.setFontSize(10);
        doc.setTextColor(100, 100, 100);
        doc.text(`Data de exportação: ${new Date().toLocaleDateString('pt-BR')} ${new Date().toLocaleTimeString('pt-BR')}`, pageWidth / 2, yPos, { align: 'center' });
        yPos += 15;
        
        // Informações básicas
        doc.setFontSize(12);
        doc.setTextColor(44, 62, 80);
        doc.setFont(undefined, 'bold');
        doc.text("INFORMAÇÕES DA NF-e", margin, yPos);
        yPos += 8;
        
        doc.setFontSize(10);
        doc.setFont(undefined, 'normal');
        doc.text(`Número: ${nfeData.nfeInfo.numero || 'N/A'}`, margin, yPos);
        doc.text(`Série: ${nfeData.nfeInfo.serie || 'N/A'}`, pageWidth / 2, yPos);
        yPos += 5;
        doc.text(`Data Emissão: ${formatDate(nfeData.nfeInfo.dataEmissao)}`, margin, yPos);
        doc.text(`Valor Total: R$ ${formatCurrency(nfeData.nfeInfo.valorTotal)}`, pageWidth / 2, yPos);
        yPos += 5;
        doc.text(`Chave de Acesso: ${nfeData.nfeInfo.chaveAcesso || 'N/A'}`, margin, yPos);
        yPos += 10;
        
        // Emitente
        doc.setFontSize(11);
        doc.setFont(undefined, 'bold');
        doc.text("EMITENTE", margin, yPos);
        yPos += 7;
        
        doc.setFontSize(9);
        doc.setFont(undefined, 'normal');
        doc.text(`Nome: ${nfeData.emitente.nome || 'N/A'}`, margin, yPos);
        yPos += 5;
        doc.text(`CNPJ: ${formatCNPJ(nfeData.emitente.cnpj) || 'N/A'}`, margin, yPos);
        yPos += 5;
        doc.text(`Endereço: ${nfeData.emitente.endereco || 'N/A'}`, margin, yPos);
        yPos += 5;
        doc.text(`Cidade/UF: ${nfeData.emitente.cidade || 'N/A'}`, margin, yPos);
        yPos += 10;
        
        // Destinatário
        doc.setFontSize(11);
        doc.setFont(undefined, 'bold');
        doc.text("DESTINATÁRIO", margin, yPos);
        yPos += 7;
        
        doc.setFontSize(9);
        doc.setFont(undefined, 'normal');
        doc.text(`Nome: ${nfeData.destinatario.nome || 'N/A'}`, margin, yPos);
        yPos += 5;
        doc.text(`CNPJ/CPF: ${formatCNPJ(nfeData.destinatario.cnpj) || 'N/A'}`, margin, yPos);
        yPos += 5;
        doc.text(`Cidade/UF: ${nfeData.destinatario.cidade || 'N/A'}`, margin, yPos);
        yPos += 15;
        
        // Tabela de produtos (EXATAMENTE COMO NA TABELA)
        doc.setFontSize(12);
        doc.setFont(undefined, 'bold');
        doc.text("PRODUTOS/SERVIÇOS", margin, yPos);
        yPos += 10;
        
        // Para PDF, usamos os dados exibidos na tabela
        const headers = [["Código", "Descrição", "Lote/Validade/Fabricação", "Qtd", "Valor Total"]];
        const data = produtosExibidos.map(prod => {
            // Formata as informações de lote para PDF (remove quebras de linha)
            const lotePDF = prod.loteDisplay.replace(/\n/g, ' | ');
            
            return [
                prod.codigo,
                prod.descricao.length > 25 ? prod.descricao.substring(0, 25) + '...' : prod.descricao,
                lotePDF.length > 30 ? lotePDF.substring(0, 30) + '...' : lotePDF,
                formatNumber(prod.quantidade),
                `R$ ${formatCurrency(prod.valorTotal)}`
            ];
        });
        
        // Adiciona total
        const totalProdutos = produtosExibidos.reduce((sum, prod) => sum + parseFloat(prod.valorTotal || 0), 0);
        data.push(["", "", "", "TOTAL:", `R$ ${formatCurrency(totalProdutos.toString())}`]);
        
        doc.autoTable({
            startY: yPos,
            head: headers,
            body: data,
            margin: { left: margin, right: margin },
            styles: { fontSize: 8, cellPadding: 2 },
            headStyles: { fillColor: [41, 128, 185], textColor: 255 },
            alternateRowStyles: { fillColor: [245, 245, 245] },
            columnStyles: {
                0: { cellWidth: 20 },
                1: { cellWidth: 35 },
                2: { cellWidth: 50 },
                3: { cellWidth: 15, halign: 'right' },
                4: { cellWidth: 25, halign: 'right' }
            },
            didDrawPage: function(data) {
                // Adiciona número da página
                doc.setFontSize(10);
                doc.text(`Página ${data.pageNumber}`, pageWidth / 2, doc.internal.pageSize.getHeight() - 10, { align: 'center' });
            }
        });
        
        // Última posição Y após a tabela
        let finalY = doc.lastAutoTable.finalY || yPos + 100;
        
        // Informações adicionais (se couber)
        if (finalY < doc.internal.pageSize.getHeight() - 50) {
            doc.setFontSize(12);
            doc.setFont(undefined, 'bold');
            doc.text("INFORMAÇÕES ADICIONAIS", margin, finalY + 15);
            
            doc.setFontSize(9);
            doc.setFont(undefined, 'normal');
            const infText = nfeData.informacoesAdicionais;
            if (infText && infText !== "Nenhuma informação adicional") {
                const lines = doc.splitTextToSize(infText, pageWidth - 2 * margin);
                doc.text(lines, margin, finalY + 25);
            } else {
                doc.text("Nenhuma informação adicional", margin, finalY + 25);
            }
        } else {
            // Adiciona nova página para informações adicionais
            doc.addPage();
            yPos = margin;
            
            doc.setFontSize(12);
            doc.setFont(undefined, 'bold');
            doc.text("INFORMAÇÕES ADICIONAIS", margin, yPos);
            yPos += 10;
            
            doc.setFontSize(9);
            doc.setFont(undefined, 'normal');
            const infText = nfeData.informacoesAdicionais;
            if (infText && infText !== "Nenhuma informação adicional") {
                const lines = doc.splitTextToSize(infText, pageWidth - 2 * margin);
                doc.text(lines, margin, yPos);
            } else {
                doc.text("Nenhuma informação adicional", margin, yPos);
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
            return `${day}/${month}/${year}`;
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
