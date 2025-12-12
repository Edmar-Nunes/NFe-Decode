<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Leitor de NF-e</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }

        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.2);
            overflow: hidden;
        }

        header {
            background: linear-gradient(to right, #2c3e50, #4a6491);
            color: white;
            padding: 25px 30px;
            text-align: center;
        }

        header h1 {
            font-size: 2.2rem;
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 15px;
        }

        header h1 i {
            color: #3498db;
        }

        header p {
            color: #bdc3c7;
            font-size: 1.1rem;
        }

        .upload-section {
            padding: 30px;
            background: #f8f9fa;
            border-bottom: 2px solid #e9ecef;
            text-align: center;
        }

        .file-input-container {
            display: inline-block;
            position: relative;
            margin-bottom: 20px;
        }

        .file-input-container input[type="file"] {
            position: absolute;
            left: 0;
            top: 0;
            opacity: 0;
            width: 100%;
            height: 100%;
            cursor: pointer;
        }

        .file-input-label {
            display: inline-flex;
            align-items: center;
            gap: 12px;
            background: #3498db;
            color: white;
            padding: 15px 30px;
            border-radius: 50px;
            font-size: 1.1rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(52, 152, 219, 0.3);
        }

        .file-input-label:hover {
            background: #2980b9;
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(52, 152, 219, 0.4);
        }

        .file-info {
            margin-top: 15px;
            color: #7f8c8d;
            font-size: 0.95rem;
        }

        .export-buttons {
            display: flex;
            gap: 15px;
            justify-content: center;
            margin-top: 20px;
        }

        .export-btn {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 12px 25px;
            border: none;
            border-radius: 8px;
            font-size: 1rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
        }

        .export-btn.excel {
            background: #27ae60;
            color: white;
        }

        .export-btn.pdf {
            background: #e74c3c;
            color: white;
        }

        .export-btn:disabled {
            background: #95a5a6;
            cursor: not-allowed;
            opacity: 0.7;
        }

        .export-btn:not(:disabled):hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
        }

        .content {
            padding: 30px;
        }

        .invoice-info {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 25px;
            margin-bottom: 30px;
        }

        .info-card {
            background: white;
            border-radius: 10px;
            padding: 20px;
            box-shadow: 0 3px 15px rgba(0, 0, 0, 0.08);
            border-left: 4px solid #3498db;
        }

        .info-card h3 {
            color: #2c3e50;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 2px solid #f1f1f1;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .info-item {
            display: flex;
            margin-bottom: 10px;
            padding: 5px 0;
        }

        .info-label {
            font-weight: 600;
            color: #34495e;
            min-width: 150px;
            flex-shrink: 0;
        }

        .info-value {
            color: #2c3e50;
            word-break: break-word;
        }

        .currency {
            font-family: 'Courier New', monospace;
            font-weight: 600;
        }

        .products-section {
            margin-top: 30px;
        }

        .section-title {
            color: #2c3e50;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #3498db;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .no-data {
            text-align: center;
            padding: 40px;
            color: #7f8c8d;
            font-style: italic;
            background: #f8f9fa;
            border-radius: 8px;
        }

        .no-data i {
            font-size: 3rem;
            margin-bottom: 15px;
            color: #bdc3c7;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
            border-radius: 8px;
            overflow: hidden;
        }

        thead {
            background: linear-gradient(to right, #2c3e50, #4a6491);
            color: white;
        }

        th {
            padding: 15px 10px;
            text-align: left;
            font-weight: 600;
            font-size: 0.95rem;
        }

        td {
            padding: 12px 10px;
            border-bottom: 1px solid #e9ecef;
            font-size: 0.9rem;
        }

        tbody tr:hover {
            background-color: #f8f9fa;
        }

        .lote-info {
            font-size: 0.85rem;
            line-height: 1.4;
            white-space: pre-line;
            max-width: 200px;
            color: #2c3e50;
        }

        footer {
            text-align: center;
            padding: 20px;
            color: #7f8c8d;
            font-size: 0.9rem;
            border-top: 1px solid #e9ecef;
            background: #f8f9fa;
        }

        @media (max-width: 768px) {
            .container {
                margin: 10px;
                border-radius: 10px;
            }
            
            .invoice-info {
                grid-template-columns: 1fr;
            }
            
            .info-item {
                flex-direction: column;
            }
            
            .info-label {
                min-width: auto;
                margin-bottom: 5px;
            }
            
            th, td {
                padding: 8px 5px;
                font-size: 0.85rem;
            }
            
            .export-buttons {
                flex-direction: column;
                align-items: center;
            }
            
            .export-btn {
                width: 100%;
                max-width: 300px;
                justify-content: center;
            }
        }
    </style>
    <!-- Bibliotecas necessárias -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.28/jspdf.plugin.autotable.min.js"></script>
</head>
<body>
    <div class="container">
        <header>
            <h1><i class="fas fa-file-invoice"></i> Leitor de NF-e</h1>
            <p>Carregue um arquivo XML de NF-e para visualizar e exportar os dados</p>
        </header>

        <div class="upload-section">
            <div class="file-input-container">
                <input type="file" id="xmlFile" accept=".xml">
                <label for="xmlFile" class="file-input-label">
                    <i class="fas fa-upload"></i> Selecionar Arquivo XML
                </label>
            </div>
            <div class="file-info">
                <i class="fas fa-info-circle"></i> Selecione um arquivo XML de Nota Fiscal Eletrônica (NF-e)
            </div>
            
            <div class="export-buttons">
                <button id="excelBtn" class="export-btn excel" disabled>
                    <i class="fas fa-file-excel"></i> Exportar para Excel
                </button>
                <button id="pdfBtn" class="export-btn pdf" disabled>
                    <i class="fas fa-file-pdf"></i> Exportar para PDF
                </button>
            </div>
        </div>

        <div class="content">
            <div id="noData" style="display: block;">
                <div class="no-data">
                    <i class="fas fa-file-invoice-dollar"></i>
                    <h3>Nenhuma NF-e Carregada</h3>
                    <p>Selecione um arquivo XML para visualizar as informações da nota fiscal</p>
                </div>
            </div>

            <div id="invoiceInfo" style="display: none;">
                <div class="invoice-info">
                    <div class="info-card">
                        <h3><i class="fas fa-building"></i> Emitente</h3>
                        <div id="emitente"></div>
                    </div>
                    
                    <div class="info-card">
                        <h3><i class="fas fa-user-tag"></i> Destinatário</h3>
                        <div id="destinatario"></div>
                    </div>
                    
                    <div class="info-card">
                        <h3><i class="fas fa-info-circle"></i> Informações da NF-e</h3>
                        <div id="nfeInfo"></div>
                    </div>
                    
                    <div class="info-card">
                        <h3><i class="fas fa-calculator"></i> Totais</h3>
                        <div id="totais"></div>
                    </div>
                </div>

                <div class="info-card">
                    <h3><i class="fas fa-sticky-note"></i> Informações Adicionais</h3>
                    <div id="infAdicionais"></div>
                </div>

                <div class="products-section">
                    <h3 class="section-title"><i class="fas fa-boxes"></i> Produtos/Serviços</h3>
                    <div id="productsTable"></div>
                </div>
            </div>
        </div>

        <footer>
            <p>Sistema de Leitura de NF-e &copy; 2025 | Desenvolvido para processamento de notas fiscais eletrônicas</p>
        </footer>
    </div>

    <script>
        // ============================================
        // VARIÁVEIS GLOBAIS
        // ============================================
        let nfeData = null;
        let xmlDoc = null;
        let produtosExibidos = [];

        // ============================================
        // INICIALIZAÇÃO
        // ============================================
        window.onload = function() {
            document.getElementById("xmlFile").addEventListener("change", handleFileSelect);
            document.getElementById("excelBtn").addEventListener("click", exportToExcel);
            document.getElementById("pdfBtn").addEventListener("click", exportToPDF);
            
            document.getElementById("excelBtn").disabled = true;
            document.getElementById("pdfBtn").disabled = true;
        };

        // ============================================
        // FUNÇÃO PRINCIPAL - CARREGAR ARQUIVO XML
        // ============================================
        function handleFileSelect(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const parser = new DOMParser();
                    xmlDoc = parser.parseFromString(e.target.result, "text/xml");
                    
                    const nfe = xmlDoc.querySelector("NFe, nfeProc");
                    if (!nfe) {
                        throw new Error("XML não é uma NF-e válida");
                    }

                    parseXML(xmlDoc);
                    
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
        // PARSE DO XML
        // ============================================
        function parseXML(xml) {
            const getTag = (tag, parent = xml) => {
                const el = parent.getElementsByTagName(tag)[0];
                return el ? el.textContent : "";
            };

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
                    
                    const loteInfo = extrairLoteValidadeFabricacao(det);
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
        // FUNÇÕES PARA LOTE/VALIDADE
        // ============================================
        function extrairLoteValidadeFabricacao(detElement) {
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
            
            if (loteInfo.validades.length > 0) {
                loteInfo.validades.forEach((val, i) => {
                    display += `VALIDADE: ${val}\n`;
                });
            }
            
            if (loteInfo.lotes.length > 0) {
                loteInfo.lotes.forEach((lote, i) => {
                    display += `LOTE: ${lote}\n`;
                });
            }
            
            if (loteInfo.fabricacoes.length > 0) {
                loteInfo.fabricacoes.forEach((fab, i) => {
                    display += `FABRICAÇÃO: ${fab}\n`;
                });
            }
            
            return display.trim() || 'N/A';
        }

        // ============================================
        // EXPORTAR PARA EXCEL
        // ============================================
        function exportToExcel() {
            if (!nfeData || !produtosExibidos.length) {
                alert('Nenhum dado disponível para exportação. Carregue um arquivo XML primeiro.');
                return;
            }
            
            try {
                const wb = XLSX.utils.book_new();
                
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
                
                if (nfeData.informacoesAdicionais && nfeData.informacoesAdicionais !== "Nenhuma informação adicional") {
                    const infLines = nfeData.informacoesAdicionais.split('\n');
                    infLines.forEach(line => {
                        infoSheetData.push([line, "", "", "", ""]);
                    });
                } else {
                    infoSheetData.push(["Nenhuma informação adicional", "", "", "", ""]);
                }
                
                const infoSheet = XLSX.utils.aoa_to_sheet(infoSheetData);
                const infoColWidths = [
                    {wch: 25}, {wch: 40}, {wch: 5}, {wch: 25}, {wch: 40}
                ];
                infoSheet['!cols'] = infoColWidths;
                
                XLSX.utils.book_append_sheet(wb, infoSheet, "Informações NF-e");
                
                const produtosSheetData = [
                    ["PRODUTOS/SERVIÇOS", "", "", "", "", "", "", "", "", ""],
                    ["", "", "", "", "", "", "", "", "", ""],
                    ["Item", "Código", "Descrição", "Validade(s)", "Lote(s)", "Qtd", "Unidade", "Valor Unitário", "Valor Total", "Informações Completas"]
                ];
                
                produtosExibidos.forEach((prod, index) => {
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
                
                const totalProdutos = produtosExibidos.reduce((sum, prod) => sum + parseFloat(prod.valorTotal || 0), 0);
                produtosSheetData.push(
                    ["", "", "", "", "", "", "", "", "", ""],
                    ["", "", "", "", "", "", "", "TOTAL:", `R$ ${formatCurrency(totalProdutos.toString())}`, ""]
                );
                
                const produtosSheet = XLSX.utils.aoa_to_sheet(produtosSheetData);
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
                
                const fileName = `NF-e_${nfeData.nfeInfo.numero || 'export'}_${new Date().toISOString().slice(0,10)}.xlsx`;
                XLSX.writeFile(wb, fileName);
                
            } catch (error) {
                alert("Erro ao exportar para Excel: " + error.message);
                console.error(error);
            }
        }

        // ============================================
        // EXPORTAR PARA PDF (CORRIGIDO)
        // ============================================
        function exportToPDF() {
            if (!nfeData || !produtosExibidos.length) {
                alert('Nenhum dado disponível para exportação. Carregue um arquivo XML primeiro.');
                return;
            }
            
            try {
                if (typeof jsPDF === 'undefined') {
                    alert('Biblioteca jsPDF não encontrada. Certifique-se de incluí-la no HTML.');
                    return;
                }
                
                const { jsPDF } = window.jspdf;
                
                // Criar documento em modo paisagem
                const doc = new jsPDF({
                    orientation: 'landscape',
                    unit: 'mm',
                    format: 'a4'
                });
                
                const pageWidth = doc.internal.pageSize.getWidth();
                const pageHeight = doc.internal.pageSize.getHeight();
                const margin = 10;
                let yPos = margin;
                
                // Título principal
                doc.setFontSize(18);
                doc.setTextColor(0, 0, 0);
                doc.setFont("helvetica", "bold");
                doc.text("NOTA FISCAL ELETRÔNICA", margin, yPos);
                yPos += 8;
                
                // Informações básicas
                doc.setFontSize(10);
                doc.setFont("helvetica", "normal");
                doc.text(`Número: ${nfeData.nfeInfo.numero || 'N/A'}`, margin, yPos);
                doc.text(`Série: ${nfeData.nfeInfo.serie || 'N/A'}`, margin + 40, yPos);
                doc.text(`Data Emissão: ${formatDate(nfeData.nfeInfo.dataEmissao) || 'N/A'}`, margin + 80, yPos);
                doc.text(`Valor Total: R$ ${formatCurrency(nfeData.nfeInfo.valorTotal) || '0,00'}`, margin + 160, yPos);
                yPos += 8;
                
                // Linha divisória
                doc.setDrawColor(0, 0, 0);
                doc.line(margin, yPos, pageWidth - margin, yPos);
                yPos += 12;
                
                // Emitente - lado esquerdo
                doc.setFontSize(11);
                doc.setFont("helvetica", "bold");
                doc.text("EMITENTE", margin, yPos);
                doc.setFontSize(10);
                doc.setFont("helvetica", "normal");
                doc.text(`${nfeData.emitente.nome || 'N/A'}`, margin + 25, yPos);
                yPos += 5;
                doc.text(`CNPJ/CPF: ${formatCNPJ(nfeData.emitente.cnpj) || 'N/A'}`, margin + 25, yPos);
                yPos += 5;
                doc.text(`Endereço: ${nfeData.emitente.endereco || 'N/A'}`, margin + 25, yPos);
                yPos += 5;
                doc.text(`Cidade/UF: ${nfeData.emitente.cidade || 'N/A'}`, margin + 25, yPos);
                
                // Destinatário - lado direito
                const rightColumn = pageWidth / 2;
                yPos = margin + 12;
                doc.setFontSize(11);
                doc.setFont("helvetica", "bold");
                doc.text("DESTINATÁRIO", rightColumn, yPos);
                doc.setFontSize(10);
                doc.setFont("helvetica", "normal");
                doc.text(`${nfeData.destinatario.nome || 'N/A'}`, rightColumn + 30, yPos);
                yPos += 5;
                doc.text(`CNPJ/CPF: ${formatCNPJ(nfeData.destinatario.cnpj) || 'N/A'}`, rightColumn + 30, yPos);
                yPos += 5;
                doc.text(`Cidade/UF: ${nfeData.destinatario.cidade || 'N/A'}`, rightColumn + 30, yPos);
                
                yPos = margin + 45;
                
                // Tabela de produtos
                doc.setFontSize(12);
                doc.setFont("helvetica", "bold");
                doc.text("PRODUTOS/SERVIÇOS", margin, yPos);
                yPos += 8;
                
                // Cabeçalhos da tabela
                const headers = [["Item", "Código", "Descrição", "Qtd", "Unid.", "Valor Unit.", "Valor Total"]];
                
                // Dados da tabela
                const tableData = produtosExibidos.map((prod, index) => {
                    let descricao = prod.descricao || 'N/A';
                    if (descricao.length > 60) {
                        descricao = descricao.substring(0, 57) + '...';
                    }
                    
                    return [
                        (index + 1).toString(),
                        prod.codigo || 'N/A',
                        descricao,
                        formatNumber(prod.quantidade || '0'),
                        prod.unidade || 'UN',
                        `R$ ${formatCurrency(prod.valorUnitario || '0')}`,
                        `R$ ${formatCurrency(prod.valorTotal || '0')}`
                    ];
                });
                
                // Adicionar linha de total
                const totalGeral = produtosExibidos.reduce((sum, prod) => sum + parseFloat(prod.valorTotal || 0), 0);
                tableData.push([
                    "", "", "TOTAL:", "", "", "",
                    `R$ ${formatCurrency(totalGeral.toString())}`
                ]);
                
                // Larguras das colunas (total 277mm em A4 paisagem)
                const columnStyles = {
                    0: { cellWidth: 15, halign: 'center' },
                    1: { cellWidth: 25, halign: 'left' },
                    2: { cellWidth: 100, halign: 'left' },
                    3: { cellWidth: 20, halign: 'right' },
                    4: { cellWidth: 20, halign: 'center' },
                    5: { cellWidth: 30, halign: 'right' },
                    6: { cellWidth: 30, halign: 'right' }
                };
                
                // Gerar tabela com autoTable
                doc.autoTable({
                    startY: yPos,
                    head: headers,
                    body: tableData,
                    margin: { left: margin, right: margin },
                    tableWidth: pageWidth - (2 * margin),
                    styles: {
                        fontSize: 9,
                        cellPadding: 3,
                        overflow: 'linebreak',
                        lineColor: [0, 0, 0],
                        lineWidth: 0.1,
                        valign: 'middle'
                    },
                    headStyles: {
                        fillColor: [220, 220, 220],
                        textColor: [0, 0, 0],
                        fontSize: 10,
                        fontStyle: 'bold',
                        halign: 'center',
                        lineWidth: 0.2
                    },
                    bodyStyles: {
                        lineWidth: 0.1
                    },
                    alternateRowStyles: {
                        fillColor: [245, 245, 245]
                    },
                    columnStyles: columnStyles,
                    didDrawCell: function(data) {
                        if (data.row.index === tableData.length - 1 && data.column.index === 2) {
                            doc.setFont("helvetica", "bold");
                            doc.setTextColor(0, 0, 0);
                        }
                        if (data.row.index === tableData.length - 1 && data.column.index === 6) {
                            doc.setFont("helvetica", "bold");
                            doc.setTextColor(0, 0, 0);
                        }
                    },
                    willDrawPage: function(data) {
                        // Borda da página
                        doc.setDrawColor(0, 0, 0);
                        doc.setLineWidth(0.5);
                        doc.rect(margin - 5, margin - 5, pageWidth - (2 * margin) + 10, pageHeight - (2 * margin) + 10);
                    },
                    didDrawPage: function(data) {
                        // Rodapé
                        doc.setFontSize(8);
                        doc.setTextColor(100, 100, 100);
                        doc.text(
                            `Página ${data.pageNumber} - Exportado em ${new Date().toLocaleDateString('pt-BR')}`, 
                            pageWidth / 2, 
                            pageHeight - 5, 
                            { align: 'center' }
                        );
                    }
                });
                
                // Salvar PDF
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
    </script>
</body>
</html>
