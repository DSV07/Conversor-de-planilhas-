const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

// Colunas que sempre devem ser números
const colunasNumericas = ['Valor', 'Saldo', 'Inicial', 'Solicitada', 'Consumida', 'Saldo Atual'];

app.post('/upload', upload.single('arquivo'), async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(req.file.path);
        const sheet = workbook.worksheets[0];

        const unidadesSet = new Set();
        sheet.eachRow(row => {
            row.eachCell(cell => {
                if (cell.value && typeof cell.value === 'string' && cell.value.trim().startsWith('SESC -')) {
                    unidadesSet.add(cell.value.trim());
                }
            });
        });

        const unidades = Array.from(unidadesSet).sort();
        res.json({ unidades, filePath: req.file.path });
    } catch (err) {
        res.status(500).send('Erro ao processar arquivo.');
    }
});

app.post('/preview', async (req, res) => {
    const { filePath, unidade } = req.body;
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const sheet = workbook.worksheets[0];

        const dados = filtrarDados(sheet, unidade);

        res.json({ cabecalho: dados.itens.length ? Object.keys(dados.itens[0]) : [], preview: dados.itens.slice(0, 5) });
    } catch (err) {
        res.status(500).send('Erro ao gerar pré-visualização.');
    }
});

app.post('/filtrar', async (req, res) => {
    const { filePath, unidade } = req.body;
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const sheet = workbook.worksheets[0];

        const dados = filtrarDados(sheet, unidade);

        if (!dados.itens.length) return res.status(400).send('Não foi possível filtrar dados.');

        const newWB = new ExcelJS.Workbook();
        const ws = newWB.addWorksheet(unidade || 'Dados Filtrados');

        // --- Título principal ---
        ws.mergeCells('A1:F1');
        ws.getCell('A1').value = 'RELATÓRIO DE DADOS FILTRADOS';
        ws.getCell('A1').font = { bold: true, size: 16 };
        ws.getCell('A1').alignment = { horizontal: 'center' };
        ws.getCell('A1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'D9D9D9' }
        };

        // --- Informações gerais ---
        const infoLabels = ['Número da Ata', 'Objeto', 'Negociação', 'Início Vigência', 'Final Vigência'];
        const infoValues = [
            dados.info.numeroAta || '-',
            dados.info.objeto || '-',
            dados.info.negociacao || '-',
            dados.info.inicioVigencia || '-',
            dados.info.finalVigencia || '-'
        ];

        // Adicionar informações com formatação
        infoLabels.forEach((label, i) => {
            const rowNumber = i + 3;
            
            // Label
            ws.getCell(`A${rowNumber}`).value = label + ':';
            ws.getCell(`A${rowNumber}`).font = { bold: true };
            ws.getCell(`A${rowNumber}`).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'F2F2F2' }
            };
            ws.getCell(`A${rowNumber}`).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            
            // Valor
            ws.getCell(`B${rowNumber}`).value = infoValues[i];
            ws.getCell(`B${rowNumber}`).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            
            // Mesclar células se necessário para valores longos
            if (label === 'Objeto' && infoValues[i].length > 50) {
                ws.mergeCells(`B${rowNumber}:F${rowNumber}`);
            } else {
                ws.mergeCells(`B${rowNumber}:C${rowNumber}`);
            }
        });

        // Espaço entre informações e tabela
        const dataStartRow = infoLabels.length + 5;

        // --- Cabeçalho da tabela ---
        const cabecalho = Object.keys(dados.itens[0]);
        const headerRow = ws.getRow(dataStartRow);
        
        cabecalho.forEach((h, i) => {
            const cell = headerRow.getCell(i + 1);
            cell.value = h;
            cell.font = { bold: true, color: { argb: 'FFFFFF' } };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '1F4E78' }
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // --- Dados da tabela ---
        dados.itens.forEach((item, rowIndex) => {
            const row = ws.getRow(dataStartRow + rowIndex + 1);
            
            cabecalho.forEach((h, colIndex) => {
                const cell = row.getCell(colIndex + 1);
                let value = item[h] || '';
                
                // Formatar números
                if (colunasNumericas.includes(h)) {
                    if (value) {
                        let num = value.toString().replace(/\s/g, '').replace(/,/g, '.').replace(/[^0-9.-]/g, '');
                        cell.value = !isNaN(Number(num)) ? Number(num) : 0;
                    } else {
                        cell.value = 0;
                    }
                    cell.alignment = { horizontal: 'right', vertical: 'middle' };
                    cell.numFmt = '#,##0.00';
                } else {
                    cell.value = value;
                    cell.alignment = { horizontal: 'left', vertical: 'middle' };
                }
                
                // Formatação visual
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                
                // Zebra stripes
                if (rowIndex % 2 === 0) {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'F9F9F9' }
                    };
                }
            });
        });

        // --- Ajustar largura das colunas ---
        ws.columns.forEach((column, i) => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, (cell) => {
                let columnLength = cell.value ? cell.value.toString().length : 10;
                if (columnLength > maxLength) {
                    maxLength = columnLength;
                }
            });
            column.width = Math.min(maxLength + 2, 50);
        });

        // --- Data de geração ---
        const lastRow = dataStartRow + dados.itens.length + 3;
        ws.mergeCells(`A${lastRow}:F${lastRow}`);
        ws.getCell(`A${lastRow}`).value = `Gerado em: ${new Date().toLocaleString('pt-BR')}`;
        ws.getCell(`A${lastRow}`).font = { italic: true, color: { argb: '666666' } };
        ws.getCell(`A${lastRow}`).alignment = { horizontal: 'right' };

        // --- Salvar e enviar ---
        const outputPath = path.join('uploads', `filtrado_${Date.now()}.xlsx`);
        await newWB.xlsx.writeFile(outputPath);

        res.download(outputPath);

    } catch (err) {
        console.error(err);
        res.status(500).send('Erro ao gerar planilha.');
    }
});

// --- Função de filtragem melhorada ---
function filtrarDados(sheet, unidadeEscolhida) {
    let capturando = false;
    let cabecalho = [];
    const itens = [];
    const info = { 
        numeroAta: '', 
        objeto: '', 
        negociacao: '', 
        inicioVigencia: '', 
        finalVigencia: '' 
    };

    let encontrouCabecalho = false;

    sheet.eachRow((row, rowNumber) => {
        const valores = row.values.slice(1);
        const textoLinha = valores.map(v => v ? v.toString() : '').join(' ').trim();

        // Buscar informações de cabeçalho apenas nas primeiras linhas
        if (rowNumber <= 20) {
            // Número da Ata
            if (!info.numeroAta && textoLinha.toLowerCase().includes('número da ata')) {
                const match = textoLinha.match(/\b[A-Z]{2}-\d{4}-[A-Z]{3}-\d{3}-\d\b/);
                if (match) info.numeroAta = match[0];
            }

            // Objeto - capturar apenas o valor após "Objeto:"
            if (!info.objeto && textoLinha.toLowerCase().includes('objeto')) {
                const objetoMatch = textoLinha.match(/objeto:\s*(.+)/i);
                if (objetoMatch && objetoMatch[1]) {
                    info.objeto = objetoMatch[1].trim();
                } else {
                    // Se não encontrar pelo padrão, pegar a próxima célula
                    const idx = valores.findIndex(v => v && v.toString().toLowerCase().includes('objeto'));
                    if (idx !== -1 && valores[idx + 1]) {
                        info.objeto = valores[idx + 1].toString().trim();
                    }
                }
            }

            // Negociação - capturar apenas o valor após "Negociação:"
            if (!info.negociacao && textoLinha.toLowerCase().includes('negociação')) {
                const negociacaoMatch = textoLinha.match(/negocia[cç][aã]o:\s*(.+)/i);
                if (negociacaoMatch && negociacaoMatch[1]) {
                    info.negociacao = negociacaoMatch[1].trim();
                } else {
                    // Se não encontrar pelo padrão, pegar a próxima célula
                    const idx = valores.findIndex(v => v && v.toString().toLowerCase().includes('negociação'));
                    if (idx !== -1 && valores[idx + 1]) {
                        info.negociacao = valores[idx + 1].toString().trim();
                    }
                }
            }

            // Início Vigência
            if (!info.inicioVigencia) {
                const inicioMatch = textoLinha.match(/in[ií]cio.*?vig[eê]ncia.*?(\d{2}\/\d{2}\/\d{4})/i);
                if (inicioMatch) {
                    info.inicioVigencia = inicioMatch[1];
                } else {
                    // Tentar padrão alternativo
                    const inicioAltMatch = textoLinha.match(/(\d{2}\/\d{2}\/\d{4}).*?in[ií]cio/i);
                    if (inicioAltMatch) {
                        info.inicioVigencia = inicioAltMatch[1];
                    }
                }
            }

            // Final Vigência
            if (!info.finalVigencia) {
                const finalMatch = textoLinha.match(/fim.*?vig[eê]ncia.*?(\d{2}\/\d{2}\/\d{4})/i);
                if (finalMatch) {
                    info.finalVigencia = finalMatch[1];
                } else {
                    // Tentar padrão alternativo
                    const finalAltMatch = textoLinha.match(/(\d{2}\/\d{2}\/\d{4}).*?fim/i);
                    if (finalAltMatch) {
                        info.finalVigencia = finalAltMatch[1];
                    } else {
                        // Buscar por "validade até"
                        const validadeMatch = textoLinha.match(/validade.*?at[eé].*?(\d{2}\/\d{2}\/\d{4})/i);
                        if (validadeMatch) {
                            info.finalVigencia = validadeMatch[1];
                        }
                    }
                }
            }
        }

        // Verificar se é uma linha de unidade
        const linhaUnidade = valores.find(v => typeof v === 'string' && v.trim().startsWith('SESC -'));
        if (linhaUnidade) {
            const unidadeLinha = linhaUnidade.trim();
            capturando = unidadeEscolhida === 'Todas' || unidadeLinha === unidadeEscolhida;
            encontrouCabecalho = false; // Resetar ao encontrar nova unidade
            return;
        }

        // Se estamos capturando dados para a unidade selecionada
        if (capturando) {
            // Procurar pelo cabeçalho da tabela
            if (!encontrouCabecalho) {
                const temCabecalho = valores.some(v => 
                    v && typeof v === 'string' && 
                    (v.toLowerCase().includes('descrição') || 
                     v.toLowerCase().includes('item') ||
                     v.toLowerCase().includes('código'))
                );
                
                if (temCabecalho) {
                    cabecalho = valores.map(v => v || '');
                    encontrouCabecalho = true;
                    return;
                }
            } else {
                // Se já encontramos o cabeçalho, capturar os dados
                const linhaVazia = valores.every(v => v === null || v === '' || v.toString().trim() === '');
                const ehRodape = textoLinha.toLowerCase().includes('itens por unidade') || 
                                textoLinha.toLowerCase().includes('total') || 
                                textoLinha.toLowerCase().includes('observações') ||
                                textoLinha.toLowerCase().includes('subtotal');
                
                if (!linhaVazia && !ehRodape) {
                    const item = {};
                    cabecalho.forEach((coluna, index) => {
                        if (coluna) {
                            item[coluna] = valores[index] || '';
                        }
                    });
                    
                    // Só adicionar se pelo menos uma célula tem valor
                    if (Object.values(item).some(v => v !== '')) {
                        itens.push(item);
                    }
                }
            }
        }
    });

    return { info, itens };
}

app.listen(3000, () => console.log('Servidor rodando em http://localhost:3000'));