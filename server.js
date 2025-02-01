const express = require('express');
const multer = require('multer');
const fs = require('fs');
const xml2js = require('xml2js');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');
const path = require('path');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Configuração para servir arquivos estáticos
app.use(express.static(path.join(__dirname, 'public')));

// Rota para a página inicial (index.html)
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Função para garantir que todas as células da linha tenham bordas
function adicionarBordas(worksheet, linha) {
    const borda = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
    };
    for (let col = 1; col <= 6; col++) {
        const cell = worksheet.getCell(`${String.fromCharCode(64 + col)}${linha}`);
        cell.border = borda;
    }
}

// Função para enviar e-mail com a planilha anexada
async function enviarEmail(caminhoArquivo, email, appPassword, recipientEmail) {
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: email,
            pass: appPassword,
        },
    });

    const mailOptions = {
        from: email,
        to: recipientEmail,
        subject: 'Planilha Packlist',
        text: 'Segue em anexo a planilha Packlist.',
        attachments: [
            {
                filename: 'Packlist.xlsx',
                path: caminhoArquivo,
            },
        ],
    };

    try {
        await transporter.sendMail(mailOptions);
        console.log('E-mail enviado com sucesso!');
    } catch (error) {
        console.error('Erro ao enviar e-mail:', error);
        throw error;
    }
}

// Função para limpar a planilha
async function limparPlanilha() {
    const caminhoArquivo = path.join(__dirname, 'Packlist.xlsx');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(caminhoArquivo);
    const worksheet = workbook.getWorksheet(1);

    let linha = 14;
    while (worksheet.getCell(`A${linha}`).value) {
        worksheet.spliceRows(linha, 1);
    }

    await workbook.xlsx.writeFile(caminhoArquivo);
    console.log('Planilha limpa com sucesso!');
}

// Função para criar a planilha inicial (se não existir)
async function criarPlanilhaInicial() {
    const caminhoArquivo = path.join(__dirname, 'Packlist.xlsx');
    if (!fs.existsSync(caminhoArquivo)) {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Packlist');

        // Adicionar cabeçalho
        worksheet.addRow(['Número da Nota Fiscal', 'Part Number', 'Volume', 'Call Num', 'Usada', 'Nova']);

        await workbook.xlsx.writeFile(caminhoArquivo);
        console.log('Planilha inicial criada com sucesso!');
    }
}

// Função para atualizar a Packlist
async function atualizarPacklist(dados) {
    const caminhoArquivo = path.join(__dirname, 'Packlist.xlsx');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(caminhoArquivo);
    const worksheet = workbook.getWorksheet(1);

    console.log('Dados para inserir na planilha:', dados);

    // Encontrar a próxima linha vazia a partir da linha 14
    let linha = 14;
    while (worksheet.getCell(`A${linha}`).value) {
        linha++;
    }

    // Inserir os dados na linha encontrada
    worksheet.getCell(`A${linha}`).value = dados.numeroNotaFiscal;
    worksheet.getCell(`B${linha}`).value = dados.partNumber;
    worksheet.getCell(`C${linha}`).value = dados.volume;
    worksheet.getCell(`D${linha}`).value = dados.callNum;

    if (dados.condicao === 'USADA') {
        worksheet.getCell(`E${linha}`).value = 'X';
    } else if (dados.condicao === 'NOVA') {
        worksheet.getCell(`F${linha}`).value = 'X';
    }

    // Aplicar bordas à linha
    adicionarBordas(worksheet, linha);

    await workbook.xlsx.writeFile(caminhoArquivo);
    console.log('Packlist atualizada com sucesso!');
}

// Função para ordenar as linhas da planilha em ordem crescente
async function ordenarPlanilha() {
    const caminhoArquivo = path.join(__dirname, 'Packlist.xlsx');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(caminhoArquivo);
    const worksheet = workbook.getWorksheet(1);

    // Obter todas as linhas a partir da linha 14
    const rows = [];
    let linha = 14;
    while (worksheet.getCell(`A${linha}`).value) {
        const rowData = {
            numeroNotaFiscal: worksheet.getCell(`A${linha}`).value,
            partNumber: worksheet.getCell(`B${linha}`).value,
            volume: worksheet.getCell(`C${linha}`).value,
            callNum: worksheet.getCell(`D${linha}`).value,
            usada: worksheet.getCell(`E${linha}`).value,
            nova: worksheet.getCell(`F${linha}`).value,
        };
        rows.push(rowData);
        linha++;
    }

    // Ordenar as linhas pelo número da nota fiscal (coluna A)
    rows.sort((a, b) => a.numeroNotaFiscal - b.numeroNotaFiscal);

    // Limpar as linhas da planilha a partir da linha 14
    let linhaAtual = 14;
    while (worksheet.getCell(`A${linhaAtual}`).value) {
        worksheet.spliceRows(linhaAtual, 1);
    }

    // Inserir as linhas ordenadas na planilha
    for (const row of rows) {
        worksheet.insertRow(linhaAtual, [
            row.numeroNotaFiscal,
            row.partNumber,
            row.volume,
            row.callNum,
            row.usada,
            row.nova,
        ]);

        // Aplicar bordas à linha
        adicionarBordas(worksheet, linhaAtual);

        linhaAtual++;
    }

    await workbook.xlsx.writeFile(caminhoArquivo);
    console.log('Planilha ordenada com sucesso!');
}

// Função para aguardar um tempo (em milissegundos)
function aguardar(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
}

// Rota para upload de XML
// Rota para upload de XML
app.post('/upload', upload.array('xmlFiles'), async (req, res) => {
    if (!req.files || req.files.length === 0) {
        return res.status(400).json({ error: 'Nenhum arquivo enviado.' });
    }

    const files = req.files;
    const email = req.body.email;
    const appPassword = req.body.appPassword;
    const recipientEmail = req.body.recipientEmail;

    try {
        // Criar a planilha inicial (se não existir)
        await criarPlanilhaInicial();

        for (let i = 0; i < files.length; i++) {
            const filePath = files[i].path;
            const fileContent = fs.readFileSync(filePath, 'utf-8');

            console.log(`Processando arquivo ${i + 1} de ${files.length}: ${files[i].originalname}`);

            const result = await xml2js.parseStringPromise(fileContent);
            const nfe = result.nfeProc.NFe[0].infNFe[0];
            const numeroNotaFiscal = nfe.ide[0].nNF[0];
            const itens = nfe.det;

            for (const item of itens) {
                const partNumber = item.prod[0].cProd[0];
                const volume = parseFloat(item.prod[0].qCom[0]);
                const descricao = item.prod[0].xProd[0];
                const infCpl = nfe.infAdic[0].infCpl[0];

                let condicao = 'NOVA';
                if (infCpl.includes('PECA USADA') || infCpl.includes('DOA')) {
                    condicao = 'USADA';
                }

                const callNumMatch = infCpl.match(/@REF (\d{9})/);
                const callNum = callNumMatch ? callNumMatch[1] : 'N/A';

                const dadosPACKLIST = {
                    numeroNotaFiscal,
                    partNumber,
                    volume,
                    callNum,
                    condicao,
                };

                console.log('Atualizando a planilha com os dados:', dadosPACKLIST);
                await atualizarPacklist(dadosPACKLIST);
            }

            const novoCaminho = path.join(__dirname, 'uploads', `${numeroNotaFiscal}.xml`);
            fs.renameSync(filePath, novoCaminho);
        }

        // Ordenar a planilha após a inserção dos dados
        await ordenarPlanilha();

        const caminhoArquivo = path.join(__dirname, 'Packlist.xlsx');
        console.log('Verificando conteúdo da planilha antes de enviar:', caminhoArquivo);

        // Verificar o conteúdo da planilha antes de enviar
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(caminhoArquivo);
        const worksheet = workbook.getWorksheet(1);
        worksheet.eachRow((row, rowNumber) => {
            console.log(`Linha ${rowNumber}:`, row.values);
        });

        // Enviar a planilha por e-mail
        await enviarEmail(caminhoArquivo, email, appPassword, recipientEmail);

        // Aguardar 5 segundos antes de limpar a planilha
        console.log('Aguardando 5 segundos antes de limpar a planilha...');
        await aguardar(5000);

        // Limpar a planilha APÓS o envio do e-mail
        await limparPlanilha();

        // Deletar os arquivos temporários
        for (const file of files) {
            const filePath = file.path;
            if (fs.existsSync(filePath)) {
                try {
                    fs.unlinkSync(filePath);
                    console.log(`Arquivo ${filePath} excluído com sucesso.`);
                } catch (err) {
                    console.error(`Erro ao excluir arquivo ${filePath}:`, err);
                }
            } else {
                console.warn(`Arquivo ${filePath} não encontrado para exclusão.`);
            }
        }

        res.json({ success: true, message: 'Arquivos processados, planilha atualizada, e-mail enviado e planilha limpa.' });
    } catch (error) {
        console.error('Erro ao processar XML:', error);
        res.status(500).json({ error: 'Erro interno ao processar o XML.' });
    }
});

// Iniciar o servidor
app.listen(3000, () => {
    console.log('Servidor rodando na porta 3000');
});