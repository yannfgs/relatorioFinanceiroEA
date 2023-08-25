const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

// Ler dados do arquivo Excel
async function readlExcelFile() {
    const workbook = new ExcelJS.Workbook(); // Instancia o ExcelJS
    await workbook.xlsx.readFile('relatorioFinanceiro.xlsx'); // Lê o arquivo Excel

    const worksheet = workbook.getWorksheet('relatorioFinanceiroGerencial'); // Pega a aba do arquivo Excel

    let processedData = {
        totalSales: 0,
        totalExpenses: 0
    };

    worksheet.eachRow((row, rowNumber) => {
        // Suponha que a coluna 2 tem vendas e a coluna 3 tem despesas
        const sales = row.getCell(2).value;
        const expenses = row.getCell(3).value;

        processedData.totalSales += sales;
        processedData.totalExpenses += expenses;
    });

    processedData.profit = processedData.totalSales - processedData.totalExpenses;

    return processedData;
}


// Gerar relatório em PDF
async function generatePDF(report) {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();

    await page.setContent(report);
    await page.pdf({ path: 'relatorioFinanceiro.pdf', format: 'A4' });

    await browser.close();
}

async function main() {
    const data = await readlExcelFile();

    // Criar relatório (texto)
    const report = `
    <html>
    <head><title>Relatório Financeiro</title></head>
    <body>
    <h1>Relatório Financeiro</h1>
    <p>Total de Vendas: ${data.totalSales}</p>
    <p>Total de Despesas: ${data.totalExpenses}</p>
    <p>Lucro: ${data.profit}</p>
    </body>
    </html>
    `;

    console.log("Dados processados:", data);
    console.log("Gerando PDF...");

    await generatePDF(report);
}

main();