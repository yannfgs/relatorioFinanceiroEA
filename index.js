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


