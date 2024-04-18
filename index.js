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
    
  let processedData = {
        totalSales: 0,
        totalExpenses: 0
    };

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
     await generatePDF(report);
}


main();

const nodemailer = require('nodemailer');

async function sendEmailWithAttachment() {
    const transporter = nodemailer.createTransport({
        service: 'gmail', // use 'yahoo' para Yahoo, 'hotmail' para Hotmail, etc.
        auth: {
            user: 'seuemail@gmail.com', // seu endereço de email
            pass: 'suasenha' // sua senha
        }
    });

    const mailOptions = {
        from: 'seuemail@gmail.com',     // seu endereço de email
        to: 'destinatario@gmail.com',   // endereço de email do destinatário
        subject: 'Relatório Financeiro',    // Linha de assunto
        text: 'Segue anexo o relatório financeiro.',    // corpo do email
        attachments: [  // arquivo(s) anexo(s)
            {
                filename: 'RelatórioFinanceiro.pdf',
                path: './RelatórioFinanceiro.pdf'
            }
        ]
    };

    return new Promise((resolve, reject) => {
        transporter.sendMail(mailOptions, (error, info) => {
            if (error) {
                reject(error);
            } else {
                resolve(info.response);
            }
        });
    });
}


async function main() {
    const data = await readExcelFile();

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

    console.log("Enviando email com o relatório em PDF...");

    try {
        const emailResponse = await sendEmailWithAttachment();
        console.log("Email enviado com sucesso:", emailResponse);
    } catch (error) {
        console.log("Erro ao enviar o email:", error);
    }

<<<<<<< HEAD
}
=======
  }

>>>>>>> fca15418043b8ba771880245df1c41a88769a8f8
