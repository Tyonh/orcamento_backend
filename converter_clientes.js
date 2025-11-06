const exceljs = require("exceljs");
const sqlite3 = require("sqlite3").verbose();
const path = require("path");

// --- CONFIGURAÇÃO ---
// 1. O nome do seu arquivo Excel com a lista de clientes
const EXCEL_PATH = path.join(__dirname, "CLIENTES.xlsx");
// 2. O nome das colunas no seu Excel
const COLUNA_ID = "CGC_CPF";
const COLUNA_NOME = "RAZAO_SOCIAL";
// 3. O nome do arquivo de banco de dados que será criado
const DB_PATH = path.join(__dirname, "clientes.db");
// --------------------

// Função assíncrona para executar o script
async function converterClientesParaSqlite() {
  console.log(`Iniciando conversão de ${EXCEL_PATH}...`);

  // Conecta ao banco de dados (ou cria se não existir)
  const db = new sqlite3.Database(DB_PATH, (err) => {
    if (err) {
      return console.error("Erro ao abrir o banco de dados:", err.message);
    }
    console.log(`Conectado ao banco de dados SQLite em ${DB_PATH}`);
  });

  // Carrega o arquivo Excel
  const workbook = new exceljs.Workbook();
  try {
    await workbook.xlsx.readFile(EXCEL_PATH);
  } catch (error) {
    console.error(
      `Erro ao ler o arquivo Excel. Verifique se o nome "${EXCEL_PATH}" está correto.`
    );
    db.close();
    return;
  }

  // Pega a primeira planilha (ou a planilha 'clientes_jer' se existir)
  const worksheet =
    workbook.getWorksheet("clientes_jer") || workbook.worksheets[0];
  if (!worksheet) {
    console.error("Nenhuma planilha encontrada no arquivo Excel.");
    db.close();
    return;
  }

  // Encontra os índices das colunas
  let colIdIndex = -1;
  let colNomeIndex = -1;

  worksheet.getRow(1).eachCell((cell, colNumber) => {
    const valorCelula = cell.value ? cell.value.toString().trim() : "";
    if (valorCelula === COLUNA_ID) {
      colIdIndex = colNumber;
    }
    if (valorCelula === COLUNA_NOME) {
      colNomeIndex = colNumber;
    }
  });

  if (colIdIndex === -1 || colNomeIndex === -1) {
    console.error(
      `Erro: Não foi possível encontrar as colunas "${COLUNA_ID}" e/ou "${COLUNA_NOME}" na primeira linha do Excel.`
    );
    db.close();
    return;
  }

  console.log(`Coluna "${COLUNA_ID}" encontrada.`);
  console.log(`Coluna "${COLUNA_NOME}" encontrada.`);

  // Prepara o banco de dados
  db.serialize(() => {
    // 1. Apaga a tabela antiga, se existir
    db.run("DROP TABLE IF EXISTS clientes", (err) => {
      if (err) return console.error("Erro ao apagar tabela:", err.message);
      console.log("Tabela 'clientes' antiga (se existia) foi apagada.");
    });

    // 2. Cria a nova tabela com 'id_cliente' e 'nome'
    // O 'nome' terá o COLLATE NOCASE para buscas case-insensitive
    db.run(
      "CREATE TABLE clientes (id_cliente TEXT PRIMARY KEY, nome TEXT NOT NULL COLLATE NOCASE)",
      (err) => {
        if (err) return console.error("Erro ao criar tabela:", err.message);
        console.log("Nova tabela 'clientes' (id_cliente, nome) criada.");
      }
    );

    // 3. Prepara o comando de inserção
    const sql = "INSERT INTO clientes (id_cliente, nome) VALUES (?, ?)";
    const stmt = db.prepare(sql, (err) => {
      if (err) return console.error("Erro ao preparar statement:", err.message);
    });

    // 4. Itera sobre as linhas do Excel
    console.log("Iniciando inserção dos clientes...");
    let contador = 0;
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return; // Pula a primeira linha (cabeçalho)

      const idCliente = row.getCell(colIdIndex).value;
      const nomeCliente = row.getCell(colNomeIndex).value;

      if (idCliente && nomeCliente) {
        // Converte para string e remove espaços extras
        const idLimpo = idCliente.toString().trim();
        const nomeLimpo = nomeCliente.toString().trim();

        // Insere os dois valores no banco de dados
        stmt.run([idLimpo, nomeLimpo], function (err) {
          if (err) {
            // Ignora erros de 'PRIMARY KEY constraint failed' (ID duplicado)
            if (!err.message.includes("constraint failed")) {
              console.warn(`Erro ao inserir "${nomeLimpo}":`, err.message);
            }
          } else {
            contador++;
          }
        });
      }
    });

    // 5. Finaliza a inserção
    stmt.finalize((err) => {
      if (err)
        return console.error("Erro ao finalizar statement:", err.message);

      console.log("--------------------------------------------------");
      console.log(`✅ Conversão de Clientes concluída!`);
      console.log(`${contador} clientes únicos foram inseridos.`);
      console.log(`O arquivo 'clientes.db' está pronto para ser usado.`);
      console.log("--------------------------------------------------");

      // 6. Fecha o banco de dados
      db.close();
    });
  });
}

// Executa a função
converterClientesParaSqlite();
