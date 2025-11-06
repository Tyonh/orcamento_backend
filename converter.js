const exceljs = require("exceljs");
const sqlite3 = require("sqlite3").verbose();
const path = require("path");

// --- CONFIGURAÇÃO ---
// 1. O nome do seu arquivo Excel (baseado no que você enviou)
const EXCEL_PATH = path.join(__dirname, "Produto.xlsx");
// 2. Os nomes das colunas no seu Excel
const COLUNA_IMAGEM = "ImagemURL";
const COLUNA_CODIGO = "Codigo";
const COLUNA_DESCRICAO = "Descrição";
// 3. O nome do arquivo de banco de dados que será criado
const DB_PATH = path.join(__dirname, "produtos.db");
// --------------------

// Função assíncrona para executar o script
async function converterExcelParaSqlite() {
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

  // Pega a primeira planilha
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    console.error("Nenhuma planilha encontrada no arquivo Excel.");
    db.close();
    return;
  }

  // Encontra os índices das colunas
  let colImgIdx = -1;
  let colCodigoIdx = -1;
  let colDescIdx = -1;

  worksheet.getRow(1).eachCell((cell, colNumber) => {
    const valorCelula = cell.value ? cell.value.toString().trim() : "";
    if (valorCelula === COLUNA_IMAGEM) {
      colImgIdx = colNumber;
    }
    if (valorCelula === COLUNA_CODIGO) {
      colCodigoIdx = colNumber;
    }
    if (valorCelula === COLUNA_DESCRICAO) {
      colDescIdx = colNumber;
    }
  });

  if (colCodigoIdx === -1 || colDescIdx === -1) {
    console.error(
      `Erro: Não foi possível encontrar as colunas "${COLUNA_CODIGO}" e/ou "${COLUNA_DESCRICAO}" na primeira linha do Excel.`
    );
    db.close();
    return;
  }
  if (colImgIdx === -1) {
    // Se não encontrar a coluna de imagem, apenas avisa, mas continua
    console.warn(
      `Aviso: Coluna "${COLUNA_IMAGEM}" não encontrada. Esta coluna será ignorada.`
    );
  }

  console.log(`Coluna "${COLUNA_CODIGO}" encontrada.`);
  console.log(`Coluna "${COLUNA_DESCRICAO}" encontrada.`);
  console.log(
    `Coluna "${COLUNA_IMAGEM}" ${
      colImgIdx > -1 ? "encontrada." : "NÃO encontrada."
    }`
  );

  // Prepara o banco de dados
  db.serialize(() => {
    // 1. Apaga a tabela antiga, se existir
    db.run("DROP TABLE IF EXISTS produtos", (err) => {
      if (err) return console.error("Erro ao apagar tabela:", err.message);
      console.log("Tabela 'produtos' antiga (se existia) foi apagada.");
    });

    // 2. Cria a nova tabela com as TRÊS colunas
    db.run(
      `CREATE TABLE produtos (
         codigo TEXT PRIMARY KEY, 
         descricao TEXT NOT NULL COLLATE NOCASE,
         imagem_url TEXT
       )`,
      (err) => {
        if (err) return console.error("Erro ao criar tabela:", err.message);
        console.log(
          "Nova tabela 'produtos' (codigo, descricao, imagem_url) criada."
        );
      }
    );

    // 3. Prepara o comando de inserção para três colunas
    const sql =
      "INSERT INTO produtos (codigo, descricao, imagem_url) VALUES (?, ?, ?)";
    const stmt = db.prepare(sql, (err) => {
      if (err) return console.error("Erro ao preparar statement:", err.message);
    });

    // 4. Itera sobre as linhas do Excel
    console.log("Iniciando inserção dos produtos...");
    let contador = 0;
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return; // Pula a primeira linha (cabeçalho)

      const codigo = row.getCell(colCodigoIdx).value;
      const descricao = row.getCell(colDescIdx).value;

      // Pega a URL da imagem (ou define como null se a coluna não foi encontrada)
      let imagemUrl = null;
      if (colImgIdx > -1) {
        imagemUrl = row.getCell(colImgIdx).value;
      }

      if (codigo && descricao) {
        const codLimpo = codigo.toString().trim();
        const descLimpa = descricao.toString().trim();

        let urlLimpa = null;
        if (imagemUrl) {
          // O Excel pode guardar URLs como { text: 'url', hyperlink: 'url' }
          if (typeof imagemUrl === "object" && imagemUrl.text) {
            urlLimpa = imagemUrl.text;
          } else {
            urlLimpa = imagemUrl.toString().trim();
          }
        }

        // Insere os três valores no banco de dados
        stmt.run([codLimpo, descLimpa, urlLimpa], function (err) {
          if (err) {
            // Ignora erros de 'PRIMARY KEY constraint failed' (código duplicado)
            if (!err.message.includes("constraint failed")) {
              console.warn(`Erro ao inserir "${descLimpa}":`, err.message);
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
      console.log(`✅ Conversão concluída!`);
      console.log(`${contador} produtos únicos foram inseridos.`);
      console.log(`O arquivo 'produtos.db' está pronto para ser usado.`);
      console.log("--------------------------------------------------");

      // 6. Fecha o banco de dados
      db.close();
    });
  });
}

// Executa a função
converterExcelParaSqlite();
