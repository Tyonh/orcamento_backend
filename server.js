const express = require("express");
const exceljs = require("exceljs");
const path = require("path");
const fs = require("fs");
const cors = require("cors");
const sqlite3 = require("sqlite3");

const app = express();
const port = process.env.PORT || 3000;

const PROD_DB_PATH = path.join(__dirname, "produtos.db");
const CLI_DB_PATH = path.join(__dirname, "clientes.db");
// Layout configurável do template (ajuste se seu template mudar)
const CLIENT_CELL = "A6"; // célula onde o nome do cliente será escrito
const ITEM_START_ROW = 9; // primeira linha disponível para itens
const COL_CODE = "A"; // coluna do código do produto
const COL_DESC = "B"; // coluna da descrição
const COL_QTD = "C"; // coluna quantidade
const COL_VAL_UNIT = "D"; // coluna valor unitário
const COL_VAL_TOTAL = "E"; // coluna valor total por linha

// Helper: consulta genérica no SQLite retornando Promise
function queryDatabase(dbPath, sql, params = []) {
  return new Promise((resolve, reject) => {
    // Verifica se o arquivo do DB existe antes de tentar abrir
    if (!fs.existsSync(dbPath)) {
      return reject(new Error(`Database file not found: ${dbPath}`));
    }

    const db = new sqlite3.Database(dbPath, sqlite3.OPEN_READONLY, (err) => {
      if (err) return reject(err);
    });

    db.all(sql, params, (err, rows) => {
      if (err) {
        db.close();
        return reject(err);
      }
      resolve(rows || []);
      db.close();
    });
  });
}

app.use(cors());
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Rota de busca de produtos (intocada)
app.get("/api/produtos/search", async (req, res) => {
  const termoBusca = (req.query.search || "").toLowerCase().trim();

  if (termoBusca.length < 2) {
    return res.json([]);
  }

  try {
    const termoLike = `%${termoBusca}%`;
    const sql =
      "SELECT codigo, descricao FROM produtos WHERE nome COLLATE NOCASE LIKE ? LIMIT 10";
    const rows = await queryDatabase(PROD_DB_PATH, sql, [termoLike]);
    return res.json(rows);
  } catch (err) {
    console.error("Erro na consulta de produtos:", err);
    return res.status(500).json({ message: "Erro ao buscar produtos." });
  }
});

// Rota de busca de clientes (autocomplete)
// Observação: assume-se que exista uma tabela `clientes` com coluna `nome`.
app.get("/api/clientes/search", async (req, res) => {
  const termoBusca = (req.query.search || "").toLowerCase().trim();

  if (termoBusca.length < 2) {
    return res.json([]);
  }

  try {
    const termoLike = `%${termoBusca}%`;
    const sql =
      "SELECT nome FROM clientes WHERE nome COLLATE NOCASE LIKE ? LIMIT 10";

    // Se o DB de clientes não existir, tentamos usar o produtos.db como fallback
    let dbToUse = CLI_DB_PATH;
    if (!fs.existsSync(dbToUse)) {
      console.warn(
        `Arquivo ${dbToUse} não encontrado. Tentando usar ${PROD_DB_PATH} como fallback.`
      );
      if (fs.existsSync(PROD_DB_PATH)) {
        dbToUse = PROD_DB_PATH;
      } else {
        console.warn(
          `Nenhum arquivo de banco de dados disponível (${CLI_DB_PATH}, ${PROD_DB_PATH}). Retornando [] para o autocomplete de clientes.`
        );
        return res.json([]);
      }
    }

    const rows = await queryDatabase(dbToUse, sql, [termoLike]);
    return res.json(rows);
  } catch (err) {
    console.error("Erro na consulta de clientes:", err);
    // Se o erro for que o arquivo não existe, devolvemos array vazio para evitar quebrar o autocomplete
    if (
      err &&
      /not found|ENOENT|unable to open database file/i.test(err.message)
    ) {
      return res.json([]);
    }
    return res.status(500).json({ message: "Erro ao buscar clientes." });
  }
});

console.log("Backend iniciado. Aguardando envios de formulário...");

const mapearCondicao = {
  avista: "À vista",
  "30d": "Cartão de Crédito",
  "60d": "Boleto",
  "30/60d": "Boleto - 30/60d",
  "30/60/90d": "Boleto - 30/60/90d",
  "30/60/90/120d": "Boleto - 30/60/90/120d",
  "Cartao de Debito": "Boleto - Cartao de Debito",
  "Cartao de Credito S/Juros": "Cartao de Credito S/Juros",
  "Cartao 1x/Juros": "Cartao de Credito 1x/Juros",
  "Cartao 2x/Juros": "Cartao de Credito 2x/Juros",
  "Cartao 3x/Juros": "Cartao de Credito 3x/Juros",
  "Boleto - c/Entrada": "Boleto - c/Entrada",
  "Cartao de Debito c/Entrada": "Cartao de Debito c/Entrada",
};

app.post("/salvar-orcamento", async (req, res) => {
  console.log("\n=== NOVA REQUISIÇÃO DE ORÇAMENTO ===");
  console.log("Headers recebidos:", req.headers);

  const dadosDoFormulario = req.body;
  console.log("\n--- DADOS RECEBIDOS DO FORMULÁRIO ---");
  console.log("Dados brutos:", JSON.stringify(dadosDoFormulario, null, 2));
  console.log("\nCampos encontrados:");
  Object.keys(dadosDoFormulario).forEach((key) => {
    console.log(`${key}:`, dadosDoFormulario[key]);
  });

  const { cliente_nome, condicao_pagamento } = dadosDoFormulario;
  console.log("\nDados principais:");
  console.log("- Cliente:", cliente_nome);
  console.log("- Condição:", condicao_pagamento);

  // Proteção contra 'undefined' se o JS do frontend falhar
  const produtos = Array.isArray(dadosDoFormulario["produto_nome"])
    ? dadosDoFormulario["produto_nome"]
    : [];

  const valores = Array.isArray(dadosDoFormulario["valor_unitario"])
    ? dadosDoFormulario["valor_unitario"]
    : [];

  const quantidades = Array.isArray(dadosDoFormulario["quantidade"])
    ? dadosDoFormulario["quantidade"]
    : [];

  console.log("\nDados dos produtos:");
  console.log("- Nomes:", produtos);
  console.log("- Quantidades:", quantidades);
  console.log("- Valores:", valores);

  // Filtra itens vazios
  const itensValidos = produtos
    .map((nome, index) => ({
      nome: nome,
      qtd: parseInt(quantidades[index] || 1),
      valor: parseFloat(valores[index] || 0),
    }))
    .filter((p) => p.nome && p.nome.trim() !== "");

  const numItens = itensValidos.length;
  console.log("\nItens após processamento:");
  console.log(`- Total de itens válidos: ${numItens}`);
  console.log("- Itens detalhados:");
  itensValidos.forEach((item, idx) => {
    console.log(
      `  ${idx + 1}. ${item.nome} (${item.qtd} x R$ ${item.valor.toFixed(2)})`
    );
  });

  try {
    const workbook = new exceljs.Workbook();
    // Para o Vercel, é melhor usar process.cwd() para garantir o caminho
    const templatePath = path.join(
      process.cwd(), // <-- Mudança para Vercel
      "templates",
      "template_orcamento.xlsx"
    );
    await workbook.xlsx.readFile(templatePath);

    const worksheet = workbook.getWorksheet("Sheet1");
    if (!worksheet) {
      throw new Error(
        "Não foi possível encontrar a planilha 'Sheet1'. Verifique o nome da aba no seu template."
      );
    }

    // 1. Preenche Cliente
    worksheet.getCell(CLIENT_CELL).value = cliente_nome;

    // 2. Define a linha inicial dos itens
    const linhaInicialItens = ITEM_START_ROW;

    // --- LÓGICA ATUALIZADA E MAIS SEGURA ---

    if (numItens === 0) {
      // Se 0 itens, escreve uma linha indicando que não há itens
      console.log("Nenhum item válido foi adicionado.");
      // Código vazio, descrição "Nenhum item", quant e valores 0
      worksheet.getCell(`${COL_CODE}${linhaInicialItens}`).value = "";
      worksheet.getCell(`${COL_DESC}${linhaInicialItens}`).value =
        "Nenhum item";
      worksheet.getCell(`${COL_QTD}${linhaInicialItens}`).value = 0;
      worksheet.getCell(`${COL_VAL_UNIT}${linhaInicialItens}`).value = 0;
      worksheet.getCell(`${COL_VAL_TOTAL}${linhaInicialItens}`).value = 0;

      // Total ficará na linha imediatamente após a linha de itens
      const linhaTotal = linhaInicialItens + 1;
      worksheet.getCell(`${COL_VAL_TOTAL}${linhaTotal}`).value = {
        formula: `0`,
      };
      // Escreve a condição de pagamento uma linha abaixo do total
      worksheet.getCell(`A${linhaTotal + 1}`).value =
        mapearCondicao[condicao_pagamento] || condicao_pagamento;
    } else {
      // Se 1 ou MAIS itens, executa a lógica de loop

      // 3. Insere novas linhas se houver MAIS de 1 item
      if (numItens > 1) {
        const linhasParaAdicionar = numItens - 1;

        const arrayDeLinhasVazias = Array.from(
          { length: linhasParaAdicionar },
          () => []
        );

        if (numItens > 1) {
          const linhasParaAdicionar = numItens - 1;

          const arrayDeLinhasVazias = Array.from(
            { length: linhasParaAdicionar },
            () => []
          );

          worksheet.insertRows(
            linhaInicialItens + 1,
            arrayDeLinhasVazias,
            "i+"
          );
        }
      }

      // 4. Preenche os dados dos itens (buscando também o código no DB de produtos)
      for (let index = 0; index < itensValidos.length; index++) {
        const item = itensValidos[index];
        const linhaNum = linhaInicialItens + index;

        // Busca o código do produto no DB (caso exista)
        let codigoProduto = "";
        try {
          const rows = await queryDatabase(
            PROD_DB_PATH,
            "SELECT codigo FROM produtos WHERE nome = ? LIMIT 1",
            [item.nome]
          );
          if (rows && rows.length > 0 && rows[0].codigo) {
            codigoProduto = rows[0].codigo;
          }
        } catch (err) {
          // se houver erro na busca do código, apenas logamos e seguimos
          console.warn(
            `Não foi possível obter código para produto '${item.nome}':`,
            err.message
          );
        }

        worksheet.getCell(`${COL_CODE}${linhaNum}`).value = codigoProduto;
        worksheet.getCell(`${COL_DESC}${linhaNum}`).value = item.nome;
        worksheet.getCell(`${COL_QTD}${linhaNum}`).value = item.qtd;
        worksheet.getCell(`${COL_VAL_UNIT}${linhaNum}`).value = item.valor;
        worksheet.getCell(`${COL_VAL_TOTAL}${linhaNum}`).value = {
          formula: `${COL_QTD}${linhaNum}*${COL_VAL_UNIT}${linhaNum}`,
        };
      }

      // 5. Preenche os campos "móveis": total e condição de pagamento
      const ultimaLinhaItem = linhaInicialItens + numItens - 1;
      const linhaTotal = ultimaLinhaItem + 1; // linha imediatamente abaixo dos itens

      worksheet.getCell(`${COL_VAL_TOTAL}${linhaTotal}`).value = {
        formula: `SUM(${COL_VAL_TOTAL}${linhaInicialItens}:${COL_VAL_TOTAL}${ultimaLinhaItem})`,
      };

      // Escreve a label TOTAL na coluna de valor unitário (coluna anterior à total)
      worksheet.getCell(`${COL_VAL_UNIT}${linhaTotal}`).value = "TOTAL:";

      // Condição de pagamento fixada uma linha abaixo do total
      worksheet.getCell(`A${linhaTotal + 1}`).value =
        mapearCondicao[condicao_pagamento] || condicao_pagamento;
    }

    // --- INÍCIO DA MODIFICAÇÃO PARA DOWNLOAD DIRETO ---

    // 6. Gerar o nome do arquivo (em memória)
    const nomeClienteFormatado = (cliente_nome || "orcamento")
      .replace(/[^a-z0-9]/gi, "_")
      .toLowerCase();
    const nomeArquivoFinal = `orcamento_${nomeClienteFormatado}_${Date.now()}.xlsx`;

    console.log(
      `Orçamento gerado: ${nomeArquivoFinal}. Salvando e enviando para download...`
    );

    // Garante que a pasta orcamentos_gerados existe
    const pastaSalvar = path.join(__dirname, "orcamentos_gerados");
    const caminhoArquivo = path.join(pastaSalvar, nomeArquivoFinal);
    console.log(`Tentando criar/verificar pasta em: ${pastaSalvar}`);

    try {
      if (!fs.existsSync(pastaSalvar)) {
        console.log(`Pasta não existe, criando...`);
        fs.mkdirSync(pastaSalvar, { recursive: true });
        console.log(`Pasta criada com sucesso`);
      } else {
        console.log(`Pasta já existe`);
      }

      // Salva o arquivo localmente
      console.log(`Tentando salvar arquivo em: ${caminhoArquivo}`);

      await workbook.xlsx.writeFile(caminhoArquivo);
      console.log(`Arquivo salvo com sucesso em: ${caminhoArquivo}`);

      // Verifica se o arquivo foi realmente criado
      if (fs.existsSync(caminhoArquivo)) {
        console.log(`Confirmado: arquivo existe no disco`);
        const stats = fs.statSync(caminhoArquivo);
        console.log(`Tamanho do arquivo: ${stats.size} bytes`);
      } else {
        throw new Error("Arquivo não encontrado após tentativa de escrita");
      }
    } catch (err) {
      console.error("Erro ao salvar arquivo:", err);
      throw err; // Re-lança o erro para ser pego pelo catch externo
    }

    // 7. Definir os cabeçalhos da resposta para o navegador
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${nomeArquivoFinal}"`
    );

    // 8. Envia o arquivo salvo para download
    // Use o caminho absoluto diretamente (não passe 'root' para evitar junção indevida em Windows)
    res.sendFile(caminhoArquivo, (err) => {
      if (err) {
        console.error("Erro ao enviar arquivo:", err);
        if (!res.headersSent) {
          res.status(500).json({
            message: "Erro ao enviar o arquivo para download.",
            details: err.message,
          });
        }
      } else {
        console.log(
          `Arquivo enviado com sucesso para download: ${caminhoArquivo}`
        );
      }
    });

    // --- FIM DA MODIFICAÇÃO ---
  } catch (error) {
    console.error("Erro ao gerar o orçamento:", error);
    // Adicionado para evitar erro caso o stream de download já tenha começado
    if (!res.headersSent) {
      res.status(500).json({
        message: "Erro ao processar o orçamento.",
        details: error.message,
      });
    }
  }
});

app.listen(port, () => {
  console.log(`Backend rodando em http://localhost:${port}`);
});

// Adicionado para Vercel
module.exports = app;
