const express = require("express");
const exceljs = require("exceljs");
const path = require("path");
const cors = require("cors");
const sqlite3 = require("sqlite3");

const app = express();
const port = process.env.PORT || 3000;

const PROD_DB_PATH = path.join(__dirname, "produtos.db");
const CLI_DB_PATH = path.join(__dirname, "clientes.db");

// --- NOVO LAYOUT DO TEMPLATE ---
const CLIENT_NAME_CELL = "C6";
const CLIENT_CNPJ_CELL = "C7";
const CLIENT_EMAIL_CELL = "F7";
const POSI_DATA = "F6";

const ITEM_START_ROW = 9; // Linha 9
const COL_IMG = "A"; // Coluna A (PRODUTO)
const COL_CODE = "B"; // Coluna B (CODIGO)
const COL_DESC = "C"; // Coluna C (DESCRIÇÃO)
const COL_QTD = "D"; // Coluna D (QTD)
const COL_VAL_UNIT = "E"; // Coluna E (VALOR UNID)
const COL_VAL_TOTAL = "F"; // Coluna F (VALOR TOTAL)
// Linhas abaixo dos itens
const TOTAL_LABEL_CELL_PREFIX = "E"; // Prefixo da célula "TOTAL:"
const TOTAL_VALUE_CELL_PREFIX = "F"; // Prefixo da célula do valor total
const VENDOR_NAME_CELL_PREFIX = "C"; // Célula para nome do Consultor
const VENDOR_PHONE_CELL_PREFIX = "C"; // Célula para Contato
const VENDOR_EMAIL_CELL_PREFIX = "C"; // Célula para Email (maiúscula)
// Coluna onde será escrita a condição de pagamento (ex: "A", "B", "C")
const COL_CONDI = "C"; // ajuste aqui se o template mudar

// --- HELPERS DO BANCO DE DADOS (MODIFICADOS) ---

// Retorna MÚLTIPLAS linhas (para autocomplete)
function queryDatabaseAll(dbPath, sql, params = []) {
  return new Promise((resolve, reject) => {
    const db = new sqlite3.Database(dbPath, sqlite3.OPEN_READONLY, (err) => {
      if (err) return reject(new Error(`DB_NOT_FOUND: ${dbPath}`));
    });
    db.all(sql, params, (err, rows) => {
      db.close();
      if (err) return reject(err);
      resolve(rows || []);
    });
  });
}

// Retorna UMA ÚNICA linha (para buscar 1 produto ou 1 cliente)
function queryDatabaseGet(dbPath, sql, params = []) {
  return new Promise((resolve, reject) => {
    const db = new sqlite3.Database(dbPath, sqlite3.OPEN_READONLY, (err) => {
      if (err) return reject(new Error(`DB_NOT_FOUND: ${dbPath}`));
    });
    db.get(sql, params, (err, row) => {
      db.close();
      if (err) return reject(err);
      resolve(row || null); // Retorna a linha ou nulo
    });
  });
}

app.use(cors());
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// --- ROTA DE PRODUTOS (CORRIGIDA) ---
// Retorna: [ { codigo: "123", nome: "Desc..." } ]
app.get("/api/produtos/search", async (req, res) => {
  const termoBusca = (req.query.search || "").toLowerCase().trim();
  if (termoBusca.length < 2) return res.json([]);

  try {
    const termoLike = `%${termoBusca}%`;
    const sql =
      "SELECT codigo, descricao AS nome FROM produtos WHERE descricao COLLATE NOCASE LIKE ? OR codigo LIKE ? LIMIT 10";
    const rows = await queryDatabaseAll(PROD_DB_PATH, sql, [
      termoLike,
      termoLike,
    ]);
    return res.json(rows);
  } catch (err) {
    console.error("Erro na consulta de produtos:", err);
    return res.status(500).json({ message: "Erro ao buscar produtos." });
  }
});

// --- ROTA DE CLIENTES (ATUALIZADA) ---
// Retorna: [ { id_cliente: "cnpj/cpf", nome: "Razao Social..." } ]
app.get("/api/clientes/search", async (req, res) => {
  const termoBusca = (req.query.search || "").toLowerCase().trim();
  if (termoBusca.length < 2) return res.json([]);

  try {
    const termoLike = `%${termoBusca}%`;
    const sql =
      "SELECT id_cliente, nome FROM clientes WHERE nome COLLATE NOCASE LIKE ? LIMIT 10";
    const rows = await queryDatabaseAll(CLI_DB_PATH, sql, [termoLike]);
    return res.json(rows);
  } catch (err) {
    console.error("Erro na consulta de clientes:", err);
    if (/DB_NOT_FOUND/i.test(err.message)) {
      console.warn(`Aviso: O arquivo ${CLI_DB_PATH} não foi encontrado.`);
      return res.json([]);
    }
    return res.status(500).json({ message: "Erro ao buscar clientes." });
  }
});

console.log("Backend iniciado. Aguardando envios de formulário...");

const mapearCondicao = {
  avista: "À vista",
  "30d": "Cartão de Crédito",
  "30d": "Boleto - 30d",
  "30/60d": "Boleto - 30/60d",
  "30/60/90d": "Boleto - 30/60/90d",
  "30/60/90/120d": "Boleto - 30/60/90/120d",
  "Cartao de Debito": "Cartao de Debito",
  "Cartao de Credito S/Juros": "Cartao de Credito S/Juros",
  "Cartao 1x/Juros": "Cartao de Credito 1x/Juros",
  "Cartao 2x/Juros": "Cartao de Credito 2x/Juros",
  "Cartao 3x/Juros": "Cartao de Credito 3x/Juros",
  "Boleto - c/Entrada": "Boleto - c/Entrada",
  "Cartao de Debito c/Entrada": "Cartao de Debito c/Entrada",
};

// --- ROTA DE SALVAR ORÇAMENTO (TOTALMENTE REESCRITA) ---
app.post("/salvar-orcamento", async (req, res) => {
  console.log("\n=== NOVA REQUISIÇÃO DE ORÇAMENTO ===");
  const dadosDoFormulario = req.body;
  console.log("\n--- DADOS RECEBIDOS DO FORMULÁRIO ---");
  console.log(JSON.stringify(dadosDoFormulario, null, 2)); // 1. Captura todos os dados do formulário (Sem alteração)

  // Extrai campos do formulário; usaremos let para poder sobrescrever com dados do banco
  let {
    cliente_nome,
    cliente_cnpj,
    cliente_email,
    id_cliente,
    vendedor,
    condicao_pagamento,
  } = dadosDoFormulario;

  // Busca id_cliente e email do banco de dados clientes pelo nome, se possível
  if (cliente_nome) {
    try {
      const sqlCliente =
        "SELECT id_cliente, email FROM clientes WHERE nome = ? COLLATE NOCASE LIMIT 1";
      const clienteDB = await queryDatabaseGet(CLI_DB_PATH, sqlCliente, [
        cliente_nome,
      ]);
      if (clienteDB) {
        if (clienteDB.id_cliente) {
          id_cliente = clienteDB.id_cliente;
          cliente_cnpj = clienteDB.id_cliente;
        }
        if (clienteDB.email) {
          cliente_email = clienteDB.email;
        }
      }
    } catch (e) {
      console.warn(
        "Não foi possível buscar dados do cliente pelo nome:",
        e.message
      );
    }
  }

  let vendedorInfo;
  try {
    vendedorInfo = JSON.parse(vendedor);
  } catch (e) {
    vendedorInfo = { nome: vendedor, email: "", fone: "" };
  }

  // 3. Captura dos Itens (===== CORREÇÃO AQUI =====)
  // Removemos os '[]' dos nomes das chaves para bater com o log
  const produtos = Array.isArray(dadosDoFormulario["produto_nome"])
    ? dadosDoFormulario["produto_nome"]
    : [];
  const valores = Array.isArray(dadosDoFormulario["valor_unitario"])
    ? dadosDoFormulario["valor_unitario"]
    : [];
  const quantidades = Array.isArray(dadosDoFormulario["quantidade"])
    ? dadosDoFormulario["quantidade"]
    : [];

  const itensValidos = produtos
    .map((nome, index) => ({
      nomeCompleto: nome,
      qtd: parseInt(quantidades[index] || 1),
      valor: parseFloat(valores[index] || 0),
    }))
    .filter((p) => p.nomeCompleto && p.nomeCompleto.trim() !== "");

  const numItens = itensValidos.length;
  console.log(`\n- Total de itens válidos: ${numItens}`);

  try {
    const workbook = new exceljs.Workbook();
    const templatePath = path.join(
      process.cwd(),
      "templates",
      "TEMPLATE-ORCAMENTEO.xlsx"
    ); // <-- Novo nome do template
    await workbook.xlsx.readFile(templatePath);

    const worksheet = workbook.getWorksheet("Sheet1");
    if (!worksheet) throw new Error("Planilha 'Sheet1' não encontrada.");

    worksheet.getCell(POSI_DATA).value = new Date();
    // 4. Preenche Dados do Cliente (Novas Células)
    worksheet.getCell(CLIENT_NAME_CELL).value = cliente_nome || ""; // C6
    // Prioriza id_cliente (quando o autocomplete retorna id_cliente), senão usa cliente_cnpj
    worksheet.getCell(CLIENT_CNPJ_CELL).value =
      id_cliente || cliente_cnpj || ""; // C7
    worksheet.getCell(CLIENT_EMAIL_CELL).value = cliente_email || ""; // F7
    // 5. Define a linha inicial dos itens
    const linhaInicialItens = ITEM_START_ROW; // Linha 9

    if (numItens === 0) {
      // (Lógica de 0 itens)
      worksheet.getCell(`${COL_DESC}${linhaInicialItens}`).value =
        "Nenhum item";
      const linhaTotal = linhaInicialItens + 1; // Linha 10
      worksheet.getCell(`${TOTAL_VALUE_CELL_PREFIX}${linhaTotal}`).value = 0; // F10

      // Preenche Vendedor
      const linhaConsultor = linhaTotal + 4; // Linha 14 (baseado no template)
      try {
        const addrName = `${VENDOR_NAME_CELL_PREFIX}${linhaConsultor}`;
        const addrPhone = `${VENDOR_PHONE_CELL_PREFIX}${linhaConsultor + 1}`;
        const addrEmail = `${VENDOR_EMAIL_CELL_PREFIX}${linhaConsultor + 2}`;
        console.log("Escrevendo vendedor em:", addrName, addrPhone, addrEmail);
        worksheet.getCell(addrName).value = vendedorInfo.nome;
        worksheet.getCell(addrPhone).value = vendedorInfo.fone;
        worksheet.getCell(addrEmail).value = vendedorInfo.email;
      } catch (vendorErr) {
        console.error("Erro ao escrever dados do vendedor. Endereços:", {
          name: `${VENDOR_NAME_CELL_PREFIX}${linhaConsultor}`,
          phone: `${VENDOR_PHONE_CELL_PREFIX}${linhaConsultor + 1}`,
          email: `${VENDOR_EMAIL_CELL_PREFIX}${linhaConsultor + 2}`,
        });
        throw vendorErr;
      }
    } else {
      // --- LÓGICA DE 1+ ITENS (COM IMAGENS) ---

      // 6. Insere novas linhas se houver MAIS de 1 item
      if (numItens > 1) {
        worksheet.insertRows(
          linhaInicialItens + 1,
          Array.from({ length: numItens - 1 }, () => []),
          "i+" // 'i+' copia a altura da linha de cima
        );
      }

      // 7. Preenche os dados dos itens (USANDO FOR...OF PARA AWAIT)
      let index = 0;
      for (const item of itensValidos) {
        const linhaNum = linhaInicialItens + index;
        const linha = worksheet.getRow(linhaNum);
        linha.height = 80; // Define a altura da linha para a imagem

        // Extrai o código do nome "CODIGO - NOME"
        let codigoProduto = "";
        let nomeProduto = item.nomeCompleto;
        if (item.nomeCompleto && item.nomeCompleto.includes(" - ")) {
          const partes = item.nomeCompleto.split(" - ");
          codigoProduto = partes[0].trim();
          nomeProduto = partes.slice(1).join(" - ").trim();
        }

        // Preenche Células de Texto/Número
        worksheet.getCell(`${COL_CODE}${linhaNum}`).value = codigoProduto; // B9
        worksheet.getCell(`${COL_DESC}${linhaNum}`).value = nomeProduto; // C9
        worksheet.getCell(`${COL_QTD}${linhaNum}`).value = item.qtd; // D9
        worksheet.getCell(`${COL_VAL_UNIT}${linhaNum}`).value = item.valor; // E9
        worksheet.getCell(`${COL_VAL_TOTAL}${linhaNum}`).value = {
          // F9
          formula: `${COL_QTD}${linhaNum}*${COL_VAL_UNIT}${linhaNum}`,
        };

        // --- LÓGICA DE IMAGEM ---
        if (codigoProduto) {
          try {
            const sql = "SELECT imagem_url FROM produtos WHERE codigo = ?";
            const produtoDB = await queryDatabaseGet(PROD_DB_PATH, sql, [
              codigoProduto,
            ]);

            if (produtoDB && produtoDB.imagem_url) {
              console.log(
                `Baixando imagem para ${codigoProduto}: ${produtoDB.imagem_url}`
              );
              const response = await fetch(produtoDB.imagem_url);
              if (!response.ok)
                throw new Error(
                  `Falha ao baixar imagem: ${response.statusText}`
                );

              const buffer = await response.arrayBuffer();
              const extensao =
                produtoDB.imagem_url.split(".").pop().toLowerCase() || "jpeg";

              const imageId = workbook.addImage({
                buffer: Buffer.from(buffer),
                extension: extensao === "jpg" ? "jpeg" : extensao,
              });

              // Adiciona a imagem à planilha (Coluna A)
              worksheet.addImage(imageId, {
                // --- MUDANÇA AQUI ---
                // Usamos as coordenadas exatas da célula (0-indexed)
                tl: { col: 0, row: linhaNum - 1 }, // Canto Superior Esquerdo
                br: { col: 1, row: linhaNum }, // Canto Inferior Direito // --- FIM DA MUDANÇA ---
                editAs: "oneCell",
              });
            }
          } catch (imgError) {
            console.error(
              `Erro ao processar imagem para ${codigoProduto}:`,
              imgError.message
            );
            worksheet.getCell(`${COL_IMG}${linhaNum}`).value = "Erro Img";
          }
        }
        // --- FIM DA LÓGICA DE IMAGEM ---

        index++;
      } // Fim do loop for...of

      // 8. Preenche os campos "móveis" (Total, Vendedor, Condições)
      const ultimaLinhaItem = linhaInicialItens + numItens - 1;
      const linhaTotal = ultimaLinhaItem + 1; // Linha 10 (ou 11, 12...)

      // Total (Ex: F10)
      worksheet.getCell(`${TOTAL_VALUE_CELL_PREFIX}${linhaTotal}`).value = {
        formula: `SUM(${COL_VAL_TOTAL}${linhaInicialItens}:${COL_VAL_TOTAL}${ultimaLinhaItem})`,
      };
      // Label "TOTAL:" (Ex: E10)
      worksheet.getCell(`${TOTAL_LABEL_CELL_PREFIX}${linhaTotal}`).value =
        "TOTAL:";

      // Posição dos campos de rodapé (Vendedor, Condições)
      const linhaObs = linhaTotal + 2; // Linha 12
      const linhaConsultor = linhaObs + 3; // Linha 15

      // Condições (não achei no template, mas vou colocar abaixo das obs)
      // Usa a constante COL_CONDI para posicionamento da coluna (flexível)
      worksheet.getCell(`${COL_CONDI}${linhaObs - 2}`).value =
        mapearCondicao[condicao_pagamento] || condicao_pagamento;

      // Vendedor (Ex: B15, B16, B17)
      worksheet.getCell(`${VENDOR_NAME_CELL_PREFIX}${linhaConsultor}`).value =
        vendedorInfo.nome;
      worksheet.getCell(
        `${VENDOR_PHONE_CELL_PREFIX}${linhaConsultor + 1}`
      ).value = vendedorInfo.fone;
      worksheet.getCell(
        `${VENDOR_EMAIL_CELL_PREFIX}${linhaConsultor + 2}`
      ).value = vendedorInfo.email;
    }

    // 9. Lógica de Download (intocada)
    const nomeClienteFormatado = (cliente_nome || "orcamento")
      .replace(/[^a-z0-9]/gi, "_")
      .toLowerCase();
    const nomeArquivoFinal = `orcamento_${nomeClienteFormatado}_${Date.now()}.xlsx`;

    console.log(
      `Orçamento gerado em memória: ${nomeArquivoFinal}. Enviando para download...`
    );

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${nomeArquivoFinal}"`
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error("Erro ao gerar o orçamento:", error);
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

// Exporta o 'app' para o Vercel (ou para rodar com 'node-main')
module.exports = app;
