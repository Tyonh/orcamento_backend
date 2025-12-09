const express = require("express");
const fs = require("fs/promises");
const puppeteer = require("puppeteer");
const path = require("path");
const cors = require("cors");
const sqlite3 = require("sqlite3");

const app = express();
const port = process.env.PORT || 3000;

const PROD_DB_PATH = path.join(__dirname, "produtos.db");
const CLI_DB_PATH = path.join(__dirname, "clientes.db");
app.use("/public", express.static(path.join(__dirname, "templates", "public")));

app.use(cors());
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Funções auxiliares para banco de dados
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

function queryDatabaseGet(dbPath, sql, params = []) {
  return new Promise((resolve, reject) => {
    const db = new sqlite3.Database(dbPath, sqlite3.OPEN_READONLY, (err) => {
      if (err) return reject(new Error(`DB_NOT_FOUND: ${dbPath}`));
    });
    db.get(sql, params, (err, row) => {
      db.close();
      if (err) return reject(err);
      resolve(row || null);
    });
  });
}

// Rota para buscar produtos
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
    return res.status(500).json({ message: "Erro ao buscar produtos." });
  }
});

// Rota para buscar clientes
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
    return res.status(500).json({ message: "Erro ao buscar clientes." });
  }
});

// Rota principal para gerar orçamento em PDF
app.post("/salvar-orcamento", async (req, res) => {
  const dados = req.body;
  let {
    cliente_nome,
    cliente_cnpj,
    cliente_email,
    id_cliente,
    vendedor,
    condicao_pagamento,
  } = dados;

  // Busca dados do cliente no banco
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
    } catch (e) {}
  }

  let vendedorInfo;
  try {
    vendedorInfo = JSON.parse(vendedor);
  } catch (e) {
    vendedorInfo = { nome: vendedor, email: "", fone: "" };
  }

  // Captura dos itens
  const produtos = Array.isArray(dados["produto_nome"])
    ? dados["produto_nome"]
    : [];
  const valores = Array.isArray(dados["valor_unitario"])
    ? dados["valor_unitario"]
    : [];
  const quantidades = Array.isArray(dados["quantidade"])
    ? dados["quantidade"]
    : [];

  const itensValidos = produtos
    .map(function (nome, index) {
      return {
        nomeCompleto: nome,
        qtd: parseInt(quantidades[index] || 1),
        valor: parseFloat(valores[index] || 0),
      };
    })
    .filter(function (p) {
      return p.nomeCompleto && p.nomeCompleto.trim() !== "";
    });

  // Monta HTML dos itens
  let itensHtml = "";
  for (const item of itensValidos) {
    let codigoProduto = "";
    let nomeProduto = item.nomeCompleto;
    if (item.nomeCompleto && item.nomeCompleto.includes(" - ")) {
      const partes = item.nomeCompleto.split(" - ");
      codigoProduto = partes[0].trim();
      nomeProduto = partes.slice(1).join(" - ").trim();
    }
    itensHtml += `
			<tr>
                <td>img</td>
				<td>${codigoProduto}</td>
				<td>${nomeProduto}</td>
				<td>${item.qtd}</td>
				<td>R$ ${item.valor.toFixed(2)}</td>
				<td>R$ ${(item.qtd * item.valor).toFixed(2)}</td>
			</tr>
		`;
  }
  if (itensValidos.length === 0) {
    itensHtml = `<tr><td colspan="5">Nenhum item</td></tr>`;
  }

  // Carrega template HTML
  const templatePath = path.join(__dirname, "templates", "template.html");
  let html;
  try {
    html = await fs.readFile(templatePath, "utf8");
  } catch (err) {
    return res.status(500).json({ message: "Template não encontrado." });
  }

  // Preenche variáveis do template
  const total = itensValidos.reduce(
    (acc, item) => acc + item.qtd * item.valor,
    0
  );
  html = html
    .replace("{{cliente_nome}}", cliente_nome || "")
    .replace("{{cliente_cnpj}}", id_cliente || cliente_cnpj || "")
    .replace("{{cliente_email}}", cliente_email || "")
    .replace("{{vendedor_nome}}", vendedorInfo.nome || "")
    .replace("{{vendedor_email}}", vendedorInfo.email || "")
    .replace("{{vendedor_fone}}", vendedorInfo.fone || "")
    .replace("{{condicao_pagamento}}", condicao_pagamento || "")
    .replace("{{data}}", new Date().toLocaleDateString("pt-BR"))
    .replace("{{itens}}", itensHtml)
    .replace("{{total}}", `R$ ${total.toFixed(2)}`);

  // Gera PDF
  try {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });
    const pdfBuffer = await page.pdf({ format: "A4" });
    await browser.close();

    const nomeClienteFormatado = (cliente_nome || "orcamento")
      .replace(/[^a-z0-9]/gi, "_")
      .toLowerCase();
    const nomeArquivoFinal = `orcamento_${nomeClienteFormatado}_${Date.now()}.pdf`;

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${nomeArquivoFinal}"`
    );
    res.send(pdfBuffer);
  } catch (error) {
    res
      .status(500)
      .json({ message: "Erro ao gerar PDF.", details: error.message });
  }
});

app.listen(port, () => {
  console.log(`Backend rodando em http://localhost:${port}`);
});

module.exports = app;
