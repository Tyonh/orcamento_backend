const express = require("express");
const fs = require("fs/promises");
const fsSync = require("fs");
const puppeteer = require("puppeteer");
const path = require("path");
const cors = require("cors");
const sqlite3 = require("sqlite3");
const axios = require("axios");

const app = express();
const port = process.env.PORT || 3000;

// --- CONFIGURAÇÕES DE CAMINHOS ---
const PROD_DB_PATH = path.join(__dirname, "produtos.db");
const CLI_DB_PATH = path.join(__dirname, "clientes.db");
const TEMPLATE_PATH = path.join(__dirname, "templates", "template.html");
const STYLE_PATH = path.join(__dirname, "templates", "public", "style.css");
const LOGO_PATH = path.join(__dirname, "templates", "public", "logo.png");
// Caminho base para as imagens dos vendedores
const PUBLIC_DIR = path.join(__dirname, "templates", "public");

const TEMP_IMG_DIR = path.join(__dirname, "temp_img");

if (!fsSync.existsSync(TEMP_IMG_DIR)) {
  console.log(`Criando pasta temporária em: ${TEMP_IMG_DIR}`);
  fsSync.mkdirSync(TEMP_IMG_DIR);
}

app.use(cors());
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use("/public", express.static(path.join(__dirname, "public")));

// ... (Funções helpers de banco de dados permanecem iguais) ...
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

// ... (Rotas de busca e função downloadImagemTemporaria permanecem iguais) ...
app.get("/api/produtos/search", async (req, res) => {
  // ... (código original)
  const termo = (req.query.search || "").toLowerCase().trim();
  if (termo.length < 2) return res.json([]);
  try {
    const termoLike = `%${termo}%`;
    const sql =
      "SELECT codigo, descricao AS nome FROM produtos WHERE descricao COLLATE NOCASE LIKE ? OR codigo LIKE ? LIMIT 10";
    const rows = await queryDatabaseAll(PROD_DB_PATH, sql, [
      termoLike,
      termoLike,
    ]);
    return res.json(rows);
  } catch (err) {
    return res.status(500).json({ message: "Erro busca produtos" });
  }
});

app.get("/api/clientes/search", async (req, res) => {
  // ... (código original)
  const termo = (req.query.search || "").toLowerCase().trim();
  if (termo.length < 2) return res.json([]);
  try {
    const termoLike = `%${termo}%`;
    const sql =
      "SELECT id_cliente, nome FROM clientes WHERE nome COLLATE NOCASE LIKE ? LIMIT 10";
    const rows = await queryDatabaseAll(CLI_DB_PATH, sql, [termoLike]);
    return res.json(rows);
  } catch (err) {
    return res.status(500).json({ message: "Erro busca clientes" });
  }
});

async function downloadImagemTemporaria(url) {
  // ... (mesma função do seu código original) ...
  const urlLimpa = url.trim();
  const uniqueId = `${Date.now()}_${Math.random().toString(36).substring(7)}`;
  const tempFilePath = path.join(TEMP_IMG_DIR, `img_${uniqueId}.jpg`);

  try {
    const response = await axios({
      url: urlLimpa,
      method: "GET",
      responseType: "arraybuffer",
      headers: {
        "User-Agent":
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
      },
      timeout: 10000,
    });

    await fs.writeFile(tempFilePath, response.data);
    const fileContent = await fs.readFile(tempFilePath);
    const contentType = response.headers["content-type"] || "image/jpeg";
    const base64 = `data:${contentType};base64,${fileContent.toString(
      "base64"
    )}`;
    await fs.unlink(tempFilePath);
    return base64;
  } catch (error) {
    console.warn(`[AVISO] Erro ao baixar imagem ${urlLimpa}:`, error.message);
    try {
      if (fsSync.existsSync(tempFilePath)) await fs.unlink(tempFilePath);
    } catch (e) {}
    return null;
  }
}

app.post("/salvar-orcamento", async (req, res) => {
  console.log("Iniciando geração de orçamento...");

  try {
    const dados = req.body;
    let {
      cliente_nome,
      cliente_cnpj,
      cliente_email,
      id_cliente,
      vendedor,
      condicao_pagamento,
    } = dados;

    let vendedorInfo = { nome: "", email: "", fone: "", imagem: "" };
    // 1. Cliente (Lógica original)
    if (cliente_nome) {
      try {
        const row = await queryDatabaseGet(
          CLI_DB_PATH,
          "SELECT id_cliente, email FROM clientes WHERE nome = ? COLLATE NOCASE LIMIT 1",
          [cliente_nome]
        );
        if (row) {
          if (row.id_cliente) {
            id_cliente = row.id_cliente;
            if (!cliente_cnpj) cliente_cnpj = row.id_cliente;
          }
          if (row.email) cliente_email = row.email;
        }
      } catch (e) {
        console.error("Erro cliente:", e);
      }
    }

    try {
      // Tenta transformar o texto do select em Objeto
      const parsed = JSON.parse(vendedor);

      // Se deu certo, preenche as variáveis
      vendedorInfo.nome = parsed.nome || "";
      vendedorInfo.imagem = parsed.imagem || ""; // Pega o nome do arquivo aqui
    } catch (e) {
      vendedorInfo.nome = vendedor;
    }

    // 3. Processa Itens (Lógica original)
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
      .map((nome, i) => ({
        nomeCompleto: nome,
        qtd: parseInt(quantidades[i] || 1),
        valor: parseFloat(valores[i] || 0),
      }))
      .filter((p) => p.nomeCompleto && p.nomeCompleto.trim() !== "");

    // --- GERAÇÃO DOS ITENS ---
    let itensHtml = "";
    for (const item of itensValidos) {
      // ... (Lógica de itens e imagem do produto original mantida) ...
      let codigo = "";
      let nome = item.nomeCompleto;

      if (item.nomeCompleto.includes(" - ")) {
        const parts = item.nomeCompleto.split(" - ");
        codigo = parts[0].trim();
        nome = parts.slice(1).join(" - ").trim();
      }

      let imgTag = `<div style="width:50px; height:50px; background:#eee; display:flex; align-items:center; justify-content:center; font-size:9px; color:#999; margin:auto;">S/ FOTO</div>`;

      if (codigo) {
        const row = await queryDatabaseGet(
          PROD_DB_PATH,
          "SELECT imagem_url FROM produtos WHERE codigo = ? LIMIT 1",
          [codigo]
        );
        if (row && row.imagem_url && row.imagem_url.startsWith("http")) {
          const base64Image = await downloadImagemTemporaria(
            row.imagem_url.trim()
          );
          if (base64Image) {
            imgTag = `<img src="${base64Image}" alt="${codigo}" style="width:50px; height:50px; object-fit:contain; display:block; margin:auto;" />`;
          }
        }
      }

      itensHtml += `
    <tr>
        <td style="width:60px; text-align:center; padding:5px;">${imgTag}</td>
        <td style="vertical-align:middle; text-align:center;">${codigo}</td>
        <td style="vertical-align:middle;">${nome}</td>
        <td style="text-align:center; vertical-align:middle;">${item.qtd}</td>
        
        <td style="vertical-align:middle;">R$ ${item.valor.toLocaleString(
          "pt-BR",
          { minimumFractionDigits: 2, maximumFractionDigits: 2 }
        )}</td>
        
        <td style="vertical-align:middle;">R$ ${(
          item.qtd * item.valor
        ).toLocaleString("pt-BR", {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        })}</td>
    </tr>`;
    }

    // 4. Carrega Template e Recursos Estáticos
    let htmlTemplate = "";
    try {
      htmlTemplate = await fs.readFile(TEMPLATE_PATH, "utf8");
    } catch (e) {
      throw new Error(`Template não encontrado em ${TEMPLATE_PATH}`);
    }

    let cssContent = "";
    try {
      cssContent = await fs.readFile(STYLE_PATH, "utf8");
    } catch (e) {}

    let logoBase64 = "";
    try {
      const logoBuffer = await fs.readFile(LOGO_PATH);
      logoBase64 = `data:image/png;base64,${logoBuffer.toString("base64")}`;
    } catch (e) {}

    // --- NOVO: LÓGICA DA IMAGEM DO VENDEDOR ---
    let vendedorImgHtml = "";

    // MUDANÇA AQUI: Usamos vendedorInfo.imagem em vez de vendedor_imagem solto
    const nomeArquivoImagem = vendedorInfo.imagem;

    if (nomeArquivoImagem && nomeArquivoImagem.trim() !== "") {
      try {
        const safeFilename = path.basename(nomeArquivoImagem);
        const vendedorImgPath = path.join(PUBLIC_DIR, safeFilename);

        const ext = path.extname(safeFilename).replace(".", "") || "png";

        const vendedorBuffer = await fs.readFile(vendedorImgPath);
        const vendedorBase64 = `data:image/${ext};base64,${vendedorBuffer.toString(
          "base64"
        )}`;

        vendedorImgHtml = `
    <div style="margin-top: auto; width: 100%; text-align: center; page-break-inside: avoid; padding-bottom: 10px;">
        
        <img src="${vendedorBase64}" style="width: 100%; max-height: 150px; object-fit: contain;">
       
    </div>`;
      } catch (e) {
        console.warn(
          `Imagem do vendedor (${nomeArquivoImagem}) não encontrada.`
        );
      }
    }
    // ------------------------------------------

    const total = itensValidos.reduce((acc, it) => acc + it.qtd * it.valor, 0);

    // Substituições no HTML
    const finalHtml = htmlTemplate
      .replace("</head>", `<style>${cssContent}</style></head>`)
      .replace(/src=".*?logo\.png"/g, `src="${logoBase64}"`)
      .replace("{{cliente_nome}}", cliente_nome || "")
      .replace("{{cliente_cnpj}}", cliente_cnpj || "")
      .replace("{{cliente_email}}", cliente_email || "")
      .replace("{{vendedor_nome}}", vendedorInfo.nome || "")

      // AQUI ENTRA A NOVA SUBSTITUIÇÃO
      .replace("{{vendedor_info}}", vendedorImgHtml)

      .replace("{{condicao_pagamento}}", condicao_pagamento || "")
      .replace("{{data}}", new Date().toLocaleDateString("pt-BR"))
      .replace("{{itens}}", itensHtml)
      .replace(
        "{{total}}",
        `R$ ${total.toLocaleString("pt-BR", {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        })}`
      );

    // 5. Puppeteer
    const browser = await puppeteer.launch({
      headless: "new",
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    });
    const page = await browser.newPage();

    await page.setContent(finalHtml, { waitUntil: "domcontentloaded" });

    const pdfBuffer = await page.pdf({
      format: "A4",
      printBackground: true,
      margin: { top: "20px", bottom: "0px", left: "20px", right: "20px" },
    });

    await browser.close();

    const filename = `orcamento_${(cliente_nome || "cliente").replace(
      /[^a-z0-9]/gi,
      "_"
    )}.pdf`;
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(pdfBuffer);
    console.log("Orçamento gerado com sucesso!");
  } catch (error) {
    console.error("ERRO FATAL NO SERVIDOR:", error);
    res.status(500).json({
      message: "Erro interno ao gerar PDF",
      erro: error.message,
      stack: error.stack,
    });
  }
});

app.listen(port, () => {
  console.log(`Servidor rodando na porta ${port}`);
});
///teste
