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
      "SELECT nome FROM produtos WHERE nome COLLATE NOCASE LIKE ? LIMIT 10";
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
    worksheet.getCell("A6").value = cliente_nome;

    // 2. Define a linha inicial dos itens
    const linhaInicialItens = 8;

    // --- LÓGICA ATUALIZADA E MAIS SEGURA ---

    if (numItens === 0) {
      // Se 0 itens, limpa a linha 8 do template e preenche 9 e 10
      console.log("Nenhum item válido foi adicionado.");
      worksheet.getRow(linhaInicialItens).getCell("A").value = "Nenhum item";
      worksheet.getRow(linhaInicialItens).getCell("B").value = 0;
      worksheet.getRow(linhaInicialItens).getCell("C").value = 0;
      worksheet.getRow(linhaInicialItens).getCell("D").value = 0;

      worksheet.getCell("D9").value = 0; // Total
      worksheet.getCell("A10").value =
        mapearCondicao[condicao_pagamento] || condicao_pagamento; // Condições
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

      // 4. Preenche os dados dos itens
      itensValidos.forEach((item, index) => {
        const linhaNum = linhaInicialItens + index;
        const linha = worksheet.getRow(linhaNum);

        linha.getCell("A").value = item.nome;
        linha.getCell("B").value = item.qtd;
        linha.getCell("C").value = item.valor;
        linha.getCell("D").value = { formula: `B${linhaNum}*C${linhaNum}` };
      });

      // 5. Preenche os campos "móveis"
      const linhaInicialTotal = 9;
      const linhaDoTotal = linhaInicialTotal + (numItens - 1);
      const ultimaLinhaItem = linhaInicialItens + numItens - 1;
      worksheet.getCell(`D${linhaDoTotal}`).value = {
        formula: `SUM(D${linhaInicialItens}:D${ultimaLinhaItem})`,
      };

      const linhaInicialCondicoes = 10;
      const linhaDasCondicoes = linhaInicialCondicoes + (numItens - 1);
      worksheet.getCell(`A${linhaDasCondicoes}`).value =
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
