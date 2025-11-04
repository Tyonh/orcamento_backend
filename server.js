const express = require("express");
const exceljs = require("exceljs");
const path = require("path");
const cors = require("cors");
const sqlite3 = require("sqlite3");

const app = express();
const port = process.env.PORT || 3000;

const DB_PATH = path.join(__dirname, "produtos.db");

app.use(cors());
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Rota de busca de produtos (intocada)
app.get("/api/produtos/search", (req, res) => {
  const termoBusca = (req.query.search || "").toLowerCase().trim();

  if (termoBusca.length < 2) {
    return res.json([]);
  }
  const db = new sqlite3.Database(DB_PATH, sqlite3.OPEN_READONLY, (err) => {
    if (err) {
      console.error("Erro ao conectar ao SQLite:", err);
      return res.status(500).json({ message: "Erro de banco de dados." });
    }
  });

  const termoLike = `%${termoBusca}%`;
  const sql =
    "SELECT nome FROM produtos WHERE nome COLLATE NOCASE LIKE ? LIMIT 10";

  db.all(sql, [termoLike], (err, rows) => {
    if (err) {
      console.error("Erro na consulta SQL:", err);
      return res.status(500).json({ message: "Erro ao buscar produtos." });
    }
    res.json(rows);
    db.close();
  });
});

console.log("Backend iniciado. Aguardando envios de formulário...");

const mapearCondicao = {
  avista: "À vista",
  "30d": "Cartão de Crédito",
  "60d": "Boleto",
};

app.post("/salvar-orcamento", async (req, res) => {
  const dadosDoFormulario = req.body;
  console.log("--- DADOS RECEBIDOS DO FORMULÁRIO ---");
  console.log(dadosDoFormulario);

  const { cliente_nome, condicao_pagamento } = dadosDoFormulario;

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

  // Filtra itens vazios
  const itensValidos = produtos
    .map((nome, index) => ({
      nome: nome,
      qtd: parseInt(quantidades[index] || 1),
      valor: parseFloat(valores[index] || 0),
    }))
    .filter((p) => p.nome && p.nome.trim() !== "");

  const numItens = itensValidos.length;
  console.log(`Número de itens válidos recebidos: ${numItens}`);

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
      `Orçamento gerado em memória: ${nomeArquivoFinal}. Enviando para download...`
    );

    // 7. Definir os cabeçalhos da resposta para o navegador
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${nomeArquivoFinal}"`
    );

    // 8. Enviar o arquivo Excel direto para o navegador
    await workbook.xlsx.write(res);
    res.end();

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
