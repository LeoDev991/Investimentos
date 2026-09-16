import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "outputs/simulador_fii";
const qaDir = "work/qa_simulador_fii";
await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(qaDir, { recursive: true });

const workbook = Workbook.create();
const dashboard = workbook.worksheets.add("Dashboard");
const premissas = workbook.worksheets.add("Premissas");
const simulacao = workbook.worksheets.add("Simulacao");
const checks = workbook.worksheets.add("Checks");
const fontes = workbook.worksheets.add("Fontes");

const horizonRows = 120;
const lastSimRow = horizonRows + 3;
const startDate = new Date(Date.UTC(2026, 8, 1));

const colors = {
  navy: "#12355B",
  teal: "#0F766E",
  mint: "#D9EAD3",
  paleBlue: "#EAF3F8",
  lightYellow: "#FFF2CC",
  lightGreen: "#E2F0D9",
  lightRed: "#FCE4D6",
  darkGray: "#404040",
  mediumGray: "#A6A6A6",
  lightGray: "#D9E2EC",
  white: "#FFFFFF",
  black: "#000000",
  inputBlue: "#0000FF",
};

function title(sheet, range, text) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format = {
    fill: colors.navy,
    font: { bold: true, color: colors.white, size: 16 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  r.format.rowHeight = 34;
}

function sectionHeader(sheet, range, text) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format = {
    fill: colors.teal,
    font: { bold: true, color: colors.white },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
}

function headerRow(range) {
  range.format = {
    fill: colors.navy,
    font: { bold: true, color: colors.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
  range.format.borders = { preset: "outside", style: "thin", color: colors.navy };
}

function frame(range) {
  range.format.borders = {
    preset: "outside",
    style: "thin",
    color: colors.mediumGray,
  };
}

function hideGridlines(...sheets) {
  for (const sheet of sheets) sheet.showGridLines = false;
}

hideGridlines(dashboard, premissas, simulacao, checks, fontes);

// Premissas
title(premissas, "A1:H1", "Simulador de Investimentos em Fundos Imobiliarios");
premissas.getRange("A2:H2").merge();
premissas.getRange("A2:H2").values = [[
  "Edite apenas as celulas amarelas. Os demais resultados sao calculados automaticamente.",
]];
premissas.getRange("A2:H2").format = {
  fill: colors.paleBlue,
  font: { italic: true, color: colors.darkGray },
  wrapText: true,
};

sectionHeader(premissas, "A4:D4", "Entradas do usuario");
premissas.getRange("A5:D5").values = [["Premissa", "Valor", "Unidade", "Observacao"]];
headerRow(premissas.getRange("A5:D5"));
premissas.getRange("A6:D14").values = [
  ["Valor inicial investido", 10000, "R$", "Capital aplicado no inicio da simulacao."],
  ["Aporte mensal", 1000, "R$/mes", "Valor aportado a cada mes."],
  ["Preco medio por cota", 100, "R$/cota", "Preco inicial estimado da cota."],
  ["Dividend yield mensal estimado", 0.008, "% a.m.", "Rendimento mensal sobre o patrimonio do mes."],
  ["Valorizacao mensal da cota", 0.002, "% a.m.", "Crescimento estimado do preco da cota."],
  ["Horizonte da simulacao", 120, "meses", "Limite operacional desta versao: 1 a 120 meses."],
  ["Reinvestir dividendos?", "Sim", "Sim/Nao", "Se Sim, dividendos compram novas cotas."],
  ["Mes inicial", startDate, "data", "Data usada para rotular os periodos mensais."],
  ["Custos/taxas mensais", 0, "% a.m.", "Opcional: custos como % do patrimonio do mes."],
];
premissas.getRange("A5:D14").format.borders = {
  insideHorizontal: { style: "thin", color: colors.lightGray },
  top: { style: "thin", color: colors.mediumGray },
  bottom: { style: "thin", color: colors.mediumGray },
  left: { style: "thin", color: colors.mediumGray },
  right: { style: "thin", color: colors.mediumGray },
};
premissas.getRange("A6:A14").format = { font: { bold: true } };
premissas.getRange("B6:B14").format = {
  fill: colors.lightYellow,
  font: { color: colors.inputBlue },
  horizontalAlignment: "right",
};
premissas.getRange("D6:D14").format = { wrapText: true, font: { color: colors.darkGray } };
premissas.getRange("B6:B8").format.numberFormat = "$#,##0;[Red]($#,##0);-";
premissas.getRange("B9:B10").format.numberFormat = "0.00%;[Red](0.00%);-";
premissas.getRange("B11").format.numberFormat = "#,##0";
premissas.getRange("B13").format.numberFormat = "mmm yyyy";
premissas.getRange("B14").format.numberFormat = "0.00%;[Red](0.00%);-";
premissas.getRange("B11").dataValidation = {
  rule: { type: "whole", operator: "between", formula1: 1, formula2: horizonRows },
};
premissas.getRange("B12").dataValidation = { rule: { type: "list", values: ["Sim", "Nao"] } };

sectionHeader(premissas, "F4:H4", "Legenda");
premissas.getRange("F5:H8").values = [
  ["Cor", "Uso", "Descricao"],
  ["", "Entrada editavel", "Altere estas celulas para testar cenarios."],
  ["", "Formula/calculo", "Valores derivados automaticamente."],
  ["", "Checagem OK", "Validacao simples sem alertas."],
];
headerRow(premissas.getRange("F5:H5"));
premissas.getRange("F6").format = { fill: colors.lightYellow, font: { color: colors.inputBlue } };
premissas.getRange("F7").format = { fill: colors.white, font: { color: colors.black } };
premissas.getRange("F8").format = { fill: colors.lightGreen };
premissas.getRange("F5:H8").format.borders = {
  insideHorizontal: { style: "thin", color: colors.lightGray },
  top: { style: "thin", color: colors.mediumGray },
  bottom: { style: "thin", color: colors.mediumGray },
  left: { style: "thin", color: colors.mediumGray },
  right: { style: "thin", color: colors.mediumGray },
};

sectionHeader(premissas, "A17:H17", "Aviso");
premissas.getRange("A18:H20").merge();
premissas.getRange("A18:H20").values = [[
  "Modelo educativo baseado em premissas informadas pelo usuario. Os resultados sao simulacoes e nao constituem recomendacao de investimento.",
]];
premissas.getRange("A18:H20").format = {
  fill: colors.paleBlue,
  wrapText: true,
  font: { color: colors.darkGray },
  verticalAlignment: "center",
};
frame(premissas.getRange("A18:H20"));

premissas.getRange("A:A").format.columnWidth = 26;
premissas.getRange("B:B").format.columnWidth = 16;
premissas.getRange("C:C").format.columnWidth = 12;
premissas.getRange("D:D").format.columnWidth = 42;
premissas.getRange("F:F").format.columnWidth = 12;
premissas.getRange("G:G").format.columnWidth = 20;
premissas.getRange("H:H").format.columnWidth = 34;
premissas.freezePanes.freezeRows(5);

// Simulacao
title(simulacao, "A1:Q1", "Simulacao mensal");
simulacao.getRange("A2:Q2").merge();
simulacao.getRange("A2:Q2").values = [[
  "Linhas ativas respeitam o horizonte definido em Premissas. Os calculos compram cotas com aporte mensal e, opcionalmente, reinvestem dividendos.",
]];
simulacao.getRange("A2:Q2").format = {
  fill: colors.paleBlue,
  font: { italic: true, color: colors.darkGray },
  wrapText: true,
};
simulacao.getRange("A3:Q3").values = [[
  "Mes",
  "Data",
  "Cotacao estimada",
  "Cotas iniciais",
  "Aporte do mes",
  "Dividendos brutos",
  "Custos/taxas",
  "Dividendos liquidos",
  "Reinvestimento",
  "Cotas compradas",
  "Cotas finais",
  "Patrimonio final",
  "Total aportado",
  "Dividendos gerados acum.",
  "Dividendos em caixa acum.",
  "Retorno total",
  "Status",
]];
headerRow(simulacao.getRange("A3:Q3"));

const rows = [];
for (let i = 1; i <= horizonRows; i += 1) rows.push([i, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null]);
simulacao.getRange(`A4:Q${horizonRows + 3}`).values = rows;

simulacao.getRange("B4").formulas = [["=IF($A4<='Premissas'!$B$11,EDATE('Premissas'!$B$13,$A4-1),\"\")"]];
simulacao.getRange(`B4:B${horizonRows + 3}`).fillDown();
simulacao.getRange("C4").formulas = [["=IF($A4<='Premissas'!$B$11,'Premissas'!$B$8*(1+'Premissas'!$B$10)^($A4-1),\"\")"]];
simulacao.getRange(`C4:C${horizonRows + 3}`).fillDown();
simulacao.getRange("D4").formulas = [["=IF($A4<='Premissas'!$B$11,IF($A4=1,'Premissas'!$B$6/'Premissas'!$B$8,K3),\"\")"]];
simulacao.getRange(`D4:D${horizonRows + 3}`).fillDown();
simulacao.getRange("E4").formulas = [["=IF($A4<='Premissas'!$B$11,'Premissas'!$B$7,\"\")"]];
simulacao.getRange(`E4:E${horizonRows + 3}`).fillDown();
simulacao.getRange("F4").formulas = [["=IF($A4<='Premissas'!$B$11,D4*C4*'Premissas'!$B$9,\"\")"]];
simulacao.getRange(`F4:F${horizonRows + 3}`).fillDown();
simulacao.getRange("G4").formulas = [["=IF($A4<='Premissas'!$B$11,D4*C4*'Premissas'!$B$14,\"\")"]];
simulacao.getRange(`G4:G${horizonRows + 3}`).fillDown();
simulacao.getRange("H4").formulas = [["=IF($A4<='Premissas'!$B$11,MAX(F4-G4,0),\"\")"]];
simulacao.getRange(`H4:H${horizonRows + 3}`).fillDown();
simulacao.getRange("I4").formulas = [["=IF($A4<='Premissas'!$B$11,IF('Premissas'!$B$12=\"Sim\",H4,0),\"\")"]];
simulacao.getRange(`I4:I${horizonRows + 3}`).fillDown();
simulacao.getRange("J4").formulas = [["=IF($A4<='Premissas'!$B$11,(E4+I4)/C4,\"\")"]];
simulacao.getRange(`J4:J${horizonRows + 3}`).fillDown();
simulacao.getRange("K4").formulas = [["=IF($A4<='Premissas'!$B$11,D4+J4,\"\")"]];
simulacao.getRange(`K4:K${horizonRows + 3}`).fillDown();
simulacao.getRange("L4").formulas = [["=IF($A4<='Premissas'!$B$11,K4*C4,\"\")"]];
simulacao.getRange(`L4:L${horizonRows + 3}`).fillDown();
simulacao.getRange("M4").formulas = [["=IF($A4<='Premissas'!$B$11,IF($A4=1,'Premissas'!$B$6+E4,M3+E4),\"\")"]];
simulacao.getRange(`M4:M${horizonRows + 3}`).fillDown();
simulacao.getRange("N4").formulas = [["=IF($A4<='Premissas'!$B$11,IF($A4=1,H4,N3+H4),\"\")"]];
simulacao.getRange(`N4:N${horizonRows + 3}`).fillDown();
simulacao.getRange("O4").formulas = [["=IF($A4<='Premissas'!$B$11,IF($A4=1,IF('Premissas'!$B$12=\"Sim\",0,H4),O3+IF('Premissas'!$B$12=\"Sim\",0,H4)),\"\")"]];
simulacao.getRange(`O4:O${horizonRows + 3}`).fillDown();
simulacao.getRange("P4").formulas = [["=IF($A4<='Premissas'!$B$11,IF(M4=0,0,(L4+O4)/M4-1),\"\")"]];
simulacao.getRange(`P4:P${horizonRows + 3}`).fillDown();
simulacao.getRange("Q4").formulas = [["=IF($A4<='Premissas'!$B$11,\"Ativo\",\"Fora\")"]];
simulacao.getRange(`Q4:Q${horizonRows + 3}`).fillDown();

simulacao.getRange(`A4:Q${horizonRows + 3}`).format.borders = {
  insideHorizontal: { style: "thin", color: colors.lightGray },
  top: { style: "thin", color: colors.mediumGray },
  bottom: { style: "thin", color: colors.mediumGray },
};
simulacao.getRange(`A4:A${horizonRows + 3}`).format.numberFormat = "#,##0";
simulacao.getRange(`B4:B${horizonRows + 3}`).format.numberFormat = "mmm yyyy";
simulacao.getRange(`C4:C${horizonRows + 3}`).format.numberFormat = "$#,##0.00;[Red]($#,##0.00);-";
simulacao.getRange(`D4:D${horizonRows + 3}`).format.numberFormat = "#,##0.0000";
simulacao.getRange(`E4:I${horizonRows + 3}`).format.numberFormat = "$#,##0;[Red]($#,##0);-";
simulacao.getRange(`J4:K${horizonRows + 3}`).format.numberFormat = "#,##0.0000";
simulacao.getRange(`L4:O${horizonRows + 3}`).format.numberFormat = "$#,##0;[Red]($#,##0);-";
simulacao.getRange(`P4:P${horizonRows + 3}`).format.numberFormat = "0.0%;[Red](0.0%);-";
simulacao.getRange(`Q4:Q${horizonRows + 3}`).format = { horizontalAlignment: "center" };
simulacao.getRange(`Q4:Q${horizonRows + 3}`).conditionalFormats.add("containsText", {
  text: "Ativo",
  format: { fill: colors.lightGreen, font: { color: "#006100" } },
});
simulacao.getRange(`Q4:Q${horizonRows + 3}`).conditionalFormats.add("containsText", {
  text: "Fora",
  format: { fill: "#F2F2F2", font: { color: colors.mediumGray } },
});
simulacao.getRange("A:Q").format.autofitColumns();
simulacao.getRange("A:A").format.columnWidth = 8;
simulacao.getRange("B:B").format.columnWidth = 12;
simulacao.getRange("C:P").format.columnWidth = 15;
simulacao.getRange("Q:Q").format.columnWidth = 10;
simulacao.freezePanes.freezeRows(3);
simulacao.tables.add(`A3:Q${horizonRows + 3}`, true, "TabelaSimulacao");

// Checks
title(checks, "A1:F1", "Checagens do modelo");
checks.getRange("A2").values = [["Status geral"]];
checks.getRange("B2").formulas = [["=IF(COUNTIF(E5:E10,\"Revisar\")=0,\"OK\",\"Revisar\")"]];
checks.getRange("A2:B2").format = {
  fill: colors.paleBlue,
  font: { bold: true },
};
checks.getRange("B2").format = { horizontalAlignment: "center", font: { bold: true } };
checks.getRange("B2").conditionalFormats.add("containsText", {
  text: "OK",
  format: { fill: colors.lightGreen, font: { color: "#006100", bold: true } },
});
checks.getRange("B2").conditionalFormats.add("containsText", {
  text: "Revisar",
  format: { fill: colors.lightRed, font: { color: "#9C0006", bold: true } },
});
checks.getRange("A4:F4").values = [["Checagem", "Atual", "Esperado", "Diferenca", "Status", "Onde ajustar"]];
headerRow(checks.getRange("A4:F4"));
checks.getRange("A5:F10").values = [
  ["Horizonte dentro do limite", null, "1 a 120 meses", null, null, "Premissas!B11"],
  ["Preco medio por cota maior que zero", null, "> 0", null, null, "Premissas!B8"],
  ["Aportes nao negativos", null, ">= 0", null, null, "Premissas!B6:B7"],
  ["Dividend yield nao negativo", null, ">= 0%", null, null, "Premissas!B9"],
  ["Valorizacao mensal acima de -100%", null, "> -100%", null, null, "Premissas!B10"],
  ["Custos/taxas nao negativos", null, ">= 0%", null, null, "Premissas!B14"],
];
checks.getRange("B5").formulas = [["='Premissas'!$B$11"]];
checks.getRange("B6").formulas = [["='Premissas'!$B$8"]];
checks.getRange("B7").formulas = [["=MIN('Premissas'!$B$6:$B$7)"]];
checks.getRange("B8").formulas = [["='Premissas'!$B$9"]];
checks.getRange("B9").formulas = [["='Premissas'!$B$10"]];
checks.getRange("B10").formulas = [["='Premissas'!$B$14"]];
checks.getRange("D5").formulas = [["=IF(AND(B5>=1,B5<=120),0,1)"]];
checks.getRange("D6").formulas = [["=IF(B6>0,0,1)"]];
checks.getRange("D7").formulas = [["=IF(B7>=0,0,1)"]];
checks.getRange("D8").formulas = [["=IF(B8>=0,0,1)"]];
checks.getRange("D9").formulas = [["=IF(B9>-100%,0,1)"]];
checks.getRange("D10").formulas = [["=IF(B10>=0,0,1)"]];
checks.getRange("E5").formulas = [["=IF(D5=0,\"OK\",\"Revisar\")"]];
checks.getRange("E5:E10").fillDown();
checks.getRange("A5:F10").format.borders = {
  insideHorizontal: { style: "thin", color: colors.lightGray },
  top: { style: "thin", color: colors.mediumGray },
  bottom: { style: "thin", color: colors.mediumGray },
  left: { style: "thin", color: colors.mediumGray },
  right: { style: "thin", color: colors.mediumGray },
};
checks.getRange("B5:B7").format.numberFormat = "#,##0.00";
checks.getRange("B8:B10").format.numberFormat = "0.00%;[Red](0.00%);-";
checks.getRange("D5:D10").format.numberFormat = "#,##0";
checks.getRange("E5:E10").format = { horizontalAlignment: "center", font: { bold: true } };
checks.getRange("E5:E10").conditionalFormats.add("containsText", {
  text: "OK",
  format: { fill: colors.lightGreen, font: { color: "#006100" } },
});
checks.getRange("E5:E10").conditionalFormats.add("containsText", {
  text: "Revisar",
  format: { fill: colors.lightRed, font: { color: "#9C0006" } },
});
checks.getRange("A:A").format.columnWidth = 36;
checks.getRange("B:D").format.columnWidth = 15;
checks.getRange("E:E").format.columnWidth = 12;
checks.getRange("F:F").format.columnWidth = 20;
checks.freezePanes.freezeRows(4);

// Fontes
title(fontes, "A1:G1", "Fontes e premissas");
fontes.getRange("A3:G3").values = [["Item", "Valor", "Unidade", "Periodo/as-of", "Tipo", "Fonte", "Observacoes"]];
headerRow(fontes.getRange("A3:G3"));
fontes.getRange("A4:G12").values = [
  ["Valor inicial investido", null, "R$", "2026-09-16", "Premissa do usuario", "Premissas!B6", "Valor editavel usado como capital inicial."],
  ["Aporte mensal", null, "R$/mes", "2026-09-16", "Premissa do usuario", "Premissas!B7", "Valor editavel de aporte recorrente."],
  ["Preco medio por cota", null, "R$/cota", "2026-09-16", "Premissa do usuario", "Premissas!B8", "Preco inicial estimado; personalize conforme o FII analisado."],
  ["Dividend yield mensal estimado", null, "% a.m.", "2026-09-16", "Premissa do usuario", "Premissas!B9", "Premissa simplificada, sem garantia de recorrencia."],
  ["Valorizacao mensal da cota", null, "% a.m.", "2026-09-16", "Premissa do usuario", "Premissas!B10", "Pode ser negativa para simular queda de preco."],
  ["Horizonte da simulacao", null, "meses", "2026-09-16", "Premissa do usuario", "Premissas!B11", "Modelo preparado para ate 120 meses."],
  ["Reinvestir dividendos?", null, "Sim/Nao", "2026-09-16", "Premissa do usuario", "Premissas!B12", "Controla se dividendos compram novas cotas."],
  ["Mes inicial", null, "data", "2026-09-16", "Premissa do usuario", "Premissas!B13", "Rotulo temporal dos periodos."],
  ["Custos/taxas mensais", null, "% a.m.", "2026-09-16", "Premissa do usuario", "Premissas!B14", "Campo opcional para custos mensais estimados."],
];
for (let row = 4; row <= 12; row += 1) {
  fontes.getRange(`B${row}`).formulas = [[`='Premissas'!$B$${row + 2}`]];
}
fontes.getRange("A4:G12").format.borders = {
  insideHorizontal: { style: "thin", color: colors.lightGray },
  top: { style: "thin", color: colors.mediumGray },
  bottom: { style: "thin", color: colors.mediumGray },
  left: { style: "thin", color: colors.mediumGray },
  right: { style: "thin", color: colors.mediumGray },
};
fontes.getRange("B4:B6").format.numberFormat = "$#,##0;[Red]($#,##0);-";
fontes.getRange("B7:B8").format.numberFormat = "0.00%;[Red](0.00%);-";
fontes.getRange("B9").format.numberFormat = "#,##0";
fontes.getRange("B11").format.numberFormat = "mmm yyyy";
fontes.getRange("B12").format.numberFormat = "0.00%;[Red](0.00%);-";
fontes.getRange("A:A").format.columnWidth = 28;
fontes.getRange("B:B").format.columnWidth = 16;
fontes.getRange("C:C").format.columnWidth = 12;
fontes.getRange("D:E").format.columnWidth = 18;
fontes.getRange("F:F").format.columnWidth = 18;
fontes.getRange("G:G").format.columnWidth = 48;
fontes.getRange("G4:G12").format = { wrapText: true };
fontes.freezePanes.freezeRows(3);

// Dashboard
title(dashboard, "A1:H1", "Painel do simulador de FIIs");
dashboard.getRange("A2:H2").merge();
dashboard.getRange("A2:H2").values = [[
  "Resumo calculado com base nas premissas e na simulacao mensal.",
]];
dashboard.getRange("A2:H2").format = {
  fill: colors.paleBlue,
  font: { italic: true, color: colors.darkGray },
  wrapText: true,
};

dashboard.getRange("A4:H4").values = [[
  "Total investido",
  "Patrimonio final",
  "Dividendos no ultimo mes",
  "Dividendos gerados",
  "Dividendos em caixa",
  "Retorno total",
  "Cotas finais",
  "Status",
]];
headerRow(dashboard.getRange("A4:H4"));
dashboard.getRange("A5:H5").formulas = [[
  `=INDEX('Simulacao'!$M$4:$M$${lastSimRow},'Premissas'!$B$11)`,
  `=INDEX('Simulacao'!$L$4:$L$${lastSimRow},'Premissas'!$B$11)`,
  `=INDEX('Simulacao'!$H$4:$H$${lastSimRow},'Premissas'!$B$11)`,
  `=INDEX('Simulacao'!$N$4:$N$${lastSimRow},'Premissas'!$B$11)`,
  `=INDEX('Simulacao'!$O$4:$O$${lastSimRow},'Premissas'!$B$11)`,
  `=INDEX('Simulacao'!$P$4:$P$${lastSimRow},'Premissas'!$B$11)`,
  `=INDEX('Simulacao'!$K$4:$K$${lastSimRow},'Premissas'!$B$11)`,
  "='Checks'!$B$2",
]];
dashboard.getRange("A5:H5").format = {
  fill: colors.white,
  font: { bold: true, size: 12 },
  horizontalAlignment: "center",
};
dashboard.getRange("A5:E5").format.numberFormat = "$#,##0;[Red]($#,##0);-";
dashboard.getRange("F5").format.numberFormat = "0.0%;[Red](0.0%);-";
dashboard.getRange("G5").format.numberFormat = "#,##0.0000";
dashboard.getRange("A4:H5").format.borders = {
  insideVertical: { style: "thin", color: colors.lightGray },
  top: { style: "thin", color: colors.mediumGray },
  bottom: { style: "thin", color: colors.mediumGray },
  left: { style: "thin", color: colors.mediumGray },
  right: { style: "thin", color: colors.mediumGray },
};
dashboard.getRange("H5").conditionalFormats.add("containsText", {
  text: "OK",
  format: { fill: colors.lightGreen, font: { color: "#006100", bold: true } },
});
dashboard.getRange("H5").conditionalFormats.add("containsText", {
  text: "Revisar",
  format: { fill: colors.lightRed, font: { color: "#9C0006", bold: true } },
});

sectionHeader(dashboard, "A8:H8", "Premissas usadas no cenario");
dashboard.getRange("A9:H10").values = [
  [
    "Valor inicial",
    "Aporte mensal",
    "Preco/cota",
    "DY mensal",
    "Valorizacao mensal",
    "Horizonte",
    "Reinvestir",
    "Custos/taxas",
  ],
  [null, null, null, null, null, null, null, null],
];
dashboard.getRange("A10:H10").formulas = [[
  "='Premissas'!$B$6",
  "='Premissas'!$B$7",
  "='Premissas'!$B$8",
  "='Premissas'!$B$9",
  "='Premissas'!$B$10",
  "='Premissas'!$B$11",
  "='Premissas'!$B$12",
  "='Premissas'!$B$14",
]];
headerRow(dashboard.getRange("A9:H9"));
dashboard.getRange("A10:C10").format.numberFormat = "$#,##0;[Red]($#,##0);-";
dashboard.getRange("D10:E10").format.numberFormat = "0.00%;[Red](0.00%);-";
dashboard.getRange("F10").format.numberFormat = "#,##0";
dashboard.getRange("H10").format.numberFormat = "0.00%;[Red](0.00%);-";
dashboard.getRange("A9:H10").format.borders = {
  insideVertical: { style: "thin", color: colors.lightGray },
  top: { style: "thin", color: colors.mediumGray },
  bottom: { style: "thin", color: colors.mediumGray },
  left: { style: "thin", color: colors.mediumGray },
  right: { style: "thin", color: colors.mediumGray },
};

dashboard.getRange("J4:L4").values = [["Mes", "Patrimonio", "Dividendos"]];
for (let row = 5; row <= horizonRows + 4; row += 1) {
  const simRow = row - 1;
  dashboard.getRange(`J${row}`).formulas = [[`=IF('Simulacao'!$A${simRow}<='Premissas'!$B$11,TEXT('Simulacao'!$B${simRow},"mmm yyyy"),"")`]];
  dashboard.getRange(`K${row}`).formulas = [[`=IF('Simulacao'!$A${simRow}<='Premissas'!$B$11,'Simulacao'!$L${simRow},NA())`]];
  dashboard.getRange(`L${row}`).formulas = [[`=IF('Simulacao'!$A${simRow}<='Premissas'!$B$11,'Simulacao'!$H${simRow},NA())`]];
}
dashboard.getRange(`K5:L${horizonRows + 4}`).format.numberFormat = "$#,##0;[Red]($#,##0);-";
dashboard.getRange("N4:O4").values = [["Mes", "Dividendos"]];
for (let row = 5; row <= horizonRows + 4; row += 1) {
  dashboard.getRange(`N${row}`).formulas = [[`=J${row}`]];
  dashboard.getRange(`O${row}`).formulas = [[`=L${row}`]];
}
dashboard.getRange(`O5:O${horizonRows + 4}`).format.numberFormat = "$#,##0;[Red]($#,##0);-";

const patrimonioChart = dashboard.charts.add("line", dashboard.getRange(`J4:K${horizonRows + 4}`));
patrimonioChart.title = "Patrimonio acumulado (R$)";
patrimonioChart.hasLegend = false;
patrimonioChart.xAxis = { axisType: "textAxis", textStyle: { fontSize: 8 }, tickLabelInterval: 12 };
patrimonioChart.yAxis = { numberFormatCode: "$#,##0" };
patrimonioChart.setPosition("A13", "H28");

const dividendosChart = dashboard.charts.add("line", dashboard.getRange(`N4:O${horizonRows + 4}`));
dividendosChart.title = "Dividendos mensais estimados (R$)";
dividendosChart.hasLegend = false;
dividendosChart.xAxis = { axisType: "textAxis", textStyle: { fontSize: 8 }, tickLabelInterval: 12 };
dividendosChart.yAxis = { numberFormatCode: "$#,##0" };
dividendosChart.setPosition("A30", "H45");

dashboard.getRange("J:O").format.columnWidth = 0;
dashboard.getRange("A:H").format.columnWidth = 16;
dashboard.getRange("A2:H2").format.rowHeight = 28;
dashboard.freezePanes.freezeRows(4);

// Comments/source notes for editable assumption cells.
workbook.comments.setSelf({ displayName: "User" });
const assumptionNotes = [
  ["B6", "Assumption: valor informado pelo usuario | Date: 2026-09-16 | Notes: capital inicial editavel."],
  ["B7", "Assumption: valor informado pelo usuario | Date: 2026-09-16 | Notes: aporte mensal editavel."],
  ["B8", "Assumption: valor informado pelo usuario | Date: 2026-09-16 | Notes: preco medio estimado por cota."],
  ["B9", "Assumption: valor informado pelo usuario | Date: 2026-09-16 | Notes: dividend yield mensal estimado."],
  ["B10", "Assumption: valor informado pelo usuario | Date: 2026-09-16 | Notes: valorizacao mensal estimada da cota."],
  ["B11", "Assumption: valor informado pelo usuario | Date: 2026-09-16 | Notes: horizonte limitado a 120 meses."],
  ["B12", "Assumption: valor informado pelo usuario | Date: 2026-09-16 | Notes: opcao de reinvestimento dos dividendos."],
  ["B13", "Assumption: valor informado pelo usuario | Date: 2026-09-16 | Notes: mes inicial da simulacao."],
  ["B14", "Assumption: valor informado pelo usuario | Date: 2026-09-16 | Notes: custos ou taxas mensais estimadas."],
];
for (const [cell, note] of assumptionNotes) {
  workbook.comments.addThread({ cell: premissas.getRange(cell) }, note);
}

// Compact visual previews for QA.
const dashboardPreview = await workbook.render({
  sheetName: "Dashboard",
  range: "A1:H45",
  scale: 1,
  format: "png",
});
const dashboardPreviewBytes = new Uint8Array(await dashboardPreview.arrayBuffer());
await fs.writeFile(`${qaDir}/dashboard_preview.png`, dashboardPreviewBytes);

for (const [sheetName, range] of [
  ["Premissas", "A1:H20"],
  ["Simulacao", "A1:Q28"],
  ["Checks", "A1:F12"],
  ["Fontes", "A1:G12"],
]) {
  const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(`${qaDir}/${sheetName.toLowerCase()}_preview.png`, new Uint8Array(await preview.arrayBuffer()));
}

const keyInspect = await workbook.inspect({
  kind: "table",
  range: "Dashboard!A1:H10",
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 8,
});
console.log(keyInspect.ndjson);

const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(errorScan.ndjson);

const checksInspect = await workbook.inspect({
  kind: "table",
  range: "Checks!A1:F10",
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 6,
});
console.log(checksInspect.ndjson);

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(`${outputDir}/simulador_investimentos_fii.xlsx`);
await fs.rm(`${outputDir}/dashboard_preview.png`, { force: true });
for (const name of ["premissas", "simulacao", "checks", "fontes"]) {
  await fs.rm(`${outputDir}/${name}_preview.png`, { force: true });
}
await fs.rm(`${outputDir}/simulador_investimentos_fii.xlsx.inspect.ndjson`, { force: true });
