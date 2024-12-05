import ExcelJS from "exceljs";


export const exportarParaExcel = (filteredData, option) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Faturamento");

  // Edita as colunas e o titulo de cada uma
  worksheet.columns = [
    { header: "Nome", key: "Nome", width: 45 },
    { header: "CPF", key: "CPF", width: 15 },
    { header: "Dias", key: "Dias", width: 6 },
    { header: "Valor", key: "Valor", width: 8 },
    { header: "Admissão", key: "Admissão", width: 15 },
    { header: "Demissão", key: "Demissão", width: 15 },
    { header: "Observação", key: "Observação", width: 15 },
  ];

  //Adicoonando Titulo e mesclando
  worksheet.mergeCells("A1:G1");
  const title = worksheet.getCell("A1");
  title.value = "Faturamento referente a"; // Texto do título
  title.alignment = { horizontal: "center", vertical: "middle" }; // Alinhamento no centro
  title.font = { bold: true, size: 20 }; // Fonte em negrito e tamanho maior
  title.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "006CE3" }, // Cor cinza claro
  };

  //Add titulo Nome
  const titleNome = worksheet.getCell("A2");
  titleNome.value = "Nome";
  titleNome.alignment = { horizontal: "center", vertical: "middle" };
  titleNome.font = { bold: true, size: 12 };
  titleNome.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "F0F0F0" }, // Cor de fundo do cabeçalho (cinza claro)
  };

  //Add titulo Nome
  const titleCPF = worksheet.getCell("B2");
  titleCPF.value = "C.P.F";
  titleCPF.alignment = { horizontal: "center", vertical: "middle" };
  titleCPF.font = { bold: true, size: 12 };
  titleCPF.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "F0F0F0" }, // Cor de fundo do cabeçalho (cinza claro)
  };

  //Add titulo Dias
  const titleDias = worksheet.getCell("C2");
  titleDias.value = "Dias";
  titleDias.alignment = { horizontal: "center", vertical: "middle" };
  titleDias.font = { bold: true, size: 12 };
  titleDias.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "F0F0F0" }, // Cor de fundo do cabeçalho (cinza claro)
  };

  //Add titulo Valor
  const titleValor = worksheet.getCell("D2");
  titleValor.value = "Valor R$";
  titleValor.alignment = { horizontal: "center", vertical: "middle" };
  titleValor.font = { bold: true, size: 12 };
  titleValor.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "F0F0F0" }, // Cor de fundo do cabeçalho (cinza claro)
  };

  //Add titulo Admissão
  const titleAd = worksheet.getCell("E2");
  titleAd.value = "Admissão";
  titleAd.alignment = { horizontal: "center", vertical: "middle" };
  titleAd.font = { bold: true, size: 12 };
  titleAd.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "F0F0F0" }, // Cor de fundo do cabeçalho (cinza claro)
  };

  //Add titulo Demissão
  const titleDm = worksheet.getCell("F2");
  titleDm.value = "Demissão";
  titleDm.alignment = { horizontal: "center", vertical: "middle" };
  titleDm.font = { bold: true, size: 12 };
  titleDm.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "F0F0F0" }, // Cor de fundo do cabeçalho (cinza claro)
  };

  //Add titulo Observação
  const titleObs = worksheet.getCell("G2");
  titleObs.value = "Observação";
  titleObs.alignment = { horizontal: "center", vertical: "middle" };
  titleObs.font = { bold: true, size: 12 };
  titleObs.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "F0F0F0" }, // Cor de fundo do cabeçalho (cinza claro)
  };

  // Aplica estilo ao cabeçalho (linha 2)
  const headerRow = worksheet.getRow(2);
  headerRow.font = { bold: true }; // Negrito
  headerRow.alignment = { horizontal: "center" }; // Centraliza o texto

  // Adicionar os dados filtrados ao arquivo Excel
  filteredData.forEach((row) => {
    worksheet.addRow({
      Nome: row.nome,
      CPF: row.cpf,
      Dias: row.Dias_Trabalhados,
      Valor: row.valorTotalEmpresa,
      Admissão: row.admissao,
      Demissão: row.demissao,
      Observação: "",
    });
  });
  worksheet.addRows(filteredData);

  // Configura a coluna 'Valor' com formato de duas casas decimais
  worksheet.getColumn("Valor").numFmt = "0.00";

  //Calcula o total
  const totalFatura = filteredData
    .reduce((v, item) => v + item.valorTotalEmpresa, 0)
    .toFixed(2);
  const ultimaLinha = filteredData.length + 3; //Ultima linha adicionada

  //Adicionar texto a ultima linha e mesclar
  worksheet.mergeCells(`A${ultimaLinha}:F${ultimaLinha}`);
  const titleFatura = worksheet.getCell(`A${ultimaLinha}`);
  titleFatura.value = "Total da fatura";
  titleFatura.alignment = { horizontal: "left", vertical: "middle" };
  titleFatura.font = { bold: true, size: 16 }; // Fonte em negrito e tamanho maior
  titleFatura.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "006CE3" }, // Cor cinza claro
  };

  //Adicionar total a celula G
  const total = worksheet.getCell(`G${ultimaLinha}`);
  total.value = "R$ " + parseFloat(totalFatura);
  total.alignment = { horizontal: "center", vertical: "middle" };
  total.font = { bold: true, size: 16 }; // Fonte em negrito e tamanho maior
  total.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "006CE3" }, // Cor cinza claro
  };
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });
  });



  // Salvar o arquivo
  workbook.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `faturamento_${option}.xlsx`;
    link.click();
  });
};
