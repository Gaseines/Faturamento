import styles from "./Planilha.module.css";

import ExcelJS from "exceljs";

import React, { useState } from "react";
import * as XLSX from "xlsx";

function Planilha() {
  const [data, setData] = useState([]); // Para armazenar os dados carregados da planilha

  const handleFileUpload = (event) => {
    const file = event.target.files[0]; // Captura o arquivo selecionado pelo usuário
    const reader = new FileReader(); // Cria um leitor de arquivos

    reader.onload = (e) => {
      const binaryStr = e.target.result; // Lê o arquivo como string binária
      const workbook = XLSX.read(binaryStr, { type: "binary" }); // Carrega o arquivo Excel
      const sheetName = workbook.SheetNames[0]; // Pega o nome da primeira aba
      const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]); // Converte para JSON
      setData(sheetData);
      // Armazena os dados no estado
    };

    reader.readAsBinaryString(file); // Lê o arquivo como binário
  };

  //Soma soma cada motorista para operador
  const somaOp = (data) => {
    return data.map((item) => ({
      result: (5.2 / 30) * item["Dias calculado"],
    }));
  };

  //Formatar data
  const excelToDate = (serial) => {
    const excelStartDate = new Date(1900, 0, 1); // Data base do Excel
    const days = parseInt(serial, 10) - 2; // Ajuste para o bug histórico do Excel
    return new Date(excelStartDate.setDate(excelStartDate.getDate() + days));
  };

  //Ajustar digitos do CPF
  const formatCPF = (cpf) => {
    if (!cpf) {
      return "CPF inválido";
    } else {
      const cpfString = cpf.toString();
      if (cpfString.length < 11) {
        return cpfString.padStart(11, "0");
      } else {
        return cpfString;
      }
    }
  };
  //Processa dados para exportar
  const exportProcess = (data) => {
    console.log(data);
    return data.map((item) => ({
      Nome: item.Nome,
      CPF: formatCPF(item["CPF"]),
      Dias: item["Dias calculado"],
      Valor: (52 / 30) * item["Dias calculado"].toFixed(2),
      Admissão: !isNaN(item.Admissão) ? excelToDate(item.Admissão) : "",
      Demissão: !isNaN(item.Demissão) ? excelToDate(item.Demissão) : "",
      Observação: "",
    }));
  };

  //Processa dados para o site
  const processarDados = (data) => {
    return data.map((item) => ({
      Dias_Trabalhados: item["Dias calculado"] || 0,
      Valor_Mot: 7,
      valorTotal: (7 / 30) * (item["Dias calculado"] || 0),
      nome: item.Nome,
      cpf: formatCPF(item["CPF"]),
      admissao: !isNaN(item.Admissão) ? excelToDate(item.Admissão) : "",
      demissao: !isNaN(item.Demissão) ? excelToDate(item.Demissão) : "",
      empresa: item.Empresa || "Não informado",
      Valor_Mot_Empresa: 52,
      valorTotalEmpresa: (52 / 30) * (item["Dias calculado"] || 0),
    }));
  };

  //Soma Total pagar ao Wagner
  const somaValorW = processarDados(data)
    .reduce((soma, item) => soma + item.valorTotal, 0)
    .toFixed(2);

  //Soma Total cobrar clientes
  const somaValorC = processarDados(data)
    .reduce((soma, item) => soma + item.valorTotalEmpresa, 0)
    .toFixed(2);

  //Soma total pagar ao operador
  const somaValorO = somaOp(data)
    .reduce((soma, item) => soma + item.result, 0)
    .toFixed(2);
  console.log(somaValorO);

  //Exporta a planilha
  const exportarParaExcel = () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Faturamento");

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

    //Add titulo Nome
    const titleNome = worksheet.getCell("A2")
    titleNome.value = "Nome"
    titleNome.alignment = { horizontal: "center", vertical: "middle" };
    titleNome.font = { bold: true, size: 12 }
    titleNome.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "F0F0F0" }, // Cor de fundo do cabeçalho (cinza claro)
      };
    

    //Add titulo Nome
    const titleCPF = worksheet.getCell("B2")
    titleCPF.value = "C.P.F"
    titleCPF.alignment = { horizontal: "center", vertical: "middle" };
    titleCPF.font = { bold: true, size: 12 }
    titleCPF.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "F0F0F0" }, // Cor de fundo do cabeçalho (cinza claro)
      };

    //Add titulo Nome
    const titleDias = worksheet.getCell("C2")
    titleDias.value = "Dias"
    titleDias.alignment = { horizontal: "center", vertical: "middle" };
    titleDias.font = { bold: true, size: 12 }
    titleDias.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "F0F0F0" }, // Cor de fundo do cabeçalho (cinza claro)
      };

    //Add titulo Nome
    const titleValor = worksheet.getCell("D2")
    titleValor.value = "Valor R$"
    titleValor.alignment = { horizontal: "center", vertical: "middle" };
    titleValor.font = { bold: true, size: 12 }
    titleValor.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "F0F0F0" }, // Cor de fundo do cabeçalho (cinza claro)
      };

    //Add titulo Nome
    const titleAd = worksheet.getCell("E2")
    titleAd.value = "Admissão"
    titleAd.alignment = { horizontal: "center", vertical: "middle" };
    titleAd.font = { bold: true, size: 12 }
    titleAd.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "F0F0F0" }, // Cor de fundo do cabeçalho (cinza claro)
      };

    //Add titulo Nome
    const titleDm = worksheet.getCell("F2")
    titleDm.value = "Demissão"
    titleDm.alignment = { horizontal: "center", vertical: "middle" };
    titleDm.font = { bold: true, size: 12 }
    titleDm.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "F0F0F0" }, // Cor de fundo do cabeçalho (cinza claro)
      };

    const titleObs = worksheet.getCell("G2")
    titleObs.value = "Observação"
    titleObs.alignment = { horizontal: "center", vertical: "middle" };
    titleObs.font = { bold: true, size: 12 }
    titleObs.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "F0F0F0" }, // Cor de fundo do cabeçalho (cinza claro)
      };

    // Aplica estilo ao cabeçalho (linha 2)
    const headerRow = worksheet.getRow(2);
    headerRow.font = { bold: true }; // Negrito
    headerRow.alignment = { horizontal: "center" }
    ; // Centraliza o texto
    

    //Adiciona os dados a planilha
    const rows = exportProcess(data);
    worksheet.addRows(rows);

    // Configura a coluna 'Valor' com formato de duas casas decimais
    worksheet.getColumn("Valor").numFmt = "0.00";

    //Calcula o total
    const totalFatura = rows.reduce((v, item) => v + item.Valor, 0).toFixed(2);
    const ultimaLinha = rows.length + 2; //Ultima linha adicionada

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

    const total = worksheet.getCell(`G${ultimaLinha}`);
    total.value = parseFloat(totalFatura);
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
      link.download = "faturamento.xlsx";
      link.click();
    });
  };

  return (
    <div className={styles.container}>
      <h1 className={styles.title}>Faturamento</h1>
      <input
        className={styles.input}
        id="file-upload"
        type="file"
        onChange={handleFileUpload}
        accept=".xlsx, .xls"
      />
      <label htmlFor="file-upload" className={styles.custom_button}>
        Selecione o arquivo
      </label>
      {data.length > 0 && (
        <div className={styles.container_planilha}>
          <h3>Dados Processados:</h3>

          <div className={styles.valor}>
            <p>
              Valor Wagner: <span>R$ {somaValorW}</span>
            </p>
            <p>
              Valor Operador: <span>R$ {somaValorO}</span>
            </p>
            <p>
              Valor Cliente: <span>R$ {somaValorC}</span>
            </p>
          </div>
          <div className={styles.planilha}>
            <div className={styles.header}>
              <p className={styles.dias}>Dias Trabalhados</p>
              <p className={styles.valor_mot}>Valor motorista</p>
              <p className={styles.total}>Valor Total</p>
              <p className={styles.nome}>Nome</p>
              <p className={styles.cpf}>CPF</p>
              <p className={styles.data}>Amissão</p>
              <p className={styles.data}>Demissão</p>
              <p className={styles.valor_mot}>Valor motorista</p>
              <p className={styles.total}>Valor Total</p>
              <p className={styles.empresa}>Empresa</p>
            </div>
            {processarDados(data).map((item, index) => (
              <div className={styles.header} key={index}>
                <p className={styles.dias}>{item.Dias_Trabalhados}</p>
                <p className={styles.valor_mot}>{item.Valor_Mot}</p>
                <p className={styles.total}>{item.valorTotal.toFixed(2)}</p>
                <p className={styles.nome}>{item.nome}</p>
                <p className={styles.cpf}>{item.cpf}</p>
                <p className={styles.data}>
                  {item.admissao instanceof Date
                    ? item.admissao.toLocaleDateString()
                    : item.admissao}
                </p>
                <p className={styles.data}>
                  {item.demissao instanceof Date
                    ? item.demissao.toLocaleDateString()
                    : item.demissao}
                </p>
                <p className={styles.valor_mot}>{item.Valor_Mot_Empresa}</p>
                <p className={styles.total}>
                  {item.valorTotalEmpresa.toFixed(2)}
                </p>
                <p className={styles.empresa}>{item.empresa}</p>
              </div>
            ))}
          </div>
          <button className={styles.custom_button} onClick={exportarParaExcel}>
            Exportar para Excel
          </button>
        </div>
      )}
    </div>
  );
}

export default Planilha;
