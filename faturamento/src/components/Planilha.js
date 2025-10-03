//STYLES
import styles from "./Planilha.module.css";

//React
import React, { useEffect, useState } from "react";
import * as XLSX from "xlsx";

//Scripts
import { exportarParaExcel } from "../scripts/Export";

//Images
import seta from "../images/seta.png";
import logo from "../images/icone_logo.png";

//Dados
import { clientMapping } from "../utils/MapaClientes";
import { ValoresClientes } from "../utils/ValoresCliente";

function Planilha() {
  const [data, setData] = useState([]); // Para armazenar os dados carregados da planilha

  //Define o Operador para calcular quanto pagara para cada um
  const [operadorValue, setOperadorValue] = useState("");
  //Define o calculo do valor por operador
  const [valorOperador, setValorOperador] = useState("");

  //UseState do Select
  const [optionCliente, setOptionCliente] = useState("");

  //HandleChange do Select
  const handleChange = (e) => {
    const value = e.target.value;
    
    setOptionCliente(value);
    console.log(optionCliente)
    console.log(operadorValue)
  };

  //Upload do arquivo
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

  //Define os clientes da Maria e do Herbert
  useEffect(() => {
    if (
      optionCliente === "barcellos" ||
      optionCliente === "dumaszak" ||
      optionCliente === "portolog" ||
      optionCliente === "werner" ||
      optionCliente === "mwm" ||
      optionCliente === "eil" ||
      optionCliente === "gh" ||
      optionCliente === "froes" ||
      optionCliente === "nardi" ||
      optionCliente === "viplog" ||
      optionCliente === "hr" ||
      optionCliente === "aj" ||
      optionCliente === "sudden" ||
      optionCliente === "elo"
    ) {
      setOperadorValue("maria");
    } else if (
      optionCliente === "picoli" || 
      optionCliente === "gbs" ||
      optionCliente === "fraga"
    ) {
      setOperadorValue("marcy");
    }
  }, [optionCliente, operadorValue]);

  //Define o calculo que será feito para valor de operador
  useEffect(() => {
    setValorOperador(0)
    if (operadorValue === "maria") {
      setValorOperador(0)
      setValorOperador(11.25 / 30);
    } else if (operadorValue === "marcy") {
      setValorOperador(0)
      setValorOperador(11.25 / 30);
    } else {
      setValorOperador(0)
      setValorOperador(9.5 / 30);
    }
  }, [operadorValue]);

  //Processa dados para o site
  const processarDados = (data) => {
    return data.map((item) => {
      const empresa = item.Empresa;
      const valCliente = ValoresClientes[empresa] || 52; //Pega o valor de cada cliente, valor padrão 52
      
      return {
        Dias_Trabalhados: item["Dias calculado"] || 0,
        Valor_Mot: 7,
        valorTotal: (7 / 30) * (item["Dias calculado"] || 0),
        nome: item.Nome,
        cpf: formatCPF(item["CPF"]),
        admissao: item.Admissão,
        demissao: item.Demissão === undefined ? "" : item.Demissão,
        empresa: item.Empresa || "Não informado",
        Valor_Mot_Empresa: valCliente,
        valorTotalEmpresa: (valCliente / 30) * (item["Dias calculado"] || 0),
        valorOp: valorOperador * (item["Dias calculado"] || 0),
      };
    });
  };

  const processarDadosFiltrados = () => {
    if (!optionCliente) return processarDados(data); // Retorna todos os dados se nenhuma opção for selecionada

    const clientes = clientMapping[optionCliente];
    if (!clientes) return []; // Retorna vazio se não houver correspondência

    // Verifica se o valor é string ou array e filtra os dados
    return processarDados(data).filter((item) => {
      if (Array.isArray(clientes)) {
        return clientes.some((cliente) => item.empresa.includes(cliente));
      }
      return item.empresa.includes(clientes);
    });
  };

  //Soma Total pagar ao Wagner
  const somaValorW = processarDadosFiltrados(data)
    .reduce((soma, item) => soma + item.valorTotal, 0)
    .toFixed(2);

  //Soma Total cobrar clientes
  const somaValorC = processarDadosFiltrados(data)
    .reduce((soma, item) => soma + item.valorTotalEmpresa, 0)
    .toFixed(2);

  //Soma total pagar ao operador
  const somaValorO = processarDadosFiltrados(data)
    .reduce((soma, item) => soma + item.valorOp, 0)
    .toFixed(2);


  return (
    <div className={styles.container}>
      <img src={logo} alt="Likizoa" className={styles.logo} />
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
            <select
              className={styles.select_cliente}
              name="opcCliente"
              id="opcCliente"
              value={optionCliente}
              onChange={handleChange}
            >
              <option value="" disabled>
                Selecione um cliente
              </option>
              <option value="aj">AJ Transportes</option>
              <option value="barcellos">Barcellos</option>
              <option value="baroncello">Baroncello</option>
              <option value="bc">BC</option>
              <option value="beviani">Beviani</option>
              <option value="betel">Betel</option>
              <option value="comex">Comexcargo</option>
              <option value="comavix">Comavix</option>
              <option value="dumaszak">Dumaszak</option>
              <option value="eil">EIL</option>
              <option value="elo">ELO</option>
              <option value="evandro">Evandro</option>
              <option value="fraga">Fraga Transportes</option>
              <option value="froes">Froes</option>
              <option value="gh">GH</option>
              <option value="gsi">GSI</option>
              <option value="gtl">GTL</option>
              <option value="hr">HR</option>
              <option value="jomar">Jomar</option>
              <option value="lf">LF Cargo</option>
              <option value="mge">MGE</option>
              <option value="mwm">MWM</option>
              <option value="mvrl">MVR LOG</option>
              <option value="mvrt">MVR Transportes</option>
              <option value="wagner">Wagner Transportes</option>
              <option value="primelog">Prime Log</option>
              <option value="itaciba">Itaciba Log</option>
              <option value="nardi">Nardi</option>
              <option value="paganini">Paganini</option>
              <option value="pedrao">Pedrão</option>
              <option value="picoli">Picoli</option>
              <option value="portolog">Portolog</option>
              <option value="portoex">Portoex</option>
              <option value="rtm">RTM</option>
              <option value="saff_fortaleza">Saff Fortaleza</option>
              <option value="saff_porto">Saff Fortaleza Porto</option>
              <option value="saff_navegantes">Saff Navegantes</option>
              <option value="saff_ipojuca">Saff Ipojuca</option>
              <option value="saff_santos">Saff Santos</option>
              <option value="saff_simoes">Saff simoes Filho</option>
              <option value="sanmartino">San Martino</option>
              <option value="santateresinha">Santa Teresinha</option>
              <option value="semfronteiras">Sem Fronteiras</option>
              <option value="simas">Simas</option>
              <option value="smlog">SmLog</option>
              <option value="sudden">Sudden</option>
              <option value="tac">Tac</option>
              <option value="transcosta">Transcosta</option>
              <option value="transmoor">Transmoor</option>
              <option value="vibelog">Vibelog</option>
              <option value="viplog">Viplog</option>
              <option value="eireli">Vip Eireli</option>
              <option value="werner">Werner</option>
              <option value="wl">W & L Transportes</option>
            </select>
            <div className={styles.drop}>
              <img src={seta} alt="Down" />
            </div>

            <button
              className={styles.export_button}
              onClick={() =>
                exportarParaExcel(processarDadosFiltrados(), optionCliente)
              }
            >
              Excel
            </button>
            
          </div>
          <div className={styles.planilha}>
            <div className={styles.header}>
              <p className={styles.dias}>Dias Trabalhados</p>
              <p className={styles.valor_mot}>Valor motorista</p>
              <p className={styles.total}>Valor Total</p>
              <p className={styles.nome}>Nome</p>
              <p className={styles.cpf}>CPF</p>
              <p className={styles.data}>Admissão</p>
              <p className={styles.data}>Demissão</p>
              <p className={styles.valor_mot}>Valor motorista</p>
              <p className={styles.total}>Valor Total</p>
              <p className={styles.empresa}>Empresa</p>
            </div>
            {processarDadosFiltrados().map((item, index) => (
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
        </div>
      )}
    </div>
  );
}
export default Planilha;
