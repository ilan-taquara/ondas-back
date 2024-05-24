import express from 'express';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import XlsxPopulate from 'xlsx-populate';
import cors from 'cors';

const port = process.env.PORT || 5000;

const app = express();
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
app.use(cors());

app.use(express.json());

app.post('/adicionar', async (req, res) => {
  try {
    const workbook = await XlsxPopulate.fromFileAsync('./cadastro.xlsx');
    const sheet = workbook.sheet(0);
    console.log(req.body);

    const {
      name,
      email,
      telephone,
      childs,
      enrollmentOrReplacement,
      preferTime,
      enrollmentBy,
      voluntary,
    } = req.body;

    if (
      !name ||
      !email ||
      !childs ||
      !enrollmentOrReplacement ||
      !preferTime ||
      !enrollmentBy ||
      !voluntary
    ) {
      throw new Error('Dados inválidos');
    }

    let lastRow = 1;
    while (sheet.cell(lastRow, 1).value()) {
      lastRow++;
    }

    console.log('Quantidade de registro:', lastRow);

    sheet.cell(lastRow, 1).value(name);
    sheet.cell(lastRow, 3).value(email);
    sheet.cell(lastRow, 5).value(telephone);
    sheet.cell(lastRow, 7).value(childs);
    sheet.cell(lastRow, 9).value(enrollmentOrReplacement);
    sheet.cell(lastRow, 11).value(preferTime);
    sheet.cell(lastRow, 13).value(enrollmentBy);
    sheet.cell(lastRow, 15).value(voluntary);

    const range = sheet.range(`A1:C${lastRow}`);
    const definedName = workbook.definedName('_xlnm._FilterDatabase');
    if (definedName) {
      definedName.value(range);
    }

    await workbook.toFileAsync('./cadastro.xlsx');

    res.send('Dados adicionados com sucesso.');
  } catch (error) {
    console.error('Erro ao adicionar dados ao arquivo XLSX:', error);
    res.status(500).send('Erro ao adicionar dados ao arquivo XLSX.');
  }
});

app.get('/download', (req, res) => {
  try {
    res.download('./cadastro.xlsx', 'cadastro.xlsx');
  } catch (error) {
    console.error('Erro ao fazer download do arquivo XLSX:', error);
    res.status(500).send('Erro ao fazer download do arquivo XLSX.');
  }
});

app.post('/addCheckinExcel', async (req, res) => {
  try {
    const workbook = await XlsxPopulate.fromFileAsync('./checkin.xlsx');
    const sheet = workbook.sheet(0);
    console.log(req.body);

    const {
      name,
      amountChilds,
      nameChilds,
      teacher,
      currentDateTime,
      numberOfClass,
    } = req.body;

    if (
      !name ||
      !amountChilds ||
      !nameChilds ||
      !teacher ||
      !currentDateTime ||
      !numberOfClass
    ) {
      throw new Error('Dados inválidos');
    }

    let lastRow = 1;
    while (sheet.cell(lastRow, 1).value()) {
      lastRow++;
    }

    console.log('Quantidade de registro:', lastRow);

    sheet.cell(lastRow, 1).value(name);
    sheet.cell(lastRow, 2).value(currentDateTime);
    sheet.cell(lastRow, 3).value(amountChilds);
    sheet.cell(lastRow, 4).value(nameChilds);
    sheet.cell(lastRow, 5).value(teacher);
    sheet.cell(lastRow, 6).value(numberOfClass);

    const range = sheet.range(`A1:C${lastRow}`);
    const definedName = workbook.definedName('_xlnm._FilterDatabase');
    if (definedName) {
      definedName.value(range);
    }

    await workbook.toFileAsync('./checkin.xlsx');

    const checkinPath = 'checkin.xlsx';
    const chamadaPath = 'chamada.xlsx';

    const checkinWorkbook = await loadWorkbook(checkinPath);
    const chamadaWorkbook = await loadWorkbook(chamadaPath);

    const checkinSheet = checkinWorkbook.sheet(0);
    const chamadaSheet = chamadaworkbook.sheet(0); // Sheet 1 conforme especificação

    const checkinData = getDataFromSheet(checkinSheet);
    const chamadaData = getDataFromSheet(chamadaSheet);

    // Processar dados, ignorando a linha de cabeçalho
    checkinData.slice(1).forEach((checkinRow) => {
      console.log(`Processando: ${checkinRow}`); // Debug para verificar cada linha processada
      updateOrAddData(chamadaSheet, chamadaData, checkinRow);
    });

    await saveWorkbook(chamadaWorkbook, chamadaPath);

    res.send('Dados adicionados com sucesso.');
  } catch (error) {
    console.error('Erro ao adicionar dados ao arquivo XLSX:', error);
    res.status(500).send('Erro ao adicionar dados ao arquivo XLSX.');
  }
});

app.get('/downloadCheckinExcel', (req, res) => {
  try {
    res.download('./checkin.xlsx', 'checkin.xlsx');
  } catch (error) {
    console.error('Erro ao fazer download do arquivo XLSX:', error);
    res.status(500).send('Erro ao fazer download do arquivo XLSX.');
  }
});

app.get('/downloadChamadaExcel', (req, res) => {
  try {
    res.download('./chamada.xlsx', 'chamada.xlsx');
  } catch (error) {
    console.error('Erro ao fazer download do arquivo XLSX:', error);
    res.status(500).send('Erro ao fazer download do arquivo XLSX.');
  }
});

app.get('/dataCheckinExcel', async (req, res) => {
  const workbook = await XlsxPopulate.fromFileAsync('./checkin.xlsx');
  const value = workbook.sheet(0).usedRange().value();
  res.send(value);
});

app.get('/dataChamadaExcel', async (req, res) => {
  const workbook = await XlsxPopulate.fromFileAsync('./chamada.xlsx');
  const value = workbook.sheet(0).usedRange().value();
  console.log(value);
  res.send(value);
});

async function loadWorkbook(filePath) {
  return await XlsxPopulate.fromFileAsync(filePath);
}

async function saveWorkbook(workbook, filePath) {
  await workbook.toFileAsync(filePath);
}

function getDataFromSheet(sheet) {
  return sheet.usedRange().value();
}

function normalizeString(str) {
  if (typeof str !== 'string') return '';
  return str.trim().toLowerCase().replace(/\s+/g, ' ');
}

function findLastRow(sheet) {
  let lastRow = sheet.usedRange().endCell().rowNumber();
  while (lastRow > 0 && !sheet.cell(lastRow, 2).value()) {
    lastRow--;
  }
  return lastRow + 1;
}

function updateOrAddData(sheet, sheetData, checkinRow) {
  const [nomeCompleto, dataHora, filhoDeMenor, nomeDosFilhos, professor, aula] =
    checkinRow;
  const normalizedCheckinName = normalizeString(nomeCompleto);

  const rowIndex = sheetData.findIndex(
    (row) => normalizeString(row[1]) === normalizedCheckinName,
  );
  if (rowIndex > -1) {
    // Atualizar dados existentes
    console.log(`Atualizando dados na linha: ${rowIndex + 1} ${aula}`); // Debug
    if (!sheetData[rowIndex][8]) {
      sheet.cell(rowIndex + 1, 9).value(filhoDeMenor); // Filhos (Coluna I)
    }
    if (!sheetData[rowIndex[3]]) {
      if (+dataHora.split(',')[1].slice(1, 3) < 12) {
        sheet.cell(rowIndex + 1, 4).value('09h');
        sheet.cell(rowIndex + 1, 16).value('MANHA');
      } else {
        sheet.cell(rowIndex + 1, 4).value('18h');
        sheet.cell(rowIndex + 1, 16).value('NOITE');
      }
    }
    if (aula == 'Aula 01' && !sheetData[rowIndex][6]) {
      sheet.cell(rowIndex + 1, 7).value(professor);
      sheet.cell(rowIndex + 1, 6).value(dataHora.split(',')[0]);
    }
    if (aula == 'Aula 02' && !sheetData[rowIndex][10]) {
      sheet.cell(rowIndex + 1, 11).value(professor);
      sheet.cell(rowIndex + 1, 10).value(dataHora.split(',')[0]);
    }
    if (aula == 'Aula 03' && !sheetData[rowIndex][12]) {
      sheet.cell(rowIndex + 1, 13).value(professor);
      sheet.cell(rowIndex + 1, 12).value(dataHora.split(',')[0]);
    }
    if (aula == 'Aula 04' && !sheetData[rowIndex][15]) {
      sheet.cell(rowIndex + 1, 15).value(dataHora.split(',')[0]);
    }
  } else {
    // Adicionar nova linha na última linha disponível
    let lastRow = findLastRow(sheet);
    console.log(`Adicionando na linha: ${lastRow}`); // Debug para verificar a última linha
    sheet.cell(lastRow, 2).value(nomeCompleto); // Nome (Coluna B)
    sheet.cell(lastRow, 9).value(filhoDeMenor); // Filhos (Coluna I)
    if (+dataHora.split(',')[1].slice(1, 3) < 12) {
      sheet.cell(lastRow, 4).value('09h');
    } else {
      sheet.cell(lastRow, 4).value('18h');
    }
    if (aula == 'Aula 01') {
      sheet.cell(lastRow, 7).value(professor); // Filhos (Coluna I)
      sheet.cell(lastRow, 6).value(dataHora.split(',')[0]);
    }
    if (aula == 'Aula 02') {
      sheet.cell(lastRow, 11).value(professor); // Filhos (Coluna I)
      sheet.cell(lastRow, 10).value(dataHora.split(',')[0]);
    }
    if (aula == 'Aula 03') {
      sheet.cell(lastRow, 13).value(professor); // Filhos (Coluna I)
      sheet.cell(lastRow, 12).value(dataHora.split(',')[0]);
    }
    if (aula == 'Aula 04') {
      sheet.cell(lastRow, 15).value(dataHora.split(',')[0]);
    }
  }
}

app.get('/updateChamadaExcel', async (req, res) => {
  try {
    const checkinPath = 'checkin.xlsx';
    const chamadaPath = 'chamada.xlsx';

    const checkinWorkbook = await loadWorkbook(checkinPath);
    const chamadaWorkbook = await loadWorkbook(chamadaPath);

    const checkinSheet = checkinWorkbook.sheet(0);
    const chamadaSheet = chamadaworkbook.sheet(0); // Sheet 1 conforme especificação

    const checkinData = getDataFromSheet(checkinSheet);
    const chamadaData = getDataFromSheet(chamadaSheet);

    // Processar dados, ignorando a linha de cabeçalho
    checkinData.slice(1).forEach((checkinRow) => {
      console.log(`Processando: ${checkinRow}`); // Debug para verificar cada linha processada
      updateOrAddData(chamadaSheet, chamadaData, checkinRow);
    });

    await saveWorkbook(chamadaWorkbook, chamadaPath);

    res.status(200).send('Planilha de chamada atualizada com sucesso.');
  } catch (err) {
    console.error(err);
    res.status(500).send('Erro ao atualizar a planilha de chamada.');
  }
});

// Função para encontrar a última linha preenchida na coluna especificada
const findLastFilledRow = (sheet, col) => {
  let rowNumber = 1; // Começa na linha 1 (a primeira linha no Excel)
  while (
    sheet.cell(`${col}${rowNumber}`).value() !== undefined &&
    sheet.cell(`${col}${rowNumber}`).value() !== null &&
    sheet.cell(`${col}${rowNumber}`).value() !== ''
  ) {
    rowNumber++;
  }
  return rowNumber - 1; // Retorna o número da última linha preenchida
};

app.post('/updateChamadaExcelFromRegisterForm', async (req, res) => {
  try {
    const workbook = await XlsxPopulate.fromFileAsync('./chamada.xlsx');
    const sheet = workbook.sheet(0); // Mudando para a segunda aba
    const {
      name,
      contact,
      hour,
      classOneDate,
      classOneTeacher,
      token,
      childs,
      classTwoDate,
      classTwoTeacher,
      classThreeDate,
      classThreeTeacher,
      fitToServe,
      classFourDate,
      shift,
      conclude,
      ministry,
      personality,
      gifts,
      birthday,
      birthyear,
      age,
      acceptJesus,
      whenAcceptJesus,
      whereAcceptJesus,
      address,
      howMeetIlan,
      introducedIlan,
      baptized,
      whenBaptized,
      whereBaptized,
      wantsBaptize,
      meetIn,
      meetedIn,
      ministryTwo,
      ministryThree,
      maritalStatus,
      nameSpouse,
      nameChilds,
      memberIlan
    } = req.body;

    const data = sheet.usedRange().value();
    const normalizedCheckinName = name.trim();

    let rowIndex = data.findIndex(
      (row) => row[1] && row[1].trim() === normalizedCheckinName, // Supondo que a coluna B contém os nomes
    );

    if (rowIndex === -1) {
      // Se o nome não for encontrado, adicionar nova linha
      rowIndex = findLastFilledRow(sheet, 'B');
      sheet.cell(`B${rowIndex + 1}`).value(name); // Adicionar o nome na nova linha
      console.log('Adicionando na linha: ', rowIndex);
    }

    console.log('Atualizando na linha: ', rowIndex + 1); // +1 porque a linha no Excel começa em 1

    // Função para atualizar ou manter valor existente
    const updateCell = (col, value) => {
      const cell = sheet.cell(`${col}${rowIndex + 1}`);
      if (value !== undefined && value !== null) {
        cell.value(value);
      } else if (cell.value() === null) {
        cell.value(''); // Manter a célula vazia se o valor existente for nulo e não for fornecido um novo valor
      }
    };

    // Atualizar ou manter os valores na linha encontrada ou nova linha
    if (contact) {
      updateCell('C', contact);
    }
    if (hour) {
      updateCell('D', hour);
    }
    if (classOneDate) {
      updateCell('F', classOneDate);
    }
    if (classOneTeacher) {
      updateCell('G', classOneTeacher);
    }
    if (token) {
      updateCell('H', token);
    }
    if (childs) {
      updateCell('I', childs);
    }
    if (classTwoDate) {
      updateCell('J', classTwoDate);
    }
    if (classTwoTeacher) {
      updateCell('K', classTwoTeacher);
    }
    if (classThreeDate) {
      updateCell('L', classThreeDate);
    }
    if (classThreeTeacher) {
      updateCell('M', classThreeTeacher);
    }

    if (fitToServe) {
      updateCell('N', fitToServe);
    }
    if (classFourDate) {
      updateCell('O', classFourDate);
    }
    if (shift) {
      updateCell('P', shift);
    }
    if (conclude) {
      updateCell('Q', conclude);
    }
    if (ministry) {
      updateCell('R', ministry);
    }
    if (personality) {
      updateCell('S', personality);
    }
    if (gifts) {
      updateCell('T', gifts);
    }
    if (birthday) {
      updateCell('U', birthday);
    }
    if (birthyear) {
      updateCell('V', birthyear);
    }
    if (age) {
      updateCell('W', age);
    }
    if (acceptJesus) {
      updateCell('X', acceptJesus);
    }
    if (whenAcceptJesus) {
      updateCell('Y', whenAcceptJesus);
    }
    if (whereAcceptJesus) {
      updateCell('Z', whereAcceptJesus);
    }
    if (address) {
      updateCell('AA', address);
    }
    if (howMeetIlan) {
      updateCell('AB', howMeetIlan);
    }
    if (introducedIlan) {
      updateCell('AC', introducedIlan);
    }
    if (baptized) {
      updateCell('AD', baptized);
    }
    if (whenBaptized) {
      updateCell('AE', whenBaptized);
    }
    if (whereBaptized) {
      updateCell('AF', whereBaptized);
    }
    if (wantsBaptize) {
      updateCell('AG', wantsBaptize);
    }
    if (meetIn) {
      updateCell('AH', meetIn);
    }
    if (meetedIn) {
      updateCell('AI', meetedIn);
    }
    if (ministryTwo) {
      updateCell('AJ', ministryTwo);
    }    
    if (ministryThree) {
      updateCell('AK', ministryThree);
    }    
    if (maritalStatus) {
      updateCell('AL', maritalStatus);
    }        
    if (nameSpouse) {
      updateCell('AM', nameSpouse);
    }        
    if (nameChilds) {
      updateCell('AN', nameChilds);
    }    
    if (memberIlan) {
      updateCell('AO', memberIlan);
    }        

    await workbook.toFileAsync('./chamada.xlsx');
    res.status(200).send('Atualização concluída com sucesso');
  } catch (error) {
    console.log(error);
    res.status(500).send('Erro ao atualizar o arquivo');
  }
});

// Coloco o nome no formulario RegisterForm e me retorna os dados do aluno de todas colunas
app.post('/dataChamadaExcelFromRegisterForm', async (req, res) => {
  try {
    const workbook = await XlsxPopulate.fromFileAsync('./chamada.xlsx');
    const sheet = workbook.sheet(0); // Mudando para a segunda aba
    const { name } = req.body;

    const data = sheet.usedRange().value();
    const normalizedCheckinName = name.trim();
    console.log('Nome sendo procurado por: ', normalizedCheckinName);

    const rowIndex = data.findIndex(
      (row) => row[1] && row[1].trim() === normalizedCheckinName, // Supondo que a coluna B contém os nomes
    );

    if (rowIndex > -1) {
      console.log('Dados encontrados na linha: ', rowIndex + 1); // +1 porque a linha no Excel começa em 1

      // Obter valores na linha encontrada
      const rowData = data[rowIndex];
      console.log('Dados da linha:', rowData);

      // Obter nome das colunas na linha 3
      const columnNames = sheet.range('B3:AI3').value()[0];

      // Retornar os dados da linha como resposta
      res.status(200).json({
        row: rowIndex + 1,
        data: rowData,
        columnNames: columnNames,
      });
    } else {
      console.log('Aluno não encontrado');
      res.status(404).send('Aluno não encontrado');
    }
  } catch (error) {
    console.log(error);
    res.status(500).send('Erro ao buscar dados no arquivo');
  }
});

app.post('/dataChamadaExcelFromRegisterFormInitial', async (req, res) => {
  try {
    const workbook = await XlsxPopulate.fromFileAsync('./chamada.xlsx');
    const sheet = workbook.sheet(0); // Mudando para a segunda aba
    const { name } = req.body;

    const data = sheet.usedRange().value();
    const normalizedInitial = name.trim().toLowerCase();

    const matchedRows = data.filter(
      (row) => row[1] && row[1].toLowerCase().startsWith(normalizedInitial),
    );

    if (matchedRows.length > 0) {
      console.log('Dados encontrados para a inicial:', name);

      // Obter nome das colunas na linha 3
      const columnNames = sheet.range('B3:AO3').value()[0];

      // Mapear os resultados, extraindo apenas valores de texto
      const results = matchedRows.map((row) => {
        const rowData = row.map((cell) =>
          cell && cell.text ? cell.text() : cell,
        );
        return {
          row: data.indexOf(row) + 1, // +1 porque a linha no Excel começa em 1
          data: rowData,
        };
      });

      // Retornar os dados como resposta
      res.status(200).json({
        columnNames: columnNames,
        results: results,
      });
    } else {
      console.log('Nenhum aluno encontrado com a inicial:', name);
      res.status(404).send('Nenhum aluno encontrado com a inicial fornecida');
    }
  } catch (error) {
    console.log(error);
    res.status(500).send('Erro ao buscar dados no arquivo');
  }
});

app.get('/names', async (req, res) => {
  try {
    const workbook = await XlsxPopulate.fromFileAsync('./chamada.xlsx');
    const sheet = workbook.sheet(0);

    // Obter todas as linhas usadas na planilha
    const usedRange = sheet.usedRange();
    const rows = usedRange.value();

    // Obter valores da coluna B, ignorando o cabeçalho (primeira linha) e células vazias
    const values = rows
      .slice(1)
      .map((row) => row[1])
      .filter((value) => value); // Coluna B é a segunda coluna (índice 1)

    // Mapear os valores para o formato desejado
    const formattedValues = values.map((name) => ({ label: name }));

    res.json(formattedValues);
  } catch (error) {
    console.error(error);
    res.status(500).send('Error reading the Excel file');
  }
});

app.post('/certificate', async (req, res) => {
  const { month, year } = req.body;
  const monthValue = month.value.slice(4, 6);
  const yearValue = year.value;

  const workbook = await XlsxPopulate.fromFileAsync('./chamada.xlsx');
  const sheet = workbook.sheet(0);

  // Obter todas as linhas usadas na planilha
  const usedRange = sheet.usedRange();
  const rows = usedRange.value().slice(1);

  const filtered = rows.filter((row) => {
    return row[5] && row[9] && row[11] && row[14];
  });

  res.send(filtered);
});

app.post('/addMembershipToChamadaExcel', async (req, res) => {
  try {
    const workbook = await XlsxPopulate.fromFileAsync('./chamada.xlsx');
    const sheet = workbook.sheet(0); // Mudando para a segunda aba
    const {
      name,
      dateOfBirth,
      telephone,
      address,
      maritalStatus,
      nameSpouse,
      nameChilds,
      acceptedJesus,
      whenAcceptedJesus,
      whereAcceptedJesus,
      baptized,
      whenBaptized,
      whereBaptized,
      completedClassOne,
      whenCompletedClassOne,
      memberIlan,
    } = req.body;

    const data = sheet.usedRange().value();
    const normalizedCheckinName = name.trim();

    let rowIndex = data.findIndex(
      (row) => row[1] && row[1].trim() === normalizedCheckinName, // Supondo que a coluna B contém os nomes
    );

    if (rowIndex === -1) {
      // Se o nome não for encontrado, adicionar nova linha
      rowIndex = findLastFilledRow(sheet, 'B');
      sheet.cell(`B${rowIndex + 1}`).value(name); // Adicionar o nome na nova linha
      console.log('Adicionando na linha: ', rowIndex);
    }

    console.log('Atualizando na linha: ', rowIndex + 1); // +1 porque a linha no Excel começa em 1

    // Função para atualizar ou manter valor existente
    const updateCell = (col, value) => {
      const cell = sheet.cell(`${col}${rowIndex + 1}`);
      if (value !== undefined && value !== null) {
        cell.value(value);
      } else if (cell.value() === null) {
        cell.value(''); // Manter a célula vazia se o valor existente for nulo e não for fornecido um novo valor
      }
    };

    // Atualizar ou manter os valores na linha encontrada ou nova linha
    if (telephone) {
      updateCell('C', telephone);
    }
    if (whenCompletedClassOne) {
      updateCell('F', whenCompletedClassOne);
    }
    if (dateOfBirth) {
      updateCell('U', dateOfBirth);
    }
    if (acceptedJesus) {
      updateCell('X', acceptedJesus);
    }
    if (whenAcceptedJesus) {
      updateCell('Y', whenAcceptedJesus);
    }
    if (whereAcceptedJesus) {
      updateCell('Z', whereAcceptedJesus);
    }
    if (address) {
      updateCell('AA', address);
    }
    if (baptized) {
      updateCell('AD', baptized);
    }
    if (whenBaptized) {
      updateCell('AE', whenBaptized);
    }
    if (whereBaptized) {
      updateCell('AF', whereBaptized);
    }
    if (maritalStatus) {
      updateCell('AL', maritalStatus);
    }
    if (nameSpouse) {
      updateCell('AM', nameSpouse);
    }
    if (nameChilds) {
      updateCell('AN', nameChilds);
    }
    if (memberIlan) {
      updateCell('AO', memberIlan);
    }

    await workbook.toFileAsync('./chamada.xlsx');
    res.status(200).send('Atualização concluída com sucesso');
  } catch (error) {
    console.log(error);
    res.status(500).send('Erro ao atualizar o arquivo');
  }
});

app.get('/', (req, res) => {
  res.send('PAssou');
});

app.listen(port, () => {
  console.log('Servidor iniciado.');
});
