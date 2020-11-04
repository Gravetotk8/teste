
function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu('MENU');
  menu.addItem('CREATE TEMPLATE', 'createTemplate').addToUi();
 
}   

function createTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var me = Session.getEffectiveUser()
//  var protection = ss.protect().setDescription('Nome do Intervalo bloqueado');

  
  
/*
 //apenas 1x
//  SpreadsheetApp.getActiveSpreadsheet().renameActiveSheet("Clientes"); //renomeia a aba
//  ss.insertSheet('Templates'); //cria uma aba
//  ss.insertSheet('Links'); //cria uma aba
//  ss.insertSheet('Forms'); //cria uma aba
//  s.deleteColumns(8, 18); //executar 1x //deletar colunas
//  var protection = ss.getSheetByName("Clientes").getRange('E9').setValue('BLOQUEADA').setFontSize(12).setFontWeight('bold').setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").protect().setDescription('Nome do Intervalo bloqueado');
//  protection.addEditor(me).removeEditors(protection.getEditors());
*/
 
//  ss.getSheetByName("Clientes").getRange('A1:H').deleteCells(SpreadsheetApp.Dimension.ROWS); //limpar linhas
  ss.getSheetByName("Clientes").setFrozenRows(1); //congelar paineis linha
  ss.getSheetByName("Clientes").setFrozenColumns(1); //congelar paineis coluna
  ss.getSheetByName("Clientes").getRange('A1').setValue('NEGRITO').setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center').setFontFamily("arial").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setVerticalAlignment("middle");
  ss.getSheetByName("Clientes").getRange('B1').setValue('ITÁLICO').setFontSize(12).setFontStyle("italic").setHorizontalAlignment('center').setFontFamily("arial").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setVerticalAlignment("middle");
  ss.getSheetByName("Clientes").getRange('C1').setValue('TACHADO').setFontSize(12).setFontLine("line-through").setHorizontalAlignment('center').setFontFamily("arial").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setVerticalAlignment("middle");
  ss.getSheetByName("Clientes").getRange('D1').setValue('SUBLINHADO').setFontSize(12).setFontLine('underline').setHorizontalAlignment('center').setFontFamily("arial").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setVerticalAlignment("middle");
  ss.getSheetByName("Clientes").getRange('E1').setValue('NUMEROS IMPARES').setFontSize(12).setHorizontalAlignment('center').setFontFamily("arial").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setVerticalAlignment("middle");
  ss.getSheetByName("Clientes").getRange('F1').setValue('NUMEROS PARES').setFontSize(12).setHorizontalAlignment('center').setFontFamily("arial").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setVerticalAlignment("middle");
  ss.getSheetByName("Clientes").getRange('G1').setFontColor("#FF0000").setFormula('={"SOMA";ARRAYFORMULA(IF(E2:E6<>"";SUMIF(IF(COLUMN(E2:F);ROW(E2:F));ROW(E2:F);E2:F);""))}').setFontSize(14).setFontFamily("arial").setHorizontalAlignment('center').setFontWeight('bold').setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setVerticalAlignment("middle");
  ss.getSheetByName("Clientes").getRange('B3:C4').setBorder(true, null, null, null, false, false, "blue", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setBorder(null, true, null, null, false, false, "red", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setBorder(null, null, true, null, false, false, "pink", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setBorder(null, null, null, true, false, false, "green", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setBorder(null, null, null, null, true, false, "#46bdc6", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setBorder(null, null, null, null, null, true, "orange", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setVerticalAlignment("middle");
  ss.getSheetByName("Clientes").getRange('G1:H1').mergeAcross().setFontSize(12).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  ss.getSheetByName("Clientes").getRange('G2:H2').mergeAcross().setFontSize(10).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(null, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  ss.getSheetByName("Clientes").getRange('G3:H3').mergeAcross().setFontSize(10).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  ss.getSheetByName("Clientes").getRange('G4:H4').mergeAcross().setFontSize(10).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  ss.getSheetByName("Clientes").getRange('G5:H5').mergeAcross().setFontSize(10).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  ss.getSheetByName("Clientes").getRange('G6:H6').mergeAcross().setFontSize(10).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  ss.getSheetByName("Clientes").getRange('B8').setValue('ESCOLHA').setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment("middle").setFontFamily("arial");
  ss.getSheetByName("Clientes").getRange('B9').setVerticalAlignment("middle").setHorizontalAlignment('left').setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(['SIM', 'NÃO'], true).build());
  ss.getSheetByName("Clientes").getRange('C9').setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  ss.getSheetByName("Clientes").getRange('D9').setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox('APROVADO','REPROVADO').build());
  ss.getSheetByName("Clientes").getRange('B11').setFontSize(12).setFontWeight('bold').setFontFamily("arial").setHorizontalAlignment('center').setFormula('=QUERY(Templates!$A$2:$D$5;"SELECT A, B, C, D LABEL A \'DIA\', B \'HORA\', C \'SIM\', D \'NÃO\'")'); //COLOCANDO FORMULA COM QUERY 
  ss.getSheetByName("Clientes").getRange('B11:E11').setFontSize(12).setFontWeight('bold').setFontFamily("arial");
  ss.getSheetByName("Clientes").getRange('B12:B15').setNumberFormat("dd/mm/yyyy hh:mm:ss");
  ss.getSheetByName("Clientes").getRange('B6').setBackground('#d4dee5');
  
  ss.getSheetByName("Clientes").getRange('E2').setValue('1').setFontSize(10).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(null, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  ss.getSheetByName("Clientes").getRange('E3').setValue('3').setFontSize(10).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  ss.getSheetByName("Clientes").getRange('E4').setValue('5').setFontSize(10).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  ss.getSheetByName("Clientes").getRange('E5').setValue('7').setFontSize(10).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  ss.getSheetByName("Clientes").getRange('E6').setFontSize(10).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  
  ss.getSheetByName("Clientes").getRange('F2').setValue('2').setFontSize(10).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(null, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  ss.getSheetByName("Clientes").getRange('F3').setValue('4').setFontSize(10).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  ss.getSheetByName("Clientes").getRange('F4').setValue('6').setFontSize(10).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  ss.getSheetByName("Clientes").getRange('F5').setValue('8').setFontSize(10).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  ss.getSheetByName("Clientes").getRange('F6').setFontSize(10).setFontFamily("arial").setHorizontalAlignment('center').setVerticalAlignment("middle").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
 
  ss.getSheetByName("Clientes").setRowHeights(1, 100, 35); //tamanho das linhas
  ss.getSheetByName("Clientes").setColumnWidth(1, 115); //tamanho da coluna
  ss.getSheetByName("Clientes").setColumnWidth(2, 130); //tamanho da coluna
  ss.getSheetByName("Clientes").setColumnWidth(3, 115); //tamanho da coluna
  ss.getSheetByName("Clientes").setColumnWidth(4, 115); //tamanho da coluna
  ss.getSheetByName("Clientes").setColumnWidth(5, 170); //tamanho da coluna
  ss.getSheetByName("Clientes").setColumnWidth(6, 150); //tamanho da coluna
  ss.getSheetByName("Clientes").setColumnWidth(7, 60); //tamanho da coluna
  ss.getSheetByName("Clientes").setColumnWidth(8, 60); //tamanho da coluna


  ss.getSheetByName("Templates").getRange('A1').setValue('DIA').setFontSize(12).setFontWeight('bold').setFontFamily("arial");
  ss.getSheetByName("Templates").getRange('B1').setValue('HORA').setFontSize(12).setFontWeight('bold').setFontFamily("arial");
  ss.getSheetByName("Templates").getRange('C1').setValue('SIM').setFontSize(12).setFontWeight('bold').setFontFamily("arial");
  ss.getSheetByName("Templates").getRange('D1').setValue('NÃO').setFontSize(12).setFontWeight('bold').setFontFamily("arial");
  
  ss.getSheetByName("Templates").getRange('B10').setValue('Poesia\n\   Poema\n\      Prosa');
  ss.getSheetByName("Templates").hideColumns(5);
  ss.getSheetByName("Templates").hideRows(8);
  
  

  
/* // CADA ITEM, {LEGENDA}
 
  ss //recebe informação da planilha
  getSheetByName("Clientes") //pega a aba na planilha
  getRange('A1') //qual celula vai colocar a informação
  setFrozenRows(1) //congela linha (numero define a quantidade de linhas)
  setFrozenColumns(1) //congela coluna (numero define a quantidade de colunas)
  setValue('NEGRITO') //define um valor. Ex. coloca um nome em uma celula
  setFontSize(12) //tamanho de fonte (o numero define o tamanho)
  setFontWeight('bold') //negrito
  setFontStyle("italic") //italico
  setFontLine("line-through") //tachado
  setFontLine('underline') //sublinhado
  setHorizontalAlignment('center') //alinhamento do texto na celula ('center') //centralizado, ('right') //a direita, ('left') //a esquerda
  setFontFamily("arial") //fonte do texto o nome fica entre aspas duplas
  setVerticalAlignment("middle") //alinhamento do texto ("top") //topo, ("middle") //centro, ("bottom") // baixo
  setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.DOTTED) //(superior, esquerdo, inferior, direito, vertical, horizontal, cor, estilo) //DOTTED Bordas pontilhadas, DASHED Bordas da linha tracejada, SOLID Bordas de linha sólida finas, SOLID_MEDIUM Bordas de linha sólida média, SOLID_THICK Bordas de linha sólida espessas, DOUBLE Duas bordas de linha sólida.
  setColumnWidths(1, 4, 115); //tamanho das colunas (o primeiro numero é a coluna, o segundo numero quantas colunas, terceiro numero qual o tamanho)
  setColumnWidth(5, 170); //tamanho da coluna (o primeiro numero é a coluna, o segundo numero é o tamanho)
  setRowHeights(1, 100, 35); //tamanho das colunas (o primeiro numero é a coluna, o segundo numero quantas colunas, terceiro numero qual o tamanho)
  setRowHeight(1, 35); //tamanho da coluna (o primeiro numero é a coluna, o segundo numero é o tamanho)
  setBackground('#d4dee5') //cor de preenchimento 
  mergeAcross() //mesclar celulas
  setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(['SIM', 'NÃO'], true).build()) //Validação de dados as opções ficam entre os [] 
  setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build()) //Cria um checkbox 
  deleteCells(SpreadsheetApp.Dimension.ROWS) //deleta o conteudo das celulas em linha 
  setFormula('=QUERY(Entrada!$A$2:$F;"SELECT B, C, D, E, F LABEL B \'MODELO DO TANQUE\', C \'DESCRIÇÃO DO ITEM\', D \'QUANTIDADE\', E \'CONTATO\',F \'CELULAR\'")') //COLOCANDO FORMULA COM QUERY 
  setNumberFormat("dd/mm/yyyy hh:mm:ss") //coloca formato data e hora na celula (Ex.) 
  hideColumns(5) //ocultar a coluna (Ex. coluna 5)
  hideRows(8) //ocultar a linha (Ex. linha 8)
  setValue('Minha\n\n\casa') //coloca um nome em uma celula o \n é uma quebra de linha
  
  //começa aqui o bloqueio da celula
  var protection = ss.protect().setDescription('Nome do Intervalo bloqueado');
  protection.addEditor(me).removeEditors(protection.getEditors());
  //termina aqui 
*/

}
