// Google Doc id from the document template
// (Get ids from the URL)
var SOURCE_TEMPLATE = "1OoDnBlQcFmfsO0Mb1mdZc31CaUw-lRaQ8YcqBQfUF8A"; // A variável recebe como valor o ID do template

// In which spreadsheet we have all the customer data
var TARGET_SHEET = "1qIUrQV7ZwJLXZURhgCCn756PYIth9_nx3D3YaCPIGyI"; // A variável recebe como valor o ID da planilha

// In which Google Drive we toss the target documents
var TARGET_FOLDER = "1mlMF9DbYGixoRL0AZSLCC_1EnAE7Etxd";

function getIdFromUrl(url) { return url.match(/[-\w]{25,}/); }

function addZero(i) {
  if (i < 10) {
    i = "0" + i;
  }
  return i;
}

function myFunction() {

  var source = DriveApp.getFileById(SOURCE_TEMPLATE); //Apenas lê o arquivo, sem liberar métodos específicos para o Docs
  var sourceSheet = DriveApp.getFileById(TARGET_SHEET); // Apenas lê o arquivo, sem liberar métodos específicos para o Spreadsheets
  var sheet = SpreadsheetApp.openById(sourceSheet.getId()); // Abre a planilha de acordo com o ID recebido
  var eachFile, idToDLET, myFolder, rtrnFromDLET, thisFile;

  myFolder = DriveApp.getFolderById(TARGET_FOLDER);
  thisFile = myFolder.getFiles();
  while (thisFile.hasNext()) 
  {//If there is another element in the iterator
    eachFile = thisFile.next();
    idToDLET = eachFile.getId();
    Logger.log('idToDLET: ' + idToDLET);
    
    //rtrnFromDLET = Drive.Files.remove(idToDLET);
  }
  var data = sheet.getDataRange().getValues(); // A variável recebe os dados presentes na planilha
  for (var j = 1; j<data.length; j++)
  {
    var newFile = source.makeCopy("S2019-2_TCC_"+data[j][2]); //Faz uma cópia do template com o nome do discente
    var targetFolder = DriveApp.getFolderById(TARGET_FOLDER);
    targetFolder.addFile(newFile);
    var doc = SlidesApp.openById(newFile.getId()); //Abre o documento de acordo com o ID recebido
    doc.replaceAllText("#NOME#", data[j][2].toUpperCase()); //Nome
    doc.replaceAllText("#CURSO#", data[j][3]); //Curso
    doc.replaceAllText("#TITULO#", data[j][5].toUpperCase()); //Titulo
    doc.replaceAllText("#ORIENT#", data[j][6]); //Orientador   
    doc.replaceAllText("#BANCA#", data[j][9]); // Banca
    doc.replaceAllText("#DATA#",  data[j][15]+" de Dezembro de 2019"); //Data
    doc.replaceAllText("#HORA#",  data[j][16]+"h"+data[j][17]); //Hora
    doc.replaceAllText("#LOCAL#", data[j][12]); //Local
    var imgSource = DriveApp.getFileById(getIdFromUrl(data[j][13]));
    var slide = doc.getSlides();
    var image = slide[0].getImages()[1];//current image
    image.replace(imgSource,true);
    //var blob = slide[0].getBlob();
    //DriveApp.createFile(blob);
    //var urlExport = "https://docs.google.com/presentation/d/"+newFile.getId()+"/export/png?id="+newFile.getId()+"&slide=id.p1";
    //data[j][15] = urlExport;
        
  }
  // A estrutura de repetição acima é responsável por fazer a mudança de texto dentro do documento de acordo com a planilha
  // Ambos tendo sido especificados anteriormente
}
