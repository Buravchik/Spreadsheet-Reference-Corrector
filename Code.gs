//в формулах не должно быть восклицательных знаков и апострофов, иначе некорректно отделяется имя листа.
//Возвращает объект, в котором названия полей берутся из списка листов, на которые в формулах ссылаются ячейки, а в качестве значений полей - массив объектов с координатами ячеек с этими формулами
//{Январь 2016=[{str=1, row=2}, {str=1, row=3}], Февраль 2016=[{str=1, row=3}]}
function getTemplateReferences(sheet) {
  sheet = sheet || SpreadsheetApp.getActive().getSheetByName("Шаблон");
  return sheet.getDataRange().getFormulas().reduce(function(acc,str,strNum){//переберем строки, результат будем накапливать в массиве-аккумуляторе acc
    str.forEach(function(cellFun,rowNum){//переберем ячейки в строках
      if (cellFun && cellFun.indexOf("!") > -1) {//найдем те, в которых используются ссылки на другие листы (там имя листа отделяется от координат знаком '!')
        var splitted = cellFun.split("!");//ссылок в одной формуле может быть несколько.
        splitted.pop();//последний элемент ссылки на лист не содержит
        splitted.forEach(function(ref){//переберем ссылки на другие листы (точнее, ту часть, где указывается имя листа)
          var firstApostrof = ref.indexOf("'");
          var secondApostrof = ref.indexOf("'", firstApostrof + 1);
          var sheetRef = ref.substring(firstApostrof + 1, secondApostrof);//имя листа заключено между двумя апострофами
          var extSS = "";
          var importRangeLocation = ref.indexOf("IMPORTRANGE")+13;
          if (importRangeLocation != -1+13) {
            extSS = ref.substring(importRangeLocation,ref.indexOf('"',importRangeLocation+1));
            var sheetNameBeg = ref.indexOf('","',ref.indexOf('"',importRangeLocation+extSS.length))+3;
            sheetRef = ref.substring(sheetNameBeg);//название листа будет до конца ref
          }
          if (!acc.hasOwnProperty(sheetRef)) acc[sheetRef] = [];
          acc[sheetRef].push({str:strNum,row:rowNum});          
          if (extSS) {//если это ссылка importRange, то добавим к свежедобавленному элементу массива еще пару полей с доп. информацией
            acc[sheetRef][acc[sheetRef].length-1].extSS = extSS;
            acc[sheetRef][acc[sheetRef].length-1].sheet = sheetRef;  //можно закомментить, возможно
          }
        });
      }
    });
    return acc;
  },{});
}
function test(){
  Logger.log(getTemplateReferences());
}
