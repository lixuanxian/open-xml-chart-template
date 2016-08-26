var ChartManager, DocUtils, JSZip, DOMParser;

DocUtils = require('./docUtils');
JSZip = require('docxtemplater').JSZip;
DOMParser = require('xmldom').DOMParser;

module.exports = ChartManager = (function () {
  function ChartManager(zip, data) {
    this.zip = zip;
    this.chartsData = data;
    if (this.chartsData) {
      this.chartsTemplates = this.zip.file(/(ppt|word)\/charts\/(chart)\d+\.xml/).map(function (file) {
        return file.name;
      });

      this.chartsTemplates.forEach(function (chartXmlFile) {
        this.replaceTag(chartXmlFile);
      }, this);
    }



  }


  ChartManager.prototype.replaceTag = function (filename) {

    var chartXml = this.zip.file(filename).asText();
    var chartData = null;

    for (var tempKey in this.chartsData) {

      if (new RegExp("{" + tempKey + "}").test(chartXml)) {
        chartData = this.chartsData[tempKey];
        break;
      }
    }

    var relsFile = this.zip.file(filename.replace(/(chart)\d+\.xml/, "_rels/$&.rels"));
    var embeddingsXLSX = null;
    if (relsFile) {
      var matchRelsFiles = relsFile.asText().match(/Target=\"..\/([\w\d_\/]+\.xlsx)/);
      if (matchRelsFiles && matchRelsFiles.length == 2) {
        var embeddingsXLSX = filename.replace(/charts\/(chart)\d+\.xml/, "") + matchRelsFiles[1];

        var excelZip = new JSZip(this.zip.file(embeddingsXLSX).asBinary());
        var sharedStringsXml = excelZip.file("xl/sharedStrings.xml").asText();
        var sheetXmls = excelZip.file(/xl\/worksheets\/sheet\d+\.xml/);
        var sheetXml = null;
        if (sheetXmls && sheetXmls.length >= 0) {
          sheetXml = sheetXmls[0].asText();
          var stringsArray = sharedStringsXml.match(/([\w\}\{_-]+)(?=<\/t>)/g);
          var tempSheetDataXml = "<sheetData>";

          tempSheetDataXml += '<row r="1">';
          for (var tmpIndex in chartData.colNames) {
            tmpIndex = parseInt(tmpIndex);

            if (stringsArray.indexOf(chartData.colNames[tmpIndex]) < 0) {
              stringsArray.push(chartData.colNames[tmpIndex]);
            }
            tempSheetDataXml += '<c r="' + (String.fromCharCode(65 + tmpIndex)) + '1" t="s">  <v>' +
              (!chartData.colNames[tmpIndex] ? 1 : stringsArray.indexOf(chartData.colNames[tmpIndex]) + 2)
              + '</v> </c>';
          }
          tempSheetDataXml += "</row>";


          for (var tmpIndex in chartData.colData) {
            tmpIndex = parseInt(tmpIndex);
            var tempColData = chartData.colData[tmpIndex];
            tempSheetDataXml += '<row r="' + (tmpIndex + 2) + '">';

            if (tempColData.rowName && stringsArray.indexOf(tempColData.rowName) < 0) {
              stringsArray.push(tempColData.rowName);
            }

            tempSheetDataXml += '<c r="A' + (tmpIndex + 2) + '" t="s">  <v>' +
              (!tempColData.rowName ? 1 : stringsArray.indexOf(tempColData.rowName) + 2)
              + '</v> </c>';

            //only support 26 cols
            tempColData.data.forEach(function (element, index) {
              tempSheetDataXml += ' <c r="' + (String.fromCharCode(65 + index + 1)) + "" + (tmpIndex + 2) + '">  <v>' + element + '</v>  </c> ';
            }, this);

            tempSheetDataXml += "</row>";
          }
          tempSheetDataXml += "</sheetData>";

          var tempStringXmls = "";
          stringsArray.forEach(function (element) {
            tempStringXmls += "<si><t>" + element + "</t></si>"
          }, this);

          sheetXml = sheetXml.replace(/<sheetData>[\w\d\W{}}=]+<\/sheetData>/g, tempSheetDataXml);

          sharedStringsXml = sharedStringsXml.replace(/uniqueCount\=\"\d+\"/g, "uniqueCount=\"" + (1 + stringsArray.length) + "\"");
          sharedStringsXml = sharedStringsXml.replace(/<si\/>[\w\d\W{}}=]+<\/sst>/, "<si/>" + tempStringXmls + "</sst>");

          excelZip.file(sheetXmls[0].name, sheetXml);
          excelZip.file("xl/sharedStrings.xml", sharedStringsXml);
        }


        this.zip.file(embeddingsXLSX, excelZip.generate({ type: "nodeBuffer" }));
        //start replace strings
      }
    }

  }

  /**
  	 * load relationships
  	 * @return {ChartManager} for chaining
   */
  ChartManager.prototype.rendered = function () {

    if (this.data)

      this.get
    return this;
  };


  return ChartManager;

})();
