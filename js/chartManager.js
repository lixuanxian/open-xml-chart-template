var ChartManager, JSZip, DOMParser;

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
    var chartSerXmls = chartXml.match(/<c:ser>.+<\/c:ser>/g);
    var chartData = null;

    for (var tempKey in this.chartsData) {

      if (new RegExp("{" + tempKey + "}").test(chartXml)) {
        chartData = this.chartsData[tempKey];
        break;
      }
    }

    if (!chartData) {
      return;
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

          //row 1  header
          var catXml = '<c:cat> <c:strRef> <c:f>Sheet1!$A$2' +
            (chartData.headerRowName && chartData.headerRowName.length <= 1 ? "" : ":$" + String.fromCharCode(65 + chartData.headerRowName.length) + "$2")
            + '</c:f><c:strCache> <c:ptCount val="' + chartData.headerRowName.length + '"/>';

          tempSheetDataXml += '<row r="1"> ';
          for (var tmpIndex in chartData.headerRowName) {
            tmpIndex = parseInt(tmpIndex);
            var isEmpty = !chartData.headerRowName[tmpIndex] || chartData.headerRowName[tmpIndex] == "";

            if (!isEmpty && stringsArray.indexOf(chartData.headerRowName[tmpIndex]) < 0) {
              stringsArray.push(chartData.headerRowName[tmpIndex]);
            }

            if (!isEmpty) {
              tempSheetDataXml += '<c r="' + (String.fromCharCode(65 + tmpIndex + 1)) + '1" t="s">  <v>' +
                (stringsArray.indexOf(chartData.headerRowName[tmpIndex]) + 1)
                + '</v> </c>';
            }
            catXml += '<c:pt idx="' + tmpIndex + '"> <c:v>' +
              (isEmpty ? "" : chartData.headerRowName[tmpIndex])
              + '</c:v> </c:pt>';
          }
          catXml += "</c:strCache></c:strRef></c:cat>";

          tempSheetDataXml += "</row>";

          //row 2 ~ n data row 
          for (var tmpIndex in chartData.rowData) {
            tmpIndex = parseInt(tmpIndex);
            var tempRowData = chartData.rowData[tmpIndex];
            var tempSerXml = chartSerXmls[tmpIndex];
            if (!tempSerXml) {
              tempSerXml = chartSerXmls[0];
              tempSerXml = tempSerXml.replace('<c:idx val="0"/>', '<c:idx val="' + tmpIndex + '"/>');
              tempSerXml = tempSerXml.replace('<c:order val="0"/>', '<c:order val="' + tmpIndex + '"/>');
              chartSerXmls[tmpIndex] = tempSerXml;
            }
            var rowMark = (String.fromCharCode(65 + tempRowData.data.length));
            var tempValXml = '<c:val><c:numRef><c:f>Sheet1!$B$' + (tmpIndex + 2) + (tempRowData.data.length <= 1 ? "" : "$" + rowMark + "$" + (tmpIndex + 2))
              + '</c:f><c:numCache><c:ptCount val="' + tempRowData.data.length + '"/>';

            tempSheetDataXml += '<row r="' + (tmpIndex + 2) + '">';

            if (tempRowData.rowName && stringsArray.indexOf(tempRowData.rowName) < 0) {
              stringsArray.push(tempRowData.rowName);
            }

            if (tempRowData.rowName) {
              tempSheetDataXml += '<c r="A' + (tmpIndex + 2) + '" t="s"><v>' +
                (stringsArray.indexOf(tempRowData.rowName) + 1)
                + '</v></c>';
              var tempSerTitleXml = '<c:tx><c:strRef><c:f>Sheet1!$A$' + (tmpIndex + 2) + '</c:f> <c:strCache><c:ptCount val="1"/><c:pt idx="0">  <c:v>' +
                tempRowData.rowName
                + '</c:v></c:pt></c:strCache></c:strRef></c:tx>';

              tempSerXml = tempSerXml.replace(/<c:tx>.+<\/c:tx>/g, tempSerTitleXml);
            }

            //only support 26 cols
            tempRowData.data.forEach(function (element, index) {
              tempSheetDataXml += ' <c r="' + (String.fromCharCode(65 + index + 1) + "" + (tmpIndex + 2)) + '">  <v>' + element + '</v></c>';
              tempValXml += '<c:pt idx="' + index + '">  <c:v>' + element + '</c:v>  </c:pt>';
            }, this);

            tempSheetDataXml += "</row>";
            tempValXml += '</c:numCache> </c:numRef>  </c:val>';

            tempSerXml = tempSerXml.replace(/<c:cat>.+<\/c:cat>/g, catXml);
            tempSerXml = tempSerXml.replace(/<c:val>.+<\/c:val>/g, tempValXml);

            //replace color  
            tempSerXml = tempSerXml.replace(/<a:solidFill>.+<\/a:solidFill>/g,
              ' <c:spPr><a:solidFill><a:srgbClr val="' + (tempRowData.color ? tempRowData.color.replace("#", "") : this.getColor(tmpIndex)) + '"/> </a:solidFill></c:spPr>'
            );

            chartSerXmls[tmpIndex] = tempSerXml;
          }
          tempSheetDataXml += "</sheetData>";

          var tempStringXmls = "";
          stringsArray.forEach(function (element) {
            tempStringXmls += "<si><t>" + element + "</t></si>"
          }, this);

          sheetXml = sheetXml.replace(/<sheetData>.+<\/sheetData>/g, tempSheetDataXml);

          sharedStringsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + (stringsArray.length + 1) + '" uniqueCount="' + (stringsArray.length + 1) + '"><si/>'
            + tempStringXmls +
            '</sst>';


          excelZip.file(sheetXmls[0].name, sheetXml);
          excelZip.file("xl/sharedStrings.xml", sharedStringsXml);
        }

        chartXml = chartXml.replace(new RegExp('<c:ser>.+<\/c:ser>'), "");

        chartSerXmls.forEach(function (element, index) {
          chartXml = chartXml.replace(/<\/c:\w+Chart>.+<\/c:plotArea>/g, element + "$&");
        }, this);

        if (chartData.variables) {
          for (var tmpKey in chartData.variables){
            chartXml = chartXml.replace("{" + tmpKey + "}", chartData.variables[tempKey]);
          }
        }

       chartXml = chartXml.replace(/{[\w\d]+}/,"");

        this.zip.file(filename, chartXml);
        this.zip.file(embeddingsXLSX, excelZip.generate({ type: "nodeBuffer" }));
        //start replace strings
      }
    }

  }


  /**
  	 * load relationships
  	 * @return {ChartManager} for chaining
   */
  ChartManager.prototype.getColor = function (index) {
    var SoftColor = ["66", "99", "cc", "ff"];
    index = index % 96;
    return SoftColor[Math.floor(index / 8) % 4] + SoftColor[index % 4] + SoftColor[Math.floor(index / 4) % 4];
  };

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
