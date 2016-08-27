var fs = require('fs');
var Docxtemplater = require('docxtemplater');
var ChartModule = require("./../js/index");
var chartModule = new ChartModule();

var OUTDIR = __dirname + "/tmp/";
var TGTDIR = __dirname + '/template/';


var testArray = [
    // {
    //     fileName: "test-ppt-chart-1-pie.pptx",
    //     data: {
    //         chart1: {
    //             variables: {
    //                 chart1: "Pie Chart"
    //             },
    //             headerRowName: ["Ireland", "Germany"],
    //             rowData: [
    //                 {
    //                     rowName: 'Oil',
    //                     data: [5, 5]
    //                 }
    //             ]
    //         }

    //     }
    // },
    {
        fileName: "test-ppt-chart-2-column.pptx",
        data: {
            chart1: {
                variables: {
                    chart1: "column Chart"
                },
                headerRowName: ["2006", "2007", "2008", "2009", "2010"],
                rowData: [
                    {
                        color: "#ff0000",
                        rowName: 'Income',
                        data: [500, 400, 300, 200, 100]
                    },
                    {
                        color: "#00ff00",
                        rowName: 'Expense',
                        data: [500, 500, 500, 500, 500]
                    }
                ]
            }
        }
    },
    // {
    //     fileName: "test-ppt-chart-3-bar.pptx",
    //     data: {
    //         chart1: {
    //             variables: {
    //                 chart1: "bar Chart"
    //             },
    //             headerRowName: ["europe", "namerica", "asia", "lamerica", "meast", "africa"],
    //             rowData: [
    //                 {
    //                     rowName: 'Y2003',
    //                     data: [5, 5, 5, 5, 5, 5]
    //                 },
    //                 {
    //                     rowName: 'Y2004',
    //                     data: [5, 4, 3, 2, 1, 0]
    //                 },
    //                 {
    //                     rowName: 'Y2005',
    //                     data: [5, 6, 7, 8, 9, 10]
    //                 }
    //             ]
    //         }
    //     }
    // }
];

var generateChart = function generateChart(fileName, data) {
    var content = fs
        .readFileSync(TGTDIR + "/" + fileName, "binary")
    var docx = new Docxtemplater()
        .attachModule(chartModule)
        .load(content)
        .setData({
            charts: data
        })
        .render();


    var buffer = docx
        .getZip()
        .generate({ type: "nodebuffer" });

    fs.writeFile(OUTDIR + "/" + fileName, buffer);
}

testArray.forEach(function (chartInfo, index) {
    generateChart(chartInfo.fileName, chartInfo.data);
});



