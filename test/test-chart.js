var fs = require('fs');
var Docxtemplater = require('docxtemplater');
var ChartModule = require("./../js/index");
var chartModule = new ChartModule();

var OUTDIR = __dirname + "/tmp/";
var TGTDIR = __dirname + '/template/';


var content = fs
    .readFileSync(TGTDIR + "/test-ppt-chart-1-pie.pptx", "binary")
var docx = new Docxtemplater()
    .attachModule(chartModule)
    .load(content)
    .setData({
        charts: {
            chart1: {
                variables:{
                   chart1:"1212"
                },
                colNames:["Oil"],
                colData: [
                    {
                        rowName: 'Ireland',
                        data: [4.3]
                             
                    },
                    {
                        rowName: 'Germany',
                        data: [9.0]
                    }
                ]
            }

        }
    })
    .render();


var buffer = docx
    .getZip()
    .generate({ type: "nodebuffer" });

fs.writeFile(OUTDIR + "/test-ppt-chart-1-pie.pptx", buffer);