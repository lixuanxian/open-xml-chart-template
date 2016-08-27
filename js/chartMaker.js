var ChartMaker;

module.exports = ChartMaker = (function() {

  ChartMaker.prototype.getLineTemplate = function(line, i) {
    var elem, result, _i, _j, _len, _len1, _ref, _ref1;
    result = "<c:ser>\n	<c:idx val=\"" + i + "\"/>\n	<c:order val=\"" + i + "\"/>\n	<c:tx>\n		<c:v>" + line.name + "</c:v>\n	</c:tx>\n	<c:marker>\n		<c:symbol val=\"none\"/>\n	</c:marker>\n	<c:cat>\n\n		<c:" + this.ref + ">\n			<c:" + this.cache + ">\n				" + (this.getFormatCode()) + "\n				<c:ptCount val=\"" + line.data.length + "\"/>\n	";
    _ref = line.data;
    for (i = _i = 0, _len = _ref.length; _i < _len; i = ++_i) {
      elem = _ref[i];
      result += "<c:pt idx=\"" + i + "\">\n	<c:v>" + elem.x + "</c:v>\n</c:pt>";
    }
    result += "		</c:" + this.cache + ">\n	</c:" + this.ref + ">\n</c:cat>\n<c:val>\n	<c:numRef>\n		<c:numCache>\n			<c:formatCode>General</c:formatCode>\n			<c:ptCount val=\"" + line.data.length + "\"/>";
    _ref1 = line.data;
    for (i = _j = 0, _len1 = _ref1.length; _j < _len1; i = ++_j) {
      elem = _ref1[i];
      result += "<c:pt idx=\"" + i + "\">\n	<c:v>" + elem.y + "</c:v>\n</c:pt>";
    }
    result += "			</c:numCache>\n		</c:numRef>\n	</c:val>\n</c:ser>";
    return result;
  };

  function ChartMaker(zip, options) {
    this.zip = zip;
    this.options = options;
    if (this.options.axis.x.type === 'date') {
      this.ref = "numRef";
      this.cache = "numCache";
    } else {
      this.ref = "strRef";
      this.cache = "strCache";
    }
  }

  ChartMaker.prototype.makeChartFile = function(lines) {
    var i, line, result, _i, _len;
    result = this.getTemplateTop();
    for (i = _i = 0, _len = lines.length; _i < _len; i = ++_i) {
      line = lines[i];
      result += this.getLineTemplate(line, i);
    }
    result += this.getTemplateBottom();
    this.chartContent = result;
    return this.chartContent;
  };

  ChartMaker.prototype.writeFile = function(path) {
    this.zip.file("word/charts/" + path + ".xml", this.chartContent, {});
  };

  return ChartMaker;

})();
