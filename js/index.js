var ChartMaker, ChartManager, ChartModule, SubContent, fs;

SubContent = require('docxtemplater').SubContent;

ChartManager = require('./chartManager');

ChartMaker = require('./chartMaker');

fs = require('fs');

ChartModule = (function() {

  /**
  	 * self name for self-identification, variable for fast changing;
  	 * @type {String}
   */
  ChartModule.prototype.name = 'chart';


  /**
  	 * initialize options with empty object if not recived
  	 * @manager = ModuleManager instance
  	 * @param  {Object} @options params for the module
   */

  function ChartModule(options) {
    this.options = options != null ? options : {};
  }

  ChartModule.prototype.handleEvent = function(event, eventData) {
    var gen;
    if (event === 'rendering') {
      this.renderingFileName = eventData;
      gen = this.manager.getInstance('gen');
      this.tags = gen.tags;
      this.chartManager = new ChartManager(gen.zip, this.tags && this.tags.charts ? this.tags.charts : null);
      return this.chartManager.rendered();
    } else if (event === 'rendered') {
      return this.finished();
    }

  };

  ChartModule.prototype.get = function(data) {
   
    return null;
  };

  ChartModule.prototype.handle = function(type, data) {
     return null;
  };

  ChartModule.prototype.finished = function() {};

  ChartModule.prototype.on = function(event, data) {
   
  };

  ChartModule.prototype.replaceBy = function(text, outsideElement) {
   
  };

  ChartModule.prototype.replaceTag = function() {

  };

  return ChartModule;

})();

module.exports = ChartModule;
