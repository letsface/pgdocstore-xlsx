var expect = require('chai').expect,
    xlsxImport = require('..');

describe('xlsx-import', function() {
  it('should export xlsxSingleWorksheetToObjects', function(done) {
    expect(xlsxImport.xlsxSingleWorksheetToObjects).to.be.a('function');
    expect(xlsxImport.xlsxSingleWorksheetToObjects).to.be.a('function');
    done();
  });
});
