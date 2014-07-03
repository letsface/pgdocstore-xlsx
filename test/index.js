var expect = require('chai').expect,
    xlsx = require('..');

describe('xlsx', function() {
  it('should export functions', function(done) {
    expect(xlsx.xlsxSingleWorksheetToObjects).to.be.a('function');
    expect(xlsx.xlsxWorksheetsToObjects).to.be.a('function');
    expect(xlsx.DataLookupInterface).to.be.a('function');
    done();
  });
});
