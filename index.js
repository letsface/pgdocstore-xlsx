'use strict';

var XLSX = require('xlsx');
var Q = require('q');
var _ = require('underscore');
var DataLookupInterface = require('./data-lookup-interface');
var checkImplements = require("checkimplements");

Q.longStackSupport = false;

/**
 * Custom error type for XLSX errors
 * @constructor XlsxError
 * @param {string} msg
 */
function XlsxError (msg) {
  this.name = 'XlsxError';
  this.message = msg;
};

XlsxError.prototype = new Error();

XlsxError.prototype.constructor = XlsxError;

/**
 * Convert worksheets to JSON object arrays.
 * @return {object} An array of JSON objects representing rows in a worksheet
 */
function worksheetsToJSON(workbook) {
  var result = [];
  workbook.SheetNames.forEach(function(sheetName) {
    var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
    if(roa.length > 0) {
      result.push(roa);
    }
  });
  return result;
}

/**
 * Fill in related EMS objects
 * @param  {object} emsObj
 * @param  {object} dataLookup
 * @return {promise}
 */
function resolveRelatedObj(emsObj, dataLookup) {
  var promise = Q();
  Object.keys(emsObj.rel).forEach(function(typeName) {
      var relObj = emsObj.rel[typeName];
      promise = promise
        .then(function() {
          if(relObj.id) {
            return dataLookup.entityById(relObj.id);
          } else if(relObj.alias) {
            return dataLookup.entityByAlias(relObj.alias);
          } else {
            return relObj;
          }
        })
        .then(function(filledRelObj) {
          emsObj.rel[typeName] = filledRelObj;
        });
  });
  return promise;
}

/**
 * Set property value in an EMS object.
 * @param {object} emsObj
 * @param {string} path
 * @param {string} value
 */
function setValue(emsObj, property, value) {
  var location = emsObj;
  var path = property.split('.')
  var valueLocation = null;
  var lastPart = null;
  path.forEach(function(part) {
    if(!location[part]) {
      location[part] = {};
    }
    valueLocation = location;
    lastPart = part;
    location = location[part];
  });
  valueLocation[lastPart] = value;
}

/**
 * Turn a column name with possible semi-column separators into a
 * valid docstore . separated path that can be used as input to setValue.
 * @param  {string} columnName
 * @return {string} a . separated path to the (possibly nested) entity doc
 */
function columnNameToPath(columnName) {
  var parts = columnName.split(':');
  var full_path = [];
  for(var i=0; i<parts.length-1; i++) {
    full_path.push('rel');
    full_path.push(parts[i]);
  }
  full_path.push('doc');
  full_path.push(parts[parts.length-1]);
  return full_path.join('.');
}


/**
 * Fill in an EMS object's properties with data from spreadsheet
 * @param  {object} emsObj
 * @param  {object} propertiesList
 * @param  {object} row
 * @return {undefined}
 */
function updateProperties(emsObj, propertiesList, row) {
  // known properties
  propertiesList.forEach(function(prop) {
    if(!row[prop.name]) {
      var idName = prop.typeName + ':id';
      var aliasName = prop.typeName + ':alias';
      if(row[idName]) {
        emsObj.rel[prop.typeName] = {id: row[idName]};
        delete row[idName];
      } else if(row[aliasName]) {
        emsObj.rel[prop.typeName] = {alias: row[aliasName]};
        delete row[aliasName];
      } else {
        if(prop.required) {
          throw new Error(
            'Missing required property [' + prop.name +
            '] on row ' + JSON.stringify(row) +
            ' and [id] is missing: ' + idName +
            ' or [alias] is missing: ' + aliasName
            );
        }
        // else ignore, not required and no way to find it again
      }
      return; // was an id or alias, skip to next
    }

    setValue(emsObj, prop.path, row[prop.name]);
    delete row[prop.name];
  });

  // special properties for the root of the object
  if(row['alias']) {
    // console.log('using alias ' + row['alias'] + ' for object ' + JSON.stringify(emsObj));
    emsObj.alias = row['alias'];
    delete row['alias'];
  }

  // whatever is left in row is to be stored in object as-is
  Object.keys(row).forEach(function(pname) {
    setValue(emsObj, columnNameToPath(pname), row[pname]);
  });
}

/**
 * Parse a JSON formatted row object into an full-blown EMS entity
 * @param  {string} typeName
 * @param  {object} dataLookup
 * @param  {object} row
 * @return {promise}
 */
function parseEmsObject(typeName, dataLookup, row) {
  return dataLookup
    .typeByName(typeName)
    .then(function(type) {
      if ( !type.doc ) {
        throw new Error('Type [' + typeName + '] does not exist.');
      }

      if( !type.doc.propertiesList ) {
        throw new Error('Type [' + typeName + '] does not have properties.');
      }

      var DEFAULT_EMS_OBJECT = { doc: {}, rel: {} };

      var emsObj = _.extend({ type: typeName }, DEFAULT_EMS_OBJECT);

      updateProperties(emsObj, type.doc.propertiesList, row);

      emsObj.mac = dataLookup.retrieveMac(emsObj.alias, typeName);

      // resolve the ids to the full object (for preview)
      return resolveRelatedObj(emsObj, dataLookup)
        .then(function() {
          return dataLookup.add(typeName, emsObj);
        });
    });
}

/**
 * Map JSON formatted rows in a worksheet to EMS objects
 * @param  {string} typeName
 * @param  {array} rows
 * @param  {object} dataLookup
 * @return {promise}
 */
function worksheetToObjects(typeName, rows, dataLookup) {
  if(!rows || !typeName || !dataLookup) {
    throw new XlsxError('Invalid parameters ' + JSON.stringify(arguments));
  }
  return Q.all(_.map(rows, parseEmsObject.bind(null, typeName, dataLookup)));
}

function rowToMac(dataLookup, row) {
  var rights = {
    query: !!row['query'],
    update: !!row['update'],
    remove: !!row['remove'],
    create: !!row['create']
  }

  if(row['Entity:alias']) {
    dataLookup.storeMacByAlias(row['Entity:alias'], row['Role:name'], rights);
  } else if(row['Type:name']) {
    dataLookup.storeMacByType(row['Type:name'], row['Role:name'], rights);
  }
}

/**
 * Import the first worksheet in a spreadsheet
 * @param  {blob} binaryData
 * @param  {string} typeName
 * @param  {object} dataLookup
 * @return {promise}
 */
function xlsxSingleWorksheetToObjects(binaryData, typeName, dataLookup) {
  checkImplements(DataLookupInterface, dataLookup);
  var workbook = XLSX.read(binaryData, {type: 'binary'});
  var worksheets = worksheetsToJSON(workbook);
  return worksheetToObjects(typeName, worksheets[0], dataLookup)
    .then(function () {
      return _.keys(worksheets[0][0]);
    });
}

function xlsxWorksheetsToObjects(binaryData, dataLookup) {
  checkImplements(DataLookupInterface, dataLookup);
  var workbook = XLSX.read(binaryData, {type: 'binary'});
  var worksheets = worksheetsToJSON(workbook);

  // start promise
  var promise = Q();

  // Mac
  if(workbook.SheetNames[0] !== 'MAC') {
    return Q.reject(new XlsxError('MAC (Mandatory Access Control) must be the first sheet'));
  }

  promise = promise.then(function() {
    return Q.all(_.map(worksheets[0], rowToMac.bind(null, dataLookup)));
  });

  // Role
  if(workbook.SheetNames[1] !== 'Role') {
    return Q.reject(new XlsxError('Role must be the second sheet, was ' + workbook.SheetNames[0]));
  }

  for(var i = 1; i < worksheets.length; i++) {
    promise = (function (typeName, worksheet, lookup) {
      return promise.then(function() {
        return worksheetToObjects(typeName, worksheet, lookup);
      });
    } (workbook.SheetNames[i], worksheets[i], dataLookup));
  }
  return promise;
}

exports.xlsxSingleWorksheetToObjects = xlsxSingleWorksheetToObjects;

exports.xlsxWorksheetsToObjects = xlsxWorksheetsToObjects;

exports.DataLookupInterface = DataLookupInterface;
