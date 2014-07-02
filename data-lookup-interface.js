'use strict';

/* jshint unused: false */

/**
  Represent a generic interface of data lookup
  This is a pseudo interface, see relevant implementation
  @see {@link module:data/data-lookup}
  @constructor
  @param {object} dbApi - a DbApi instance
 */

module.exports = function DataLookupInterface(dbApi) {
  // get type by name
  this.typeByName = function(typeName) {
  };
  // find entity by id
  this.entityById = function(typeId) {
  };
  // find entity by alias
  this.entityByAlias = function(alias) {
  };
  // create new entity
  this.add = function(typeName, entity) {
  };
  // store mac of an entity by alias
  this.storeMacByAlias = function(alias, roleName, mac) {
  };
  // store mac of entities by type
  this.storeMacByType = function(typeName, roleName, mac) {
  };
  // retrieve mac for alias or type
  this.retrieveMac = function(alias, typeName) {
  };
};
