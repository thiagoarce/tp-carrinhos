// Mocks mínimos das APIs do Apps Script para testar funções do Code.gs
// que tocam a planilha sem precisar de Google real.

function makeSheet(name, rows) {
  rows = rows || [];
  var sheet = {
    _name: name,
    _data: rows,
    getName: function() { return name; },
    getLastRow: function() { return this._data.length; },
    getLastColumn: function() {
      var max = 0;
      this._data.forEach(function(r) { if (r.length > max) max = r.length; });
      return max;
    },
    getDataRange: function() {
      var self = this;
      return {
        getValues: function() { return self._data.map(function(r) { return r.slice(); }); }
      };
    },
    getRange: function(row, col, numRows, numCols) {
      var self = this;
      numRows = numRows || 1;
      numCols = numCols || 1;
      return {
        getValues: function() {
          var out = [];
          for (var i = 0; i < numRows; i++) {
            var r = self._data[row - 1 + i] || [];
            out.push(r.slice(col - 1, col - 1 + numCols));
          }
          return out;
        },
        getValue: function() {
          var r = self._data[row - 1];
          return r ? r[col - 1] : null;
        },
        setValue: function(v) {
          while (self._data.length < row) self._data.push([]);
          var r = self._data[row - 1];
          while (r.length < col) r.push("");
          r[col - 1] = v;
          return this;
        },
        setValues: function(vals) {
          for (var i = 0; i < vals.length; i++) {
            while (self._data.length < row + i) self._data.push([]);
            var r = self._data[row - 1 + i];
            for (var j = 0; j < vals[i].length; j++) {
              while (r.length < col + j) r.push("");
              r[col - 1 + j] = vals[i][j];
            }
          }
          return this;
        }
      };
    },
    appendRow: function(row) { this._data.push(row.slice()); },
    deleteRow: function(rowIdx) { this._data.splice(rowIdx - 1, 1); },
    setFrozenRows: function() { return this; } // no-op
  };
  return sheet;
}

function makeSpreadsheet(sheets) {
  var byName = {};
  sheets.forEach(function(s) { byName[s.getName()] = s; });
  return {
    getSheetByName: function(n) { return byName[n] || null; },
    insertSheet: function(n) { var s = makeSheet(n, []); byName[n] = s; sheets.push(s); return s; }
  };
}

function installMocks(ctx, sheets) {
  var ss = makeSpreadsheet(sheets);
  ctx.SpreadsheetApp = { getActiveSpreadsheet: function() { return ss; } };
  ctx.LockService = {
    getScriptLock: function() {
      return {
        tryLock: function() { return true; },
        releaseLock: function() {}
      };
    }
  };
  ctx.Utilities = {
    formatDate: function(d, tz, fmt) {
      var dt = (d instanceof Date) ? d : new Date(d);
      var yyyy = dt.getFullYear();
      var mm = String(dt.getMonth() + 1).padStart(2, '0');
      var dd = String(dt.getDate()).padStart(2, '0');
      if (fmt === 'yyyy-MM-dd') return yyyy + '-' + mm + '-' + dd;
      var hh = String(dt.getHours()).padStart(2, '0');
      var mi = String(dt.getMinutes()).padStart(2, '0');
      return yyyy + '-' + mm + '-' + dd + ' ' + hh + ':' + mi;
    }
  };
  ctx.PropertiesService = {
    _props: {},
    getScriptProperties: function() {
      var self = this;
      return {
        getProperty: function(k) { return self._props[k] || null; },
        setProperty: function(k, v) { self._props[k] = v; }
      };
    }
  };
  ctx.CacheService = {
    getScriptCache: function() {
      var store = {};
      return {
        get: function(k) { return store[k] || null; },
        put: function(k, v) { store[k] = v; },
        remove: function(k) { delete store[k]; }
      };
    }
  };
  ctx.Session = { getScriptTimeZone: function() { return 'GMT-3'; } };
  ctx.MailApp = { _sent: [], sendEmail: function(opts) { this._sent.push(opts); } };
  return ss;
}

module.exports = { makeSheet, makeSpreadsheet, installMocks };
