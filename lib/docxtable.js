module.exports = {

  // assume passed in an array of row objects
  getTable: function(rows, tblOpts) {
    var xmlbuilder = require('xmlbuilder');

    if ( !String.prototype.encodeHTML ) {
      String.prototype.encodeHTML = function () {
        return this.replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
      };
    }
    tblOpts = tblOpts || {};

    var self = this;

    return self._getBase(
      rows.map(function(row) {
        return self._getRow(
          row.map(function(cell) {
            cell = cell || {};
            if (typeof cell === 'string' || typeof cell === 'number') {
              var val = cell;
              cell = {
                val: val
              };
            }

            return self._getCell(cell, tblOpts);
          }),
          tblOpts
        );
      }),
      self._getColSpecs(rows, tblOpts),
      tblOpts
    );
  },

  _getBase: function(rowSpecs, colSpecs, opts) {
    var self = this;
    // var outString ='<w:tbl> <w:tblPr> <w:tblStyle w:val="';
    // outString += opts.cellColWidth || tblOpts.tableColWidth || "0";
    // outString += '"/> <w:tblW w:w="0" w:type="auto"/>';
    // outString += '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" ';
    // outString += 'w:lastColumn="0" w:noHBand="0" w:noVBand="1"/> <w:tblBorders>';
    // outString += '<w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>';
    // outString += '<w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>';
    // outString += '<w:left w:val="single" w:sz="12" w:space="0" w:color="000000"/>';
    // outString += '<w:right w:val="single" w:sz="12" w:space="0" w:color="000000"/>';
    // outString += '</w:tblBorders> </w:tblPr> <w:tblGrid> <w:gridCol w:w="';
    // outString += colSpecs + '"/>';

    var baseTable = {
      "w:tbl": {
        "w:tblPr": {
          "w:tblStyle": {
            "@w:val": "a3"
          },
          "w:tblW": {
            "@w:w": "0",
            "@w:type": "auto"
          },
          "w:tblLook": {
            "@w:val": "04A0",
            "@w:firstRow": "1",
            "@w:lastRow": "0",
            "@w:firstColumn": "1",
            "@w:lastColumn": "0",
            "@w:noHBand": "0",
            "@w:noVBand": "1"
          }
        },
        "w:tblGrid": {
          "#list": colSpecs
        },
        "#list": [rowSpecs]
      }
    };
    if (opts.borders) {
      baseTable["w:tbl"]["w:tblPr"]["w:tblBorders"] = {
        "w:top": {
          "@w:val": "single",
          "@w:sz": "12",
          "@w:space": "0",
          "@w:color": "000000"
        },
        "w:bottom": {
          "@w:val": "single",
          "@w:sz": "12",
          "@w:space": "0",
          "@w:color": "000000"
        },
        "w:left": {
          "@w:val": "single",
          "@w:sz": "12",
          "@w:space": "0",
          "@w:color": "000000"
        },
        "w:right": {
          "@w:val": "single",
          "@w:sz": "12",
          "@w:space": "0",
          "@w:color": "000000"
        }
      };
    }

    return baseTable;
  },

  _getColSpecs: function(cols, opts) {
    var self = this;
    return cols[0].map(function(val, idx) {
      return self._tblGrid(opts);
    });
  },

  // TODO 
  _tblGrid: function(opts) {
    return {
      "w:gridCol": {
        "@w:w": opts.tableColWidth || "1"
      }
    };
  },


  _getRow: function(cells, opts) {
    return {
      "w:tr": {
        "@w:rsidR": "00995B51",
        "@w:rsidTr": "007F1D13",
        "#list": [cells] // populate this with an array of table cell objects
      }
    };
  },

  _getCell: function(cell, tblOpts) {
    var val = cell[0].text;
    opts = cell[0].options || {};
    // var b = {};

    // if (opts.b) {
    //   b = {
    //     "w:tc": {
    //       "w:p": {
    //         "w:r": {
    //           "w:rPr": {
    //             "w:b": {}
    //           }
    //         }
    //       }
    //     }
    //   }
    // }

    // var altCellObj = '';
    // altCellObj += '<w:tc> <w:tcPr> <w:tcW w:w="' + (opts.cellColWidth || tblOpts.tableColWidth || "0");
    // altCellObj += '" w:type="dxa"/> <w:vAlign w:val="' + (opts.vAlign || "top");
    // altCellObj += '"/> <w:shd w:val="clear" w:color="auto" w:fill="' + (opts.shd && opts.shd.fill || "");
    // altCellObj += '" w:themeFill="' + (opts.shd && opts.shd.themeFill || "");
    // altCellObj += '" w:themeFillTint="' + (opts.shd && opts.shd.themeFillTint || "");
    // altCellObj += '"/> </w:tcPr> <w:p w:rsidR="00995B51" w:rsidRPr="00722E63" w:rsidRDefault="00995B51"> <w:pPr>' +
    //               '<w:jc w:val="';
    // altCellObj += (opts.align || tblOpts.tableAlign || "center") + '/> </w:pPr>';
    // altCellObj += parseCellContent(cell);
    // alCellObject += ' </w:p> </w:tc>';
    var cellContentsList = parseCellContent(cell.data);
    var cellObj = {
      "w:tc": {
        "w:tcPr": {
          "w:tcW": {
            "@w:w": cell.options.cellColWidth || tblOpts.tableColWidth || "0",
            "@w:type": "dxa"
          },
          "w:vAlign": {
            "@w:val": cell.options.vAlign || "top"
          },
          "w:shd": {
            "@w:val": "clear",
            "@w:color": "auto",
            "@w:fill": cell.options.shd && cell.options.shd.fill || "",
            "@w:themeFill": cell.options.shd && cell.options.shd.themeFill || "",
            "@w:themeFillTint": cell.options.shd && cell.options.shd.themeFillTint || ""
          }
        },
        "w:p": {
          "@w:rsidR": "00995B51",
          "@w:rsidRPr": "00722E63",
          "@w:rsidRDefault": "00995B51",
          "w:pPr": {
            "w:jc": {
              "@w:val": cell.options.align || tblOpts.tableAlign || "center"
            }
          },
          "#list": [cellContentsList]
        }
      }
    };

    function parseCellContent(cell) {
      var cellContents = [];
      var _ = require('lodash');

      _.forEach(cell, function (content) {
        var outString = {};
          if (content.text) {
            var rExtra = '';
            var tExtra = '';
            var rPrData = {};

            if (content.text && content.link_rel_id) {
              var linkOptions = content.options;
              if (linkOptions === undefined) {
                linkOptions = {};
              }

              linkOptions.underline = true;
              linkOptions.color = '426EDD';
              content.options = linkOptions;
            }

            rPrData = {
              "w:color": {
                "@w:val": content.options.color || "000"
              }
            };
            if (content.options) {
              if (content.options.back) {
                rPrData["w:shd"] = {
                  "@w:val": "clear",
                  "@w:color": "auto",
                  "@w:fill": content.options.back
                };
              } // Endif.

              if (content.options.bold) {
                rPrData["w:b"] = {
                  "w:val": content.options.bold
                };
              } // Endif.

              if (content.options.italic) {
                rPrData["w:i"] = {};
              } // Endif.

              if (content.options.underline) {
                rPrData["w:u"] = {
                  "@w:val": "single"
                };
              } // Endif.

              if (content.options.font_face) {
                rPrData["w:rFonts"] = {
                  "@w:ascii": content.options.font_face,
                  "@w:hAnsi": content.options.font_face,
                  "@w:cs": content.options.font_face
                };
              } // Endif.

              if (content.options.font_size) {
                var fontSizeInHalfPoints = 2 * content.options.font_size;
                rPrData["w:sz"] = {
                    "@w:val": fontSizeInHalfPoints
                  };
                rPrData["w:szCs"] = {
                  "@w:val": fontSizeInHalfPoints
                };
              } // Endif.
            }

            if ( content.text ) {
              content.text = content.text.toString();
              if ( (content.text[0] == ' ') || (content.text[content.text.length - 1] == ' ') ) {
                tExtra += ' xml:space="preserve"';
              } // Endif.

              if (content.link_rel_id){
                outString["w:hyperlink"] = {
                  "@r:id": "rId" + content.link_rel_id
                };
                rPrData["w:rStyle"] = {
                  "@w:val": "Hyperlink"
                };
              }

              if ( rPrData ) {
                outString["w:rPr"] = rPrData;
              } // Endif.
              outString["w:t"] = content.text.encodeHTML ();
              if(tExtra) {
                outString["w:t"]["@xml:space"] = "preserve";
              }

            } else if ( content.page_break ) {
              outString["w:r"] = {
                "w:br": {
                  "@w:type": "page"
                }
              };

            } else if ( content.line_break ) {
              outString["w:r"] = {
                "w:br": {}
              };
            }
          }

        var xmlCell = {};
        xmlCell["w:r"] = outString;
        cellContents.push(xmlCell);
      });
      return cellContents;
    }
    return cellObj;
  }
};
