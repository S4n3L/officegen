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
          row.data.map(function(cell) {
            cell = cell || {};
            if (typeof cell === 'string' || typeof cell === 'number') {
              var val = cell;
              cell = {
                val: val
              };
            }

            return self._getCell(cell, tblOpts);
          }),
          row.options
        );
      }),
      self._getColSpecs(rows, tblOpts),
      tblOpts
    );
  },

  _getBase: function(rowSpecs, colSpecs, opts) {
    var self = this;
    var baseTable = {
      "w:tbl": {
        "w:tblPr": {
          "w:tblStyle": {
            "@w:val": "a3"
          },
          "w:tblW": {
            "@w:w": "9638",
            "@w:type": "dxa"
          },
          "w:tblLook": {
            "@w:val": "04A0",
            "@w:firstRow": "1",
            "@w:lastRow": "0",
            "@w:firstColumn": "1",
            "@w:lastColumn": "0",
            "@w:noHBand": "0",
            "@w:noVBand": "1"
          },
          "w:tblBorders": {
            "w:top": {
              "@w:val": "single",
              "@w:sz": "6",
              "@w:space": "0",
              "@w:color": "000000"
            },
            "w:bottom": {
              "@w:val": "single",
              "@w:sz": "6",
              "@w:space": "0",
              "@w:color": "000000"
            },
            "w:left": {
              "@w:val": "single",
              "@w:sz": "6",
              "@w:space": "0",
              "@w:color": "000000"
            },
            "w:right": {
              "@w:val": "single",
              "@w:sz": "6",
              "@w:space": "0",
              "@w:color": "000000"
            },
            "w:insideH": {
              "@w:val": "single",
              "@w:sz": "6",
              "@w:space": "0",
              "@w:color": "000000"
            },
            "w:insideV": {
              "@w:val": "single",
              "@w:sz": "6",
              "@w:space": "0",
              "@w:color": "000000"
            }
          }
        },
        "w:tblGrid": {
          "#list": colSpecs
        },
        "#list": rowSpecs
      }
    };

    return baseTable;
  },

  _getColSpecs: function(cols, opts) {
    var self = this;
    return cols[0].data.map(function(val, idx) {
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
        "w:trPr": {
          "w:trHeight": {
            "@w:val": opts.height || "0",
            "@w:hRule": "atLeast"
          }
        },
        "#list": cells // populate this with an array of table cell objects
      }
    };
  },

  _getCell: function(cell, tblOpts) {
    var paragraphList = parseCellContent(cell.data);

    var cellObj = {
      "w:tc": {
        "w:tcPr": {
          "w:tcW": {
            "@w:w": cell.options.width || "2000",
            "@w:type": "dxa"
          },
          "w:vAlign": {
            "@w:val": cell.options.vAlign || "top"
          },
          // "w:shd": {
          //   "@w:val": "clear",
          //   "@w:color": "auto",
          //   "@w:fill": cell.options.shd && cell.options.shd.fill || "",
          //   "@w:themeFill": cell.options.shd && cell.options.shd.themeFill || "",
          //   "@w:themeFillTint": cell.options.shd && cell.options.shd.themeFillTint || ""
          // }
       // <w:shd w:val="clear" w:color="auto" w:fill="FF0000">
          "w:shd": {
            "@w:val": "clear",
            "@w:color": "auto",
            "@w:fill": "FFFFFF"
          }
        },
        "#list": paragraphList
        }
    };


    function parseCellContent(cellData) {
      var _ = require('lodash');
      var paragraphList = []; //array to be populated with <w:p> objects

      _.forEach(cellData, function (paragraph) {
          paragraphList.push(parseParagraphContent(paragraph));
      });

      function parseParagraphContent (paragraph) {
        var paragraphObj = {};
        var contentsList = []; //array to be populated with <w:r> objects
        _.forEach(paragraph.data, function(content) {
          contentsList.push(parseContent(content));
        });
        paragraphObj = makeParagraphObject(paragraph.options, contentsList);

        function makeParagraphObject(paragraphOptions, contentsList) {
          var paragraphObj = {
            "w:p": {
              "w:pPr": {
                "w:jc": {
                  "@w:val": paragraphOptions.align || tblOpts.tableAlign || "center"
                }
              },
              "#list": contentsList
            }
          };

          if(paragraphOptions.list_type === "1") {
            paragraphObj["w:p"]["w:pPr"]["w:pStyle"] = {
              "@w:val": "ListParagraph"
            };
            paragraphObj["w:p"]["w:pPr"]["w:numPr"] = {
              "w:ilvl": {"@w:val": 0},
              "w:numId": {"@w:val": 1}
            }
          }
          return paragraphObj;
        }

        function parseContent(content) { //in W:r tags
          var contentObj = {};
          var finalContentObj = {};
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
                contentObj["w:hyperlink"] = {
                  "@r:id": "rId" + content.link_rel_id
                };
                rPrData["w:rStyle"] = {
                  "@w:val": "Hyperlink"
                };
              }

              if ( rPrData ) {
                contentObj["w:rPr"] = rPrData;
              } // Endif.

              contentObj["w:t"] = content.text.encodeHTML ();
              if(tExtra) {
                contentObj["w:t"]["@xml:space"] = "preserve";
              }

              finalContentObj["w:r"] = contentObj;


            }
          } else if ( content.options.line_break ) {
            finalContentObj["w:r"] = contentObj;
            finalContentObj["w:r"]["w:br"] = {};
          }
          return finalContentObj;
        }

        return paragraphObj;
      }

      return paragraphList;
    }

    return cellObj;
  }
};
