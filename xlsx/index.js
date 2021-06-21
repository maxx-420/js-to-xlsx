import templateStrings from "./xml-template-strings";
import { toArrayOfArray, createNode, addToZip, escape } from "./util";
import JSZip from "jszip";

const CELL_DATA_TYPE = {
  string: "string",
  number: "number",
};

const STYLE_NODE_TYPE = {
  xf: "cellXfs",
  font: "fonts",
  fill: "fills",
};

const CELL_STYLES = {
  noStyle: 0,
  boldText: 2,
  wrapText: 10,
  wrappedBoldText: 11,
};

const FONT_TYPES = {
  normal: 0,
  bold: 2,
  italic: 3,
  underlined: 4,
};

const TEXT_ALIGNMENT = {
  right: "right",
  left: "left",
  justify: "fill",
  center: "center",
};

const CELL_STYLE_ATTRIBUTES = {
  [CELL_STYLES.noStyle]: {
    fontId: FONT_TYPES.normal,
  },
  [CELL_STYLES.boldText]: {
    fontId: FONT_TYPES.bold,
  },
  [CELL_STYLES.wrapText]: {
    wrap: true,
  },
  [CELL_STYLES.wrappedBoldText]: {
    fontId: FONT_TYPES.bold,
    wrap: true,
  },
};

const SPECIAL_CHARS = [
  {
    match: /^\-?\d+\.\d%$/,
    style: 16,
    fmt: function (d) {
      return d / 100;
    },
  }, // Precent with d.p.
  {
    match: /^\-?\d+\.?\d*%$/,
    style: 12,
    fmt: function (d) {
      return d / 100;
    },
  }, // Percent
  { match: /^\-?\$[\d,]+.?\d*$/, style: 13 }, // Dollars
  { match: /^\-?£[\d,]+.?\d*$/, style: 14 }, // Pounds
  { match: /^\-?€[\d,]+.?\d*$/, style: 15 }, // Euros
  { match: /^\-?\d+$/, style: 21 }, // Numbers without thousand separators
  { match: /^\-?\d+\.\d{2}$/, style: 22 }, // Numbers 2 d.p. without thousands separators
  {
    match: /^\([\d,]+\)$/,
    style: 17,
    fmt: function (d) {
      return -1 * d.replace(/[\(\)]/g, "");
    },
  }, // Negative numbers indicated by brackets
  {
    match: /^\([\d,]+\.\d{2}\)$/,
    style: 18,
    fmt: function (d) {
      return -1 * d.replace(/[\(\)]/g, "");
    },
  }, // Negative numbers indicated by brackets - 2d.p.
  { match: /^\-?[\d,]+$/, style: 19 }, // Numbers with thousand separators
  { match: /^\-?[\d,]+\.\d{2}$/, style: 20 },
  {
    match: /^[\d]{4}\-[\d]{2}\-[\d]{2}$/,
    style: 23,
    fmt: function (d) {
      return Math.round(25569 + Date.parse(d) / (86400 * 1000));
    },
  }, //Date yyyy-mm-dd
];

export class XLSX {

  _DEFAULT_STYLE = 0;
  _displayedColumns;
  _boldHeader;
  _wrapAll;
  __getCellStyleFn;

  _customXfNodes = [];
  _customFillNodes = [];
  _customFontNodes = [];
  _xfsCount;
  _fillsCount;
  _fontsCount;

  _colConfig = [];

  create = (aoo = [], config = {}) => {
    this._displayedColumns = config.displayedColumns || [];
    this._boldHeader = config.boldHeader || true;
    this._wrapAll = config.wrapAll || false;
    this.__getCellStyleFn = config.getCellStyle || 0;
    this._colConfig = config.colConfig || [];

    this._setDisplayedColumns(aoo);
    const aoa = toArrayOfArray(aoo, this._displayedColumns);
    const worksheetTemplate = this._getTemplate("xl/worksheets/sheet1.xml")
      .replace("{placeholder}", this._createWorksheetTemplate(aoa))
      .replace("{columnConfig}", this.createColumnConfig(aoa));

    const { xfTemplate, fontTemplate, fillTemplate } =
      this._updateStylesTemplate();
    const updatedStylesTemplate = this._getTemplate("xl/styles.xml")
      .replace("{xf}", xfTemplate)
      .replace(/(<cellXfs count=")(\d*)(?=")/g, this._xfsCount? `$1${this._xfsCount}` : '$1$2')
      .replace("{font}", fontTemplate)
      .replace(/(<fonts count=")(\d*)(?=")/g, this._fontsCount? `$1${this._fontsCount}` : '$1$2')
      .replace("{fill}", fillTemplate)
      .replace(/(<fills count=")(\d*)(?=")/g, this._fillsCount? `$1${this._fillsCount}` : '$1$2');

      console.log(updatedStylesTemplate)

    // define file structure
    const xlsx = {
      _rels: {
        ".rels": this._getTemplate("_rels/.rels"),
      },
      xl: {
        _rels: {
          "workbook.xml.rels": this._getTemplate("xl/_rels/workbook.xml.rels"),
        },
        "workbook.xml": this._getTemplate("xl/workbook.xml"),
        "styles.xml": updatedStylesTemplate,
        worksheets: {
          "sheet1.xml": worksheetTemplate,
        },
      },
      "[Content_Types].xml": this._getTemplate("[Content_Types].xml"),
    };

    let zip = new JSZip();
    zip = addToZip(zip, xlsx);
    return zip
      .generateAsync({
        type: "blob",
        mimeType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      })
      .catch((err) => {
        throw new Error(err);
      });
  };

  _getTemplate(key) {
    return templateStrings[key];
  }

  _setDisplayedColumns = (aoo) => {
    if (this._displayedColumns.length > 0) {
      return;
    }
    let result = [];
    aoo.forEach((row) => {
      Object.keys(row)
        .sort()
        .forEach((key, index) => {
          if (result.indexOf(key) == -1) {
            result.splice(index, 0, key);
          }
        });
    });
    this._displayedColumns = result;
  };

  _createWorksheetTemplate = (aoa) => {
    let header = this._displayedColumns.map((colKey) => colKey.toUpperCase());

    let rows = [header, ...aoa]
      .map((rowData, rowIndex) => this._createRows(rowData, rowIndex + 1))
      .join("");
    return rows;
  };

  _createRows = (rowData, rowIndex) => {
    let cells = rowData
      .map((cellData, cellIndex) =>
        this._createCell(cellData, cellIndex, rowIndex)
      )
      .join("");
    return `<row r="${rowIndex}">${cells}</row>`;
  };

  _createCell = (cellData, cellIndex, rowIndex) => {
    if (cellData == null || cellData == undefined) {
      return "";
    }
    let cellPostion = this._getCellPos(cellIndex, rowIndex);

    // Special formatting options
    let specialCell;
    for (var j = 0, jen = SPECIAL_CHARS.length; j < jen; j++) {
      var special = SPECIAL_CHARS[j];
      if (
        cellData.match &&
        !cellData.match(/^0\d+/) &&
        cellData.match(special.match)
      ) {
        let val = cellData.replace(/[^\d\.\-]/g, "");

        if (special.fmt) {
          val = special.fmt(val);
        }

        specialCell = createNode({
          nodeName: "c",
          attr: {
            r: cellPostion,
            s: special.style,
          },
          children: [{ nodeName: "v", cellContent: val }],
        });

        break;
      }
    }
    if (specialCell) return specialCell;

    const cellType = this._getCellType(cellData);
    const cellStyle = this._getCellStyle(rowIndex, cellIndex + 1, cellData);

    if (cellType == CELL_DATA_TYPE.number) {
      return createNode({
        nodeName: "c",
        attr: {
          t: "n",
          r: cellPostion,
          s: cellStyle,
        },
        children: [{ nodeName: "v", cellContent: cellData }],
      });
    }

    cellData = escape(
      cellData.toString().replace(/[\x00-\x09\x0B\x0C\x0E-\x1F\x7F-\x9F]/g, "")
    );
    return createNode({
      nodeName: "c",
      attr: {
        t: "inlineStr",
        r: cellPostion,
        s: cellStyle,
      },
      children: [
        {
          nodeName: "is",
          children: [
            {
              nodeName: "t",
              cellContent: cellData,
              attr: {
                "xml:space": "preserve",
              },
            },
          ],
        },
      ],
    });
  };

  _getCellStyle(rowIndex, cellIndex, cellData) {

    let cellStyle = typeof this.__getCellStyleFn == 'function'?  this.__getCellStyleFn(rowIndex, cellIndex, cellData): this._DEFAULT_STYLE;

    if(cellStyle == null || cellStyle == undefined){
      cellStyle = this._DEFAULT_STYLE;
    }

    const isPrefixQuoteRequired = cellData.toString().match(/^([+,=,-])/g);

    if (typeof this.__getCellStyleFn == "function") {
        if (typeof cellStyle == "number") {
          // styles upto 12th xfs are in use
          let styleIndex = Math.min(Math.abs(cellStyle), 11);
          cellStyle = {
            ...CELL_STYLE_ATTRIBUTES[styleIndex]
          };
        }

        if(rowIndex == 1){
          cellStyle = {
            ...cellStyle,
            fontId: FONT_TYPES.bold
          }
        }

      return isPrefixQuoteRequired? this._createNewXfNode({ ...cellStyle, quotePrefix: true }): this._createNewXfNode(cellStyle)
    }

    return isPrefixQuoteRequired
      ? this._createNewXfNode({
          ...CELL_STYLE_ATTRIBUTES[this._DEFAULT_STYLE],
          quotePrefix: true,
        })
      : this._DEFAULT_STYLE;
  }

  _getCellPos(cellIndex, rowIndex) {
    var ordA = "A".charCodeAt(0);
    var ordZ = "Z".charCodeAt(0);
    var len = ordZ - ordA + 1;
    var s = "";

    while (cellIndex >= 0) {
      s = String.fromCharCode((cellIndex % len) + ordA) + s;
      cellIndex = Math.floor(cellIndex / len) - 1;
    }

    return s + rowIndex;
  }

  _getCellType(cellData) {
    if (
      typeof cellData === "number" ||
      (cellData.match &&
        cellData.match(/^-?\d+(\.\d+)?$/) &&
        !cellData.match(/^0\d+/))
    ) {
      return CELL_DATA_TYPE.number;
    } else {
      return CELL_DATA_TYPE.string;
    }
  }

  _createNewXfNode({
    fontType = 0,
    border = false,
    fontColor = null,
    bgColor = null,
    wrap = this._wrapAll,
    alignment = null,
    quotePrefix = false,
  }) {
    if (!this._xfsCount) {
      this._xfsCount = this._getStyleNodeCount(STYLE_NODE_TYPE.xf);
    }
    let existingStyle = this._customXfNodes.find(({ data }) => {
      return (
        data.fontId == fontType &&
        data.borderId == (border? 1: 0) &&
        data.fontColor == fontColor &&
        data.bgColor == bgColor &&
        data.wrap == wrap &&
        data.alignment == alignment &&
        data.quotePrefix == quotePrefix
      );
    });

    if (existingStyle) {
      return existingStyle.xfId;
    } else {
      this._customXfNodes.push({
        xfId: this._xfsCount,
        data: {
          fontId: fontType,
          borderId: border? 1: 0,
          fontColor,
          bgColor,
          wrap,
          alignment,
          quotePrefix,
        },
      });

      return this._xfsCount++;
    }
  }

  _getStyleNodeCount(nodeType) {
    let styleTemplate = this._getTemplate("xl/styles.xml");
    let nodeCount;
    let regExp;
    switch (nodeType) {
      case STYLE_NODE_TYPE.xf:
        regExp = new RegExp(
          `(<${STYLE_NODE_TYPE.xf} count=")\\d*(?=")`,
          "g"
        );
        nodeCount = +styleTemplate.match(regExp).join('').replace(`<${STYLE_NODE_TYPE.xf} count="`,'');
        break;
      case STYLE_NODE_TYPE.font:
        regExp = new RegExp(
          `(<${STYLE_NODE_TYPE.font} count=")\\d*(?=")`,
          "g"
        );
        nodeCount = +styleTemplate.match(regExp).join('').replace(`<${STYLE_NODE_TYPE.font} count="`,'');;
        break;
      case STYLE_NODE_TYPE.fill:
        regExp = new RegExp(
          `(<${STYLE_NODE_TYPE.fill} count=")\\d*(?=")`,
          "g"
        );
        nodeCount = +styleTemplate.match(regExp).join('').replace(`<${STYLE_NODE_TYPE.fill} count="`,'');;
        break;
      default:
        nodeCount = 0;
        break;
    }

    return nodeCount;
  }

  _updateStylesTemplate() {
    this._customXfNodes.forEach((xf) => {
      if (xf.data.fontColor) {
        xf.data.fontId = this._createNewFontNode({
          fontId: xf.data.fontId,
          color: xf.data.fontColor,
        });
      }
      if (xf.data.bgColor) {
        xf.data.fillId = this._createNewFillNode({ fgColor: xf.data.bgColor });
      }
    });
    //xf
    let xfTemplate = this._customXfNodes
      .map(({ data, xfId }) => {
        let attrs = {
          numFmtId: 0,
          fontId: data.fontId,
          fillId: data.fillId || 0,
          borderId: data.borderId || 0,
          applyFont: 1,
          applyFill: 1,
          applyBorder: 1,
          xfId: 0,
        };
        if (data.alignment) {
          attrs.applyAlignment = 1;
        }
        if (data.quotePrefix) {
          attrs.quotePrefix = 1;
        }
        return createNode({
          nodeName: "xf",
          attr: { ...attrs },
          children: attrs.applyAlignment
            ? [
                {
                  nodeName: "alignment",
                  attr: {
                    horizontal: data.alignment,
                    wrapText: data.wrap ? 1 : 0,
                  },
                },
              ]
            : [],
        });
      })
      .join("");

    ///font
    let fontTemplate = this._customFontNodes
      .map(({ data, fontId }) => {
        let children = [
          {
            nodeName: "sz",
            attr: {
              val: 11,
            },
          },
          {
            nodeName: "name",
            attr: {
              val: "Calibri",
            },
          },
        ];
        // add fontType
        switch (data.fontType) {
          case "bold":
            children.push({
              nodeName: "b",
            });
            break;
          case "underlined":
            children.push({
              nodeName: "u",
            });
            break;
          case "italic":
            children.push({
              nodeName: "i",
            });
            break;
          default:
            break;
        }

        // add fontColor
        children.push({
          nodeName: "color",
          attr: {
            rgb: data.color.replace("#", "FF").toUpperCase(),
          },
        });
        return createNode({
          nodeName: "font",
          children,
        });
      })
      .join("");

    //fill
    let fillTemplate = this._customFillNodes
      .map(({ data, fillId }) => {
        let children = [
          {
            nodeName: "patternFill",
            attr: {
              patternType: "solid",
            },
            children: [
              {
                nodeName: "fgColor",
                attr: {
                  rgb: data.fgColor.replace("#", "FF").toUpperCase(),
                },
              },
              {
                nodeName: "bgColor",
                attr: {
                  indexed: 64,
                },
              },
            ],
          },
        ];
        return createNode({
          nodeName: "fill",
          children,
        });
      })
      .join("");

    return {
      xfTemplate,
      fontTemplate,
      fillTemplate
    };
  }

  _createNewFontNode({ fontId, color }) {
    if (!this._fontsCount) {
      this._fontsCount = this._getStyleNodeCount(STYLE_NODE_TYPE.font);
    }
    let fontType = Object.keys(FONT_TYPES).find(
      (key) => FONT_TYPES[key] == fontId
    );
    let existingFont = this._customFontNodes.find(({ data }) => {
      return data.fontType == fontType && data.color == color;
    });

    if (existingFont) {
      return existingFont.fontId;
    } else {
      this._customFontNodes.push({
        fontId: this._fontsCount,
        data: {
          fontType,
          color,
        },
      });

      return this._fontsCount++;
    }
  }

  _createNewFillNode({ fgColor }) {
    if (!this._fillsCount) {
      this._fillsCount = this._getStyleNodeCount(STYLE_NODE_TYPE.fill);
    }

    let existingFill = this._customFillNodes.find(({ data }) => {
      return data.fgColor == fgColor;
    });

    if (existingFill) {
      return existingFill.fillId;
    } else {
      this._customFillNodes.push({
        fillId: this._fillsCount,
        data: {
          fgColor,
        },
      });

      return this._fillsCount++;
    }
  }

  createColumnConfig(aoa) {
    const cols = this._displayedColumns.map((colKey, colIndex)=>{
        const {width} = this._colConfig.find(config=> config.colKey == colKey) || { width: 0};
        return createNode({
          nodeName: 'col',
          attr: {
            min: colIndex+1,
            max: colIndex+1,
            width: this.calculateColumnWidth(aoa,colIndex,colKey, width),
            customWidth: 1
          }
        })
    }).join('');

    return `<cols>${cols}</cols>`
  }

  calculateColumnWidth(aoa,colIndex, colKey, customWidth) {
    let max = colKey.length;
    let len, lineSplit, str;

    customWidth *= 1.35;

    for (var i = 0, ien = aoa.length; i < ien; i++) {
      var point = aoa[i][colIndex];
      str = point !== null && point !== undefined ? point.toString() : "";

      // If there is a newline character, workout the width of the column
      // based on the longest line in the string
      if (str.indexOf("\n") !== -1) {
        lineSplit = str.split("\n");
        lineSplit.sort(function (a, b) {
          return b.length - a.length;
        });

        len = lineSplit[0].length;
      } else {
        len = str.length;
      }

      if (len > max) {
        max = len;
      }

      // Max width rather than having potentially massive column widths
      if (max > 40) {
        return customWidth? customWidth : 54; 
      }
    }

    max *= 1.35;

    // And a min width
    return max > 6 ? customWidth? customWidth: max  : 6;
  }
}
