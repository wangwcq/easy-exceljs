const _ = require('lodash');
const ExcelJS = require('exceljs');

class SpreadSheet {
  static ensureArray(strOrArray = []) {
    let ret = strOrArray;
    if (!_.isArray(ret)) {
      ret = [ret];
    }
    return _.flattenDeep(ret);
  }
  static getStyles(stylesDefinition = {}, styles = []) {
    let vStyles = this.ensureArray(styles);
    const mergedStyles = _.merge({}, ..._.map(vStyles, style => _.get(stylesDefinition, style, {})));
    return mergedStyles;
  }
  static applyStylesOnItem(item = null, stylesDefinition = {}, styles = []) {
    if (!item) return;
    const mergedStyles = this.getStyles(stylesDefinition, styles);
    _.forEach(mergedStyles, (attributes, key) => {
      item[key] = attributes || {};
    });
  }
  constructor(options = {}) {
    const {
      styles = {},
      pageSetup = {},
    } = options;

    this.workbook = new ExcelJS.Workbook();
    this.currentSheet = null;
    this.currentRow = 1;
    this.currentCol = 1;
    this.lastPos = { row: 1, col: 1, };
    this.lastStyles = [];
    this.pageBreaks = {};
    this.pageSetup = _.merge({}, {
      paperSize: 9,
      orientation: 'landscape',
      fitToWidth: 1,
    }, pageSetup);

    this.styles = styles;
  }
  addSheet(title = 'Worksheet') {
    const sheet = this.workbook.addWorksheet(title, {
      pageSetup: this.pageSetup,
    });
    this.pageBreaks[sheet.id] = [];
    this.currentSheet = sheet;
    this.currentRow = 0;
    this.nextRow();
    return this;
  }
  getSheet(sheetName = null) {
    if (sheetName) {
      this.currentSheet = this.workbook.getWorksheet(sheetName);
    }
    return this.currentSheet;
  }
  pos(row, col) {
    this.lastPos = {
      row: this.currentRow,
      col: this.currentCol,
    };
    this.currentRow = row;
    this.currentCol = col;
    return this;
  }
  posLast() {
    const { row, col } = this.lastPos;
    return this.pos(row, col);
  }
  posRel(dRow = 0, dCol = 0) {
    return this.pos(this.currentRow + dRow, this.currentCol + dCol);
  }
  getCurrentRow() {
    return this.currentRow;
  }
  getCurrentCol() {
    return this.currentCol;
  }
  getRow() {
    return this.currentSheet.getRow(this.currentRow);
  }
  nextRow(styles = []) {
    this.pos(this.currentRow + 1, 1);
    if (this.styles.row || styles.length) {
      this.constructor.applyStylesOnItem(this.getRow(), this.styles, ['row', styles])
    }
    return this;
  }
  addPageBreak() {
    this.getRow().addPageBreak();
    this.pageBreaks[this.currentSheet.id].push(this.getCurrentRow());
    this.pageBreaks[this.currentSheet.id] = this.pageBreaks[this.currentSheet.id].sort();
    return this.nextRow();
  }
  getCurrentPageNumber() {
    let ret = 1;
    _.forEach(this.pageBreaks[this.currentSheet.id], rowIndex => {
      if (this.currentRow <= rowIndex) { return false; }
      if (this.currentRow > rowIndex) { ret += 1; return true; }
    });
    return ret;
  }
  getTotalPages() {
    return this.pageBreaks[this.currentSheet.id].length;
  }
  merge(dRow = 1, dCol = 1, posNext = true) {
    this.currentSheet.mergeCells(
      this.currentRow,
      this.currentCol,
      this.currentRow + dRow - 1,
      this.currentCol + dCol - 1,
    );
    return this.pos(
      this.currentRow + dRow - 1,
      this.currentCol + dCol - 1 + (posNext ? 1 : 0),
    );
  }
  getCell() {
    const cell = this.currentSheet.getRow(this.currentRow).getCell(this.currentCol);
    return cell;
  }
  fill(text = '', styles = null, posNext = true) {
    const cell = this.getCell();
    cell.value = text;
    this.applyStyles(styles);
    return this.pos(this.currentRow, this.currentCol + (posNext ? 1 : 0));
  }
  fillAoa(aoa = [], options = {}) {
    const {
      pageBreakAt,
      pageBreak,
    } = options;

    _.forEach(aoa, (row, rowIndex) => {
      _.forEach(row, cell => {
        this.fill(cell);
      });
      if (rowIndex < aoa.length - 1) {
        this.nextRow();
        if (rowIndex % pageBreakAt === pageBreakAt - 1) {
          pageBreak(this);
        }
      }
    });

    return this;
  }
  addImage(imageIdType, imageId, imageExtension, rowSpan, colSpan, width, height) {
    const image = this.workbook.addImage({
      [imageIdType]: imageId,
      extension: imageExtension,
    });
    const imagePos = {
      tl: { col: this.currentCol - 1 + 0.1, row: this.currentRow - 1 + 0.1 },
      br: { col: this.currentCol + colSpan - 1 - 0.1, row: this.currentRow + rowSpan - 1 - 0.1},
      editAs: 'absolute',
    };
    if (width && height) {
      delete imagePos.br;
      imagePos.ext = {
        width,
        height,
      };
    }
    this.currentSheet.addImage(image, imagePos);
    return this;
  }
  applyStyles(styles = null) {
    const cell = this.getCell();
    const vStyles = styles || this.lastStyles;
    this.constructor.applyStylesOnItem(cell, this.styles, this.constructor.ensureArray([
      'base',
      this.constructor.ensureArray(vStyles),
    ]));
    this.lastStyles = vStyles;
    return this;
  }
  columnWidths(widths = []) {
    this.currentSheet.columns = _.map(widths, width => ({ width }));
    return this;
  }
  async writeXlsx(path) {
    await this.workbook.xlsx.writeFile(path);
  }
  async toXlsxBuffer() {
    const buffer = await this.workbook.xlsx.writeBuffer();
    return buffer;
  }
}

module.exports = SpreadSheet;
