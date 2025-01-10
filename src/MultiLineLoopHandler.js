import { TagUtil } from './TagUtil.js';
import { default as AtTagHandler } from './AtTagHandler.js';
import { default as ExcelCopier } from './ExcelCopier.js';

class MultiLineLoopHandler {
  constructor(worksheet, data) {
    this.worksheet = worksheet;
    this.data = data;
    this._reset();
  }

  _reset() {
    this.startRowIndex = null;
    this.endRowIndex = null;
    this.loopTag = null;
    this.loopDataArray = null;
  }

  /**
   * 查找多行循环的起始和结束位置
   * @returns {boolean} 是否找到循环标记
   */
  _findLoopRange() {
    const range = this.worksheet.dimensions;
    if (!range) return false;

    // 遍历所有行寻找循环开始标记
    for (let row = range.top; row <= range.bottom; row++) {
      const currentRow = this.worksheet.getRow(row);
      let startTag = null;

      // 检查这一行的每个单元格
      currentRow.eachCell({ includeEmpty: false }, (cell) => {
        if (!startTag) {
          const value = cell.value?.toString() || "";
          const match = value.match(/\{#(\w+)\}/);
          if (match) {
            startTag = match[1];
            this.startRowIndex = row;
            this.loopTag = startTag;
          }
        }
      });

      // 如果找到了开始标记，继续寻找结束标记
      if (this.startRowIndex) {
        for (let endRow = row; endRow <= range.bottom; endRow++) {
          const endRowObj = this.worksheet.getRow(endRow);
          let foundEnd = false;

          endRowObj.eachCell({ includeEmpty: false }, (cell) => {
            const value = cell.value?.toString() || "";
            if (value.includes(`{/${startTag}}`)) {
              this.endRowIndex = endRow;
              foundEnd = true;
            }
          });

          if (foundEnd) break;
        }

        if (this.endRowIndex) {
          return true;
        }
      }
    }

    return false;
  }

  /**
   * 处理循环数据
   */
  async _handleLoopData() {
    // 从数据中获取循环数组
    this.loopDataArray = this.data[this.loopTag] || [];
    if (!Array.isArray(this.loopDataArray)) {
      this.loopDataArray = [this.loopDataArray];
    }

    const rowCount = this.endRowIndex - this.startRowIndex + 1;

    if (this.loopDataArray.length === 0) {
      // 删除循环行
      this.worksheet.spliceRows(this.startRowIndex, rowCount);
    } else if (this.loopDataArray.length > 1) {
      // 复制行
      const copier = new ExcelCopier(this.worksheet);
      copier.copyRows(
        this.startRowIndex,
        this.endRowIndex,
        this.endRowIndex + 1
      );
    }
  }

  /**
   * 移除循环标记
   * @param {*} startRow 开始行
   * @param {*} endRow 结束行
   */
  async _removeLoopTag(startRow, endRow) {
    for (let row = startRow; row <= endRow; row++) {
      const currentRow = this.worksheet.getRow(row);

      currentRow.eachCell({ includeEmpty: false }, async (cell) => {
        if (row === startRow) {
          let valueStr = cell.value?.toString() || "";
          let pattern = `\{#${this.loopTag}\}`;
          if (valueStr.match(pattern)) {
            cell.value = (cell.value?.toString() || "").replace(
              `{#${this.loopTag}}`,
              ""
            );
          }
        }
        if (row === endRow) {
          let valueStr = cell.value?.toString() || "";
          let pattern = `\{\/${this.loopTag}\}`;
          if (valueStr.match(pattern)) {
            cell.value = (cell.value?.toString() || "").replace(
              `{/${this.loopTag}}`,
              ""
            );
          }
        }
      });
    }
  }
  /**
   * 处理标记替换
   */
  async _handleTags() {
    const rowCount = this.endRowIndex - this.startRowIndex + 1;

    for (let i = 0; i < this.loopDataArray.length; i++) {
      const currentData = this.loopDataArray[i];
      const startRow = this.startRowIndex + i * rowCount;
      const endRow = startRow + rowCount - 1;

      // 移除循环标记
      await this._removeLoopTag(startRow, endRow);

      // 处理@标记
      await this._processAtTags(startRow, endRow, currentData);

      // 处理普通标记
      await this._processNormalTags(startRow, endRow, currentData);
    }
  }

  /**
   * 处理@标记
   * @param {*} startRow 开始行
   * @param {*} endRow 结束行
   * @param {*} currentData 当前数据
   */
  async _processAtTags(startRow, endRow, currentData) {
    // 处理@标记
    const atHandler = new AtTagHandler(currentData);
    await this.iterateAllCell(startRow, endRow, async (cell) => {
      await atHandler.next(cell);
    });
  }

  async iterateAllCell(startRow, endRow, handler){
    for (let row = startRow; row <= endRow; row++) {
      const currentRow = this.worksheet.getRow(row);
      let cellCount = currentRow.cellCount
      for (let columnIndex = 1; columnIndex <= cellCount; ++columnIndex) {
        let cell = currentRow.getCell(columnIndex)
        await handler(cell)
      }
    }
  }
  /**
   * 处理普通标记
   * @param {*} startRow 开始行
   * @param {*} endRow 结束行
   * @param {*} currentData 当前数据
   */
  async _processNormalTags(startRow, endRow, currentData) {
    await this.iterateAllCell(startRow, endRow, async (cell) => {
      // 处理普通标记和图片标记
      if (cell.value) {
        await TagUtil.replaceCell(cell, currentData);
      }
    });
  }

  /**
   * 主处理函数
   */
  async handle() {
    if (!this._findLoopRange()) {
      return false;
    }

    // 如果是单行循环，则不处理
    if(this.endRowIndex == this.startRowIndex){
      return false;
    }

    await this._handleLoopData();

    if (this.loopDataArray.length > 0) {
      await this._handleTags();
    }

    this._reset();
    return true;
  }
}

export { MultiLineLoopHandler as default };
