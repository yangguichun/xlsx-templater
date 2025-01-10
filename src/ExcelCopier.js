class ExcelCopier {
  constructor(worksheet) {
    this.worksheet = worksheet;
  }

  /**
   * 复制指定范围的行到目标位置
   * @param {number} startRow 起始行号
   * @param {number} endRow 结束行号
   * @param {number} targetRow 目标插入位置的行号
   */
  copyRows(startRow, endRow, targetRow) {
    const rowCount = endRow - startRow + 1;
    
    // 1. 先将目标位置及以下的行向下移动
    this._shiftRowsDown(targetRow, rowCount);
    
    // 2. 复制每一行
    for (let i = 0; i < rowCount; i++) {
      const sourceRowNum = startRow + i;
      const targetRowNum = targetRow + i;
      this._copyRow(sourceRowNum, targetRowNum);
    }

    // 3. 处理合并单元格
    this._handleMerges(startRow, endRow, targetRow);

    // 4. 更新插入位置之后的公式
    this._updateFormulasBelow(targetRow + rowCount, rowCount);

    // 5. 处理条件格式
    this._handleConditionalFormatting(startRow, endRow, targetRow, rowCount);
  }

  /**
   * 将指定行及以下的所有行向下移动
   */
  _shiftRowsDown(startRow, shiftCount) {
    this.worksheet.spliceRows(startRow, 0, ...Array(shiftCount).fill(null));
  }

  /**
   * 复制单行
   */
  _copyRow(sourceRowNum, targetRowNum) {
    const sourceRow = this.worksheet.getRow(sourceRowNum);
    const targetRow = this.worksheet.getRow(targetRowNum);

    // 复制行高
    targetRow.height = sourceRow.height;

    sourceRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const targetCell = targetRow.getCell(colNumber);
      
      // 复制值和公式
      if (cell.formula) {
        targetCell.value = {formula: this._adjustFormula(cell.formula, sourceRowNum, targetRowNum), value: cell.value};
      } else {
        targetCell.value = cell.value;
      }

      // 复制样式
      targetCell.style = JSON.parse(JSON.stringify(cell.style));
    });
  }

  /**
   * 处理合并单元格
   */
  _handleMerges(startRow, endRow, targetRow) {
    const rowDiff = targetRow - startRow;
    
    // 获取所有合并单元格
    const merges = Object.keys(this.worksheet._merges || {}).map(key => {
      const range = this.worksheet._merges[key];
      return {
        top: range.top,
        left: range.left,
        bottom: range.bottom,
        right: range.right
      };
    });

    // 先解除目标区域的所有合并单元格
    merges.forEach(range => {
      const newTop = range.top + rowDiff;
      const newBottom = range.bottom + rowDiff;
      
      try {
        this.worksheet.unMergeCells(
          newTop,
          range.left,
          newBottom,
          range.right
        );
      } catch (e) {
        // 忽略未合并的单元格错误
      }
    });

    // 然后添加新的合并单元格
    merges.forEach(range => {
      // 检查合并单元格是否在源范围内
      if (range.top >= startRow && range.bottom <= endRow) {
        const newTop = range.top + rowDiff;
        const newBottom = range.bottom + rowDiff;
        
        try {
          // 添加新的合并单元格
          this.worksheet.mergeCells({
            top: newTop,
            left: range.left,
            bottom: newBottom,
            right: range.right
          });
        } catch (e) {
          console.warn('Failed to merge cells:', e.message);
        }
      }
    });
  }

  /**
   * 调整公式中的行号引用
   */
  _adjustFormula(formula, sourceRow, targetRow) {
    const rowDiff = targetRow - sourceRow;
    
    // 使用正则表达式查找并替换公式中的行号
    return formula.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
      const newRow = parseInt(row) + rowDiff;
      return `${col}${newRow}`;
    });
  }

  /**
   * 更新指定行之后的所有公式
   */
  _updateFormulasBelow(startRow, shiftCount) {
    // 获取工作表的使用范围
    const range = this.worksheet.dimensions;
    if (!range) return;

    // 遍历startRow之后的所有行
    for (let row = startRow; row <= range.bottom; row++) {
      const currentRow = this.worksheet.getRow(row);
      
      currentRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (cell.formula) {
          // 调整公式中的行号引用，所有行号都增加shiftCount
          const newFormula = cell.formula.replace(/([A-Z]+)(\d+)/g, (match, col, rowNum) => {
            const row = parseInt(rowNum);
            // 只调整startRow及之后的行号引用
            if (row >= startRow - shiftCount) {
              return `${col}${row + shiftCount}`;
            }
            return match;
          });
          
          // 设置新公式
          cell.value = {formula: newFormula, value: cell.value}
        }
      });
    }
  }

  /**
   * 处理条件格式
   */
  _handleConditionalFormatting(startRow, endRow, targetRow, rowCount) {
    if (!this.worksheet.conditionalFormattings) return;

    const newFormattings = [];
    
    // 遍历所有条件格式
    this.worksheet.conditionalFormattings.forEach(cf => {
      const refs = cf.ref.split(' ');  // 可能有多个引用范围
      const newRefs = [];

      refs.forEach(ref => {
        // 先尝试匹配范围格式 (A1:B2)
        let rangeMatch = ref.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        
        if (rangeMatch) {
          // 处理范围引用
          const [, startCol, startRowNum, endCol, endRowNum] = rangeMatch;
          const cfStartRow = parseInt(startRowNum);
          const cfEndRow = parseInt(endRowNum);

          // 1. 如果条件格式范围在复制的源范围内，创建新的条件格式
          if (cfStartRow >= startRow && cfEndRow <= endRow) {
            const newStartRow = cfStartRow + (targetRow - startRow);
            const newEndRow = cfEndRow + (targetRow - startRow);
            newRefs.push(`${startCol}${newStartRow}:${endCol}${newEndRow}`);
          }

          // 2. 如果条件格式范围在插入位置之后，更新行号
          if (cfStartRow >= targetRow) {
            const newStartRow = cfStartRow + rowCount;
            const newEndRow = cfEndRow + rowCount;
            newRefs.push(`${startCol}${newStartRow}:${endCol}${newEndRow}`);
          } else {
            // 保留原来的范围
            newRefs.push(ref);
          }
        } else {
          // 尝试匹配单个单元格格式 (A1)
          const cellMatch = ref.match(/([A-Z]+)(\d+)/);
          if (cellMatch) {
            const [, col, rowNum] = cellMatch;
            const cfRow = parseInt(rowNum);

            // 1. 如果单元格在复制的源范围内，创建新的引用
            if (cfRow >= startRow && cfRow <= endRow) {
              const newRow = cfRow + (targetRow - startRow);
              newRefs.push(`${col}${newRow}`);
            }

            // 2. 如果单元格在插入位置之后，更新行号
            if (cfRow >= targetRow) {
              const newRow = cfRow + rowCount;
              newRefs.push(`${col}${newRow}`);
            } else {
              // 保留原来的引用
              newRefs.push(ref);
            }
          } else {
            // 如果既不是范围也不是单个单元格，保留原引用
            newRefs.push(ref);
          }
        }
      });

      if (newRefs.length > 0) {
        // 创建新的条件格式
        const newCf = {
          ...cf,
          ref: newRefs.join(' ')
        };

        // 更新条件格式中的公式（如果有）
        if (newCf.formula) {
          newCf.formula = this._adjustFormula(newCf.formula, startRow, targetRow);
        }
        if (newCf.formulae) {
          newCf.formulae = newCf.formulae.map(f => 
            f ? this._adjustFormula(f, startRow, targetRow) : f
          );
        }

        newFormattings.push(newCf);
      }
    });

    // 清除原有的条件格式
    this.worksheet.conditionalFormattings = [];

    // 添加更新后的条件格式
    newFormattings.forEach(cf => {
      this.worksheet.addConditionalFormatting(cf);
    });
  }

  /**
   * 复制单元格的样式
   * @param {*} sourceCell 
   * @param {*} targetCell 
   */
  _copyCellStyle(sourceCell, targetCell) {
    // ... implementation ...
  }

  /**
   * 复制条件格式
   * @param {*} startRow 
   * @param {*} endRow 
   * @param {*} targetRow 
   */
  _copyConditionalFormatting(startRow, endRow, targetRow) {
    // ... implementation ...
  }
}

export { ExcelCopier as default }; 