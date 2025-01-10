import cloneDeep from 'lodash/cloneDeep';

class ExcelUtil {
  /**
   * 判断指定的单元是否是被合并的单元格
   * @param {*} cell 
   * @returns boolean
   */
  static isMergedCell(cell) {
    return cell != cell.master
  }

  /**
  * 将一个常规的excel范围字符串，格式化为一个数字格式的行列对象
  * @param {*} range 类似于 A12:D33这样的字符串
  * @return {row: {start: 12, end: 33}, column: {start: 1, end: 4}}
  */
  static convertRangeToRowColumnSpan(range) {
    let cellList = range.split(':')
    let lt = this.getRowColumn(cellList[0])
    let rb = this.getRowColumn(cellList[1])
    return {
      row: {
        start: lt.r,
        end: rb.r
      },
      column: {
        start: lt.c,
        end: rb.c
      }
    }
  }

  /**
   * 将Excel单元格地址转换为行列数值
   * @param {*} address，类似于A2, C33这样的Excel单元格地址
   * @returns，格式为 {r: 2, c: 1}
   */
  static getRowColumn(address) {
    let column = address.substring(0, 1).toUpperCase()
    let columnCode = column.codePointAt(0) - 'A'.codePointAt(0) + 1
    let row = parseInt(address.substring(1))
    return {
      r: row,
      c: columnCode
    }
  }

  /**
   * 将Excel单元格地址转换为行列的字符串
   * @param {*} address，类似于A2, C33这样的Excel单元格地址
   * @returns，格式为 {c: 'C', r: '33'}
   */
  static splitRowColumn(address) {
    let column = address.substring(0, 1).toUpperCase()
    let row = address.substring(1)
    return {
      r: row,
      c: column
    }
  }

  /**
   * 给指定的条件格式的ref和公式内部所引用的单元格添加指定的行
   * @param {*} cd，条件格式对象   
   * @param {*} rowOffset，要添加的行数 
   */
  static _offsetConditionFormular(cd, rowOffset){
    let origRowIndex = 0
    // 修改ref
    let newRef = cd.ref.split(':').map(item=>{
      let rc = ExcelUtil.splitRowColumn(item)
      origRowIndex = parseInt(rc.r)
      return rc.c + (parseInt(rc.r) + rowOffset)
    }).join(':')
    cd.ref = newRef
    
    // 修改rules的内容
    let newRowIndex = origRowIndex + rowOffset
    cd.rules.forEach(rule => {
      rule.formulae = rule.formulae.map(formulae=>{
        let regx = /\$\d+/ig        
        return formulae.replace(regx, `$${newRowIndex}`)
      })
    });
  }
  
  /**
   * 根据增加或者删除行，调整条件格式的样式
   * 这里假设条件格式是针对同一行的单元格的，不会跨行，包括ref和 rules->formulae的内容都不会跨行
   * @param {*} rowIndex，要调整的行号，从1开始，如果action是add，则这个rowIndex是要复制的源行号，如果action是del，则是要删除的行号
   * @param {*} count, 如果是del，则count是负数，如果是add，则count是正数，表示要添加的行数
   * @param {*} action ，动作，可以还add, del
   */
  static adjustConditionFormatters(rowIndex, count, worksheet, action='add'){
    console.log('adjustConditionFormatters', rowIndex, count, action)
    if(!worksheet.conditionalFormattings || !worksheet.conditionalFormattings.length){
      return
    }

    if(action == 'del'){
      worksheet.conditionalFormattings
      // 删除ref包含这一行的条件格式
      for(let i = worksheet.conditionalFormattings.length-1; i>=0; i--){
        let cd = worksheet.conditionalFormattings[i]        
        let currentRowIndex = ExcelUtil._getConditionFormatterRowIndex(cd)
        if(currentRowIndex == rowIndex){
          worksheet.conditionalFormattings.splice(i, 1)
        }        
        // 对于行号大于要删除的行的条件格式，要调整他们内部的单元格地址的行号
        // 对于del操作，外面传进来的count是负数
        if(currentRowIndex > rowIndex){          
          ExcelUtil._offsetConditionFormular(cd, count)
        }
      }
    }else{
      // 主要做两件事情
      // 1. 找出rowIndex下方行的条件格式，将他们的ref和rules->formulae内的行号加上count
      for(let i = 0; i<worksheet.conditionalFormattings.length; i++){
        let cd = worksheet.conditionalFormattings[i]
        if(ExcelUtil._getConditionFormatterRowIndex(cd) > rowIndex){
          ExcelUtil._offsetConditionFormular(cd, count)
        }
      }
      // 2. 找出rowIndex对应行的条件格式，也复制count份
      //    对于每一份，调整他的ref和rules->formulae内的单元格的行号，复制的第一行就+1，复制的第二行就加2      
      for(let i = 0; i<worksheet.conditionalFormattings.length; i++){
        let cd = worksheet.conditionalFormattings[i]        
        if(ExcelUtil._getConditionFormatterRowIndex(cd) == rowIndex){
          for(let j = 0; j < count; j++){
            let newCd = cloneDeep(cd)          
            ExcelUtil._offsetConditionFormular(newCd, j+1)
            worksheet.conditionalFormattings.push(newCd)
          }          
        }        
      }
    }
  }
  
  /**
   * 获取指定的条件格式所属的行号
   * @param {*} cd 
   */
  static _getConditionFormatterRowIndex(cd){
    let refCellAddress = cd.ref.split(':')[0]
    let rc = ExcelUtil.getRowColumn(refCellAddress)
    return rc.r
  }

  /**
   * 复制行，并且复制行单元格的样式，包括哪些单元格和哪些合并
   * @param {*} worksheet 
   * @param {*} rowIndex 
   * @param {*} count 
   * @returns 
   */
  static dupliateRowAndCopyStyle(worksheet, rowIndex, count) {
    if (count <= 0) {
      return
    }

    worksheet.duplicateRow(rowIndex, count, true)
    let start = rowIndex + 1
    let end = rowIndex + count
    let originalRow = worksheet.getRow(rowIndex)
    for (let i = start; i <= end; ++i) {
      let row = worksheet.getRow(i)
      let mergeStart = ''
      let mergeEnd = ''
      let preCellAddress = ''

      for (let ci = 1; ci <= originalRow.cellCount; ++ci) {
        let originalCell = originalRow.getCell(ci)
        let cell = row.getCell(ci)        
        
        if(originalCell.isMerged){
          worksheet.unMergeCells(cell.address)
          // console.log('cell isMerged', cell.address, cell.isMerged, cell.master.address)
          if(!mergeStart){
            mergeStart = cell.address
          }
        }else{
          if(mergeStart){
            mergeEnd = preCellAddress
          }
          if(mergeStart && mergeEnd){
            // 做合并操作          
            worksheet.mergeCells(`${mergeStart}:${mergeEnd}`)
            mergeStart = ''
            mergeEnd = ''
          }
        }
        preCellAddress = cell.address
        // cell.master = rc1.c + i
        // cell.isMerged = originalCell.isMerged
        // TODO
      }
    }

    ExcelUtil.adjustConditionFormatters(rowIndex, count, worksheet)
  }
}

export { ExcelUtil as default };