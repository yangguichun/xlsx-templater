import { TagUtil } from './TagUtil.js';
import { default as AtTagHandler } from './AtTagHandler.js';
import { default as ExcelUtil } from './ExcelUtil.js';
/**
 * 用于处理所有循环 {#xxxx}
 * 允许在内部的单元格嵌套 @标记， 
 * 允许在内部的单元格嵌套 innerloop
 * 但是不能嵌套普通循环
 */
class LoopRowHandler{
  constructor(data){    
    this._outerData = data
    this._reset()
  }
  
  _reset(){
    // 开始结束标记之间的单元格，包含头尾
    this._containedCellList = []
    this._startCell = null
    this._endCell = null

    // 标记 {tag: 'defects', tagH: '{#defects}'}
    this._wrapTag = null
    this._innerData = []
  }

  async _handleLoop(row, loopInfo){
    // console.log(row.number)
    // console.log(row.worksheet)
    if(loopInfo.tagName in this._outerData){
      // 根据循环标记，从outerData中找到要循环的数据
      let data = this._outerData[loopInfo.tagName]
      if(data == null){
        data = []
      }      
      // 如果只是个对象，则把他格式化为数组
      // 有了这句话，就相当于在当行内实现了@标记，也就是对一个对象应用了循环标记，就相当于对一个长度为1的数组应用循环标记。
      // 这个与@标记的区别就在于，@标记支持多行，而循环标记应用在对象的时候只支持单行。
      if(!Array.isArray(data)){
        data = [data]
      }
      this._innerData = data
    }
        
    let worksheet = row.worksheet
    if(this._innerData.length == 0){
      // 如果没有数据，则行也删除了
      worksheet.spliceRows(row.number, 1)      
      ExcelUtil.adjustConditionFormatters(row.number, -1, worksheet, 'del')
    }
    if(this._innerData.length>1){
      // worksheet.duplicateRow(row.number, this._innerData.length-1, true)
      ExcelUtil.dupliateRowAndCopyStyle(worksheet, row.number, this._innerData.length-1)      
    }
    
    let tagName = loopInfo.tagName
    // console.log('startCell', loopInfo.startCell.address, loopInfo.startCell.col, loopInfo.endCell.col)
    let startIndex = loopInfo.startCell.col
    let endIndex = loopInfo.endCell.col
    // if(this._innerData.length == 0){
    //   new LoopRowReplacer(row, {}, tagName, startIndex, endIndex)
    // }else{
      for(let index=0; index<this._innerData.length; ++index){        
        let newRow = worksheet.getRow(row.number + index)         
        let replacer = new LoopRowReplacer(newRow, this._innerData[index], tagName, startIndex, endIndex)
        await replacer.handle()
      }
    // }
  }

  async handle(row){
    // 首先把这一行过一遍，看看是不是loop row
    // 如果不是，则返回1，到下一行
    // 如果是，则根据tag找到循环的数据
    // 然后遍历该数据，对于每个数据项复制一行，然后这个数据项对这一行数据做处理    
    // if(row.number == 34){
    //   console.log(row)
    // }
    let loopInfo = TagUtil.isLoopRow(row)
    if(!loopInfo){
      return 1
    }

    // console.log('row', row)
    await this._handleLoop(row, loopInfo)
    // 如果数据是空，则返回1
    let count = this._innerData.length>0?this._innerData.length: -1
    this._reset()
    return count
  }
}

/**
 * 将指定的行内的loopTag内的所有标记用data的数据替换
 * 允许有innerloop
 * 允许有普通标记
 */
class LoopRowReplacer{
  constructor(row, data, loopTagName, startColumnIndex, endColumnIndex){
    this._row = row
    this._data = data
    this._loopTagName = loopTagName
    this._startColumnIndex =  startColumnIndex
    this._endColumnIndex =  endColumnIndex
    this._containedCellList = []
  }  

  /**
   * 将首尾的loopTag去掉
   * 然后把loopTag范围内的cell添加到 _containedCellList
   */
  _formatCellAndFillContainedCellList(){
    for(let i = this._startColumnIndex; i<=this._endColumnIndex; ++i){
      let cell = this._row.getCell(i)
      if(i == this._startColumnIndex){
        cell.value = cell.value.replace(`{#${this._loopTagName}}`, '')
      }
      if(i == this._endColumnIndex){
        cell.value = cell.value.replace(`{/${this._loopTagName}}`, '')
      }
      this._containedCellList.push(cell)
    }
  }

  async handle(){    
    this._formatCellAndFillContainedCellList()    
    await this._handleAtTag()
    await TagUtil.replaceCellList(this._containedCellList, this._data)
  }
  async _handleAtTag(){
    let atTagHandler = new AtTagHandler(this._data)
    for(let cell of this._containedCellList){
      await atTagHandler.next(cell)
    }
  }
}
export { LoopRowHandler as default };