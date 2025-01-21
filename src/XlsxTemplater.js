import ExcelJS from 'exceljs'
import { TagUtil } from './TagUtil.js'
import { default as ExcelUtil } from './ExcelUtil.js'
import { default as AtTagHandler } from './AtTagHandler.js'
import { default as LoopRowHandler } from './LoopRowHandler.js'
import { default as MultiLineLoopHandler } from './MultiLineLoopHandler.js'

/**
 * 
 */
class XlsxTemplater {
  constructor(worksheet, data) {
    this._worksheet = worksheet
    this._data = data
  }
  /**
   * 使用data渲染 filePath路径下的文件，并返回渲染后的workbook   
   * @param {*} filePath 
   * @param {*} data 
   * @param {*} worksheetNameList ，要渲染的worksheet名称列表，如果不指定，默认就渲染第一个worksheet
   * @returns 渲染后的workbook
   */
  static async renderFromFile(filePath, data, worksheetNameList=[]){
    let workbook = new ExcelJS.Workbook()
    // workbook.getWorksheet().addConditionalFormatting
    await workbook.xlsx.readFile(filePath)
    await XlsxTemplater._findWorksheetAndRender(workbook, data, worksheetNameList)
    return workbook
  }
  /**
   * 解析从buffer中读取的Excel文件，用data渲染，然后再返回为buffer
   * @param {*} buffer 
   * @param {*} data 
   * @param {*} worksheetNameList ，要渲染的worksheet名称列表，如果不指定，默认就渲染第一个worksheet
   * @returns 返回一个Buffer
   */
    static async renderFromBuffer(buffer, data, worksheetNameList=[]){
      let workbook = new ExcelJS.Workbook()
      await workbook.xlsx.load(buffer)
      await XlsxTemplater._findWorksheetAndRender(workbook, data, worksheetNameList)    
      return await workbook.xlsx.writeBuffer()
    }
  /**
   * 
   * @param {*} workbook，要渲染的workbook
   * @param {*} data，用于渲染的json格式数据
   * @param {*} worksheetNameList ，要渲染的worksheet名称列表
   */
  static async _findWorksheetAndRender(workbook, data, worksheetNameList){
    let renderWorksheetList = []
    if(worksheetNameList == null || worksheetNameList === undefined || worksheetNameList.length == 0){
      renderWorksheetList = [workbook.worksheets[0]]
    }else{
      workbook.worksheets.forEach(item=>{
        if(worksheetNameList.indexOf(item.name)>=0){
          renderWorksheetList.push(item)
        }
      })
    }
    for(let i = 0; i<renderWorksheetList.length; i++){
      let templater = new XlsxTemplater(renderWorksheetList[i], data)
      await templater.render()
    }   
  }
  /**
   * 便利数据表内的所有单元格
   */
  _renderNormalTag() {
    this._worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (ExcelUtil.isMergedCell(cell)) {
          return
        }
        TagUtil.replaceCellNormalTag(cell, this._data)
      })
    })
  }

  /**
   * 处理被 {@xxx} ... {/xxx}包括的所有单元格，
   * 在这个内部，
   * 不允许嵌套普通循环，
   * 允许嵌套@标记
   * 允许嵌套单元格内的循环
   */
  async _renderAtTag() {
    let tagHandler = new AtTagHandler(this._data)
    let worksheet = this._worksheet
    for (let rowIndex = 1; rowIndex <= worksheet.rowCount; ++rowIndex) {
      let row = worksheet.getRow(rowIndex)
      let cellCount = row.cellCount
      for (let columnIndex = 1; columnIndex <= cellCount; ++columnIndex) {
        let cell = row.getCell(columnIndex)
        await tagHandler.next(cell)
      }
    }
  }

  /**
   * 替换掉最顶层的所有单元格内循环
   */
  async _renderInnerLoopTag() {
    let worksheet = this._worksheet
    for (let rowIndex = 1; rowIndex <= worksheet.rowCount; ++rowIndex) {
      let row = worksheet.getRow(rowIndex)
      let cellCount = row.cellCount
      for (let columnIndex = 1; columnIndex <= cellCount; ++columnIndex) {
        let cell = row.getCell(columnIndex)
        await TagUtil.replaceCellInnerLoopTag(cell, this._data)
      }
    }
  }

  /**
   * 替换整个worksheet的图片标记
   */
  async _renderImageTag() {
    let worksheet = this._worksheet
    for (let rowIndex = 1; rowIndex <= worksheet.rowCount; ++rowIndex) {
      let row = worksheet.getRow(rowIndex)
      let cellCount = row.cellCount
      for (let columnIndex = 1; columnIndex <= cellCount; ++columnIndex) {
        let cell = row.getCell(columnIndex)
        await TagUtil.replaceImageCellTag(cell, this._data)
      }
    }
  }

  async _renderLoopTag() {
    // console.log('_renderLoopTag begin');
    let loopHandler = new LoopRowHandler(this._data)
    let worksheet = this._worksheet
    for (let rowIndex = 1; rowIndex <= worksheet.rowCount;) {
      let row = worksheet.getRow(rowIndex)
      rowIndex += await loopHandler.handle(row)
      // console.log('rowIndex', rowIndex, worksheet.rowCount)            
    }
  }

  /**
   * 处理多行循环标记
   */
  async _renderMultiLineLoopTag() {
    // console.log('_renderLoopTag begin');
    let multiLineLoopHandler = new MultiLineLoopHandler(this._worksheet, this._data)    
    await multiLineLoopHandler.handle()    
  }

  /**
   * 总的入口
   */
  async render(){
    await this._renderMultiLineLoopTag()
    await this._renderLoopTag()
    await this._renderAtTag()
    await this._renderInnerLoopTag()
    this._renderNormalTag()
    await this._renderImageTag()
  }
}

export { XlsxTemplater as default };