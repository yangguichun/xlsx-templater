let TagUtil = require('./TagUtil')
let ExcelUtil = require('./ExcelUtil')
/**
 * 用于处理所有@标记{@xxx}...{/xxx}
 * @标记 主要用于方便定位一个对象内部的属性。
 * 允许在内部嵌套innerloop
 * 但是不能嵌套普通循环
 */
class AtTagHandler{
  constructor(data){    
    this._containedCellList = []
    this._wrapTag = null
    this._outerData = data
    this._innerData = {}
  }
  
  _reset(){
    this._containedCellList = []
    this._wrapTag = null
    this._innerData = {}
  }
  // 判断是否有嵌套的@标记
  _hasSubAtTag(){
    for(let cell of this._containedCellList){
      if(TagUtil.getAtStartTag(cell.value).tag){
        return true
      }
    }
    return false
  }
  async _handle(){
    if(this._hasSubAtTag()){
      let tagHandler = new AtTagHandler(this._innerData)
      for(let cell of this._containedCellList){
        await tagHandler.next(cell)
      }
    }
    await TagUtil.replaceCellList(this._containedCellList, this._innerData)
  }


  async next(cell){
    if (ExcelUtil.isMergedCell(cell)) {
      return
    }
    if(!this._wrapTag){
      let startTag = TagUtil.getAtStartTag(cell.value)
      if(startTag.tag){
        // console.log('startTag', startTag)
        this._wrapTag = startTag
        cell.value = cell.value.replace(startTag.tagH, '')
      }
    }    

    if(this._wrapTag){          
      let endTag = TagUtil.getEndTag(cell.value, this._wrapTag.tag)
      if(endTag.tag){
        cell.value = cell.value.replace(endTag.tagH, '')
      }
      this._containedCellList.push(cell)
      if(endTag.tag){
        // 找到结束标记之后，说明找到了所有@标记内的单元格，可以开始执行内部替换了
        this._innerData = {}
        if(this._wrapTag.tag in this._outerData){
          this._innerData = this._outerData[this._wrapTag.tag]
        }
        await this._handle()        
        // 这一组替换之后，要充值缓存，继续寻找下一组
        this._reset()
      }
    }
  }
}

module.exports = AtTagHandler