let ImageUtil = require('./ImageUtil')
let ExcelUtil = require('./ExcelUtil')
/**
 * 封装了用于匹配各种tag的正则
 */
 class TagUtil {
  
  /**
   * 将这个数组内的所有标记都替换了
   * 这个数组可能来自于行循环 {#}，也可能来自于At标记 {@}
   * @param {*} cellList 
   * @param {*} data 
   */
   static async replaceCellList(cellList, data){
    for(let cell of cellList){
      // 如果这个不先替换，则innerloop内部的normalTag会被replaceCellNormalTag替换掉
      // 所以可以认为innerloop的优先级更高
      await TagUtil.replaceCell(cell, data)      
    }
  }

  /**
   * 替换单个单元格内的所有标记
   * @param {*} cell 
   * @param {*} data 
   */
  static async replaceCell(cell, data){    
    // 如果这个不先替换，则innerloop内部的normalTag会被replaceCellNormalTag替换掉
    // 所以可以认为innerloop的优先级更高
    await TagUtil.replaceCellInnerLoopTag(cell, data)
    TagUtil.replaceCellNormalTag(cell, data)
    await TagUtil.replaceImageCellTag(cell, data)    
  }

  /**
   * 替换该单元格内的图片标记
   * 对于图片url，会从网络上读取相关的图片
   * @param {*} cell 
   * @param {*} data 
   * @param {*} imageCount, 这个单元格中总共有几个图片，主要针对innerloop的情况
   * @param {*} imageIndexInCell, 该图片是单元格中的第几个
   */
   static async replaceImageCellTag(cell, data, imageCount=1, imageIndexInCell=0){
    let res= TagUtil.getImageTag(cell.value)
    if (res.tag) {      
      let workbook = cell.worksheet.workbook
      for(let index=0; index<res.tag.length; index++){
        let tag = res.tag[index]      
        let imageUrl = ''
        if (tag in data) {
          imageUrl = data[tag]
          try {
            let imageData = await ImageUtil.getImageData(imageUrl)
            let imageId = workbook.addImage({
              buffer: imageData,
              extension: ImageUtil.getImageExt(imageUrl)
            })
            cell.value = cell.value.replace(res.tagH[index], '')
            let rc = ExcelUtil.getRowColumn(cell.address)                 
            // 实现的效果就是，图片始终限制在这个单元格内，
            // 右下角对齐，第一章图片撑满，后面每个图片都缩小20%
            let columnWidthAdjust = 0.2*imageIndexInCell
            let rowWidthAdjust = 0.2*imageIndexInCell
            let param = {
              tl: {col: rc.c-1 + columnWidthAdjust, row: rc.r-1 + rowWidthAdjust}, 
              br: {col: rc.c, row: rc.r}
            }
            // console.log('image replacer param', cell.address, param)            
            cell.worksheet.addImage(imageId, param)            
            // cell.worksheet.addImage(imageId, `${cell.address}:${cell.address}`)
          } catch (error) {
            console.error('replace image faild', imageUrl, error)
            continue
          }
        }
      }
    }
  }
  /**
   * 把单个单元格内的normalTag都替换为对应的数据
   * @param {*} cell，要替换的单元格 
   * @param {*} data，数据寻找的范围，有这个之后，就可以在循环内的单元格中使用，只要传入循环的数据即可
   */
   static replaceCellNormalTag(cell, data) {
    let res = TagUtil.getNormalTag(cell.value)
    if (res.tag) {      
      res.tag.forEach((tag, index)=>{
        let tagVal = ''
        
        if (typeof data === 'object' && tag in data) {
          tagVal = data[tag]
          // 如果是数组，则特殊处理下
          if(Array.isArray(tagVal)){
            tagVal = tagVal.join(',')
          }
          cell.value = cell.value.replace(res.tagH[index], tagVal)
        }
      })
    }
  }

  /**
   * 判断当前cell中是否包含innerloop，如果包含，就用data中的数据来替换
   * @param {*} cell 
   * @param {*} data 
   * data中与loop tag对应的数据可以是数组，也可以是对象，像下面这样
   * defects:[{}, {}]
   * defects: {}
   */
  static async  replaceCellInnerLoopTag(cell, data){
    if(!ExcelUtil.isMergedCell(cell)){
      let res = TagUtil.getInnerLoopTag(cell.value)
      if(res.tag){            
        let loopData = []
        if(res.tag in data){
          loopData = data[res.tag]
          if(!Array.isArray(loopData)){
            // 如果他不是个数组，那就把它装到数组内
            loopData = [loopData]
          }
        }else{
          // 如果这个标记不是当前data的，就跳过，不做任何处理
          return
        }
        
        let value = ''
        for(let i= 0; i < loopData.length; i++){
          let dataItem = loopData[i]        
          // 去掉innerloop的标签
          cell.value = res.tagInner
          TagUtil.replaceCellNormalTag(cell, dataItem)
          await TagUtil.replaceImageCellTag(cell, dataItem, loopData.length, i)
          value += cell.value
          cell.value = res.tagInner
        }
        cell.value = value
      }
    }
  }
  /**
   * 用于匹配基础标记，支持一个字符串内包含多个基础标识
   * @param {*} value ，值的格式
   * {contactName}
   * {@basic}{contactName}
   * {@basic}{index}.{rectifyPlan};{/basic}       
   * @returns {tag, tagH}，tag和tagH都是数组
   * tag是不包含括号的，如 ['index', 'rectifyPlan']
   * tagH是包含括号和标记的，如 ['{index}', '{rectifyPlan}']
   */
  static getNormalTag(value) {    
    // let pattern = /\{(.+?)\}/g
    // let pattern = /(?<=\{)(.+?)(?=\})/g
    let pattern = /(\{([^#/%@][^\{\}]+)\})/g
    let matches = pattern.exec(value)
    let res = {tag: [], tagH: []}
    while(matches && matches.length){      
      // console.log('matches', matches);
      res.tag.push(matches[2])
      res.tagH.push(matches[1])
      matches = pattern.exec(value)
    }
    if(res.tag.length){
      return res
    }else{
      return {}
    }
  }

  /**
   * 用于匹配图片标记，支持一个单元格内包含多个图片标记
   * @param {*} value，单元格的值，格式包括
   * {%imagTag} 
   * {%imageTag}{normalTag}
   * {@basic}{%imageTag}{/basic}
   * @returns，tag和tagH都是数组
   * tag是不包含括号的，如 ['beforePicUrl', 'afterPicUrl']
   * tagH是包含括号和标记的，如 ['{%beforePicUrl}', '{%afterPicUrl}']
   */
  static getImageTag(value){
    // let pattern = /\{%(.+)\}/g
    // let pattern = /(?<=\{%)(.+?)(?=\})/g
    let pattern = /(\{%([^\{\}]+)\})/g
    let matches = pattern.exec(value)
    let res = {tag: [], tagH: []}
    while(matches && matches.length){      
      // console.log('matches', matches);
      res.tag.push(matches[2])
      res.tagH.push(matches[1])
      matches = pattern.exec(value)
    }
    if(res.tag.length){
      return res
    }else{
      return {}
    }
  }
  
  /**
   * 用于匹配对象标记
   * @param {} value ，格式为  
   * {@basic}
   * {@basic}{contactName}
   * @returns 
   */
  static getAtStartTag(value) {
    // let pattern = /\{\@(.+)\}/g
    // let pattern = /(?<=\{@)(.+?)(?=\})/g
    let pattern = /(\{@([^\{\}]+)\})/g
    let matches = pattern.exec(value)
    if (matches && matches.length) {
      // console.log('getAtStartTag', matches)
      return {tag: matches[2], tagH: matches[1]} 
    }
    return {}
  }

  /**
   * 获取结束标记，可以是AtTag的标记，也可以是循环tag的标记
   * @param {*} value ，格式是 {/xxxx}
   * @param {*} tagName 
   * @returns 
   */
  static getEndTag(value, tagName) {
    // let pattern = /\{\/(\w+)\}/ig
    let patternStr = `\\{\\/(${tagName})\\}`
    let pattern = new RegExp(patternStr, 'ig')    
    let matches = pattern.exec(value)
    if (matches && matches.length) {
      return {tag: matches[1], tagH: matches[0]} 
    }
    return {}
  }

  /**
   * 判断当前行是否有循环标记
   * @param {*} row 
   * @returns 
   */
  static isLoopRow(row){
    let startCell = null
    let tagName = ''
    let endCell = null
    for(let index = 1; index <=row.cellCount; ++index){
      let cell = row.getCell(index)
      if(ExcelUtil.isMergedCell(cell)){
        continue
      }
      if(!startCell){
        let startRes = TagUtil.getLoopTag(cell.value)
        if(startRes.tag){
          startCell = cell
          tagName = startRes.tag
          // continue
        }
      }
      if(startCell){
        let endRes = TagUtil.getEndTag(cell.value, tagName)
        if(endRes.tag){
          endCell = cell
        }
      }      
    }
    // 开始结尾都找到，就说明当前行包含了loop
    if(startCell && endCell){
      return {
        // 有了这个cell，可以通过 cell.columnIndex找到该单元格所属的列
        startCell, endCell, tagName
      }
    }
    return null
  }
   /**
   * 用于匹配循环标记
   * @param {} value ，格式为  
   * {@basic}
   * {@basic}{contactName}
   * @returns 
   */
  static getLoopTag(value) {
    // let pattern = /\{\#(.+)\}/g
    // let pattern = /(?<=\{#)(.+?)(?=\})/g    
    let pattern = /(\{#([^\{\}]+)\})/g
    let matches = pattern.exec(value)
    if (matches && matches.length) {
      // console.log('getLoopTag', matches)
      return {tag: matches[2], tagH: matches[1]} 
    }
    return {}
  }
  
  /**
   * 判断是否为行内循环标记
   * @param {*} value 
   * @returns 
   */
  static getInnerLoopTag(value) {
    // 参考了这里L: https://blog.csdn.net/u013299635/article/details/125717591，这个勉强也能实现，但是不够好，已经废弃了
    // let pattern = /((?<=\{)#(.+?)(?=\}))(.+)(\{\/\})/g
    // 下面这个实现，参考了《正则表达式经典实例(第二版)》5.4查找某个单词之外的任意单词，其中有一句话这么说：[^cat]是一个合法的正则式，但是它会匹配除了c、a或t之外的任意字符。
    // 那我只要把cat换成\{\}就可以了
    // 返回的matches中，
    // 第一个是整个正则表达式匹配的内容
    // 第二个是第一个左括号匹配的内容
    // 第三个是第二个左括号匹配的内容
    // 以此类推
    /**
     * 比如对于这个输入：'{@outer}{#中文}{attach}--{haha};{/}{/outer}'
     * 对应的matches是这样
     * [
        '{#中文}{attach}--{haha};{/}',
        '{#中文}',
        '中文',
        '{attach}--{haha};',
        '{/}',
        index: 8,
        input: '{@outer}{#中文}{attach}--{haha};{/}{/outer}',
        groups: undefined
      ]
     */
    let pattern = /(\{#([^\{\}]+)\})(.+)(\{\/\})/g
    let matches = pattern.exec(value)
    if (matches && matches.length) {
      // console.log('getInnerLoopTag', matches)      
      return {tag: matches[2], tagH: matches[1], tagInner: matches[3]} 
    }
    return {}
  }
}

module.exports = TagUtil