import ExcelJS from 'exceljs';
import fetch from 'node-fetch';
import cloneDeep from 'lodash/cloneDeep';

/**
 * 图像的辅助工具
 */
class ImageUtil{
  /**
   * 通过url获取对应的图像文件数据
   * @param {*} url 必须是绝对路径
   * @returns 
   */
  static getImageData(url){
    let urlPattern = /^http[s]{0,1}:\/\/.+/i;
    if(!urlPattern.test(url)){
      return Promise.reject('url格式不正确，必须是绝对路径')
    }

    return new Promise((resolve, reject)=>{
      console.log('image loading', url);
      fetch(url).then((res) => {
        res.blob().then((blob) => {
          blob.arrayBuffer().then((arr)=>{
            console.log('image loaded', url);
            resolve(arr);
          });      
        });
      })
      .catch(err=>{
        reject(err);
      });
    })
  }

  /**
   * 获取指定图片url的后缀名
   * @param {*} url 例如
   * https://www.baidu.com/img/PCtm_d9c8750bed0b3c7d089fa7d55720d6cf.png
   * @returns png, jpg, jpeg之类的
   */
  static getImageExt(url){
    let pattern = /.+\.(\w+)$/i;
    let matches = pattern.exec(url);
    if(matches && matches.length>=2){
      return matches[1]
    }
    return 'jpg'
  }
}

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
    let cellList = range.split(':');
    let lt = this.getRowColumn(cellList[0]);
    let rb = this.getRowColumn(cellList[1]);
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
    let column = address.substring(0, 1).toUpperCase();
    let columnCode = column.codePointAt(0) - 'A'.codePointAt(0) + 1;
    let row = parseInt(address.substring(1));
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
    let column = address.substring(0, 1).toUpperCase();
    let row = address.substring(1);
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
    let origRowIndex = 0;
    // 修改ref
    let newRef = cd.ref.split(':').map(item=>{
      let rc = ExcelUtil.splitRowColumn(item);
      origRowIndex = parseInt(rc.r);
      return rc.c + (parseInt(rc.r) + rowOffset)
    }).join(':');
    cd.ref = newRef;
    
    // 修改rules的内容
    let newRowIndex = origRowIndex + rowOffset;
    cd.rules.forEach(rule => {
      rule.formulae = rule.formulae.map(formulae=>{
        let regx = /\$\d+/ig;        
        return formulae.replace(regx, `$${newRowIndex}`)
      });
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
    console.log('adjustConditionFormatters', rowIndex, count, action);
    if(!worksheet.conditionalFormattings || !worksheet.conditionalFormattings.length){
      return
    }

    if(action == 'del'){
      worksheet.conditionalFormattings;
      // 删除ref包含这一行的条件格式
      for(let i = worksheet.conditionalFormattings.length-1; i>=0; i--){
        let cd = worksheet.conditionalFormattings[i];        
        let currentRowIndex = ExcelUtil._getConditionFormatterRowIndex(cd);
        if(currentRowIndex == rowIndex){
          worksheet.conditionalFormattings.splice(i, 1);
        }        
        // 对于行号大于要删除的行的条件格式，要调整他们内部的单元格地址的行号
        // 对于del操作，外面传进来的count是负数
        if(currentRowIndex > rowIndex){          
          ExcelUtil._offsetConditionFormular(cd, count);
        }
      }
    }else {
      // 主要做两件事情
      // 1. 找出rowIndex下方行的条件格式，将他们的ref和rules->formulae内的行号加上count
      for(let i = 0; i<worksheet.conditionalFormattings.length; i++){
        let cd = worksheet.conditionalFormattings[i];
        if(ExcelUtil._getConditionFormatterRowIndex(cd) > rowIndex){
          ExcelUtil._offsetConditionFormular(cd, count);
        }
      }
      // 2. 找出rowIndex对应行的条件格式，也复制count份
      //    对于每一份，调整他的ref和rules->formulae内的单元格的行号，复制的第一行就+1，复制的第二行就加2      
      for(let i = 0; i<worksheet.conditionalFormattings.length; i++){
        let cd = worksheet.conditionalFormattings[i];        
        if(ExcelUtil._getConditionFormatterRowIndex(cd) == rowIndex){
          for(let j = 0; j < count; j++){
            let newCd = cloneDeep(cd);          
            ExcelUtil._offsetConditionFormular(newCd, j+1);
            worksheet.conditionalFormattings.push(newCd);
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
    let refCellAddress = cd.ref.split(':')[0];
    let rc = ExcelUtil.getRowColumn(refCellAddress);
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

    worksheet.duplicateRow(rowIndex, count, true);
    let start = rowIndex + 1;
    let end = rowIndex + count;
    let originalRow = worksheet.getRow(rowIndex);
    for (let i = start; i <= end; ++i) {
      let row = worksheet.getRow(i);
      let mergeStart = '';
      let mergeEnd = '';
      let preCellAddress = '';

      for (let ci = 1; ci <= originalRow.cellCount; ++ci) {
        let originalCell = originalRow.getCell(ci);
        let cell = row.getCell(ci);        
        
        if(originalCell.isMerged){
          worksheet.unMergeCells(cell.address);
          // console.log('cell isMerged', cell.address, cell.isMerged, cell.master.address)
          if(!mergeStart){
            mergeStart = cell.address;
          }
        }else {
          if(mergeStart){
            mergeEnd = preCellAddress;
          }
          if(mergeStart && mergeEnd){
            // 做合并操作          
            worksheet.mergeCells(`${mergeStart}:${mergeEnd}`);
            mergeStart = '';
            mergeEnd = '';
          }
        }
        preCellAddress = cell.address;
        // cell.master = rc1.c + i
        // cell.isMerged = originalCell.isMerged
        // TODO
      }
    }

    ExcelUtil.adjustConditionFormatters(rowIndex, count, worksheet);
  }
}

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
      await TagUtil.replaceCell(cell, data);      
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
    await TagUtil.replaceCellInnerLoopTag(cell, data);
    TagUtil.replaceCellNormalTag(cell, data);
    await TagUtil.replaceImageCellTag(cell, data);    
  }

  /**
   * 替换该��元格内的图片标记
   * 对于图片url，会从网络上读取相关的图片
   * @param {*} cell 
   * @param {*} data 
   * @param {*} imageCount, 这个单元格中总共有几个图片，主要针对innerloop的情况
   * @param {*} imageIndexInCell, 该图片是单元格中的第几个
   */
  static async replaceImageCellTag(cell, data, imageCount=1, imageIndexInCell=0){
    let res= TagUtil.getImageTag(cell.value);
    if (res.tag) {      
      let workbook = cell.worksheet.workbook;
      for(let index=0; index<res.tag.length; index++){
        let tag = res.tag[index];      
        let imageUrl = '';
        if (tag in data) {
          imageUrl = data[tag];
          try {
            let imageData = await ImageUtil.getImageData(imageUrl);
            let imageId = workbook.addImage({
              buffer: imageData,
              extension: ImageUtil.getImageExt(imageUrl)
            });
            cell.value = cell.value.replace(res.tagH[index], '');
            let rc = ExcelUtil.getRowColumn(cell.address);                 
            // 实现的效果就是，图片始终限制在这个单元格内，
            // 右下角对齐，第一章图片撑满，后面每个图片都缩小20%
            let columnWidthAdjust = 0.2*imageIndexInCell;
            let rowWidthAdjust = 0.2*imageIndexInCell;
            let param = {
              tl: {col: rc.c-1 + columnWidthAdjust, row: rc.r-1 + rowWidthAdjust}, 
              br: {col: rc.c, row: rc.r}
            };
            // console.log('image replacer param', cell.address, param)            
            cell.worksheet.addImage(imageId, param);            
            // cell.worksheet.addImage(imageId, `${cell.address}:${cell.address}`)
          } catch (error) {
            console.error('replace image faild', imageUrl, error);
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
    let res = TagUtil.getNormalTag(cell.value);
    if (res.tag) {      
      res.tag.forEach((tag, index)=>{
        let tagVal = '';
        
        if (typeof data === 'object' && tag in data) {
          tagVal = data[tag];
          // 如果是数组，则特殊处理下
          if(Array.isArray(tagVal)){
            tagVal = tagVal.join(',');
          }
          cell.value = cell.value.replace(res.tagH[index], tagVal);
        }
      });
    }
  }

  /**
   * 判断当前cell中是否包含innerloop，如果包含，就用data中的数据来替换
   * @param {*} cell 
   * @param {*} data 
   * data中与loop tag对应的数据可以是数组，也可以是对象，像��面这样
   * defects:[{}, {}]
   * defects: {}
   */
  static async  replaceCellInnerLoopTag(cell, data){
    if(!ExcelUtil.isMergedCell(cell)){
      let res = TagUtil.getInnerLoopTag(cell.value);
      if(res.tag){            
        let loopData = [];
        if(res.tag in data){
          loopData = data[res.tag];
          if(!Array.isArray(loopData)){
            // 如果他不是个数组，那就把它装到数组内
            loopData = [loopData];
          }
        }else {
          // 如果这个标记不是当前data的，就跳过，不做任何处理
          return
        }
        
        let value = '';
        for(let i= 0; i < loopData.length; i++){
          let dataItem = loopData[i];        
          // 去掉innerloop的标签
          cell.value = res.tagInner;
          TagUtil.replaceCellNormalTag(cell, dataItem);
          await TagUtil.replaceImageCellTag(cell, dataItem, loopData.length, i);
          value += cell.value;
          cell.value = res.tagInner;
        }
        cell.value = value;
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
    let pattern = /(\{([^#/%@][^\{\}]+)\})/g;
    let matches = pattern.exec(value);
    let res = {tag: [], tagH: []};
    while(matches && matches.length){      
      // console.log('matches', matches);
      res.tag.push(matches[2]);
      res.tagH.push(matches[1]);
      matches = pattern.exec(value);
    }
    if(res.tag.length){
      return res
    }else {
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
    let pattern = /(\{%([^\{\}]+)\})/g;
    let matches = pattern.exec(value);
    let res = {tag: [], tagH: []};
    while(matches && matches.length){      
      // console.log('matches', matches);
      res.tag.push(matches[2]);
      res.tagH.push(matches[1]);
      matches = pattern.exec(value);
    }
    if(res.tag.length){
      return res
    }else {
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
    let pattern = /(\{@([^\{\}]+)\})/g;
    let matches = pattern.exec(value);
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
    let patternStr = `\\{\\/(${tagName})\\}`;
    let pattern = new RegExp(patternStr, 'ig');    
    let matches = pattern.exec(value);
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
    let startCell = null;
    let tagName = '';
    let endCell = null;
    for(let index = 1; index <=row.cellCount; ++index){
      let cell = row.getCell(index);
      if(ExcelUtil.isMergedCell(cell)){
        continue
      }
      if(!startCell){
        let startRes = TagUtil.getLoopTag(cell.value);
        if(startRes.tag){
          startCell = cell;
          tagName = startRes.tag;
          // continue
        }
      }
      if(startCell){
        let endRes = TagUtil.getEndTag(cell.value, tagName);
        if(endRes.tag){
          endCell = cell;
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
    let pattern = /(\{#([^\{\}]+)\})/g;
    let matches = pattern.exec(value);
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
    let pattern = /(\{#([^\{\}]+)\})(.+)(\{\/\})/g;
    let matches = pattern.exec(value);
    if (matches && matches.length) {
      // console.log('getInnerLoopTag', matches)      
      return {tag: matches[2], tagH: matches[1], tagInner: matches[3]} 
    }
    return {}
  }
}

/**
 * 用于处理所有@标记{@xxx}...{/xxx}
 * @标记 主要用于方便定位一个对象内部的属性。
 * 允许在内部嵌套innerloop
 * 但是不能嵌套普通循环
 */
class AtTagHandler{
  constructor(data){    
    this._containedCellList = [];
    this._wrapTag = null;
    this._outerData = data;
    this._innerData = {};
  }
  
  _reset(){
    this._containedCellList = [];
    this._wrapTag = null;
    this._innerData = {};
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
      let tagHandler = new AtTagHandler(this._innerData);
      for(let cell of this._containedCellList){
        await tagHandler.next(cell);
      }
    }
    await TagUtil.replaceCellList(this._containedCellList, this._innerData);
  }


  async next(cell){
    if (ExcelUtil.isMergedCell(cell)) {
      return
    }
    if(!this._wrapTag){
      let startTag = TagUtil.getAtStartTag(cell.value);
      if(startTag.tag){
        // console.log('startTag', startTag)
        this._wrapTag = startTag;
        cell.value = cell.value.replace(startTag.tagH, '');
      }
    }    

    if(this._wrapTag){          
      let endTag = TagUtil.getEndTag(cell.value, this._wrapTag.tag);
      if(endTag.tag){
        cell.value = cell.value.replace(endTag.tagH, '');
      }
      this._containedCellList.push(cell);
      if(endTag.tag){
        // 找到结束标记之后，说明找到了所有@标记内的单元格，可以开始执行内部替换了
        this._innerData = {};
        if(this._wrapTag.tag in this._outerData){
          this._innerData = this._outerData[this._wrapTag.tag];
        }
        await this._handle();        
        // 这一组替换之后，要充值缓存，继续寻找下一组
        this._reset();
      }
    }
  }
}

/**
 * 用于处理所有循环 {#xxxx}
 * 允许在内部的单元格嵌套 @标记， 
 * 允许在内部的单元格嵌套 innerloop
 * 但是不能嵌套普通循环
 */
class LoopRowHandler{
  constructor(data){    
    this._outerData = data;
    this._reset();
  }
  
  _reset(){
    // 开始结束标记之间的单元格，包含头尾
    this._containedCellList = [];
    this._startCell = null;
    this._endCell = null;

    // 标记 {tag: 'defects', tagH: '{#defects}'}
    this._wrapTag = null;
    this._innerData = [];
  }

  async _handleLoop(row, loopInfo){
    // console.log(row.number)
    // console.log(row.worksheet)
    if(loopInfo.tagName in this._outerData){
      // 根据循环标记，从outerData中找到要循环的数据
      let data = this._outerData[loopInfo.tagName];
      if(data == null){
        data = [];
      }      
      // 如果只是个对象，则把他格式化为数组
      // 有了这句话，就相当于在当行内实现了@标记，也就是对一个对象应用了循环标记，就相当于对一个长度为1的数组应用循环标记。
      // 这个与@标记的区别就在于，@标记支持多行，而循环标记应用在对象的时候只支持单行。
      if(!Array.isArray(data)){
        data = [data];
      }
      this._innerData = data;
    }
        
    let worksheet = row.worksheet;
    if(this._innerData.length == 0){
      // 如果没有数据，则行也删除了
      worksheet.spliceRows(row.number, 1);      
      ExcelUtil.adjustConditionFormatters(row.number, -1, worksheet, 'del');
    }
    if(this._innerData.length>1){
      // worksheet.duplicateRow(row.number, this._innerData.length-1, true)
      ExcelUtil.dupliateRowAndCopyStyle(worksheet, row.number, this._innerData.length-1);      
    }
    
    let tagName = loopInfo.tagName;
    // console.log('startCell', loopInfo.startCell.address, loopInfo.startCell.col, loopInfo.endCell.col)
    let startIndex = loopInfo.startCell.col;
    let endIndex = loopInfo.endCell.col;
    // if(this._innerData.length == 0){
    //   new LoopRowReplacer(row, {}, tagName, startIndex, endIndex)
    // }else{
      for(let index=0; index<this._innerData.length; ++index){        
        let newRow = worksheet.getRow(row.number + index);         
        let replacer = new LoopRowReplacer(newRow, this._innerData[index], tagName, startIndex, endIndex);
        await replacer.handle();
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
    let loopInfo = TagUtil.isLoopRow(row);
    if(!loopInfo){
      return 1
    }

    // console.log('row', row)
    await this._handleLoop(row, loopInfo);
    // 如果数据是空，则返回1
    let count = this._innerData.length>0?this._innerData.length: -1;
    this._reset();
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
    this._row = row;
    this._data = data;
    this._loopTagName = loopTagName;
    this._startColumnIndex =  startColumnIndex;
    this._endColumnIndex =  endColumnIndex;
    this._containedCellList = [];
  }  

  /**
   * 将首尾的loopTag去掉
   * 然后把loopTag范围内的cell添加到 _containedCellList
   */
  _formatCellAndFillContainedCellList(){
    for(let i = this._startColumnIndex; i<=this._endColumnIndex; ++i){
      let cell = this._row.getCell(i);
      if(i == this._startColumnIndex){
        cell.value = cell.value.replace(`{#${this._loopTagName}}`, '');
      }
      if(i == this._endColumnIndex){
        cell.value = cell.value.replace(`{/${this._loopTagName}}`, '');
      }
      this._containedCellList.push(cell);
    }
  }

  async handle(){    
    this._formatCellAndFillContainedCellList();    
    await this._handleAtTag();
    await TagUtil.replaceCellList(this._containedCellList, this._data);
  }
  async _handleAtTag(){
    let atTagHandler = new AtTagHandler(this._data);
    for(let cell of this._containedCellList){
      await atTagHandler.next(cell);
    }
  }
}

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
          cell.value = {formula: newFormula, value: cell.value};
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
      let cellCount = currentRow.cellCount;
      for (let columnIndex = 1; columnIndex <= cellCount; ++columnIndex) {
        let cell = currentRow.getCell(columnIndex);
        await handler(cell);
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

/**
 * 
 */
class XlsxTemplater {
  constructor(worksheet, data) {
    this._worksheet = worksheet;
    this._data = data;
  }
  /**
   * 使用data渲染 filePath路径下的文件，并返回渲染后的workbook   
   * @param {*} filePath 
   * @param {*} data 
   * @param {*} worksheetNameList ，要渲染的worksheet名称列表，如果不指定，默认就渲染第一个worksheet
   * @returns 渲染后的workbook
   */
  static async renderFromFile(filePath, data, worksheetNameList=[]){
    let workbook = new ExcelJS.Workbook();
    // workbook.getWorksheet().addConditionalFormatting
    await workbook.xlsx.readFile(filePath);
    await XlsxTemplater._findWorksheetAndRender(workbook, data, worksheetNameList);
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
      let workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      await XlsxTemplater._findWorksheetAndRender(workbook, data, worksheetNameList);    
      return await workbook.xlsx.writeBuffer()
    }
  /**
   * 
   * @param {*} workbook，要渲染的workbook
   * @param {*} data，用于渲染的json格式数据
   * @param {*} worksheetNameList ，要渲染的worksheet名称列表
   */
  static async _findWorksheetAndRender(workbook, data, worksheetNameList){
    let renderWorksheetList = [];
    if(worksheetNameList == null || worksheetNameList === undefined || worksheetNameList.length == 0){
      renderWorksheetList = [workbook.worksheets[0]];
    }else {
      workbook.worksheets.forEach(item=>{
        if(worksheetNameList.indexOf(item.name)>=0){
          renderWorksheetList.push(item);
        }
      });
    }
    for(let i = 0; i<renderWorksheetList.length; i++){
      let templater = new XlsxTemplater(renderWorksheetList[i], data);
      await templater.render();
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
        TagUtil.replaceCellNormalTag(cell, this._data);
      });
    });
  }

  /**
   * 处理被 {@xxx} ... {/xxx}包括的所有单元格，
   * 在这个内部，
   * 不允许嵌套普通循环，
   * 允许嵌套@标记
   * 允许嵌套单元格内的循环
   */
  async _renderAtTag() {
    let tagHandler = new AtTagHandler(this._data);
    let worksheet = this._worksheet;
    for (let rowIndex = 1; rowIndex <= worksheet.rowCount; ++rowIndex) {
      let row = worksheet.getRow(rowIndex);
      let cellCount = row.cellCount;
      for (let columnIndex = 1; columnIndex <= cellCount; ++columnIndex) {
        let cell = row.getCell(columnIndex);
        await tagHandler.next(cell);
      }
    }
  }

  /**
   * 替换掉最顶层的所有单元格内循环
   */
  async _renderInnerLoopTag() {
    let worksheet = this._worksheet;
    for (let rowIndex = 1; rowIndex <= worksheet.rowCount; ++rowIndex) {
      let row = worksheet.getRow(rowIndex);
      let cellCount = row.cellCount;
      for (let columnIndex = 1; columnIndex <= cellCount; ++columnIndex) {
        let cell = row.getCell(columnIndex);
        await TagUtil.replaceCellInnerLoopTag(cell, this._data);
      }
    }
  }

  /**
   * 替换整个worksheet的图片标记
   */
  async _renderImageTag() {
    let worksheet = this._worksheet;
    for (let rowIndex = 1; rowIndex <= worksheet.rowCount; ++rowIndex) {
      let row = worksheet.getRow(rowIndex);
      let cellCount = row.cellCount;
      for (let columnIndex = 1; columnIndex <= cellCount; ++columnIndex) {
        let cell = row.getCell(columnIndex);
        await TagUtil.replaceImageCellTag(cell, this._data);
      }
    }
  }

  async _renderLoopTag() {
    // console.log('_renderLoopTag begin');
    let loopHandler = new LoopRowHandler(this._data);
    let worksheet = this._worksheet;
    for (let rowIndex = 1; rowIndex <= worksheet.rowCount;) {
      let row = worksheet.getRow(rowIndex);
      rowIndex += await loopHandler.handle(row);
      // console.log('rowIndex', rowIndex, worksheet.rowCount)            
    }
  }

  /**
   * 处理多行循环标记
   */
  async _renderMultiLineLoopTag() {
    // console.log('_renderLoopTag begin');
    let multiLineLoopHandler = new MultiLineLoopHandler(this._worksheet, this._data);    
    await multiLineLoopHandler.handle();    
  }

  /**
   * 总的入口
   */
  async render(){
    await this._renderMultiLineLoopTag();
    await this._renderLoopTag();
    await this._renderAtTag();
    await this._renderInnerLoopTag();
    this._renderNormalTag();
    await this._renderImageTag();
  }
}

export { XlsxTemplater as default };
