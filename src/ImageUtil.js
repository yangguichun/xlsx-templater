let fetch = require('node-fetch')

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
      console.log('image loading', url)
      fetch(url).then((res) => {
        res.blob().then((blob) => {
          blob.arrayBuffer().then((arr)=>{
            console.log('image loaded', url)
            resolve(arr)
          })      
        })
      })
      .catch(err=>{
        reject(err)
      })
    })
  }

  /**
   * 获取指定图片url的后缀名
   * @param {*} url 例如
   * https://www.baidu.com/img/PCtm_d9c8750bed0b3c7d089fa7d55720d6cf.png
   * @returns png, jpg, jpeg之类的
   */
  static getImageExt(url){
    let pattern = /.+\.(\w+)$/i
    let matches = pattern.exec(url)
    if(matches && matches.length>=2){
      return matches[1]
    }
    return 'jpg'
  }
}

module.exports = ImageUtil