# 2024年12月19日17:20:31 
修改循环标签，使其支持多行循环，可能遇到的问题：
- 首先要弄清楚，要循环的是几行；
- 目前ExcelJs的复制行的api只能复制单行，

# 2024年12月20日14:27:59
参考我之前写的  @LoopRowHandler.js ，这个是用来处理单个循环标记的，现在我们另外写一个 MultiLineLoopHandler类，初始化的入参是worksheet和要替换的数据data，内部的处理逻辑与LoopRawHandler有些类似。首先要找出在worksheet内部哪几行是要循环的，得到startRowIndex, endRowIndex还有循环标记loopTag，然后在data中根据loopTag找到对应的数据loopDataArray，看需要循环几次，
1）如果loopDataArray是空数据或者null之类的，则删除这些需要循环的行；
2）如果loopDataArray长度大于1，则将 startRowIndex和endRowIndex之间的行复制loopDataArray.length-1遍，，插入到endRowIndex下方，用刚才新写的ExcelCopier来复制这些行；
3）如果数据的长度是1，则不用复制，也不用删除，可以继续后面的处理。

然后再回到startRowIndex，根据loopDataArray来处理循环标记，对于每一段 startRowIndex和endRowIndex之间的行，都要处理里面的 NormalTag、ImageTag和AtTag
