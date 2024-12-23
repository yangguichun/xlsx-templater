# xlsxtemplater 是什么
xlsxtemplater 使用 JavaScript 对象或 JSON 作为数据输入来渲染 xlsx 文件。

xlsxtemplater 基于标记的概念，每种类型的标记都为模板编写者提供特定的功能。

xlsxtemplater 的标记语法与 [docxtemplater 的 xlsx 模块](https://docxtemplater.com/modules/xlsx/#loops) 兼容，但只支持标记语法的一个子集。最重要的是，xlsxtemplater 是免费的。

xlsxtemplater 基于 [ExcelJS](https://github.com/exceljs/exceljs) 构建，这是一个功能强大的 Excel 文件读取、操作和写入库。

# 安装

```bash
npm install @sailimuhu/xlsxtemplater
# 或者
yarn add @sailimuhu/xlsxtemplater
```

## 依赖
- ExcelJS: ^4.3.0

## 使用方法

### Node.js (CommonJS)
```javascript
const XlsxTemplater = require('@sailimuhu/xlsxtemplater');

// 创建模板实例
const templater = new XlsxTemplater('./template.xlsx');
```

### Node.js (ESM)
```javascript
import XlsxTemplater from '@sailimuhu/xlsxtemplater';

// 创建模板实例
const templater = new XlsxTemplater('./template.xlsx');
```

### 浏览器
```html
<!-- UMD 版本 -->
<script src="node_modules/@sailimuhu/xlsxtemplater/dist/xlsxtemplater.umd.js"></script>
<script>
  const templater = new XlsxTemplater(file); // file 可以是 Blob 或 ArrayBuffer
</script>

<!-- ES Module -->
<script type="module">
  import XlsxTemplater from '@sailimuhu/xlsxtemplater';
  const templater = new XlsxTemplater(file);
</script>
```

### 浏览器使用注意事项
在浏览器环境中使用时：
1. 输入文件应该是 Blob 或 ArrayBuffer 类型，而不是文件路径
2. 使用 `templater.renderToBlob()` 替代 `templater.save()`
```javascript
// 浏览器示例
const response = await fetch('template.xlsx');
const templateFile = await response.arrayBuffer();

const templater = new XlsxTemplater(templateFile);
await templater.render(data);

// 获取结果为 Blob
const resultBlob = await templater.renderToBlob();
// 或直接下载
const link = document.createElement('a');
link.href = URL.createObjectURL(resultBlob);
link.download = 'result.xlsx';
link.click();
```

# 快速开始
```javascript
const XlsxTemplater = require('@sailimuhu/xlsxtemplater');

// 创建模板实例
const templater = new XlsxTemplater('./template.xlsx');

// 使用数据渲染
await templater.render({
  company: '华杰数能',
  date: '2024-03-20',
  items: [
    { name: '产品1', price: 100 },
    { name: '产品2', price: 200 }
  ]
});

// 保存渲染后的文件
await templater.save('./output.xlsx');
```

支持的标记：
- 普通标记，`{tagName}`
- 循环标记，`{#loopTag} ... {/loopTag}`
- 内部循环标记，`{#loopTag} ... {/}`
- 图片标记，`{%imageTag}`

扩展标记：
- 对象标记，`{@tagName} ...{/tagName}`，用于使用 JSON 数据中的嵌套对象。

# 待办事项
- 当数据是数组且作为最终数据使用时，应该将其连接并展开。

# 如何使用
## 概述
简单来说，支持5种类型的标记：`普通标记`、`对象标记`、`循环标记`、`内部循环标记`和`图片标记`。
- `普通标记`和`内部循环标记`最灵活，可以嵌套在`对象标记`和`循环标记`内部。
- `循环标记`可以包含嵌套的`对象标记`、`内部循环标记`、`图片标记`，但不能嵌套另一个`循环标记`。
- `对象标记`可以包含嵌套的`对象标记`、`图片标记`、`内部循环标记`，但不能嵌套`循环标记`，并且可以跨多行。

## 示例
以下是��本用法。
这个例子展示了所有支持的数据格式：
```js
let XlsxTemplater = require('XlsxTemplater')
let templater = new XlsxTemplater('./data/month_sale_report.xlsx')
templater.render({
    company: '华杰',
    createTime: '2022-12-09 05:25:00',
    reporters: ['张三', '李四'],
    summary:{
        salesAmount: 5200000,
        newCustomer: 3,
        orderAmount: 6200000,
        productList: ['产品A', '产品B', '产品C']
    }, 
    orders: [
        {
            date: '2022-11-1', 
            number: 'X221101001',
            customer: '大族',
            products: ['产品A', '产品B'],
            salesAmount: 520000,
            remark: ''
        },
        {
            date: '2022-11-3', 
            number: 'X221103002',
            customer: 'vivo',
            products: ['产品A', '产品C'],
            salesAmount: 320000,
            remark: ''
        }
    ]
})
```

[继续下一部分...] 

# 标记语法
## 1. 普通标记 `{xxx}`
使用大括号包围字段名称，如：
```js
{someTag}
```

| {hostCompany} |{createTime}|{productList}|
|:-|:-|:-

数据：
```js
templater.render({
  hostCompany: '华杰',
  createTime: '2022-12-09 05:25:00',
  productList:['产品A', '产品B']
})
```
渲染后，注意 productList 是一个数组，会被转换��逗号分隔的字符串：

|  华杰  | 2022-12-09 05:25:00 |'产品A', '产品B'
|:-|:-|:-

### 缺失标记自动填充空字符串
在下面的模板中，如果实际数据中缺少 'remark' 字段，它将被替换为空字符串：

| {hostCompany} |{createTime}|{remark}|
|:---|:---|:---|

```js
templater.render({
  hostCompany: '华杰',
  createTime: '2022-12-09 05:25:00'
})
```
渲染后：

|  华杰  | 2022-12-09 05:25:00 ||
|:---|:----|:---|

## 2. 循环标记 `{#xxx}...{/xxx}`
循环标记用于处理数组数据，支持单行循环和多行循环两种方式。
循环标记包含开始和结束标记：
- 开始标记：大括号内以 `#` 开头，例如 `{#someTag}`
- 结束标记：大括号内以 `/` 开头，例如 `{/someTag}`

### 2.1 单行循环
单行循环的开始和结束标记在同一行内，用于处理简单的数组数据：
```js
{#items} {name}| {quantity} | {@price}{value} {type}{/price} | {/items}
```

示例数据：
```js
templater.render({
    items: [
        {
            name: "产品A",
            quantity: 5,
            price: {
                type: "元",
                value: 10,
            },
        },
        {
            name: "产品B",
            quantity: 1,
            price: {
                type: "元",
                value: 20,
            },
        },
    ],
});
```

结果：
|产品A | 5 | 10 元|
|:-|:-|:-
|产品B | 1 | 20 元|

### 2.2 多行循环
多行循环允许开始和结束标记在不同行，适用于需要重复整块内容的场景：

| 问题描述 | 整改措施 | 负责人 |
|---------|----------|--------|
| {#defects}{description} | {rectify_plan} | {responsible_party.name} |
| 联系电话：{responsible_party.tel} | | |
| 整改结果：{rectify_result} {/defects}|| |

示例数据：
```js
{
  defects: [
    {
      description: '问题1',
      rectify_plan: '整改1',
      rectify_result: '已完成',
      responsible_party: {
        name: '张三',
        tel: '13800000001'
      }
    },
    {
      description: '问题2',
      rectify_plan: '整改2',
      rectify_result: '进行中',
      responsible_party: {
        name: '李四',
        tel: '13800000002'
      }
    }
  ]
}
```

多行循环特性：
1. 自动复制循环标记之间的所有行
2. 保持单元格格式（包括合并单元格、样式等）
3. 正确处理公式引用
4. 支持在循环内部使用其他类型的标记（对象标记、普通标记、图片标记）
5. 自动调整后续行的公式引用和条件格式

注意事项：
>1. 循环支持单行和多行两种模式
>2. 循环内部不能嵌套普���循环标记
>3. 循环内部可以引用外层对象的属性
>4. 循环内部支持嵌套[对象标记](#4-对象标记-xxxxxx)
>5. 循环内部支持嵌套[内部循环标记](#3-内部循环标记)
>6. 循环内部支持嵌套[图片标记](#5-图片标记-tag)
>7. 如果循环标记的目标是一个对象，它的作用相当于单行的对象标记

## 3. 内部循环标记
用于需要在单个单元格内填充数组中每个项目的值时使用。
"内部"意味着开始和结束标记必须在同一个单元格内。

内部循环标记包括：
- 开始标记：与普通循环相同，`{#someTag}`
- 结束标记：仅包含 `{/}`

示例：
|{#items} {name} | {#tags}{value},{/} |  {quantity} | {price} | {/items}                
|:-|:-|:-|:-|:-

数据：
```js
templater.render({
    "items": [
        {
            "name": "软件A",
            "tags": [{ "value": "好用" }, { "value": "优秀" }],
            "quantity": 10,
            "price": "￥100"
        }
    ]
})
```

结果：
|软件A | 好用,优秀 | 10 | ￥100 |
|:-|:-|:-|:-|

>1. 内部循环支持嵌套[`普通标记`](#1普通标记-xxx)和[`图片标记`](#5-图片标记-tag)

## 4. 对象标记 `{@xxx}...{/xxx}`
用于访问 JSON 中的嵌套对象数据。包含开始和结束标记：
- 开始标记：大括号内以 @ 开头，例如 `{@someTag}`
- 结束标记：大括号内以 / 开头，例如 `{/someTag}`

示例：
|{@basic}{hostCompany}| | | |
|:-|:-|:-|:-|
|{createTime}|{contactName}|{contactPhone}|{/basic}|

数据：
```js
templater.render({
  basic:{
    hostCompany: '华杰',
    contactName: '张三',
    contactPhone: '13088888888',
    createTime: '2022-12-09 05:25:00'
  }
})
```

注意事项：
>1. 对象标记的开始、结束和内部的普通标记可以出现在同一个单元格中
>2. 对象标记可以跨越多行
>3. 对象标记支持嵌套其他对象标记
>4. 对象标记支持[内部循环标记](#内部循环标记)但不支持[循环标记](#2-循环标记)

## 5. 图片标记 `{%tag}`
用于在单元格中插入图片。示例：
|{%beforePic}|{%afterPic}|

渲染后，这些标记将被替换为目标图片。图片会填充整个单元格，所以需要根据需要调整单元格大小。

# 条件格式支持
当使用循环标记添加或删除行时，工作表中的条件格式可能会受到影响。因此：
- 复制行时，会复制相应的条件格式并调整行号
- 删除行时，会删除相应的条件格式并调整行号

>注意：假设条件格式的引用和表达式仅使用同一行内的单元格，不存在跨行引用。 
