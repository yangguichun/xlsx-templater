中文文档请参考 [readme.zh-cn.md](readme.zh-cn.md)
# What is xlsxtemplater
xlsxtemplater uses JavaScript Object or JSON as data input to render an xlsx file.

xlsxtemplater is based on the concepts of tags, and each type of tag exposes a feature to the user writing the template.

xlsxtemplater's tag is compatible with the [docxtemplater's xlsx module](https://docxtemplater.com/modules/xlsx/#loops), but only supports a subset of the tag grammar. The most important thing is that xlsxtemplater is free.

xlsxtemplater is built on top of [ExcelJS](https://github.com/exceljs/exceljs), a robust library for reading, manipulating and writing Excel files.

# Installation

```bash
npm install @sailimuhu/xlsxtemplater
# or
yarn add @sailimuhu/xlsxtemplater
```

## Dependencies
- ExcelJS: ^4.3.0

## Usage

### Node.js (CommonJS)
```javascript
const XlsxTemplater = require('@sailimuhu/xlsxtemplater');

// Create a new instance with your template file
const templater = new XlsxTemplater('./template.xlsx');
```

### Node.js (ESM)
```javascript
import XlsxTemplater from '@sailimuhu/xlsxtemplater';

// Create a new instance with your template file
const templater = new XlsxTemplater('./template.xlsx');
```

### Browser
```html
<!-- UMD version -->
<script src="node_modules/@sailimuhu/xlsxtemplater/dist/xlsxtemplater.umd.js"></script>
<script>
  const templater = new XlsxTemplater(file); // file can be Blob/ArrayBuffer
</script>

<!-- ES Module -->
<script type="module">
  import XlsxTemplater from '@sailimuhu/xlsxtemplater';
  const templater = new XlsxTemplater(file);
</script>
```

### Browser Usage Notes
When using in browser environment:
1. Input file should be Blob or ArrayBuffer instead of file path
2. Use `templater.renderToBlob()` instead of `templater.save()`
```javascript
// Browser example
const response = await fetch('template.xlsx');
const templateFile = await response.arrayBuffer();

const templater = new XlsxTemplater(templateFile);
await templater.render(data);

// Get result as Blob
const resultBlob = await templater.renderToBlob();
// Or download directly
const link = document.createElement('a');
link.href = URL.createObjectURL(resultBlob);
link.download = 'result.xlsx';
link.click();
```

# Quick Start
```javascript
const XlsxTemplater = require('@sailimuhu/xlsxtemplater');

// Create a new instance with your template file
const templater = new XlsxTemplater('./template.xlsx');

// Render with your data
await templater.render({
  company: 'Acme Corp',
  date: '2024-03-20',
  items: [
    { name: 'Item 1', price: 100 },
    { name: 'Item 2', price: 200 }
  ]
});

// Save the rendered file
await templater.save('./output.xlsx');
```

The supported tags: 
- normal tag, `{tagName}`
- loop tag, `{#loopTag} ... {/loopTag}`
- innerloop tag, `{#loopTag} ... {/}`
- image tag, `{%imageTag}`

Extended tag:
- object tag, `{@tagName} ...{/tagName}`, for using nested objects in JSON data.

# How to Use
## Summary
Simply put, it supports 5 types of tags: `normal tag`, `object tag`, `loop tag`, `innerloop tag`, and `image tag`.
- `normal tag` and `innerloop tag` are the most flexible, can be nested inside `object tag` and `loop tag`.
- `loop tag` can contain nested `object tag`, `innerloop tag`, `image tag`, but cannot nest another `loop tag`.
- `object tag` can contain nested `object tag`, `image tag`, `innerloop`, but cannot nest `loop tag`, and can span multiple rows.

## Example
Here's the basic usage.
The following example demonstrates all supported data formats:
```js
let XlsxTemplater = require('XlsxTemplater')
let templater = new XlsxTemplater('./data/month_sale_report.xlsx')
templater.render({
    company: 'Huajie',
    createTime: '2022-12-09 05:25:00',
    reporters: ['Zhangsan', 'Lisi'],
    summary:{
        salesAmount: 5200000,
        newCustomer: 3,
        orderAmount: 6200000,
        productList: ['productA', 'productB', 'productC']
    }, 
    orders: [
        {
            date: '2022-11-1', 
            number: 'X221101001',
            customer: 'dazu',
            products: ['productA', 'productB'],
            salesAmount: 520000,
            remark: ''
        },
        {
            date: '2022-11-3', 
            number: 'X221103002',
            customer: 'vivo',
            products: ['productA', 'productC'],
            salesAmount: 320000,
            remark: ''
        }
    ]
})
```

[继续下一部分...]

# Tag Grammar
## 1. Normal Tag, `{xxx}`
Use curly braces to enclose a field name, like this:
```js
{someTag}
```

| {hostCompany} |{createTime}|{productList}|
|:-|:-|:-

Data:
```js
templater.render({
  hostCompany: 'Huajie',
  createTime: '2022-12-09 05:25:00',
  productList:['productA', 'productB']
})
```
After rendering, note that productList is an array and will be converted to a comma-separated string:

|  Huajie  | 2022-12-09 05:25:00 |'productA', 'productB'
|:-|:-|:-

### Auto-filling Empty String for Missing Tags
In the template below, if the 'remark' field is missing from the actual data, it will be replaced with an empty string:

| {hostCompany} |{createTime}|{remark}|
|:---|:---|:---|

```js
templater.render({
  hostCompany: 'Huajie',
  createTime: '2022-12-09 05:25:00'
})
```
After rendering:

|  Huajie  | 2022-12-09 05:25:00 ||
|:---|:----|:---|

## 2. Loop Tag, `{#xxx}...{/xxx}`
Loop tags are used to generate content from array data. They support both single-line and multi-line loops.
A loop tag consists of start and end tags:
- Start tag: begins with `#` inside curly braces, e.g., `{#someTag}`
- End tag: begins with `/` inside curly braces, e.g., `{/someTag}`

### 2.1 Single-line Loop
Single-line loops have start and end tags on the same line, used for simple array data:
```js
{#items} {name}| {quantity} | {@price}{value} {type}{/price} | {/items}
```

Example data:
```js
templater.render({
    items: [
        {
            name: "Product A",
            quantity: 5,
            price: {
                type: "currency",
                value: 10,
            },
        },
        {
            name: "Product B",
            quantity: 1,
            price: {
                type: "currency",
                value: 20,
            },
        },
    ],
});
```

Result:
|Product A | 5 | 10 currency|
|:-|:-|:-
|Product B | 1 | 20 currency|

### 2.2 Multi-line Loop
Multi-line loops allow start and end tags to be on different rows, useful for repeating blocks of content:

| Description | Action | Responsible |
|------------|--------|-------------|
| {#defects}{description} | {rectify_plan} | {responsible_party.name} |
| Contact: {responsible_party.tel} | | |
| Result: {rectify_result} {/defects}|| |

Example data:
```js
{
  defects: [
    {
      description: 'Issue 1',
      rectify_plan: 'Fix 1',
      rectify_result: 'Completed',
      responsible_party: {
        name: 'Zhang San',
        tel: '13800000001'
      }
    },
    {
      description: 'Issue 2',
      rectify_plan: 'Fix 2',
      rectify_result: 'In Progress',
      responsible_party: {
        name: 'Li Si',
        tel: '13800000002'
      }
    }
  ]
}
```

Multi-line Loop Features:
1. Automatically copies all rows between loop tags
2. Preserves cell formatting (including merged cells, styles)
3. Correctly handles formula references
4. Supports other tag types within the loop (object tags, normal tags, image tags)
5. Automatically adjusts formula references and conditional formatting for subsequent rows

Notes:
>1. Loops support both single-line and multi-line modes
>2. Loops cannot contain nested loop tags
>3. Loops can reference outer object properties
>4. Loops support nested [object tags](#4-object-tag-xxxxxx)
>5. Loops support nested [innerloop tags](#3-innerloop-tag)
>6. Loops support nested [image tags](#5-image-tag-tag)
>7. If a loop tag's target is an object, it acts like a single-line object tag

## 3. Innerloop Tag
Used when you need to fill a single cell with values from each item in an array.
'Inner' means both start and end tags must be within the same cell.

Innerloop tags include:
- Start tag: same as normal loop, `{#someTag}`
- End tag: just `{/}` in curly braces

Example:
|{#items} {name} | {#tags}{value},{/} |  {quantity} | {price} | {/items}                
|:-|:-|:-|:-|:-

Data:
```js
templater.render({
    "items": [
        {
            "name": "AcmeSoft",
            "tags": [{ "value": "fun" }, { "value": "awesome" }],
            "quantity": 10,
            "price": "$100"
        }
    ]
})
```

Result:
|AcmeSoft | fun,awesome | 10 | $100 |
|:-|:-|:-|:-|

>1. Innerloop supports nested [`normal tags`](#1general-tag) and [`image tags`](#5-image-tag)

## 4. Object Tag, `{@xxx}...{/xxx}`
Used to access nested object data in JSON. Includes start and end tags:
- Start tag: begins with @ inside curly braces, e.g., `{@someTag}`
- End tag: begins with / inside curly braces, e.g., `{/someTag}`

Example:
|{@basic}{hostCompany}| | | |
|:-|:-|:-|:-|
|{createTime}|{contactName}|{contactPhone}|{/basic}|

Data:
```js
templater.render({
  basic:{
    hostCompany: 'Huajie',
    contactName: 'Zhangsan',
    contactPhone: '13088888888',
    createTime: '2022-12-09 05:25:00'
  }
})
```

Notes:
>1. Object tag start, end, and inner normal tags can appear in the same cell
>2. Object tags can span multiple rows
>3. Object tags support nesting other object tags
>4. Object tags support [innerloop tags](#innerloop-tag) but not [loop tags](#2-loop-tag)

## 5. Image Tag, `{%tag}`
Used to insert images into cells. Example:
|{%beforePic}|{%afterPic}|

After rendering, these tags will be replaced with the target images. Images will fill the entire cell, so adjust cell size as needed.

# Conditional Formatting Support
When using loop tags that add or delete rows, conditional formatting in the worksheet might be affected. Therefore:
- When copying rows, corresponding conditional formatting is copied and row numbers are adjusted
- When deleting rows, corresponding conditional formatting is removed and row numbers are adjusted

>Note: Assumes conditional formatting references and expressions only use cells within the same row, no cross-row references.

# Resolve an error of the Exceljs library: The Worksheet.spiceRows() function will unmerge all subsequent merged cells when deleting rows.
If you are using exceljs version 4.4.0 or earlier, please go to node_modules/exceljs/lib/doc/worksheet.js and then find the spiceRows() function.
In the following branch, it was originally like this.


```js
if (nExpand < 0) {
      // remove rows
      if (start === nEnd) {
        this._rows[nEnd - 1] = undefined;
      }
      for (i = nKeep; i <= nEnd; i++) {
        rSrc = this._rows[i - 1];
        if (rSrc) {
          const rDst = this.getRow(i + nExpand);
          rDst.values = rSrc.values;
          rDst.style = rSrc.style;
          rDst.height = rSrc.height;
          // eslint-disable-next-line no-loop-func
          rSrc.eachCell({includeEmpty: true}, (cell, colNumber) => {
            rDst.getCell(colNumber).style = cell.style;
            // remerge cells accounting for insert offset
            if (cell._value.constructor.name === 'MergeValue') {
              const cellToBeMerged = this.getRow(cell._row._number + nExpand).getCell(colNumber);
              const prevMaster = cell._value._master;
              const newMaster = this.getRow(prevMaster._row._number + nExpand).getCell(prevMaster._column._number);
              cellToBeMerged.merge(newMaster);
            }
          });
          this._rows[i - 1] = undefined;
        } else {
          this._rows[i + nExpand - 1] = undefined;
        }
      }
    }
```
Change it to the following, mainly adding the processing logic for merged cells.
```js
if (nExpand < 0) {
      // remove rows
      if (start === nEnd) {
        this._rows[nEnd - 1] = undefined;
      }
      for (i = nKeep; i <= nEnd; i++) {
        rSrc = this._rows[i - 1];
        if (rSrc) {
          const rDst = this.getRow(i + nExpand);
          rDst.values = rSrc.values;
          rDst.style = rSrc.style;
          rDst.height = rSrc.height;
          // eslint-disable-next-line no-loop-func
          rSrc.eachCell({includeEmpty: true}, (cell, colNumber) => {
            rDst.getCell(colNumber).style = cell.style;
            // new added, fix the unmerged cell bug
            // remerge cells accounting for insert offset
            if (cell._value.constructor.name === 'MergeValue') {
              const cellToBeMerged = this.getRow(cell._row._number + nExpand).getCell(colNumber);
              const prevMaster = cell._value._master;
              const newMaster = this.getRow(prevMaster._row._number + nExpand).getCell(prevMaster._column._number);
              cellToBeMerged.merge(newMaster);
            }
          });
          this._rows[i - 1] = undefined;
        } else {
          this._rows[i + nExpand - 1] = undefined;
        }
      }
    }
```