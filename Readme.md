# @ys/easy-exceljs

A private wrapper for npm module exceljs.

```javascript
  const workbook = new SpreadSheet({
    styles: {
      row: {
        height: 21.95,
      },
      base: {
        font: {
          size: 12,
        },
        alignment: {
          vertical: 'middle',
        },
      },
    },
  });
  workbook
    .addSheet('New worksheet')
    .fill('Hello world')
  ;
  await workbook.writeXlsx('./1.xlsx');
```

## Publish

```shell script
npm publish --access public
```
