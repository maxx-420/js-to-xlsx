

<h3 align="center">js-to-xlsx</h3>


---


## 📝 Table of Contents

- [About](#about)
- [Usage](#usage)
- [Built Using](#built_using)
- [Acknowledgments](#acknowledgement)

## 🧐 About <a name = "about"></a>

This project showcases downloading client side .xlsx files with a few customization option. Unlike SheetJS, this is a lightweight solution(<100kb). 
Lots of room for improvement. :)

## 🏁 Usage <a name = "usage"></a>

<code>npm run serve</code> - To serve.

IE supported but requires prod build - <code>npm run preview</code>

```

const data = [{...}] (Array of objects)

const config = {
  displayedColums: ['objKey'] // Order of columns to be displayed. If null/empty then alphabetical order is maintained.
  boldHeaders: boolean // default: true,
  upperCaseHeaders: boolean //default: false
  wrapAll: boolean // default: false,
  getCellStyle: (rowIndex, colIndex, cellData) => number/obj // expects a predefined style number or object.
  colConfig: [
    {
      colKey: 'key', //column key
      width: number  // width
    }....
  ]

}
const xlsx = new XLSX().create(data, config)
```

## Dependencies

-JSZip

-FileSaver(can be replaced with a custom solution)


### Cell Styling

getCellStyle expects one of the following:

1. Number:

    normal: 0

    bold text: 2

    wrap text: 10

    wrap & bold text: 11

2. Object literal containing following properties(optional):

       
        fontType: 0/2/10/11
        border: true/false,
        fontColor: '#FCFCFC',
        bgColor: '#EEEEEE',
        wrap: true/false,
        alignment: 'right/left/center/flll'





## ⛏️ Built Using <a name = "built_using"></a>

- Vite JS

## 🎉 Acknowledgements <a name = "acknowledgement"></a>

- ExcelHTML5
- OpenXML
