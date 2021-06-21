import './style.css'
import { XLSX } from './xlsx';
import json from './mock';
import { saveAs } from "file-saver";

const data = json;

const config = {
    boldHeader: true,
    wrapAll: true,
    colConfig: [
        {
            colKey: 'phone',
            width: 40
        }
    ],
    getCellStyle: (rowIndex, colIndex, cellData) => {      
        if(rowIndex == 1) return {
            bgColor: '#00FFF0',
            border: true
        }
    },
}

document.querySelector('#download').addEventListener('click',()=> {
    new XLSX().create(data,config).then(data=> saveAs(data, 'sheet.xlsx'))
})





