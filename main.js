import './style.css'
import { XLSX } from './src';
import json from './mock';
import { saveAs } from "file-saver";

const data = json;

const config = {
    boldHeader: true,
    wrapAll: true,
    colConfig: [
        {
            colKey: 'phone',
            width: 7
        }
    ],
    getCellStyle: (rowIndex, colIndex, cellData) => {
        if(colIndex == 3)
        return {
            fontId: 0,
            borderId: 1,
            bgColor: '#FF0000',
            alignment: 'left' 
        }
        return 0
    },
}
new XLSX().create(data,config).then(data=> saveAs(data, 'test.xlsx'))




