import { excelColumn } from 'excel-to-object-decorator';


export class ResultWithHeadersType {
      @excelColumn({
        targetPropertyName: 'name'
      }, val => val.toUpperCase())
      name: string;


      @excelColumn({
        targetPropertyName: 'price'
      }, val => +val * 10)
      total: number;
}
