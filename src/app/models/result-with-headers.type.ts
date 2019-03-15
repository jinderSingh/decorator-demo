import { excelColumn } from 'excel-to-object-decorator';


export class ResultWithHeadersType {
      @excelColumn({
        header: 'name'
      }, val => val.toUpperCase())
      name: string;


      @excelColumn({
        header: 'price'
      }, val => +val * 10)
      total: number;
}
