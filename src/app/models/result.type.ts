import { excelColumn } from 'excel-to-object-decorator';


export class ResultType {

  @excelColumn({
    columnNumber: 0
  }, val => val.toUpperCase())
  name: string;


  @excelColumn({
    columnNumber: 1
  }, val => +val * 10)
  total: number;
}
