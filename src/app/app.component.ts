import { Component } from '@angular/core';
import { excelData, excelToObjectParser } from 'excel-to-object-decorator';
import { objectToExcel } from 'excel-to-object-decorator/dist/object-to-excel';
import * as XLSX from 'xlsx';
import { ResultWithHeadersType } from './models/result-with-headers.type';
import { ResultType } from './models/result.type';


@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'decorator-demo';

  results: any;

  resultsWithHeaders: any;



  /**
   *  SOURCE GITHUB REPO: https://github.com/SheetJS/js-xlsx/tree/1eb1ec985a640b71c5b5bbe006e240f45cf239ab/demos/angular2
   **/
  readExcelFileWithHeaders(evt): void {
    const target: DataTransfer = (evt.target) as DataTransfer;
    if (target.files.length !== 1) {
      throw new Error('Cannot use multiple files');
    }
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, {
        type: 'binary'
      });

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      /* save data */
      const data = (XLSX.utils.sheet_to_json(ws, {
        header: 1
      }));
      this.resultsWithHeaders = this.handleData(data);
    };
    reader.readAsBinaryString(target.files[0]);
  }


  @excelToObjectParser(ResultWithHeadersType, {headerRowIndex: 0}) // or
  // @excelToObjectParser(ResultWithHeadersType, {headers: ['name', 'price']})
  private handleData(@excelData data: any) {
    return data;
  }

  /**
   *  SOURCE GITHUB REPO: https://github.com/SheetJS/js-xlsx/tree/1eb1ec985a640b71c5b5bbe006e240f45cf239ab/demos/angular2
   **/
  readExcelFile(evt): void {
    const target: DataTransfer = (evt.target) as DataTransfer;
    if (target.files.length !== 1) {
      throw new Error('Cannot use multiple files');
    }
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, {
        type: 'binary'
      });

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      /* save data */
      const data = (XLSX.utils.sheet_to_json(ws, {
        header: 1
      }));
      this.results = this.handleDataWithoutHeaders(data);
    };
    reader.readAsBinaryString(target.files[0]);
  }

  @excelToObjectParser(ResultType)
  private handleDataWithoutHeaders(@excelData data: any) {
    // do business logic with mapper data
    data.forEach(element => console.log(element.name));
  }


  export(): void {
    const data = objectToExcel(ResultWithHeadersType)(this.resultsWithHeaders);
    const dataWithoutHeaders = objectToExcel(ResultType)(this.results);

    /* generate worksheet */
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(dataWithoutHeaders);
    const wsWithHeaders: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(data);

    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Without headers');
    XLSX.utils.book_append_sheet(wb, wsWithHeaders, 'With headers');

    /* save to file */
    XLSX.writeFile(wb, 'SheetJS.xlsx');
    return;
  }

}
