import { Component } from '@angular/core';
import { excelRows } from 'excel-to-object-decorator';
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


  @excelRows(ResultType)
  results: any;

  @excelRows(ResultWithHeadersType)
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
      this.resultsWithHeaders = {
        headers: data[0],
        results: data.slice(1)
      };
    };
    reader.readAsBinaryString(target.files[0]);
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
      this.results = data;
    };
    reader.readAsBinaryString(target.files[0]);
  }

}
