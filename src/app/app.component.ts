import { Component } from '@angular/core';

import * as XLSX from 'xlsx';


type AOA = any[][];

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'sheetjs';

  data: AOA = [['name', 'age', 'DOB', 'isMan'], ['keatkeat', 11, new Date(1987, 11, 15, 3, 4, 22), true]];
  wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
  fileName = 'SheetJS.xlsx';

  export(): void {
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);
    const wb: XLSX.WorkBook = XLSX.utils.book_new();

    const a =  ws.A2;
    console.log(a);

    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, this.fileName);
  }

}
