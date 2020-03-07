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

  data: AOA = [
    ['name', 'age', 'DOB', 'isMan'],
    ['keatkeat', 1111.3, new Date(1987, 11, 15, 3, 4, 22), true],
    ['keatkeat', 11.38, new Date(2014, 11, 15, 3, 4, 22), false],
    ['keatkeat', 157993.75, new Date(2014, 11, 15, 15, 4, 22), false]
  ];
  wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
  fileName = 'SheetJS.xlsx';

  export(): void {
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);


    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    const range = XLSX.utils.decode_range(ws['!ref']);
    console.log('range', range);
    for (let row = 2; row <= range.e.r + 1; row++) {
      const cell: XLSX.CellObject = ws[`B${row}`];
      cell.z = '#,##0.00';
    }
    for (let row = 2; row <= range.e.r + 1; row++) {
      const cell: XLSX.CellObject = ws[`C${row}`];
      cell.z = 'dd-MM-yyyy hh:mm:ss AM/PM';
    }
    console.log('ws', ws);

    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, this.fileName);
  }

}
