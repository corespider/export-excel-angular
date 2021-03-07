import { Component } from '@angular/core';
import * as Excel from 'exceljs';
import * as fs from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'export-excel';
  emprow: number =1;
  employees = [{
    "name": "Akash Jain",
    "designation": "Developer",
    "Country": "India",
    "email": "akash@gmail.com",
    "age": 25
  },
  {
    "name": "John Smith",
    "designation": "Solution Architect",
    "Country": "England",
    "email": "john.smith@gmail.com",
    "age": 36
  },
  {
    "name": "Raghav Rana",
    "designation": "Developer",
    "Country": "India",
    "email": "raghav.ran@hotmail.com",
    "age": 33
  },
  {
    "name": "Julie Roberts",
    "designation": "Developer",
    "Country": "USA",
    "email": "juile.roberts@gmail.com",
    "age": 36
  },
  {
    "name": "Sundar Raghav",
    "designation": "Solution Consultant",
    "Country": "India",
    "email": "sundar113@gmail.com",
    "age": 38
  },
  {
    "name": "Reba sen",
    "designation": "IT Manager",
    "Country": "Africa",
    "email": "reba.sen@yahoo.com",
    "age": 42
  }
  ]
  downloadExcel() {

   
    //create new excel work book
    let workbook = new Excel.Workbook();

    //add name to sheet
    let worksheet = workbook.addWorksheet("Employee Data");

    //add column name
    let header = ["Name", "Designation","Country","Email","Age"]
    let headerRow = worksheet.addRow(header);
    //headerRow.font = { size: 14, bold: true };

    for (let x1 of this.employees) {
      let x2 = Object.keys(x1);
      let temp = []
      for (let y of x2) {
        temp.push(x1[y])
      }
      worksheet.addRow(temp)
    }
    //set downloadable file name
    let fname = "Employee Sheet"

    //add data and file name and download
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob, fname + '-' + new Date().valueOf() + '.xlsx');
    });

  }
}