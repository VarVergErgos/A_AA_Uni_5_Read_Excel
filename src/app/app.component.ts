import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import * as FileSaver from 'file-saver'; 

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  title = 'read-excel-in-angular8';
  storeData: any;
  csvData: any;
  jsonData: any;
  textData: any;
  htmlData: any;
  fileUploaded: any = File;
  worksheet: any;
  first_sheet_name :any;
  
  uploadedFile(event: any) {
    this.fileUploaded = event.target.files[0];
    this.readExcel();
  }
  readExcel() {
    let readFile = new FileReader();
    readFile.onload = (e) => {
      this.storeData = readFile.result;
      var data = new Uint8Array(this.storeData);
      var arr = new Array();
      for (var i = 0; i != data.length; ++i)
        arr[i] = String.fromCharCode(data[i]);
      var bstr = arr.join('');
      var workbook = XLSX.read(bstr, { type: 'binary' });
      var first_sheet_name = workbook.SheetNames[0];

      this.worksheet = workbook.Sheets[first_sheet_name];
    };
    readFile.readAsArrayBuffer(this.fileUploaded);
  }
  readAsCSV() {
    this.csvData = XLSX.utils.sheet_to_csv(this.worksheet);
    const data: Blob = new Blob([this.csvData], {
      type: 'text/csv;charset=utf-8;',
    });
    FileSaver.saveAs(data, 'CSVFile' + new Date().getTime() + '.csv');
  }
  readAsJson() {
    this.jsonData = XLSX.utils.sheet_to_json(this.worksheet, { raw: false });
    this.jsonData = JSON.stringify(this.jsonData);
    const data: Blob = new Blob([this.jsonData], { type: 'application/json' });
    // FileSaver.saveAs(data, 'JsonFile' + new Date().getTime() + '.json');
    FileSaver.saveAs(data, 'JsonFile' + 'D:/Demo.json');
  }

  readAsJsonZieleinkommen() {
    this.jsonData = XLSX.utils.sheet_to_json(this.worksheet, { raw: false });
    this.jsonData = JSON.stringify(this.jsonData);
    const data: Blob = new Blob([this.jsonData], { type: 'application/json' });
    // FileSaver.saveAs(data, 'JsonFile' + new Date().getTime() + '.json');
    FileSaver.saveAs(data, 'JsonFile' + 'D:/A_AA_Json_DateienD/ZEK_G24.json');
  }

  readAsHTML() {
    this.htmlData = XLSX.utils.sheet_to_html(this.worksheet);
    const data: Blob = new Blob([this.htmlData], {
      type: 'text/html;charset=utf-8;',
    });
    FileSaver.saveAs(data, 'HtmlFile' + new Date().getTime() + '.html');
  }
  readAsText() {
    this.textData = XLSX.utils.sheet_to_txt(this.worksheet);
    const data: Blob = new Blob([this.textData], {
      type: 'text/plain;charset=utf-8;',
    });
    FileSaver.saveAs(data, 'TextFile' + new Date().getTime() + '.txt');
  }
}   

