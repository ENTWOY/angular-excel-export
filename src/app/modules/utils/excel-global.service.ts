import { DatePipe } from '@angular/common';
import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';

@Injectable({
  providedIn: 'root'
})
export class ExcelGlobalService {

  constructor(
    private _datePipe: DatePipe
  ) { }

  public createExcel(info_excel:any) {
    let workbook = new Workbook();
    let worksheet = workbook.addWorksheet("Listado de Modificaciones");
    let mapa: Map<any, any> = info_excel.mapa
    let arr: any = [...mapa.keys()]
    let titulos = [].concat(arr)
    let data: any[] = info_excel.data || []
    let fecha = this._datePipe.transform(info_excel.fecha, 'dd/MM/yyyy')

    let file_name =  `Archivo_Excel_de_${fecha}`;

    let cell_title = worksheet.getCell('A1:D1')
    let merTitle = worksheet.mergeCells('A1:D1')
    let title =  `Reporte Excel API de ${fecha}`;
    cell_title.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '4167b8' },
      bgColor: { argb: '' }
    }
    cell_title.font = {
      bold: true,
      color: { argb: 'FFFFFF' },
      size: 17
    }
    cell_title.alignment = {
      vertical: 'middle',
      horizontal: 'center',
    }
    cell_title.border = {
      top: { style: 'thin', color: { argb: 'FFFFFF' } },
      left: { style: 'thin', color: { argb: 'FFFFFF' } },
      bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
      right: { style: 'thin', color: { argb: 'FFFFFF' } }
    }
    cell_title.value = title
    let headerRow = worksheet.addRow(titulos);

    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '5175c2' },
        bgColor: { argb: '' }
      }
      cell.font = {
        bold: true,
        color: { argb: 'FFFFFF' },
        size: 12
      }
      cell.alignment = {
        vertical: 'middle',
        horizontal: 'center',
      }
    });
    data.forEach(item => {
      let row = []
      for (let index = 0; index < titulos.length; index++) {
        row[index] = item[mapa.get(titulos[index]) || ''] || ''
      }
      let insertRow = worksheet.addRow(row)
      this.putBorder(insertRow)
    })

    let widhts = info_excel.widths

    for (let index = 0; index < widhts.length; index++) {
      worksheet.getColumn(index + 1).width = widhts[index];
    }

    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob, file_name + '.xlsx');
    })
  }

  public createExcel2(info_excel:any) {
    let workbook = new Workbook();
    let worksheet = workbook.addWorksheet("DATA API");
    let mapa: Map<any, any> = info_excel.mapa
    let arr:any = [...mapa.keys()]
    let titulos = [].concat(arr)
    let data: any[] = info_excel.data || []
    let query = info_excel.fields

    let total_title = titulos.length
    let cell_title = worksheet.getCell('A1')
    let title_cell = worksheet.mergeCells(1, 1, 2, total_title)
    /* let merTitle = worksheet.mergeCells('A1:AD2') */
    let title = "Lista de datos API V2"
    cell_title.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '4167B8' },
      bgColor: { argb: '' }
    }
    cell_title.font = {
      bold: true,
      color: { argb: 'FFFFFF' },
      size: 20
    }
    cell_title.alignment = {
      vertical: 'middle',
      horizontal: 'center',
    }
    cell_title.border = {
      top: { style: 'thin', color: { argb: 'FFFFFF' } },
      left: { style: 'thin', color: { argb: 'FFFFFF' } },
      bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
      right: { style: 'thin', color: { argb: 'FFFFFF' } }
    }
    cell_title.value = title

    let fecha = []
    fecha[1] = "Fecha"
    fecha[2] = new Date
    
    worksheet.addRow([])
    let row = worksheet.addRow(fecha)
    worksheet.addRow([])
    let headerRow = worksheet.addRow(titulos);
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '4167B8' },
        bgColor: { argb: '' }
      }
      cell.font = {
        bold: true,
        color: { argb: 'FFFFFF' },
        size: 12
      }
      cell.alignment = {
        vertical: 'middle',
        horizontal: 'center',
      }
    });
    let campo = 0
    data.forEach(item => {
      let row = []
      row[1] = ++campo
      for (let index = 1; index < titulos.length; index++) {
        if (index == 4) {
          row[index + 1] = this._datePipe.transform((item[mapa.get(titulos[index])]), 'dd/MM/yyyy')
        }
        else {
          row[index + 1] = item[mapa.get(titulos[index]) || ''] || ''
        }
      }
      let insertRow = worksheet.addRow(row)
      this.putBorder(insertRow)
    })

    worksheet.getColumn(1).width = 10;
    worksheet.getColumn(2).width = 30;
    worksheet.getColumn(3).width = 30;
    worksheet.getColumn(4).width = 25;
    worksheet.getColumn(5).width = 20;

    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob, 'Archivo Data Excel V2' + '.xlsx');
    })
  }

  private putBorder(row: any) {
    row.eachCell((cell: any, number: number) => {
      cell.border = {
        top: { style: 'thin', color: { argb: '000000' } },
        left: { style: 'thin', color: { argb: '000000' } },
        bottom: { style: 'thin', color: { argb: '000000' } },
        right: { style: 'thin', color: { argb: '000000' } }
      }
    });
  }
}