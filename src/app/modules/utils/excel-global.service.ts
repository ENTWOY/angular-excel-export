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