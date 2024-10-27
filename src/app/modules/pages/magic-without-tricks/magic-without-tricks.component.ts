import { Component, Inject, OnInit } from '@angular/core';
import { GeneralService } from '../../../services/general.service';
import { CommonModule, DatePipe } from '@angular/common';
import { ExcelGlobalService } from '../../utils/excel-global.service';
import { SweetAlertService } from '../../utils/sweet-alert.service';

@Component({
  selector: 'app-magic-without-tricks',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './magic-without-tricks.component.html',
  styleUrl: './magic-without-tricks.component.scss',
  providers: [ExcelGlobalService, DatePipe]
})
export default class MagicWithoutTricksComponent implements OnInit {
  data: any[] = [];
  date = new Date;

  constructor(
    private _generalService: GeneralService,
    private _excelService: ExcelGlobalService,
    private _sweetService: SweetAlertService
  ) { }

  ngOnInit(): void {
    this._generalService.getData().subscribe((res:any) => {
      this.data = res?.results||null;
      console.log(this.data)
    })
  }

  exportToExcel() {
    const confirmExport = confirm("¿Desea exportar los datos a un archivo Excel?");

    if (!this.data || this.data.length === 0) {
      this._sweetService.alertToast("No hay datos disponibles para exportar a Excel", "info", 3000)
      return
    }

    if (confirmExport) {
      let arreglo = this.data
      let mapa : Map<any,any> = new Map( [      
        ['ID','id'],
        ['NAME','name'],
        ['STATUS','status'],
        ['SPECIES','species']
      ])

      let info_excel  = {
        mapa : mapa,
        title : 'Reporte de API',
        data: arreglo, 
        widths : [10,24,10,10],
        fecha: new Date
      }
      this._excelService.createExcel(info_excel)
    } else {
      this._sweetService.alertToast("Exportación cancelada.", "info", 3000)
    }
  }

  exportToExcel2() {
    const confirmExport = confirm("¿Desea exportar los datos a un archivo Excel?");

    if (!this.data || this.data.length === 0) {
      this._sweetService.alertToast("No hay datos disponibles para exportar a Excel", "info", 3000)
      return
    }

    if (confirmExport) {
      let arreglo = this.data
      let mapa : Map<any,any> = new Map( [      
        ['ID','id'],
        ['NAME','name'],
        ['STATUS','status'],
        ['SPECIES','species'],
        ['CREATED','created']
      ])

      let info_excel  = {
        mapa : mapa,
        title : 'Reporte de API V2',
        data: arreglo, 
        widths : [10,24,10,10,10],
        fecha: new Date
      }
      this._excelService.createExcel2(info_excel)
    } else {
      this._sweetService.alertToast("Exportación cancelada.", "info", 3000)
    }
  }
}