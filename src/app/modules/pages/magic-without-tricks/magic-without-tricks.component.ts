import { Component, Inject, OnInit } from '@angular/core';
import { GeneralService } from '../../../services/general.service';
import { CommonModule, DatePipe } from '@angular/common';
import { HttpClient } from '@angular/common/http';
import { BrowserModule } from '@angular/platform-browser';
import { BehaviorSubject } from 'rxjs';
import { ExcelGlobalService } from '../../utils/excel-global.service';

@Component({
  selector: 'app-magic-without-tricks',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './magic-without-tricks.component.html',
  styleUrl: './magic-without-tricks.component.scss',
  providers: [ExcelGlobalService, DatePipe]
})
export class MagicWithoutTricksComponent implements OnInit {

  private dataSubject = new BehaviorSubject<any[]>([
    {
      "id": 1,
      "name": "John Doe",
      "email": "john.doe@example.com"
    },
    {
      "id": 2,
      "name": "Jane Smith",
      "email": "jane.smith@example.com"
    },
    {
      "id": 3,
      "name": "Alice Johnson",
      "email": "alice.johnson@example.com"
    }
  ]);
  
  data$ = this.dataSubject.asObservable();
  data: any[] = [];

  constructor(
    private _generalService: GeneralService,
    private _excelService: ExcelGlobalService
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
      console.log("No hay datos disponibles para exportar a Excel.");
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
      console.log("Exportación cancelada.");
    }
  }

  getJson() {
    // Suscribirse a los datos
    this.data$.subscribe((data) => {
      this.data = data;
      console.log(this.data);
    });
  }
}