import { Routes } from '@angular/router';

export const routes: Routes = [
    {
        path: '',
        loadChildren: () => import('./modules/pages/magic-without-tricks/magic-without-tricks.component').then(m => m.default)  
    }
];
