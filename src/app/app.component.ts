import { Component } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { MagicWithoutTricksComponent } from './modules/pages/magic-without-tricks/magic-without-tricks.component';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet, MagicWithoutTricksComponent],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent {
  title = 'angular-export-excel';
}
