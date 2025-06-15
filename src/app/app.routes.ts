import { Routes } from '@angular/router';
import { LoginComponent } from './login/login.component';
import { TablasComponent } from './tablas/tablas.component';

export const routes: Routes = [
  { path: '', component: LoginComponent },
  { path: 'tablas', component: TablasComponent }
];