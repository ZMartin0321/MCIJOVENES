import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { Router } from '@angular/router';
import { HttpClient, HttpClientModule } from '@angular/common/http';
import { AfterViewInit, Inject, PLATFORM_ID } from '@angular/core';
import { isPlatformBrowser } from '@angular/common';

@Component({
  selector: 'app-login',
  standalone: true,
  imports: [CommonModule, FormsModule, HttpClientModule], // <-- AGREGA HttpClientModule AQUÍ
  templateUrl: './login.component.html',
  styleUrl: './login.component.css'
})
export class LoginComponent implements AfterViewInit {
  usuario = '';
  contrasena = '';
  error = '';

  constructor(@Inject(PLATFORM_ID) private platformId: Object,private router: Router, private http: HttpClient) {}

  login() {
    this.http.post<any>('http://localhost:3000/api/login', {
      usuario: this.usuario,
      contrasena: this.contrasena
    }).subscribe({
      next: resp => {
        if (resp.ok) {
          localStorage.setItem('usuario', this.usuario); 
          this.router.navigate(['/tablas']);
        } else {
          this.error = 'Usuario o contraseña incorrectos';
        }
      },
      error: () => {
        this.error = 'Usuario o contraseña incorrectos';
      }
    });
  }

ngAfterViewInit() {
    if (isPlatformBrowser(this.platformId)) {
      const logo = document.getElementById('logo');
      if (!logo) return;

      function spinPauseCycle() {
        if (!logo) return;
        logo.style.animationPlayState = 'running';
        setTimeout(() => {
          if (!logo) return;
          logo.style.animationPlayState = 'paused';
          logo.style.transform = 'rotateY(0deg)';
          setTimeout(spinPauseCycle, 3000);
        }, 8000);
      }

      spinPauseCycle();
    }
  }
}