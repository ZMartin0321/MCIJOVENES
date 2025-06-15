import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import * as XLSX from 'xlsx';
import * as FileSaver from 'file-saver';
import * as ExcelJS from 'exceljs';
import { HttpClient } from '@angular/common/http'; // Agrega esto
import { firstValueFrom } from 'rxjs';


import { OnInit } from '@angular/core';



@Component({
  selector: 'app-tablas',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './tablas.component.html',
  styleUrl: './tablas.component.css'
})

export class TablasComponent {
  mostrarRegistrarOfrenda = false;
    mergesHoja2: any[] = [];
  data: any[][] = [];
  dataHoja2: any[][] = [];
  mostrarRegistrarMotivoTodos = false;
  motivoSeleccionado: string = '';
  mostrarSidebar: boolean = true;
  apartadoSeleccionado = 'Celugrama'; 
  mostrarEntornoTrabajo = false;
  mostrarRegistrarAsistencia = false;
  mostrarFormulario: boolean = false;
  nombresSeleccionados: string[] = [];
  agregarPersonaCE_CL() {
    if (this.nuevoNombreApellido && this.nuevaEdad) {
      this.data.push([this.nuevoNombreApellido, this.nuevaEdad]);
      this.nuevoNombreApellido = '';
      this.nuevaEdad = '';
      this.mostrarFormulario = false;
    }
  }

mostrarEditarEliminar = false;
editandoUsuario = false;
accionUsuario: 'agregar' | 'modificar' | 'eliminar' = 'agregar';
usuarioSeleccionado: string = '';
nuevoNombreApellido: string = '';
nuevaEdad: string = '';
nuevoNombreEditar: string = '';
nuevaEdadEditar: string = '';


limpiarCamposUsuario() {
  this.usuarioSeleccionado = '';
  this.nuevoNombreApellido = '';
  this.nuevaEdad = '';
  this.nuevoNombreEditar = '';
  this.nuevaEdadEditar = '';
}

prepararEdicion() {
  this.editandoUsuario = true;
  this.nuevoNombreEditar = this.usuarioSeleccionado;
  const rangos = [
    { inicio: 6, fin: 40 },
    { inicio: 53, fin: 87 },
    { inicio: 103, fin: 137 }
  ];
  for (const rango of rangos) {
    for (let i = rango.inicio; i <= rango.fin; i++) {
      if (this.data[i]?.[2] === this.usuarioSeleccionado) {
        this.nuevaEdadEditar = this.data[i][9]; // Debe ser [9]
        return;
      }
    }
  }
}

  originalFileBuffer: ArrayBuffer | undefined;
    constructor(private http: HttpClient) {}

    ngOnInit() {
  if (typeof window !== 'undefined') {
    this.CargarMiArchivo();
  }
   
  
}

merges: any[] = [];

toggleSeleccionDisciple(nombre: string) {
  const idx = this.nombresSeleccionados.indexOf(nombre);
  if (idx === -1) {
    this.nombresSeleccionados.push(nombre);
  } else {
    this.nombresSeleccionados.splice(idx, 1);
  }
}

hojaActual: 1 | 2 = 1;

verHoja(numero: 1 | 2) {
  this.hojaActual = numero;
}

registrarMotivoATodos(semana: number, mes: string, tabla: number, motivo: string) {
  if (!motivo) {
    alert('Selecciona un motivo.');
    return;
  }
  type Mes = 'enero' | 'febrero' | 'marzo' | 'abril' | 'mayo' | 'junio' | 'julio' | 'agosto' | 'septiembre' | 'octubre' | 'noviembre' | 'diciembre';
  const columnasMes: Record<Mes, number> = {
    enero: 10, febrero: 15, marzo: 20, abril: 25, mayo: 30, junio: 35, julio: 40,
    agosto: 45, septiembre: 50, octubre: 55, noviembre: 60, diciembre: 65
  };
  const mesKey = mes.toLowerCase() as Mes;
  const colInicio = columnasMes[mesKey];
  if (colInicio === undefined) return;

  const columnaAsistencia = colInicio + (semana - 1);

  const rangos = [
    { inicio: 6, fin: 40 },
    { inicio: 53, fin: 87 },
    { inicio: 103, fin: 137 }
  ];
  const rango = rangos[tabla - 1];

  if (rango) {
    for (let i = rango.inicio; i <= rango.fin; i++) {
      const nombre = this.data[i]?.[2];
      if (nombre && nombre.trim() !== '' && nombre !== 'NOMBRES Y APELLIDOS') {
        this.data[i][columnaAsistencia] = motivo;
      }
    }
  }
  this.mensajeAsistencia = `Se registró "${motivo}" para todos los discípulos en la semana ${semana} de ${mes} (${this.tablaNombres[tabla - 1]}).`;
  setTimeout(() => this.mensajeAsistencia = '', 3000);
}
//Carga tabla
onFileChange(evt: any) {
  const target: DataTransfer = <DataTransfer>(evt.target);
  if (target.files.length !== 1) throw new Error('Solo se permite un archivo');
  const file = target.files[0];
  const reader: FileReader = new FileReader();
  reader.onload = (e: any) => {
    this.originalFileBuffer = e.target.result;
    const wb: XLSX.WorkBook = XLSX.read(e.target.result, { type: 'array' });
    const wsname: string = wb.SheetNames[0];
    const ws: XLSX.WorkSheet = wb.Sheets[wsname];
    this.data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    this.merges = ws['!merges'] || [];
    if (wb.SheetNames.length > 1) {
      const wsname2: string = wb.SheetNames[1];
      const ws2: XLSX.WorkSheet = wb.Sheets[wsname2];
      this.dataHoja2 = XLSX.utils.sheet_to_json(ws2, { header: 1 });
      this.mergesHoja2 = ws2['!merges'] || [];
    }
  };
  reader.readAsArrayBuffer(file);
}
isMergedCell(row: number, col: number): boolean {
  return this.merges.some(
    merge =>
      row >= merge.s.r && row <= merge.e.r &&
      col >= merge.s.c && col <= merge.e.c &&
      !(row === merge.s.r && col === merge.s.c)
  );
}

getColspan(row: number, col: number): number {
  const merge = this.merges.find(
    m => m.s.r === row && m.s.c === col
  );
  return merge ? merge.e.c - merge.s.c + 1 : 1;
}

getRowspan(row: number, col: number): number {
  const merge = this.merges.find(
    m => m.s.r === row && m.s.c === col
  );
  return merge ? merge.e.r - merge.s.r + 1 : 1;
}


// Descarga el archivo original
descargarOriginal() {
  if (!this.originalFileBuffer) {
    throw new Error('No se ha cargado ningún archivo.');
  }
  FileSaver.saveAs(
    new Blob([this.originalFileBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    'archivo_original.xlsx'
  );
}


 
// Agrega nombre y edad en C y J de las tres tablas automáticamente
agregarPersonaEnTodasLasTablas() {
  const rangos = [
    { inicio: 6, fin: 40 },    // Celugrama
    { inicio: 53, fin: 87 },   // Somos Uno
    { inicio: 103, fin: 137 }  // Intercesión
  ];

  for (const rango of rangos) {
    // Verifica si ya existe el nombre en este rango
    const yaExiste = this.data
      .slice(rango.inicio, rango.fin + 1)
      .some(row => row[2] && row[2].trim() === this.nuevoNombreApellido.trim());

    if (yaExiste) continue;

    let filaLibre = -1;
    for (let i = rango.inicio; i <= rango.fin; i++) {
      if (!this.data[i][2] && !this.data[i][9]) {
        filaLibre = i;
        break;
      }
    }
    if (filaLibre === -1) {
      const nuevaFila = Array(Math.max(this.data[0]?.length || 20, 10)).fill('');
      nuevaFila[2] = this.nuevoNombreApellido;
      nuevaFila[9] = this.nuevaEdad;
      this.data.splice(rango.fin + 1, 0, nuevaFila);
    } else {
      if (this.data[filaLibre].length < 10) {
        this.data[filaLibre].length = 10;
      }
      this.data[filaLibre][2] = this.nuevoNombreApellido;
      this.data[filaLibre][9] = this.nuevaEdad;
    }
  }

  this.nuevoNombreApellido = '';
  this.nuevaEdad = '';
  this.mostrarFormulario = false;
}

// Editar persona en todas las tablas (por índice de fila y nuevos valores)
editarUsuarioEnTodasLasTablas(nombreOriginal: string, nuevoNombre: string, nuevaEdad: string) {
  const rangos = [
    { inicio: 6, fin: 40 },
    { inicio: 53, fin: 87 },
    { inicio: 103, fin: 137 }
  ];
  for (const rango of rangos) {
    for (let i = rango.inicio; i <= rango.fin; i++) {
      if (this.data[i]?.[2] === nombreOriginal) {
        this.data[i][2] = nuevoNombre;
        this.data[i][9] = nuevaEdad;
      }
    }
  }
}

// Eliminar persona en todas las tablas (por índice de fila)
eliminarUsuarioEnTodasLasTablas(nombre: string) {
  const rangos = [
    { inicio: 6, fin: 40 },
    { inicio: 53, fin: 87 },
    { inicio: 103, fin: 137 }
  ];
  for (const rango of rangos) {
    for (let i = rango.inicio; i <= rango.fin; i++) {
      if (this.data[i]?.[2] === nombre) {
        this.data[i][2] = '';
        this.data[i][9] = '';
      }
    }
  }
}


tablaSeleccionada: number = 1;
nombreBusqueda: string = '';
mesSeleccionado: string = '';
semanaSeleccionada: number = 0;
mensajeAsistencia: string = '';

private tablaNombres = ['Celugrama', 'Somos Uno', 'Intercesión'];

registrarAsistencia(nombres: string[], semana: number, mes: string = 'enero', tabla: number = 1) {
  type Mes = 'enero' | 'febrero' | 'marzo' | 'abril' | 'mayo' | 'junio' | 'julio' | 'agosto' | 'septiembre' | 'octubre' | 'noviembre' | 'diciembre';
  const columnasMes: Record<Mes, number> = {
    enero: 10, febrero: 15, marzo: 20, abril: 25, mayo: 30, junio: 35, julio: 40,
    agosto: 45, septiembre: 50, octubre: 55, noviembre: 60, diciembre: 65
  };
  const mesKey = mes.toLowerCase() as Mes;
  const colInicio = columnasMes[mesKey];
  if (colInicio === undefined) return;

  const columnaAsistencia = colInicio + (semana - 1);

  // Rangos de las tablas
  const rangos = [
    { inicio: 6, fin: 40 },    // Tabla 1
    { inicio: 53, fin: 87 },   // Tabla 2
    { inicio: 103, fin: 137 }  // Tabla 3
  ];

  const rango = rangos[tabla - 1];
  let registrados: string[] = [];
  let noEncontrados: string[] = [];

  if (rango) {
    for (const nombre of nombres) {
      let encontrado = false;
      for (let i = rango.inicio; i <= rango.fin; i++) {
        if (this.data[i]?.[2] === nombre) {
          this.data[i][columnaAsistencia] = '✅';
          encontrado = true;
        }
      }
      if (encontrado) {
        registrados.push(nombre);
      } else {
        noEncontrados.push(nombre);
      }
    }
  }

  if (registrados.length > 0) {
    this.mensajeAsistencia = `Se registró asistencia para: ${registrados.join(', ')} en la semana ${semana} de ${mes} (${this.tablaNombres[tabla - 1]}).`;
  }
  if (noEncontrados.length > 0) {
    this.mensajeAsistencia += ` No se encontró: ${noEncontrados.join(', ')}.`;
  }

  // Limpia el formulario
  this.nombreBusqueda = '';
  this.mesSeleccionado = '';
  this.semanaSeleccionada = 0;
  this.tablaSeleccionada = 0;

  setTimeout(() => this.mensajeAsistencia = '', 3000);
}

// Nuevo método para eliminar asistencia
eliminarAsistencia(nombre: string, semana: number, mes: string = 'enero', tabla: number = 1) {
  type Mes = 'enero' | 'febrero' | 'marzo' | 'abril' | 'mayo' | 'junio' | 'julio' | 'agosto' | 'septiembre' | 'octubre' | 'noviembre' | 'diciembre';
  const columnasMes: Record<Mes, number> = {
    enero: 10, febrero: 15, marzo: 20, abril: 25, mayo: 30, junio: 35, julio: 40,
    agosto: 45, septiembre: 50, octubre: 55, noviembre: 60, diciembre: 65
  };
  const mesKey = mes.toLowerCase() as Mes;
  const colInicio = columnasMes[mesKey];
  if (colInicio === undefined) return;

  const columnaAsistencia = colInicio + (semana - 1);

  const rangos = [
    { inicio: 6, fin: 40 },
    { inicio: 53, fin: 87 },
    { inicio: 103, fin: 137 }
  ];

  const rango = rangos[tabla - 1];
  let eliminado = false;
  if (rango) {
    for (let i = rango.inicio; i <= rango.fin; i++) {
      if (this.data[i]?.[2] === nombre) {
        this.data[i][columnaAsistencia] = '';
        eliminado = true;
      }
    }
  }

  if (eliminado) {
    this.mensajeAsistencia = `Se eliminó la asistencia de ${nombre} en la semana ${semana} de ${mes} (${this.tablaNombres[tabla - 1]}).`;
    this.nombreBusqueda = '';
    this.mesSeleccionado = '';
    this.semanaSeleccionada = 0;
    this.tablaSeleccionada = 0;
  } else {
    this.mensajeAsistencia = `No se encontró el usuario "${nombre}" en ${this.tablaNombres[tabla - 1]}.`;
  }
  setTimeout(() => this.mensajeAsistencia = '', 3000);
}

getUsuariosUnicos(): string[] {
  if (!this.tablaSeleccionada) return [];

  const rangos = [
    { inicio: 6, fin: 40 },    // Celugrama
    { inicio: 53, fin: 87 },   // Somos Uno
    { inicio: 103, fin: 137 }  // Intercesión
  ];
  const rango = rangos[this.tablaSeleccionada - 1];
  if (!rango) return [];

  const nombresSet = new Set<string>();
  for (let i = rango.inicio; i <= rango.fin; i++) {
    const nombre = this.data[i]?.[2];
    if (nombre && nombre.trim() !== '') {
      nombresSet.add(nombre.trim());
    }
  }
  return Array.from(nombresSet);
}


// Exporta la tabla editada a un nuevo archivo Excel usando ExcelJS
async exportToExcel(): Promise<void> {
  if (!this.originalFileBuffer) {
    alert('Primero debes cargar un archivo Excel.');
    return;
  }

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(this.originalFileBuffer);
    const worksheet = workbook.worksheets[0];
    const worksheet2 = workbook.worksheets[1];

    // Normaliza las filas para la hoja 1
    const maxCols1 = Math.max(...this.data.map(row => row.length));
    const dataNormalizada1 = this.data.map(row => {
      const nuevaFila = [...row];
      while (nuevaFila.length < maxCols1) nuevaFila.push('');
      return nuevaFila;
    });
    this.data = dataNormalizada1;

    // Normaliza las filas para la hoja 2 (si existe)
    if (worksheet2 && this.dataHoja2) {
      const maxCols2 = Math.max(...this.dataHoja2.map(row => row.length));
      const dataNormalizada2 = this.dataHoja2.map(row => {
        const nuevaFila = [...row];
        while (nuevaFila.length < maxCols2) nuevaFila.push('');
        return nuevaFila;
      });
      this.dataHoja2 = dataNormalizada2;
    }

    // Actualiza ambas hojas
    this.actualizarDatosEnHoja(worksheet, this.data); // Hoja 1 (igual que antes)
      if (worksheet2 && this.dataHoja2) {
      this.actualizarDatosEnHoja2(worksheet2, this.dataHoja2); // Hoja 2 (nueva lógica)
    }
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    FileSaver.saveAs(blob, 'tabla_editada.xlsx');
  } catch (error: any) {
    console.error('Error al exportar a Excel:', error);
    alert('Hubo un problema al generar el archivo Excel.');
  }
}

async guardarEnBaseDeDatos(): Promise<void> {
  if (!this.originalFileBuffer) {
    alert('Primero debes cargar un archivo Excel.');
    return;
  }

  const usuario = localStorage.getItem('usuario');
  if (!usuario) {
    alert('No hay usuario logueado.');
    return;
  }

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(this.originalFileBuffer);
    const worksheet = workbook.worksheets[0];

    // Rangos de las tablas
    const rangos = [
      { inicio: 6, fin: 40 },
      { inicio: 53, fin: 87 },
      { inicio: 103, fin: 137 }
    ];

    // Actualiza SOLO las filas originales, respetando su posición
    for (const rango of rangos) {
      for (let i = rango.inicio; i <= rango.fin; i++) {
        const dataRow = this.data[i];
        if (!dataRow) continue;
        const nombre = dataRow[2]?.trim();
        // Si hay nombre, actualiza; si no, limpia
        worksheet.getRow(i + 1).getCell(3).value = nombre ? dataRow[2] : null;
        worksheet.getRow(i + 1).getCell(10).value = dataRow[9] ? dataRow[9] : null;
        worksheet.getRow(i + 1).commit();
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const file = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    const formData = new FormData();
    formData.append('file', file, 'tabla_editada.xlsx');
    formData.append('nombre', usuario);

    await firstValueFrom(this.http.post('http://localhost:3000/api/guardar-excel', formData));
    alert('Archivo guardado correctamente en la base de datos.');
  } catch (error: any) {
    console.error('Error al guardar el archivo en la base de datos:', error);
    alert('Ocurrió un error al guardar el archivo en el servidor.');
  }
}

private actualizarDatosEnHoja(worksheet: ExcelJS.Worksheet, data: any[][]): void {
  const filaInicio = 7;
  const nombresEscritos = new Set<string>();
  for (let i = 0; i < this.data.length; i++) {
    const dataRow = this.data[i];
    if (!dataRow) continue;
    const nombre = dataRow[2]?.trim();
    if (!nombre || nombresEscritos.has(nombre)) continue; // Solo la primera vez

    const excelRow = worksheet.getRow(i + 1); // i+1 porque this.data[7] es fila 8 en Excel
    excelRow.getCell(3).value = dataRow[2] === '' ? null : dataRow[2];
    excelRow.getCell(10).value = dataRow[9] === '' ? null : dataRow[9];
    excelRow.commit();
    nombresEscritos.add(nombre);
  }
}

private actualizarDatosEnHoja2(worksheet: ExcelJS.Worksheet, data: any[][]): void {
  for (let i = 0; i < data.length; i++) {
    const dataRow = data[i];
    if (!dataRow) continue;
    const excelRow = worksheet.getRow(i + 1);

    // Copia la columna A a la B
    excelRow.getCell(2).value = dataRow[0] === '' ? null : dataRow[0];

    // Copia el resto de columnas en su lugar (B, C, D, ...)
    for (let j = 1; j < dataRow.length; j++) {
      excelRow.getCell(j + 1).value = dataRow[j] === '' ? null : dataRow[j];
    }
    excelRow.commit();
  }
}
// Exporta la tabla actual a un archivo Excel usando XLSX (sin formato ni merges avanzados)
exportarExcel(): void {
  if (!this.data || this.data.length === 0) {
    alert('No hay datos para exportar.');
    return;
  }

  // Normaliza las filas para que todas tengan la misma cantidad de columnas
  const maxCols = Math.max(...this.data.map(row => row.length));
  const dataNormalizada = this.data.map(row => {
    const nuevaFila = [...row];
    while (nuevaFila.length < maxCols) nuevaFila.push('');
    return nuevaFila;
  });

  const worksheet: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(dataNormalizada);
  const workbook: XLSX.WorkBook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Datos');

  // Aplica los merges si existen
  if (this.merges && this.merges.length > 0) {
    worksheet['!merges'] = this.merges;
  }

  const excelBuffer: ArrayBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  const blob: Blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  FileSaver.saveAs(blob, 'datos_exportados.xlsx');
}

CargarMiArchivo() {
  // Verifica que window esté definido (solo en navegador)
  if (typeof window === 'undefined') {
    return;
  }
  const usuario = localStorage.getItem('usuario');
  if (!usuario) {
    alert('No hay usuario logueado.');
    return;
  }
  this.http.get(`http://localhost:3000/api/obtener-excel/${usuario}`, { responseType: 'arraybuffer' })
    .subscribe({
      next: (arrayBuffer) => {
        this.originalFileBuffer = arrayBuffer;
        const wb: XLSX.WorkBook = XLSX.read(arrayBuffer, { type: 'array' });
        const wsname: string = wb.SheetNames[0];
        const ws: XLSX.WorkSheet = wb.Sheets[wsname];
        this.data = XLSX.utils.sheet_to_json(ws, { header: 1 });

        const maxCols = Math.max(...this.data.map(row => row.length));
        this.data = this.data.map(row => {
          while (row.length < maxCols) row.push('');
          return row;
        });

        this.merges = ws['!merges'] || [];
        console.log('Datos cargados:', this.data);
      },
      error: () => alert('Archivo no encontrado o error en la carga')
    });
}

isMergedCellHoja2(row: number, col: number): boolean {
  return this.mergesHoja2.some(
    merge =>
      row >= merge.s.r && row <= merge.e.r &&
      col >= merge.s.c && col <= merge.e.c &&
      !(row === merge.s.r && col === merge.s.c)
  );
}
getColspanHoja2(row: number, col: number): number {
  const merge = this.mergesHoja2.find(
    m => m.s.r === row && m.s.c === col
  );
  return merge ? merge.e.c - merge.s.c + 1 : 1;
}
getRowspanHoja2(row: number, col: number): number {
  const merge = this.mergesHoja2.find(
    m => m.s.r === row && m.s.c === col
  );
  return merge ? merge.e.r - merge.s.r + 1 : 1;
 }

 mesOfrenda: string = 'enero';
semanaOfrenda: number = 1;
cantidadOfrenda: number = 0;
cantidadPersonasOfrenda: number = 0; // <-- Agrega esta línea

public marcarPersonasOfrendaHoja2(mes: string, semana: number, cantidad: number): void {
  const columnasMes: Record<string, number> = {
    enero: 3,     // D = 4, pero índice base 0, D=3
    febrero: 8,   // I = 9, índice 8
    marzo: 13,    // N = 14, índice 13
    abril: 18,    // S = 19, índice 18
    mayo: 23,     // X = 24, índice 23
    junio: 28,    // AC = 29, índice 28
    // ...agrega más meses si tienes
  };

  const mesKey = mes.toLowerCase();
  const colInicio = columnasMes[mesKey];
  if (colInicio === undefined) return;

  // Fila 29 (índice 28) es la fila de los círculos azules
  const filaCirculos = 28;
  // Columna de la semana (horizontal)
  const columnaInicio = colInicio + (semana - 1);

  // Limpia primero las celdas de la semana seleccionada (por si se vuelve a registrar)
  for (let i = 0; i < 5; i++) {
    if (!this.dataHoja2[filaCirculos]) this.dataHoja2[filaCirculos] = [];
    this.dataHoja2[filaCirculos][colInicio + i] = '';
  }

  // Inserta los círculos azules según la cantidad
  for (let i = 0; i < cantidad; i++) {
    this.dataHoja2[filaCirculos][columnaInicio + i] = '🔵';
  }
}

public registrarOfrendaHoja2(mes: string, semana: number, cantidad: number): void {
  // Mapeo de columnas de inicio por mes (ajusta según tu archivo)
  const columnasMes: Record<string, number> = {
    enero: 3,     // D = 4, pero índice base 0, D=3
    febrero: 8,   // I = 9, índice 8
    marzo: 13,    // N = 14, índice 13
    abril: 18,    // S = 19, índice 18
    mayo: 23,     // X = 24, índice 23
    junio: 28,    // AC = 29, índice 28
    // ...agrega más meses si tienes
  };

  const mesKey = mes.toLowerCase();
  const colInicio = columnasMes[mesKey];
  if (colInicio === undefined) return;

  // Fila 30 (índice 29) es la fila de "Ofrenda"
  const filaOfrenda = 29;
  // Columna de la semana (horizontal)
  const columnaOfrenda = colInicio + (semana - 1);

  // Inserta la cantidad en la celda correspondiente
  if (!this.dataHoja2[filaOfrenda]) this.dataHoja2[filaOfrenda] = [];
  this.dataHoja2[filaOfrenda][columnaOfrenda] = cantidad;
}

}



