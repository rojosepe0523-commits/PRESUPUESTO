/**
 * Repara la hoja "Plantilla_Presupuesto" y la deja en pesos colombianos (COP).
 *
 * Qué hace:
 *  1. Configura la hoja en español (Colombia): coma decimal y punto de miles.
 *  2. Vuelve a escribir todas las fórmulas de "Categorias" y "Resumen" con setFormula(),
 *     que siempre entiende la sintaxis en inglés (comas), sin importar la configuración regional.
 *     Las fórmulas que estaban escritas como texto daban #ERROR! (error de análisis).
 *  3. Pone los presupuestos mensuales por categoría (incluida "Deudas").
 *  4. Da formato de pesos sin centavos ($ 1.250.000) a montos, presupuestos y saldos.
 *  5. Rehace el gráfico de "Resumen" para que lea solo números.
 *
 * Cómo usarlo: pega este archivo en Extensiones > Apps Script (como archivo nuevo),
 * elige la función "repararHoja" y pulsa Ejecutar. Se puede ejecutar las veces que quieras.
 */

const REP_PRESUPUESTOS = [
  ['Comida', 300000],
  ['Supermercado', 200000],
  ['Transporte', 150000],
  ['Casa', 200000],
  ['Servicios', 150000],
  ['Salud', 100000],
  ['Entretenimiento', 100000],
  ['Ropa', 100000],
  ['Educación', 100000],
  ['Otros', 100000],
  ['Mascotas', 200000],
  ['Ahorro de emergencia', 100000],
  ['Deudas', 2000000]
];

const REP_FORMATO_COP = '"$"#,##0';
const REP_FORMATO_FECHA = 'dd/mm/yyyy';
const REP_FILAS_RESUMEN = 30; // Resumen!A8:C37: espacio para hasta 30 categorías

function repararHoja() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetLocale('es_CO');

  repCategorias_(ss);
  repResumen_(ss);
  repMovimientos_(ss);
  repGrafico_(ss);

  SpreadsheetApp.flush();
  ss.toast('Hoja reparada: fórmulas, pesos colombianos y presupuestos actualizados.', 'Presupuesto', 8);
}

function repHoja_(ss, nombre) {
  const h = ss.getSheetByName(nombre);
  if (!h) throw new Error('No encuentro la pestaña "' + nombre + '".');
  return h;
}

function repCategorias_(ss) {
  const h = repHoja_(ss, 'Categorias');
  const n = REP_PRESUPUESTOS.length;
  const ultima = Math.max(h.getLastRow(), n + 1);

  h.getRange(1, 1, 1, 5).setValues([['Categoría', 'Presupuesto mensual', 'Gastado este periodo', 'Disponible', '% usado']]);
  h.getRange(2, 1, ultima - 1, 5).clearContent();
  h.getRange(2, 1, n, 2).setValues(REP_PRESUPUESTOS);

  const formulas = [];
  for (let f = 2; f <= n + 1; f++) {
    formulas.push([
      '=SUMIFS(Gastos!D:D,Gastos!C:C,A' + f + ',Gastos!A:A,">="&Resumen!$B$2)',
      '=B' + f + '-C' + f,
      '=IF(B' + f + '=0,0,C' + f + '/B' + f + ')'
    ]);
  }
  h.getRange(2, 3, n, 3).setFormulas(formulas);

  h.getRange(2, 2, n, 3).setNumberFormat(REP_FORMATO_COP);
  h.getRange(2, 5, n, 1).setNumberFormat('0%');
}

function repResumen_(ss) {
  const h = repHoja_(ss, 'Resumen');

  h.getRange('A1').setValue('Resumen del periodo (desde el último salario)');
  h.getRange('A2:A5').setValues([['Inicio del periodo'], ['Ingresos del periodo'], ['Total gastado'], ['Saldo']]);
  h.getRange('B2:B5').setFormulas([
    ['=IF(COUNTIFS(Ingresos!C:C,"Salario")=0,EOMONTH(TODAY(),-1)+1,MAXIFS(Ingresos!A:A,Ingresos!C:C,"Salario"))'],
    ['=SUMIFS(Ingresos!D:D,Ingresos!A:A,">="&B2)'],
    ['=SUMIFS(Gastos!D:D,Gastos!A:A,">="&B2)'],
    ['=B3-B4']
  ]);
  h.getRange('B2').setNumberFormat(REP_FORMATO_FECHA);
  h.getRange('B3:B5').setNumberFormat(REP_FORMATO_COP);

  h.getRange('A7:C7').setValues([['Categoría', 'Gastado', 'Presupuesto']]);

  // Una fórmula por fila (sin FILTER en A8, que chocaba con las filas de abajo).
  // Las filas sin categoría quedan vacías de verdad (IF(...,,...)), así el gráfico solo ve números.
  const rango = h.getRange(8, 1, REP_FILAS_RESUMEN, 3);
  rango.clearContent();
  const formulas = [];
  for (let i = 0; i < REP_FILAS_RESUMEN; i++) {
    const f = 8 + i;      // fila en Resumen
    const c = 2 + i;      // fila en Categorias
    formulas.push([
      '=IF(Categorias!A' + c + '="",,Categorias!A' + c + ')',
      '=IF(A' + f + '="",,SUMIFS(Gastos!D:D,Gastos!C:C,A' + f + ',Gastos!A:A,">="&$B$2))',
      '=IF(A' + f + '="",,IFERROR(VLOOKUP(A' + f + ',Categorias!A:B,2,FALSE),0))'
    ]);
  }
  rango.setFormulas(formulas);
  h.getRange(8, 2, REP_FILAS_RESUMEN, 2).setNumberFormat(REP_FORMATO_COP);
  h.getRange(8, 1, REP_FILAS_RESUMEN, 1).setHorizontalAlignment('left');
}

function repMovimientos_(ss) {
  ['Gastos', 'Ingresos'].forEach(nombre => {
    const h = repHoja_(ss, nombre);
    const filas = h.getMaxRows() - 1;
    h.getRange(2, 1, filas, 1).setNumberFormat(REP_FORMATO_FECHA);
    h.getRange(2, 4, filas, 1).setNumberFormat(REP_FORMATO_COP);
  });
}

function repGrafico_(ss) {
  const h = repHoja_(ss, 'Resumen');
  h.getCharts().forEach(g => h.removeChart(g));
  const grafico = h.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(h.getRange(7, 1, REP_FILAS_RESUMEN + 1, 2)) // Categoría + Gastado
    .setNumHeaders(1)
    .setPosition(2, 5, 0, 0)
    .setOption('title', 'Gasto del periodo por categoría')
    .setOption('pieSliceText', 'percentage')
    .setOption('width', 560)
    .setOption('height', 360)
    .build();
  h.insertChart(grafico);
}
