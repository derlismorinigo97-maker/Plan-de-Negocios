# Revisión del informe de Alianza del Acero (RV00)

Fecha de revisión: 09/10/2026. Archivos revisados:

- `INFORME_ALIANZA_DEL_ACERO_RV00.xlsx`: 13 hojas, 1.467 fórmulas.
- `ALIANZA_Informe_Ejecutivo_RV00.pptx`: 9 diapositivas.

Resultado: `ALIANZA_Informe_Ejecutivo_RV01.pptx`, con correcciones de cifras y textos (sección 5). El Excel no se modificó.

## 1. Conclusión

**El cálculo es consistente, pero el informe todavía no está listo para enviarse como cifra cerrada.**

- **Lo que está bien.** Las cifras financieras del PPT (diapositivas 4 a 8) coinciden con el Excel al guaraní, con diferencias de redondeo de ±1. Esto incluye los 29 puntos del flujo, la serie de TC, el puente de facturación y las 23 partidas del aporte. Las fórmulas se reprodujeron por fuera y no hay errores ni vínculos externos. Las diapositivas 2, 3 y 9, y el cuadro de capital de la 8, no tienen fuente en el Excel: se verificaron solo por consistencia interna.
- **Tres puntos pueden cambiar el mensaje** (sección 3):
  1. una partida del aporte sin respaldo, de Gs 287 M;
  2. el flete y despacho de la chapa CUMIC, de USD 557.616, que sigue sin incluirse;
  3. una fecha de autonomía que contradice el propio cronograma.
- **Cuatro criterios conviene explicitarlos ante la Gerencia** (sección 4). Los tres primeros suman, sin duplicar, unos USD 170 mil a favor de CIE:
  1. intereses al 9% del plan frente a la tasa bancaria real;
  2. overhead sobre intereses;
  3. IVA cobrado como costo;
  4. moneda de facturación.

Lectura financiera para la Gerencia. Después de cobrar todo, la caja de CIE termina en **+USD 19.511**, y CIE conserva su participación de Gs 1.945 M. El resultado se explica así:

| Componente | USD |
|---|---:|
| Overhead 7% (no es salida de caja) | +198.392 |
| Diferencia entre intereses al 9% e intereses pagados al banco | +73.600 |
| Efecto TC (gastos en Gs dolarizados al TC de cierre) | +59.810 |
| Cobro de ventas al TC del mes de cobro frente al TC de factura | +19.807 |
| Aporte en especie: se descuenta y no vuelve en efectivo | −332.098 |
| **Caja final** | **+19.511** |

En otras palabras, **el overhead y el diferencial de intereses son los que financian el aporte de capital de CIE**. Si los socios negocian a la baja cualquiera de los dos, la caja de CIE queda negativa.

## 2. Qué se verificó y está correcto

| Verificación | Resultado |
|---|---|
| Errores en valores guardados del Excel | 0 (los «#REF!» de Claude Log son texto). Sin vínculos externos. |
| Diap. 4 y 7: total Gs 17.742.564.318 = USD 3.032.559; saldo Gs 10.902.067.391 = USD 1.863.381 | Coinciden con INFORME AA!G64, H64, G72 y H72 (TC 5.850,69). |
| Diap. 5: flujo acumulado | Los 29 meses coinciden con FLUJO fila 49 (oct-26: 731.896; nov-26: −22.746; ene-27: 19.511). |
| Diap. 6: TC mensual y recupero a TC de cierre y a TC del flujo | Coinciden con FLUJO fila 36 y C39 a C44. Las diferencias +79.683, +4.067 y +59.810 son correctas. |
| Diap. 8: 23 partidas del aporte | Coinciden una por una con APORTE ALIANZA E3:E25 y suman Gs 1.943.000.000. |
| Diap. 8 y 9: capital social | Las sumas y los porcentajes cuadran (49% = 1.945/3.970; 49% = 643/1.311). |
| Fórmulas matriciales del INFORME AA (activos por año, obras civiles, IVA de producción, intereses al 9%) | Recalculadas por fuera desde ACTIVOS FIJOS, STOCK, FLUJO y PRESTAMO: coinciden. |
| Intereses al 9% | 2.500.000 × 9% × (270 + 241) / 360 = USD 319.375 + IVA 31.937,50 = USD 351.312,50. |
| Intereses reales | Préstamo 1: 7,85%, USD 145.171,23. Préstamo 2: 6,50%, USD 107.294,52. Coinciden con PRESTAMO. |
| Conciliación FLUJO (Gs) frente al INFORME AA | **Implantación:** diferencia de Gs 1.472.423. Se explica porque la PC de 2019 está en el informe y no en el flujo (Gs 5.038.036 + IVA), y porque el flujo carga IVA sobre sueldos (Gs 7.014.262). **Producción:** diferencia de Gs 413.301.931. Se explica porque el informe valúa la chapa CUMIC por los USD pagados al TC de cierre, y el flujo la valúa al costo SAP. |

Nota sobre el método: LibreOffice no recalcula este libro de forma fiable, porque renombra «Total general» de las tablas dinámicas y evalúa distinto las matriciales. Por eso los cálculos se reprodujeron por fuera, en Python.

## 3. Puntos a resolver antes de enviar

### 3.1 Aporte en especie: partida sin respaldo (Gs 287.238.199)

- **Qué ocurre.** La línea «Adecuaciones e infraestructura» (APORTE ALIANZA!E21) está tipeada y es **exactamente el residuo** para llegar a Gs 1.943.000.000:
  - las otras 22 partidas suman Gs 1.655.761.801;
  - 1.943.000.000 − 1.655.761.801 = 287.238.199.
- **Por qué no se puede respaldar con lo ejecutado.** El bloque «Adecuaciones infraestructura» del INFORME AA, sin instalaciones eléctricas ni red de aire (que ya están en el aporte), suma solo Gs 232.821.812. Además, es casi todo gasto:

  | Concepto | Gs |
  |---|---:|
  | Gastos administrativos | 137,9 M |
  | Sueldos | 65,4 M |
  | EPIs | 16,1 M |
  | Obras civiles | 8,7 M |
  | Mano de obra de puesta a punto | 4,7 M |

  Nada de esto es un bien que pueda aportarse en especie.
- **Inconsistencias del cuadro.**
  - El flete de la perfiladora (Gs 150.771.654) está en el cuadro, pero la hoja de selección lo marca «NO» (GASTOS DE IMPLANTACION!I7).
  - El control de conciliación (J35) solo cubre E23:E25, así que las filas 21 y 22 quedan fuera del control.
- **Impacto.** Si la partida no se respalda, el aporte en especie baja a Gs 1.655.761.801 y el aporte en efectivo sube a Gs 289.238.199. El saldo a facturar aumenta en Gs 287.238.199 (USD 49.095). Para CIE el neto es el mismo, pero el cuadro debe poder sostenerse ante los socios y ante el escribano.

### 3.2 Flete y despacho de la chapa CUMIC (USD 557.616): sigue sin incluirse

- **Qué ocurre.** Figura como pendiente en el Claude Log desde el turno 23. El informe solo incluye el flete de la perfiladora.
- **Impacto.** Si lo pagó CIE, el saldo a recuperar sube **USD 596.649** con overhead (+32%), hasta unos USD 2,46 M.
- **Acción.** Confirmar con Comercio Exterior y Contabilidad antes de enviar. Si no corresponde, dejar escrito por qué se excluye.

### 3.3 Fecha de autonomía contradictoria (diapositivas 2 y 3)

- La diapositiva 2 fija la autonomía el 31/10/2026 («1 mes y 1 día»).
- En la diapositiva 3, la Habilitación Municipal (11%) y la Licencia Ambiental (56%) cierran el 14/01/2027.
- **Acción.** Mover la fecha de autonomía, o aclarar que esos dos trámites no condicionan la operación independiente.
- **Medidor del 56%.** No sale de la tabla: el promedio simple de los 12 rubros es 62%, y el Excel no tiene la fuente. Definir la ponderación o ajustar el medidor.

## 4. Criterios a explicitar ante la Gerencia

| # | Tema | Situación | Impacto en el saldo |
|---|---|---|---:|
| 4.1 | Intereses | Se cobra el 9% del plan sobre USD 2,5 M por 511 días, con base 360: USD 351.313 con IVA. Al banco se pagó 7,85% y 6,50%, con base 365: USD 277.712 con IVA. El cálculo incluye 5 días de superposición entre préstamos (del 01/04 al 06/04/2026, USD 3.125). | +73.600 (+78.752 con overhead) a favor de CIE |
| 4.2 | Overhead 7% | La base de producción incluye intereses e IVA (INFORME AA!F62 = (F32+F56)×7%). | USD 24.592 sobre intereses; USD 3.046 sobre IVA de compras |
| 4.3 | IVA como costo | 10% plano, no tomado de facturas: implantación USD 34.544, producción USD 8.968 e intereses USD 31.938. Si CIE lo computó como crédito fiscal, o si la factura a Alianza lleva IVA, se recupera dos veces. Validar con Impuestos. | Hasta −80.731 con overhead |
| 4.4 | Moneda de facturación (diap. 7) | Facturar en Gs protege los costos pagados en Gs. La chapa CUMIC (USD 1,91 M) y los intereses se pagaron en USD, y el préstamo vence en USD el 28/11/2026. Valuada en Gs, la chapa es Gs 11.186 M al TC de cierre frente a unos Gs 14.258 M al TC de cierre del mes de pago, así que la frase «recupera exactamente lo registrado en la contabilidad» no se cumple para esa parte. | Riesgo de TC hasta el cobro |
| 4.5 | TC de octubre provisorio (5.850,69, tipeado en FLUJO!AA36) | Cada 100 Gs de TC mueven el saldo en dólares unos USD 32 mil: con 5.750, USD 1.896.012; con 5.950, USD 1.832.280. Reemplazar por el cierre oficial del BCP al 31/10. | ±32.000 |
| 4.6 | Valuación del aporte | Se valúa a costo de adquisición SAP. Incluye una PC de 2019, un monitor de 2017 y perfiladoras con unos 22 meses de uso. También incluye gastos (instalaciones, flete). Los bienes están en los libros de CIE S.A., pero el accionista es CIE Maquiladora EAS. Validar con Legal y Contabilidad la base de valuación (costo, valor neto contable o tasación), la documentación y el tratamiento impositivo de la transferencia. | Cualitativo |
| 4.7 | Aporte en efectivo de Gs 2.000.000 | La diapositiva 8 dice «una vez cobrado», pero el cuadro de capital muestra que CIE ya integró Gs 2.000.000. Una de las dos afirmaciones es incorrecta. | Cualitativo |
| 4.8 | Moneda de las ventas | Si las exportaciones se facturaron en USD, el cobro real es USD 813.140 (diap. 6) y no USD 832.947 (FLUJO fila 39). En ese caso, la caja final pasa de +19.511 a unos −300. Confirmar la moneda de factura. | −19.807 en la caja final |

## 5. Correcciones aplicadas en la RV01

| Diap. | RV00 | RV01 | Motivo |
|---|---|---|---|
| 2 | Total transcurrido «1 año y 3 meses» | «1 año y 4 meses» | Recorrido de 17 meses − 1 mes restante = 16 meses (misma convención que «9 meses» de producción). |
| 3 | «AUTONOMIA OPERATIVA ALIANZA DEL ACERO» | «AUTONOMÍA OPERATIVA» | Tilde, y el título pasaba a dos líneas y salía de la banda superior. |
| 4 | Título «INVERSIÓN Y GASTOS TOTALES: USD 3.032.559» | «INVERSIÓN Y GASTOS: USD 3.032.559» | Desborde del título. |
| 4 | «Pagado por CIE S.A al 30/09/2026…» y un punto suelto | «Pagos de CIE S.A. al 30/09/2026 + proyección oct-26, overhead 7% e intereses al 9% del plan · TC provisorio» | El total incluye overhead e intereses al 9%, que no son pagos. |
| 4 | Recuadro «Proyección a ejecutar» | «Movimientos de oct-26». Se aclara que la chapa CUMIC ya está pagada y se agregan los gastos generales y el IVA | Las partidas sumaban 178.796, no 179.778: faltaban gastos generales (USD 564) e IVA (USD 419). La chapa consumida en octubre no cambia el total, porque el informe valúa CUMIC por pagos. |
| 5 | Eje vertical con mínimo de −250.000 | −500.000 | May y jun-25 (−251.604) quedaban fuera del gráfico. |
| 5 | «270 / 240 días · 9% anual», «Abril / Noviembre 2026» | «270 / 241 días · banco 7,85% / 6,50% · cobro 9% s/plan», «06/04/2026 / 28/11/2026» | Datos de PRESTAMO. El 9% no es la tasa bancaria. |
| 5 | Etiquetas sobre la línea del gráfico | Reubicadas. Tildes en «RENOVACIÓN» y «PRÉSTAMO» | Legibilidad. |
| 7 | «PRÓXIMOS PASOS: FACTURACIÓN A ALIANZA» | «FACTURACIÓN A ALIANZA» | Desborde del título. |
| 7 | «Aporte de CIE (maquinarias)» | «Aporte de CIE en especie» | El aporte incluye instalaciones, flete y adecuaciones. |
| 7 | Recomendación de facturación | Se mantiene facturar en Gs y se agrega la salvedad de la parte en USD (punto 4.4) | Precisión. |
| 9 | «Aportado en las tres maquilas: Gs 3.180.000.000» | «Capital suscrito en las tres maquilas: Gs 3.178.000.000» | 643 + 1.945 + 590 = 3.178. En Alianza solo hay Gs 2 M integrados. |
| 9 | Montos cortados en las barras de Alianza y TCM | Textos subidos dentro de la barra | Legibilidad. |

**No se modificó**, porque requiere tu decisión:

- el cuadro del aporte (diap. 8);
- el medidor del 56% y las fechas de autonomía (diap. 2 y 3);
- la columna «TC flujo» de la diapositiva 6: mezcla las ventas al TC de factura con el flujo al TC de cobro, y no recalcula el overhead a ese TC;
- el hito de compras de dic-2024: el flujo empieza en nov-24 y las perfiladoras se compraron el 05/12/2024, antes del «Inicio de implantación Jun/2025».

## 6. Sensibilidad del saldo a recuperar (base USD 1.863.381)

| Ajuste | Impacto USD | Saldo USD |
|---|---:|---:|
| Incluir el flete y despacho de CUMIC, si lo pagó CIE (con overhead) | +596.649 | 2.460.030 |
| Sacar la partida de Gs 287 M del aporte en especie | +49.095 | 1.912.476 |
| Intereses al costo bancario real (con IVA y overhead) | −78.752 | 1.784.629 |
| Sin overhead sobre intereses | −24.592 | 1.838.789 |
| Sin IVA, si es crédito fiscal (con overhead) | −80.731 | 1.782.650 |
| TC de oct-26 a 5.750 | +32.630 | 1.896.012 |
| TC de oct-26 a 5.950 | −31.101 | 1.832.280 |

Los ajustes no son aditivos: por ejemplo, el IVA de los intereses está tanto en la fila de intereses como en la de IVA.

## 7. Cambios sugeridos en el Excel (no aplicados)

- **APORTE ALIANZA E21:E22.** Reemplazar los montos tipeados por partidas con respaldo y extender el control J35 a todo el cuadro.
- **GASTOS DE IMPLANTACION K5:K31 e I7.** Actualizar las notas: siguen diciendo «queda dentro del aporte en efectivo» y citan «filas 22 a 24», cuando hoy son las filas 23 a 25. Corregir también el «NO» del flete.
- **FLUJO D10:AF10.** Excluir los sueldos de la base del IVA, igual que en INFORME AA fila 30.
- **FLUJO AA14, AA22 y AA36.** La mano de obra y el TC de octubre están tipeados: marcarlos como datos de entrada (proyección).
- **INFORME AA C74:C75.** Alinear con el PPT: «Datos reales al 30/09/2026 + proyección oct-26». La fecha de informe 02/10/2026 está tipeada.
- **STOCK.** La tabla muestra saldos negativos por material (por ejemplo, Z120 0,48: −1.473.669.353), así que son movimientos y no existencias. En el total no tiene efecto, porque el stock CUMIC se resta en F35 y se suma en F44. Aun así, la línea «Stock de materia prima USD 93.772» no es un saldo de inventario.
- **COMPROMETIDO.** El monto «por ejecutar» neta órdenes sobreejecutadas (por ejemplo, CIE do Brasil −USD 668,93). El efecto es menor.

## Fuentes de tipo de cambio (consulta del 09/10/2026)

No se pudo acceder a la serie oficial del BCP. Las referencias de prensa:

- La cotización de referencia de cierre de septiembre de 2026 se informa en torno a G. 5.873; el libro usa 5.833,63 ([MarketData](https://marketdata.com.py/noticias/dolar-baja-a-g-5-873-y-retoma-tendencia-bajista-al-cierre-de-septiembre-156007/)). El valor de septiembre solo afecta la columna «TC flujo».
- El interbancario cerró por debajo de G. 5.850 el 29/09 ([La Nación](https://www.lanacion.com.py/negocios/2026/09/29/dolar-cierra-el-mercado-interbancario-por-debajo-de-los-g-5850/)).
- A comienzos de octubre, el interbancario rondaba G. 5.700–5.750 ([MarketData](https://marketdata.com.py/noticias/dolar-continua-su-baja-y-llega-a-g-5-700-mientras-el-guarani-profundiza-su-apreciacion-156193)).
- La apreciación acumulada en 2026 superaba el 10% ([Infobae/EFE, 03/08/2026](https://www.infobae.com/america/agencias/2026/08/03/la-moneda-paraguaya-se-aprecia-mas-de-un-10-frente-al-dolar-en-lo-que-va-de-2026/)).

Antes de facturar, conviene validar la serie contra [bcp.gov.py](https://www.bcp.gov.py).
