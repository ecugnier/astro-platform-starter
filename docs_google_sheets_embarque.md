# Google Sheets + Apps Script para seguimiento de embarques

## 1) Estructura del archivo en Google Sheets

Crea un Google Sheets con estas hojas:

### Hoja: `embarques`
Cada fila es una orden/booking/contenedor.

Encabezados sugeridos (fila 1):

- `order_number`
- `booking_number`
- `container_number`
- `cliente`
- `origen`
- `puerto_origen`
- `fecha_llegada_puerto` (YYYY-MM-DD)
- `buque`
- `fecha_embarque` (YYYY-MM-DD)
- `puerto_transbordo`
- `fecha_transbordo` (YYYY-MM-DD)
- `puerto_destino`
- `eta_destino` (YYYY-MM-DD)
- `estado_actual`
- `ultima_actualizacion`
- `notas`

### Hoja: `busqueda`
- `B2`: campo para que el cliente escriba **número de orden**, **booking** o **contenedor**.
- `A5:B20`: espacio para mostrar resultados.

### Hoja: `catalogo_estados` (opcional)
Lista de estados permitidos:
- En origen
- En puerto de origen
- Embarcado
- En transbordo
- En tránsito
- Arribado
- Entregado

---

## 2) Apps Script

En el menú de Google Sheets: **Extensiones > Apps Script** y pega este código en `Code.gs`.

```javascript
/**
 * Busca un embarque por número de orden, booking o contenedor.
 * Se usa desde la hoja "busqueda", celda B2.
 */
function buscarEmbarque() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shData = ss.getSheetByName('embarques');
  const shSearch = ss.getSheetByName('busqueda');

  const query = (shSearch.getRange('B2').getValue() || '').toString().trim().toUpperCase();

  // Limpia resultados previos
  shSearch.getRange('A5:B30').clearContent();

  if (!query) {
    shSearch.getRange('A5').setValue('Escribe un número de orden, booking o contenedor en B2.');
    return;
  }

  const values = shData.getDataRange().getValues();
  if (values.length < 2) {
    shSearch.getRange('A5').setValue('No hay datos en la hoja "embarques".');
    return;
  }

  const headers = values[0];
  const rows = values.slice(1);

  const idx = {
    order: headers.indexOf('order_number'),
    booking: headers.indexOf('booking_number'),
    container: headers.indexOf('container_number'),
  };

  if (idx.order === -1 || idx.booking === -1 || idx.container === -1) {
    shSearch.getRange('A5').setValue('Faltan columnas obligatorias: order_number, booking_number, container_number.');
    return;
  }

  const match = rows.find((r) => {
    const order = (r[idx.order] || '').toString().trim().toUpperCase();
    const booking = (r[idx.booking] || '').toString().trim().toUpperCase();
    const container = (r[idx.container] || '').toString().trim().toUpperCase();
    return order === query || booking === query || container === query;
  });

  if (!match) {
    shSearch.getRange('A5').setValue('No se encontró información para: ' + query);
    return;
  }

  // Mostrar resultado clave-valor
  const output = [];
  headers.forEach((h, i) => {
    output.push([h, match[i]]);
  });

  shSearch.getRange(5, 1, output.length, 2).setValues(output);
  shSearch.getRange('A5:B5').setFontWeight('bold');
}

/**
 * Menú personalizado al abrir el archivo.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Seguimiento')
    .addItem('Buscar embarque', 'buscarEmbarque')
    .addToUi();
}
```

---

## 3) Cómo usarlo (flujo del cliente)

1. El cliente abre el Google Sheets compartido en modo solo lectura.
2. Va a la hoja `busqueda`.
3. Escribe en `B2` cualquiera de estos identificadores:
   - número de orden,
   - número de booking,
   - número de contenedor.
4. Ejecuta el menú **Seguimiento > Buscar embarque**.
5. Ve en pantalla todo el estado del envío.

> Si quieres que la búsqueda sea automática al escribir en `B2`, también te lo puedo dejar con trigger `onEdit(e)`.

---

## 4) Importar desde tu Excel de Drive

Opciones:

- Manual: sube el Excel a Drive y ábrelo como Google Sheets.
- Semiautomático: mantén el Excel en Drive y usa una hoja intermedia de importación.

Recomendación práctica:
- Convierte el Excel principal a Google Sheets.
- Mantén una hoja limpia `embarques` con columnas fijas.
- Protege la hoja para que solo tu equipo edite.

---

## 5) Mejoras recomendadas

- Historial de eventos por contenedor (hoja `eventos`).
- Semáforo de estado con colores.
- Enlace a mapa del buque (si tienes IMO/MMSI).
- Envío automático de correo al cliente cuando cambie `estado_actual`.

Si quieres, en el siguiente paso te preparo una **versión avanzada** con:
- múltiples coincidencias,
- timeline del embarque,
- panel bonito para cliente.
