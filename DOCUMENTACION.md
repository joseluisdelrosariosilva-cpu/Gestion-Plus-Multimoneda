# GESTIÓN PLUS 2.0 - MINI

Sistema de gestión empresarial en Excel VBA con soporte **doble moneda (MN/USD)**.
Rework de una versión anterior monomoneda.

---

## ÍNDICE

1. [Propósito](#1-propósito)
2. [Arquitectura General](#2-arquitectura-general)
3. [Hojas del Libro](#3-hojas-del-libro)
4. [Módulos de Código](#4-módulos-de-código)
5. [Formularios (UserForms)](#5-formularios-userforms)
6. [Sistema de Doble Moneda](#6-sistema-de-doble-moneda)
7. [Flujo de Trabajo Principal](#7-flujo-de-trabajo-principal)
8. [Seguridad y Licencia](#8-seguridad-y-licencia)
9. [Respaldos](#9-respaldos)
10. [Funcionalidades Implementadas vs Pendientes](#10-funcionalidades-implementadas-vs-pendientes)

---

## 1. PROPÓSITO

Sistema POS + inventario + gastos para pequeños negocios en Cuba, donde conviven dos monedas (MN y USD). Permite:

- Gestionar inventario con productos en MN, USD o mixtos (costo en una moneda, venta en otra)
- Registrar ventas (POS) con pago combinado (efectivo MN, transferencia MN, USD)
- Facturación electrónica con ID auto-generado
- Control de gastos variables y fijos (salarios)
- Elaboración de productos por recetas con descuento de ingredientes
- Dashboard con indicadores
- Respaldo automático de datos

---

## 2. ARQUITECTURA GENERAL

### Stack

| Componente | Tecnología |
|------------|------------|
| Plataforma | Microsoft Excel (VBA) |
| Archivo | `.xlsm` (Libro habilitado para macros) |
| Base de datos | Tablas de Excel (`ListObject`) dentro del mismo libro |
| UI | UserForms VBA con botones overlay (reposo/active) |

### Convenciones de código

| Aspecto | Convención |
|---------|------------|
| Monedas | `MN` = Moneda Nacional, `USD` = Dólar |
| Tipo de cambio | Celda `G3` en hoja `Gestión de Almacén` |
| IDs de producto | Formato `Pr_00001` (generado automáticamente) |
| IDs de factura | Prefijo = fecha serial (5 dígitos) + sufijo consecutivo (5 dígitos) |
| Contraseñas | Hardcodeadas (ver sección Seguridad) |
| Navegación | Botones con overlay: `_reposo` (fondo) + `_covered` (frente visible en hover) |

### Estructura de archivos

```
Gestion Plus 2.0 - Mini/
├── Archivos/
│   ├── Forms/          → 14 UserForms (.frm + .frx)
│   ├── Hojas/          → 13 módulos de hoja (.cls)
│   └── Modulos/        → 18 módulos estándar (.bas)
├── Respaldos/          → backups automáticos
├── Gestion Plus 2.0 - Mini.xlsm
└── Gestion Plus 2.0 - Mini Prueba.xlsm
```

---

## 3. HOJAS DEL LIBRO

| Hoja | Nombre Técnico | Tabla(s) | Propósito |
|------|---------------|----------|-----------|
| Gestión de Almacén | Hoja4 | `INVENTARIO` | Inventario principal (16 columnas) |
| Gestión de Ingredientes | Hoja5 | `INGREDIENTES` | Materias primas e insumos |
| Elaboración de Productos | Hoja8 | `Recetas` | Recetas de producción |
| Ingredientes | Hoja9 | `Ingredientes_Receta` | Ingredientes por receta |
| Historial de ventas | Hoja6 | `HISTORIALVENTAS` | Registro de ventas (9 columnas) |
| Historial de Gastos | Hoja7 | `Gastos` | Gastos fijos y variables (5 columnas) |
| Registro de facturas | Hoja? | `HISTORIALVENTAS4` | Facturas con desglose de pago |
| Dashboard | Hoja3 | — | Dashboard con tablas dinámicas |
| Gestión de Trabajadores | Hoja? | `trabajadores` | Empleados, cargos, salarios |
| Fechas | Hoja? | `FechasUnicas`, `Productos_Unicos` | Catálogos auxiliares |
| Hidden | Hoja? | — | Fecha de expiración y licencia |
| Blanco | Hoja? | — | Hoja de respaldo visible al cerrar |

### Estructura de la tabla INVENTARIO (columna → índice)

| Col | Índice | Campo | Descripción |
|-----|--------|-------|-------------|
| A | 1 | Código | ID único `Pr_XXXXX` |
| B | 2 | Producto | Nombre del producto |
| C | 3 | Fecha Entrada | Fecha de ingreso |
| D | 4 | Cantidad Inicial | Cantidad registrada originalmente |
| E | 5 | Stock | Stock actual (fórmula `=actualizar_almacen(...)`) |
| F | 6 | Precio Venta | Precio unitario de venta |
| G | 7 | Moneda Venta | `MN` o `USD` |
| H | 8 | Precio Costo | Costo unitario |
| I | 9 | Moneda Costo | `MN` o `USD` |
| J | 10 | Fondo | Costo total = PrecioCosto × Stock |
| K | 11 | Moneda Fondo | Moneda del costo |
| L | 12 | Ingreso Esperado | Ingreso total = PrecioVenta × Stock |
| M | 13 | Moneda Ingreso | Moneda de venta |
| N | 14 | Ganancia | Ingreso - Fondo (con conversión si aplica) |
| O | 15 | Moneda Ganancia | Moneda de la ganancia |
| P | 16 | — | Reservado |

---

## 4. MÓDULOS DE CÓDIGO

### Módulos Estándar (`Archivos/Modulos/`)

| Módulo | Tipo | Funcionalidad |
|--------|------|---------------|
| `calling` | Público | Punto de entrada principal. Llama formularios desde botones de hoja |
| `Módulo1` | Público | CRUD trabajadores, pagar salarios, refresh dashboard, utilidades |
| `Módulo2` | Público | Respaldos (diario, cierre, limpieza de BD) |
| `Módulo3` | Público | Validación de repeticiones, protección libro/hojas, CRUD ingredientes |
| `Generador_codigo` | Público | Genera IDs de producto (`Pr_XXXXX`) |
| `ID_Factura` | Público | Genera IDs de factura basados en fecha + consecutivo |
| `Dolarizacion` | Público | Cambia tipo de cambio y recalcula ganancias del inventario |
| `Conversor` | Público | Conversión de unidades (kg↔g↔lb↔L↔mL, etc.) |
| `Calculo_precio_UBase` | Público | Calcula precio por unidad base en ingredientes |
| `Caluculo_de_total` | Público | Calcula total a pagar en MN y USD en facturación |
| `descuento_almacen` | Público | Función `actualizar_almacen()` para stock en tiempo real |
| `busqueda_factura` | Público | Buscador simple de facturas por código |
| `Eliminar_producto_seleccionado` | Público | Elimina productos del inventario o recetas |
| `ModuloCalendario` | Público | Control de calendario UI (fechas) |
| `Validar_campos_vacios` | Público | Valida campos vacíos en `ingresar_form` |
| `Validar_vacios_2` | Público | Valida campos vacíos en `ingresar_ingr_nuevo_form` |
| `Validar_vacios_3` | Público | Valida campos vacíos en `ingresar_ingr_existente_form` |
| `MermasAbastecimiento` | Público | Gestión de mermas y reabastecimiento de inventario. Incluye `ObtenerMermasProducto()` (UDF), `RegistrarMerma()` (manual) y `ReabastecerProducto()` (manual) |

### Módulos de Hoja (`Archivos/Hojas/`)

| Hoja | Funcionalidad clave |
|------|---------------------|
| `ThisWorkbook` | `Workbook_Open`: login, verificación dispositivo, vigencia, protección |
| | `Workbook_BeforeClose`: respaldo final, ocultar hojas, proteger |
| | Verificación de integridad de fecha, respaldo diario |

---

## 5. FORMULARIOS (USERFORMS)

| Formulario | Propósito |
|------------|-----------|
| `facturacion` | **POS principal**. Búsqueda de productos, carrito, totales, procesar pago |
| `ingresar_form` | Alta/modificación de productos en inventario |
| `Elaborar` | Producción por receta (selecciona receta, lotes, descuenta ingredientes) |
| `Gasto_Form` | Registro de gastos variables (con moneda) |
| `Trabajador` | Alta de trabajadores |
| `Trabajador_Mod` | Modificación de trabajadores |
| `Login` | Acceso administrador (revela hojas ocultas) |
| `Busqueda` | Búsqueda de gastos por rango de fechas + descripción |
| `Busqueda_Ventas` | Búsqueda de ventas por rango de fechas |
| `Busqueda_Facturacion` | Búsqueda de facturas por rango de fechas |
| `Ingr_Receta` | Alta/modificación de recetas de producción |
| `ingresar_ingr_nuevo_form` | Alta de nuevos ingredientes |
| `ingresar_ingr_existente_form` | Registrar compras de ingredientes existentes |
| `frmCalendario` | Selector de fecha universal |

### Formulario `facturacion` — Detalle

Flujo:

```
1. UserForm_Initialize → carga productos en productos_busqueda_ListBox
2. UserForm_Activate → fecha de hoy, oculta admin si no está en modo admin
3. Usuario busca producto (nombre_producto_busqueda_TextBox_Change)
4. Double-click en producto → carga datos a campos de facturación
5. boton_add_covered_Click → valida y agrega al carrito (productos_facturacion_ListBox)
6. calcular_total → total en MN + USD con conversión
7. boton_procesar_pago_covered_Click → bloquea edición, muestra campos de pago
8. Checkboxes: efectivo MN, transferencia MN, USD
9. ActualizarSaldos → calcula restante con conversión en tiempo real
10. boton_registrar_ventas_covered_Click → registra en Historial de ventas + Facturas + descuenta stock
```

**Campos de pago:**

| Campo | Moneda | Tipo |
|-------|--------|------|
| `ingreso_cliente_pago_TextBox` | MN | Efectivo |
| `cantidad_a_pagar_transeferencia_TextBox` | MN | Transferencia |
| `ingreso_cliente_pago_USD_TextBox` | USD | Billetes USD |

---

## 6. SISTEMA DE DOBLE MONEDA

### 6.1. Concepto

Cada producto puede tener **costo en una moneda** y **venta en otra**, generando **4 combinaciones posibles**:

| # | Costo | Venta | Ganancia en | Cálculo de Ganancia |
|---|-------|-------|-------------|---------------------|
| 1 | MN | MN | MN | `Ingreso - Fondo` |
| 2 | MN | USD | USD | `Ingreso - (Fondo / TC)` |
| 3 | USD | MN | MN | `Ingreso - (Fondo × TC)` |
| 4 | USD | USD | USD | `Ingreso - Fondo` |

Donde **TC** = Tipo de Cambio (celda `G3` de Gestión de Almacén).

### 6.2. Componentes del sistema dual

#### a) Moneda en producto (INVENTARIO)
- Columnas G (Moneda Venta) e I (Moneda Costo) → combo box con `MN`/`USD`
- Todas las combinaciones generan Ingreso, Ganancia, Fondo con su moneda correspondiente

#### b) Moneda en gastos
- Columna 5 de la tabla Gastos → `MN` o `USD`

#### c) Pago en facturación
- Hasta 3 formas de pago en una misma venta:
  - Efectivo MN
  - Transferencia MN
  - USD (billetes)
- El POS calcula el restante en ambas monedas usando el TC

#### d) Dashboard
- Indicadores basados en dual-moneda (totales convertidos)

### 6.3. Funciones clave para dual-currency

| Función/Sub | Ubicación | Qué hace |
|-------------|-----------|----------|
| `Cambiar_Precio_Dolar` | Dolarizacion.bas | Actualiza TC en G3 y recalcula ganancias |
| `calcular_total` | Caluculo_de_total.bas | Suma carrito en MN y USD con conversión |
| `ActualizarSaldos` | facturacion.frm | Calcula saldo restante en tiempo real |
| `setear_funcion` | calling.bas | Recalcula todo el inventario (stock = inicial − vendido − mermas, fondo, ingreso, ganancia) |
| `Restar_Cantidad` | facturacion.frm | Descuenta stock y recalcula con moneda |
| `Moneda_Costo_Change` / `Moneda_Venta_Change` | ingresar_form.frm | Recálculo en vivo al cambiar moneda |

---

## 7. FLUJO DE TRABAJO PRINCIPAL

```
APERTURA DEL LIBRO
│
├─ Workbook_Open
│   ├─ Verificar dispositivo autorizado
│   ├─ Verificar integridad de fecha
│   ├─ Verificar vigencia (licencia)
│   ├─ Proteger hojas
│   └─ Mostrar facturacion (POS)
│
├─ LOGIN ADMIN (desde POS)
│   ├─ Revela hojas ocultas (Almacén, Ingredientes, etc.)
│   └─ Activa menú completo
│
├─ GESTIÓN DE ALMACÉN
│   ├─ ingresar_form → Alta/Modificación de productos
│   ├─ Genera_codigo → ID automático Pr_XXXXX
│   └─ setear_funcion → Recalcula fondo/ingreso/ganancia
│
├─ ELABORACIÓN DE PRODUCTOS
│   ├─ Ingr_Receta → Alta de recetas con ingredientes
│   ├─ Elaborar → Produce lotes, descuenta ingredientes
│   └─ Calculo_precio_UBase → Costo por unidad base
│
├─ VENTAS (POS) → facturacion
│   ├─ Buscar producto → Agregar al carrito
│   ├─ Procesar pago → Elegir métodos
│   └─ Registrar → Historial + Factura + Descuento stock
│
├─ GASTOS
│   ├─ Gasto_Form → Gastos variables (con moneda)
│   ├─ Módulo1 → Gestión trabajadores + pago salarios
│   └─ Busqueda → Consulta por fechas
│
├─ CONSULTAS
│   ├─ Busqueda → Gastos
│   ├─ Busqueda_Ventas → Ventas
│   └─ Busqueda_Facturacion → Facturas
│
└─ CIERRE
    ├─ crearRespaldoFinal → Backup antes de cerrar
    ├─ Ocultar hojas de datos
    └─ Proteger libro
```

---

## 8. SEGURIDAD Y LICENCIA

### Contraseñas (hardcodeadas)

| Propósito | Contraseña | Variable |
|-----------|-----------|----------|
| Protección de hojas | `theSheetProtectpassisthe011rd` | — |
| Protección del libro | `theBookProtectpassisthe010nd` | — |

⚠️ **ATENCIÓN**: Las contraseñas están hardcodeadas en el código fuente. Cualquier usuario con acceso al VBA puede verlas.

### Licencia

- **Dispositivos autorizados**: 3 equipos (DESKTOP-FFH10GP, DESKTOP-27I02D0, DESKTOP-6Q2K16M)
- **Identificación**: COMPUTERNAME + USERNAME + SerialNumber del disco C:
- **Vigencia**: Fecha de expiración en hoja `Hidden` celda A1
- **Códigos**: Hash simple (algoritmo propio) para renovación mensual/anual
- **Integridad**: Detecta retroceso en la fecha del sistema (celda A2 de Hidden)

### Visibilidad de hojas

| Estado | Hojas visibles |
|--------|---------------|
| Usuario normal (POS) | Solo `facturacion` |
| Admin (Login) | Almacén, Ingredientes, Elaboración, Ventas, Gastos, Facturas |
| Al cerrar | Solo `Blanco` (visible) |

---

## 9. RESPALDOS

| Tipo | Cuándo | Nombre | Destino |
|------|--------|--------|---------|
| Diario | `verificarIntegridadFecha` (si pasó 1 día) | `Respaldo_diario_DD-MMM-YYYY.xlsm` | `./Respaldos/` |
| Cierre | `Workbook_BeforeClose` | `Ultimo_respaldo_DD-MMM-YYYY_HH-MM.xlsm` | `./Respaldos/` |
| Manual (viejo) | Botón en UI | `Respaldo_FECHA_MIN_FECHA_MAX.xlsx` | `D:\Work\Respaldos\` |
| Manual (nuevo) | Botón en UI | `Respaldo_FECHA_MIN_FECHA_MAX.xlsm` | `./Respaldos/` |

El respaldo manual nuevo:
1. Guarda copia del libro
2. Elimina botones de la copia
3. Extiende la licencia +10000 días
4. Limpia tablas de ventas, gastos, facturas y fechas
5. Resetea stock inicial = stock actual

---

## 10. FUNCIONALIDADES IMPLEMENTADAS VS PENDIENTES

### ✅ Implementado

- [x] **Almacén**: CRUD de productos con doble moneda (costo/venta independientes)
- [x] **POS (Ventas)**: Carrito, búsqueda en vivo, pago combinado (efectivo MN + transferencia MN + USD)
- [x] **Facturación**: ID auto-generado, registro de facturas con desglose de pago
- [x] **Gastos**: Registro de gastos variables y fijos (con moneda)
- [x] **Trabajadores**: CRUD, pago de salarios (se registra como gasto)
- [x] **Elaboración**: Recetas con ingredientes, producción por lotes, descuento de stock
- [x] **Dashboard**: Tablas dinámicas (requiere refresh manual)
- [x] **Búsquedas**: Por rango de fechas en gastos, ventas y facturas
- [x] **Respaldos**: Diario, al cerrar y manual (con limpieza de BD)
- [x] **Seguridad**: Protección de hojas/libro, licencia por dispositivo, vigencia
- [x] **Conversión de unidades**: kg/g/lb/oz/L/mL entre ingredientes
- [x] **Tipo de cambio**: Global (G3), recálculo masivo de ganancias
- [x] **Mermas**: Registro manual desde inventario como gastos categoría "Merma". Se descuentan automáticamente del stock vía `setear_funcion`
- [x] **Reabastecimiento**: Suma manual de cantidad a un producto desde botón en inventario. Actualiza Cantidad Inicial y Fecha de Entrada

### 🔲 Pendiente / En desarrollo

- [ ] **Módulo de Reportes**: Reportes avanzados (PDF, exportación)
- [ ] **Módulo de Proveedores**: Gestión de proveedores
- [ ] **Módulo de Clientes**: Base de datos de clientes recurrente
- [ ] **Dashboard mejorado**: Refresco automático, gráficos adicionales
- [ ] **Notificaciones**: Stock bajo, vencimientos, etc.
- [ ] **Multi-sucursal**: Soporte para múltiples puntos de venta
- [ ] **Historial de precios**: Tracking de cambios de precio en el tiempo
- [ ] **Cálculo de costo promedio**: vs. costo unitario actual
- [ ] **Exportación a PDF**: Facturas imprimibles
- [ ] **Configuración de impuestos**: IVA, etc.
- [ ] **Separación de responsabilidades**: Refactorizar módulos monolíticos
- [ ] **Manejo de errores**: Sistema robusto vs. `On Error Resume Next`

---

## NOTAS TÉCNICAS IMPORTANTES

### Validaciones de campos
Los módulos de validación (`Validar_campos_vacios`, `Validar_vacios_2`, `Validar_vacios_3`) tienen el código de resaltado visual COMENTADO (BackColor). Solo cuentan campos vacíos.

### Dependencia de nombres de hoja
Muchos módulos referencian hojas por su **nombre en español** (ej: `"Gestión de Almacén"`) en lugar de usar `CodeName`. Si se renombra una hoja, el código se rompe.

### Stock como valor estático
La celda de Stock (columna E) **no usa fórmula**. Se escribe como valor estático desde `setear_funcion`. Esto evita lentitud con miles de registros pero requiere recalcular manualmente.

### Stock = Inicial − Vendido − Mermas
`setear_funcion` calcula: `stock = CantidadInicial − TotalVendido(código) − ObtenerMermasProducto(código)`. Las mermas registradas como gasto "Merma" se descuentan automáticamente.

### Recálculo manual
`setear_funcion` debe ejecutarse después de cualquier cambio en inventario (alta, modificación, merma, reabastecimiento) para que los valores de stock, fondo, ingreso y ganancia se actualicen.

---

*Documentación generada el 27 de mayo de 2026*
*Proyecto: Gestion Plus 2.0 - Mini*
