# Complemento Excel — Aplicar Color con Atajos de Teclado

Complemento de Excel (`.xlam`) desarrollado en VBA que resuelve una limitación de Excel: **cuando tienes un filtro activo, el atajo `F4` deja de funcionar para repetir el último color aplicado**, obligando al usuario a aplicar el color manualmente celda por celda.

Este complemento permite guardar un color de referencia y reutilizarlo rápidamente con atajos de teclado personalizados, incluyendo un sistema de **deshacer múltiple**.

---

## 🚀 ¿Qué problema resuelve?

En Excel existe `F4` para repetir la última acción, incluyendo aplicar un color. Sin embargo, **cuando hay un filtro activo, `F4` deja de funcionar** para esta tarea.

Este complemento lo soluciona:

* El usuario elige su color normalmente desde la barra de Excel.
* Guarda ese color con un atajo (`Ctrl + Shift + K`).
* Puede aplicarlo después cuantas veces quiera con `Ctrl + M`, incluso con filtros activos.
* Puede deshacer cambios con `Ctrl + Shift + Z`.
* Conserva hasta **10 niveles de historial de deshacer**.
* Restaura correctamente colores previos y celdas sin relleno.

---

## 📦 Instalación del Complemento

### 1. Descargar el archivo

Descarga el archivo `ColorCeldas.xlam` de este repositorio.

### 2. Instalar en Excel

1. Abre Excel.
2. Ve a **Archivo → Opciones → Complementos**.
3. En la parte inferior selecciona **Complementos de Excel** y haz clic en **Ir...**.
4. Haz clic en **Examinar** y selecciona `ColorCeldas.xlam`.
5. Marca el complemento con ✅ y pulsa **Aceptar**.

> ✅ Los atajos se activan automáticamente al abrir Excel.

---

## ⌨️ Atajos disponibles

| Atajo              | Acción                                   |
| ------------------ | ---------------------------------------- |
| `Ctrl + Shift + K` | Guardar el color de la celda activa      |
| `Ctrl + M`         | Aplicar el color guardado a la selección |
| `Ctrl + Shift + Z` | Deshacer el último cambio realizado      |

---

## 🧩 Flujo de uso

| Paso | Acción                                              |
| ---- | --------------------------------------------------- |
| 1    | Aplica un color manualmente a una celda desde Excel |
| 2    | Selecciona esa celda                                |
| 3    | Presiona `Ctrl + Shift + K` para guardar ese color  |
| 4    | Selecciona las celdas destino                       |
| 5    | Presiona `Ctrl + M` para aplicar el color guardado  |
| 6    | Si necesitas revertir, usa `Ctrl + Shift + Z`       |

> Si cambias de color, simplemente vuelve a guardarlo con `Ctrl + Shift + K`.

---

## 🧠 ¿Cómo funciona internamente?

```text
GuardarColorElegido:
  → Lee el color de la celda activa
  → Guarda el valor RGB en memoria
  → Confirma al usuario el color guardado

AplicarColorCeldas (Ctrl + M):
  → Verifica que exista un color guardado
  → Guarda el estado actual de cada celda seleccionada
  → Registra color, ColorIndex y dirección de cada celda
  → Aplica el color guardado

RestaurarUnPaso (Ctrl + Shift + Z):
  → Recupera el último estado guardado
  → Restaura cada celda a su estado original
  → Elimina ese nivel del historial
```

El historial funciona como una **pila (stack)** de hasta **10 acciones**.

Cuando la pila llega al límite, la acción más antigua se descarta automáticamente.

---

## 🗂️ Estructura del Proyecto

```text
📁 complemento-color-celdas/
├── ColorCeldas.xlam        # Complemento listo para instalar
├── ModuloColor.bas         # Código VBA principal
├── ThisWorkbook.cls        # Atajos y eventos del libro
└── README.md               # Documentación
```

---

## ⚠️ Limitaciones conocidas

* El color guardado **no se conserva al cerrar Excel**.
* El historial de deshacer también se reinicia al cerrar Excel.
* Los atajos afectan todos los libros abiertos mientras el complemento esté cargado.
* El máximo de deshacer es **10 niveles** (editable desde `MAX_UNDO`).
* Este sistema usa un **deshacer personalizado**, no el `Ctrl + Z` nativo de Excel.

---

## 📋 Requisitos

* Microsoft Excel 2016 o superior.
* Macros habilitadas.
* Pestaña **Programador** habilitada (opcional, para editar código).

Para habilitar la pestaña Programador:

**Archivo → Opciones → Personalizar cinta de opciones → ✅ Programador**

---

📄 Licencia
BSD 3-Clause
Este proyecto está bajo la licencia BSD 3-Clause. Esto significa que puedes usar, modificar y distribuir este 
código libremente, pero no puedes usar el nombre del autor para promocionar productos derivados sin permiso previo por escrito.

---

*Desarrollado en VBA para Microsoft Excel*
