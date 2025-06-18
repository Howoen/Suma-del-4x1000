# 📊 Macro VBA - Suma del 4x1000

Esta macro en VBA está diseñada para **filtrar y sumar automáticamente el valor del impuesto "IMPTO GOBIERNO 4X1000"** a partir de extractos bancarios importados a Excel desde archivos PDF.

## 🚀 ¿Qué hace esta macro?

- Aplica un filtro en la columna B buscando `"IMPTO GOBIERNO 4X1000"`.
- Recorre las celdas visibles en la columna D (valores monetarios del impuesto).
- Limpia los valores (elimina puntos y cambia comas por puntos para que sean numéricos).
- Suma los valores detectados correctamente.
- Muestra el resultado en la celda `H2` y lo formatea como moneda.
- Muestra un `MsgBox` con el total calculado.

## 🧩 Requisitos

- El archivo Excel debe tener una hoja llamada `"datos"`.
- Las columnas deben estar organizadas de forma que:
  - **Columna B** contenga el concepto de cada movimiento.
  - **Columna D** contenga el valor del movimiento (puede tener separadores de miles y decimales).
- La cabecera debe estar en la fila 1.

## 📥 Cómo usarla

1. Abre tu archivo Excel con los datos.
2. Presiona `ALT + F11` para abrir el **Editor de Visual Basic (VBA)**.
3. Inserta un **módulo nuevo** (`Insertar > Módulo`).
4. Pega el contenido de la macro `Filtro_Suma_4X1000`.
5. Ejecuta la macro con `F5` o desde Excel (puedes asignarla a un botón).

## 📌 Resultado

- El total del 4x1000 se mostrará en la celda `H2` de la hoja `"datos"`.
- El mensaje emergente (MsgBox) también mostrará el resultado.

## 🛑 Importante

- Este repositorio **no incluye datos reales ni archivos Excel por confidencialidad**.
- Solo se publica el **código fuente de la macro** (`.bas`), sin datos bancarios.
