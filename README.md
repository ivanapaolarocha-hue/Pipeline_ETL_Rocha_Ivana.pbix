# Pipeline ETL en Power BI — Análisis de Ventas

**Alumna:** Ivana Paola Rocha  
**Herramienta:** Power BI Desktop / Power Query (Lenguaje M)  
**Archivo del proyecto:** `Pipeline_ETL_Rocha_Ivana.pbix`

---

## 1. ¿De qué se trata este proyecto?
En este trabajo armé un proceso de limpieza y preparación de datos (**ETL**) usando **Power Query** en Power BI Desktop, partiendo del archivo `Pipeline_ETL_Dataset.xlsx`.

La idea fue dejar las cuatro tablas limpias, sin datos duplicados ni errores de formato, y conectadas bajo un **modelo estrella** para poder armar reportes y gráficos sin problemas.

---

## 2. Tablas del Modelo Estrella

| Consulta | Tipo de Tabla | Qué contiene |
| :--- | :--- | :--- |
| **`Fact_Ventas`** | Tabla de Hechos | Tiene el detalle de cada venta (`id_venta`, `fecha_venta`, `cantidad`, `precio_unitario`, `descuento`, `total_venta`, `canal`) junto con el nombre y la categoría del producto. |
| **`Dim_Clientes`** | Dimensión | Guarda los datos únicos de cada cliente (`id_cliente`, `nombre_cliente`, `email`, `ciudad`, `pais`, `segmento`, `fecha_registro`). |
| **`Dim_Productos`** | Dimensión | Es el catálogo de productos (`id_producto`, `nombre_producto`, `id_categoria`, `precio_lista`). |
| **`Dim_Categorias`** | Dimensión | Lista las categorías comerciales (`id_categoria`, `categoria`). |

---

## 3. ¿Qué fui limpiando y cómo lo resolví?

* **Orden en los nombres:** Renombré todas las consultas usando `Dim_` para las dimensiones y `Fact_` para la tabla de hechos, así es más fácil ubicar cada tabla en el modelo.
* **Filas vacías y duplicados:**
  * En `Fact_Ventas` saqué las filas que no tenían `id_venta` y borré los duplicados antes de cruzar tablas, para que no se sumen ventas repetidas en los totales.
  * En `Dim_Clientes` filtré los registros sin `id_cliente` y dejé un único registro por persona.
* **Valores nulos o en blanco (`"Sin Datos"`):** En `Dim_Clientes` encontré un `email` nulo (fila 9) y una `ciudad` vacía (fila 11). En vez de borrar esos clientes —que ya tenían compras hechas—, reemplacé esos espacios por `"Sin Datos"` para no perder sus ventas en la tabla principal.
* **Unión de tablas (`Merge`):** Combiné `Fact_Ventas` con `Dim_Productos` usando `id_producto` y un *Left Outer Join* para no perder ninguna venta histórica, y traje las columnas `nombre_producto` y `categoria` para poder filtrar más rápido en los gráficos.
* **Formato de fechas y moneda:** Convertí los números que venían de Excel en fechas reales (`type date`) para `fecha_venta

  ## 4. Documentación en Código M (Editor Avanzado)

Renombré pasos clave y agregué comentarios con `//` explicando el motivo de cada cambio en las dos consultas principales:

### Consulta: `Fact_Ventas`
```powerquery
let
    Origen = Excel.Workbook(File.Contents("D:\Documents\Downloads\Pipeline_ETL_Dataset.xlsx"), null, true),
    #"Navegación 1" = Origen{[Item = "ventas", Kind = "Sheet"]}[Data],
    #"Encabezados promovidos" = Table.PromoteHeaders(#"Navegación 1", [PromoteAllScalars = true]),
    #"Tipo de columna cambiado" = Table.TransformColumnTypes(#"Encabezados promovidos", {{"id_venta", Int64.Type}, {"fecha_venta", Int64.Type}, {"id_cliente", Int64.Type}, {"id_producto", Int64.Type}, {"cantidad", Int64.Type}, {"precio_unitario", Int64.Type}, {"descuento", type number}, {"total_venta", type number}, {"canal", type text}}, "es"),
    // Saco las filas que no tienen id_venta para que no queden registros incompletos o sin identificar en los cálculos
    #"Filas filtradas" = Table.SelectRows(#"Tipo de columna cambiado", each [id_venta] <> null and [id_venta] <> ""),
    // Se deduplica por id_venta antes de combinar para impedir que se dupliquen los montos facturados
    Quitar_Duplicados_Venta = Table.Distinct(#"Filas filtradas", {"id_venta"}),
    // Uno la tabla con Dim_Productos usando id_producto y Left Outer Join para no perder ninguna venta registrada
    #"Consultas combinadas" = Table.NestedJoin(Quitar_Duplicados_Venta, {"id_producto"}, Dim_Productos, {"id_producto"}, "Dim_Productos", JoinKind.LeftOuter),
    // Traigo nombre_producto y categoria directamente a Fact_Ventas para poder armar los filtros y gráficos del reporte más rápido
    Expandir_Producto_Categoria = Table.ExpandTableColumn(#"Consultas combinadas", "Dim_Productos", {"nombre_producto", "categoria"}, {"nombre_producto", "categoria"}),
    // Cambio fecha_venta a formato fecha y precio_unitario a moneda para poder analizar las ventas por mes o año y sumar bien los montos
    #"Tipo de columna cambiado 1" = Table.TransformColumnTypes(Expandir_Producto_Categoria, {{"fecha_venta", type date}, {"precio_unitario", Currency.Type}})
in
    #"Tipo de columna cambiado 1"
