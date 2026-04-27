# Análisis de Datos de Ventas de Supermercado (SuperStore US 2015)

## Objetivo
Proyecto universitario de Ingeniería de Software / Inteligencia Artificial para el análisis de datos de ventas de un supermercado, utilizando técnicas de limpieza, análisis y visualización de datos con Python.

## Dataset
Archivo: `SuperStoreUS-2015(Orders).csv`
- 1952 registros, 25 columnas
- Datos de pedidos de supermercado en Estados Unidos (2015)
- Formato: Separador de columnas `;`, decimal `,`
- Contiene información de clientes, productos, fechas, ventas, ganancias y descuentos

## Pasos Realizados
1. **Carga de Datos**: Lectura del CSV en un DataFrame de pandas, verificación de filas y columnas.
2. **Exploración de Datos**: Muestra de primeras filas, información general (tipos de datos, valores no nulos), estadísticas descriptivas y conteo de valores nulos.
3. **Limpieza de Datos**:
   - Conversión de columnas de fecha (`Order Date`, `Ship Date`) a formato datetime
   - Relleno de 16 valores nulos en `Product Base Margin` con la mediana (0.53)
   - Aseguramiento de tipos de datos numéricos correctos
4. **Análisis de Datos (4 conjuntos requeridos)**:
   - Promedio de ventas por categoría de producto
   - Top 5 ciudades con mayor ganancia (adaptado de países, ya que el dataset solo contiene datos de EE.UU.)
   - Top 5 clientes con mayores compras totales
   - Correlación entre descuentos y ganancias
5. **Análisis Adicional**:
   - Ventas totales por mes
   - Identificación de productos con pérdidas (`Profit < 0`)
6. **Visualizaciones**: Generación de 5 gráficas en formato PNG usando matplotlib y seaborn:
   - Barras: Promedio de ventas por categoría
   - Barras: Top 5 ciudades con mayor ganancia
   - Línea: Ventas por mes
   - Dispersión: Relación descuento vs ganancia
   - Barras: Top 10 productos con mayores pérdidas
7. **Reporte Word**: Generación automática de `Analisis_Supermercado.docx` con título, introducción, explicación de análisis, resultados, tablas, gráficas y conclusiones generales.

## Dependencias
Instalar las siguientes librerías antes de ejecutar:
```bash
pip install pandas matplotlib seaborn python-docx
```

## Cómo Ejecutar
1. Guardar el script `analisis_supermercado.py` en la misma carpeta que el archivo CSV
2. Ejecutar el script:
```bash
python analisis_supermercado.py
```

## Archivos Generados
- `analisis_supermercado.py`: Script principal de análisis
- `promedio_ventas_categoria.png`: Gráfica de barras promedio de ventas por categoría
- `top_ciudades.png`: Gráfica de barras top 5 ciudades con mayor ganancia
- `ventas_por_mes.png`: Gráfica de línea ventas por mes
- `descuento_vs_ganancia.png`: Gráfica de dispersión descuento vs ganancia
- `productos_perdida.png`: Gráfica de barras top 10 productos con pérdidas
- `Analisis_Supermercado.docx`: Reporte en Word con todos los resultados

## Conclusiones Clave
- La categoría con mayor promedio de ventas es `Furniture` (1651.76)
- `Washington` es la ciudad con mayor ganancia total (11677.36)
- `Kristine Connolly` es la cliente con mayores compras totales (50475.31)
- Existe una correlación negativa débil entre descuentos y ganancias (-0.0618)
- Abril es el mes con mayores ventas, marzo el más bajo
- Los productos con mayores pérdidas suman -108563.76 en el top 10
