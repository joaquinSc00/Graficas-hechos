# Graficas-hechos
Esto es lo que entiendo —con lujo de detalle— que querés lograr:

### Meta del proyecto

Automatizar el **armado preliminar** de cada edición del periódico en **InDesign** para:

* **Volcar** en la maqueta, **dentro de zonas válidas**, todas las **notas** (título + cuerpo + foto).
* **Probar** distribuciones inteligentes hasta que “entre” lo máximo posible **sin overset**, respetando tu diseño.
* **Ahorrarte** la primera pasada de medición/acomodo; vos luego hacés el retoque fino.

### Estructura del repositorio

```
├── src/
│   ├── Prediseño_automatizado.py
│   └── solver.py
└── scripts/
    └── indesign/
        └── auto_layout_cierre_nov.jsx
```

* `src/Prediseño_automatizado.py` contiene el andamiaje principal del pipeline en Python. Las funciones siguen siendo stubs
  documentados, pero concentran toda la lógica de preprocesamiento pensada para ejecutarse desde la línea de comandos.
  Ejecutalo como módulo (`python -m src.Prediseño_automatizado`) para garantizar que los imports relativos funcionen.
* `src/solver.py` aloja el beam-search responsable de evaluar combinaciones de notas dentro de cada página.
* `scripts/indesign/auto_layout_cierre_nov.jsx` es el ExtendScript que se ejecuta dentro de InDesign para volcar el resultado.

Todos los scripts prototipo y archivos de prueba aislados se eliminaron para dejar solo los componentes necesarios del flujo actual.

### Estructura de entrada

* Carpeta raíz del cierre con **subcarpetas por página** (ej.: `Pagina 2 Matorrales`, `Pagina 5 y 6 Vdr`, etc.).
* Dentro de cada subcarpeta:

  * **1 o más .docx** con varias notas.

    * Cada **nota** se identifica por un **título en negrita** (primer párrafo en bold) y su **cuerpo** debajo.
  * **Imágenes** opcionales para cada nota con convención tipo `nota1_foto1.*`, `nota1_foto2.*`, etc.
* Pueden existir notas “de más” en una carpeta (material de reserva). **El script no rellena sobrantes** en otras páginas; solo mide/maqueta lo correspondiente a esa página.

### Documento InDesign

* Usás siempre el **mismo diseño** (5 **columnas**).
* El archivo **de edición** lo abrís vos antes de correr el script.
* En las páginas donde **sí** puede ir contenido dejaste **rectángulos** (huecos) dibujados por vos; si **no hay** rectángulo es porque esa página lleva un aviso u otros elementos fijos y **no se toca**.
* Esos rectángulos definen el **área disponible** para ubicar **título + cuerpo + foto** (“slots”).
  (Etiquetados con “root” o con estilo `SLOT_*`, pero lo importante es que el script los detecte.)

### Reglas tipográficas y de imagen

* **Cuerpo**: tamaño base fijo con **tolerancia ±0,5 pt** (p. ej. 9,5 ± 0,5).
* **Títulos**: tamaño base con **tolerancia ±1 pt** (p. ej. 25 → 24–26).
* **Foto 2 columnas**: **ancho fijo 10,145 cm**, **alto mínimo 5,35 cm**; ancho **no** se toca (para no “dejar ¼ de columna”). Alto puede variar **un poco** si hace falta.
* Preferencias de párrafo razonables (no viudas/huérfanas, keeps básicos).

### Qué debe hacer el script (ráfaga)

1. **Detectar slots** en cada página (los rectángulos que dibujaste).
2. **Leer** cada subcarpeta de esa página, abrir los **.docx** en un marco temporal, **segmentar** las notas por **título en negrita**.
3. **Asociar** imágenes por nombre a cada nota (si existen).
4. **Iterar layouts dentro de cada slot**:

   * Distribuciones: posición de **foto** (arriba/abajo), **span** del **título** (1–5 col dentro del ancho del slot), variaciones de **tamaño** de título/cuerpo dentro de sus límites.
   * El **slot** es la “caja madre”: el script **reserva** internamente el espacio de foto (10,145 × ≥5,35) y reparte el resto entre **título** y **cuerpo**.
   * **Objetivo**: cero **overset**. Si no hay solución, elegir la combinación con **menor excedente** y marcarla (para que vos ajustes).
   * Para páginas con dos números (ej. **5–6 VdR**) se pueden tratar como **spread** y considerar los slots de ambas.
5. **Volcar de verdad** en el documento (no copias de prueba): crear marcos finales (título/cuerpo/foto) **dentro del slot**, aplicar estilos y tamaños elegidos.
6. **Respetar** páginas sin slot (aviso, membrete): **no las modifica**.
7. **Reportar**: generar `reporte.csv` con métricas (página, chars por nota, tamaños usados, overflow, advertencias) y, opcionalmente, PDFs de muestra por página.

### Qué **no** querés

* Que el script “adivine” espacios donde **vos no dibujaste** un slot.
* Que cambie el **ancho** de la foto 2 col (siempre 10,145 cm).
* Que redistribuya notas “de reserva” a otras páginas.
* Que rompa la estética/estilos definidos.

### Tecnología y precisión

* Script en **ExtendScript (JSX)**, corriendo **dentro de InDesign**.
* Coordenadas siempre con **RulerOrigin.PAGE_ORIGIN** y `zeroPoint=[0,0]`; medidas en **pt** (conversión fiable desde cm).
* Detección de slots por: **label “root”** y/o **estilo de objeto** `SLOT_*` (y podemos incluir tag XML “root” como respaldo).
* **Solver** con intentos acotados (más intentos cuanto más notas) hasta encontrar una combinación válida.

### Herramientas auxiliares

* `scripts/indesign/auto_layout_cierre_nov.jsx`: script de InDesign que recibe el plan generado en Python y crea los marcos
  finales dentro de cada slot.

El resto de utilidades que se usaban como prototipos quedaron fuera del repositorio para simplificar la puesta en marcha y
evitar archivos de datos obsoletos.

---

Si esto refleja exactamente tu idea, paso el encargo a Codex con estas reglas (y mantengo que ya **no necesitamos** `layout_slots_report.json`: leemos **directo** del documento). Si querés, añado “bonus” para que el solver intente **dividir un slot grande** en sub-bloques si detecta que la foto no entra abajo/arriba en ninguna combinación.
