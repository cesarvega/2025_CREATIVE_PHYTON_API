# Diagrama de Flujo - Integración de Categorías BSR

## Flujo Visual Completo

```
┌─────────────────────────────────────────────────────────────────────┐
│                    POST /presentations/create-bsr                    │
│                                                                       │
│  Payload:                                                             │
│  ┌─────────────────────────────────────────────────────────────┐   │
│  │ metadata: {                                                  │   │
│  │   project_name: "BSR_Project"                                │   │
│  │   display_name: "Q1_Dashboard"                               │   │
│  │   slide_number: 3                                            │   │
│  │   ...                                                         │   │
│  │   categories: {          ← NUEVO (opcional)                 │   │
│  │     add_categories: true                                     │   │
│  │     mode: "both"                                             │   │
│  │     category1: { name: "Region", elements: [...] }          │   │
│  │     category2: { name: "Channel", elements: [...] }         │   │
│  │   }                                                           │   │
│  │ }                                                             │   │
│  │ pptx_file: <archivo>                                         │   │
│  └─────────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────┐
│                   ROUTE: create_bsr_presentation()                   │
│                  (presentation_routes.py:558-638)                    │
└─────────────────────────────────────────────────────────────────────┘
                                    │
                  ┌─────────────────┴─────────────────┐
                  ▼                                     ▼
        ┌──────────────────┐              ┌──────────────────────┐
        │ Validar archivo  │              │ Pydantic valida      │
        │ PPTX             │              │ metadata.categories  │
        │ (tipo, tamaño)   │              │ con @model_validator │
        └──────────────────┘              └──────────────────────┘
                  │                                     │
                  └─────────────────┬─────────────────┘
                                    ▼
┌─────────────────────────────────────────────────────────────────────┐
│              SERVICE: create_bsr_presentation()                      │
│            (presentation_service.py:49-171)                          │
│                                                                       │
│  Parámetros:                                                          │
│  - project_name, display_name, slide_number, ...                    │
│  - categories: Optional[Dict] ← NUEVO                               │
└─────────────────────────────────────────────────────────────────────┘
                                    │
                    ┌───────────────┴───────────────┐
                    ▼                               ▼
        ┌────────────────────────┐    ┌────────────────────────┐
        │ 1. Limpiar             │    │ 2. Validar proyecto    │
        │    display_name        │    │    no existe           │
        └────────────────────────┘    └────────────────────────┘
                    │                               │
                    └───────────────┬───────────────┘
                                    ▼
                    ┌───────────────────────────────┐
                    │ 3. Convertir PPTX a imágenes  │
                    │    (pptx_service)             │
                    │    → 001.jpg, 002.jpg, ...    │
                    └───────────────────────────────┘
                                    │
                                    ▼
                    ┌───────────────────────────────┐
                    │ 4. Extraer títulos de slides  │
                    │    (python-pptx library)      │
                    └───────────────────────────────┘
                                    │
                                    ▼
        ┌───────────────────────────────────────────────────────┐
        │ 5. Insertar bsr_Master                                 │
        │    SP: bsr_InsertPresentationMaster                   │
        │    ↓                                                   │
        │    Retorna: presentation_id = 12345                   │
        └───────────────────────────────────────────────────────┘
                                    │
                                    ▼
        ┌───────────────────────────────────────────────────────┐
        │ 6. Insertar bsr_Details                                │
        │    SP: bsr_InsertPresentationDetail (loop)            │
        │    - Slides antes del summary                          │
        │    - Slide summary (tipo: NameSummary)               │
        │    - Slides después del summary                        │
        │    ↓                                                   │
        │    Retorna: total_slides = 6                          │
        └───────────────────────────────────────────────────────┘
                                    │
                                    ▼
╔═══════════════════════════════════════════════════════════════════╗
║              7. PROCESO DE CATEGORÍAS (NUEVO)                      ║
║                                                                     ║
║  if categories AND categories.add_categories == true:              ║
╚═══════════════════════════════════════════════════════════════════╝
                                    │
                    ┌───────────────┴───────────────┐
                    ▼                               ▼
        ┌─────────────────────────┐   ┌─────────────────────────┐
        │ _check_existing_        │   │ Si ya existen:          │
        │ categories()            │   │ → ValueError            │
        │                         │   │ → Log error             │
        │ SELECT COUNT(*)         │   │ → categories_created=[] │
        │ FROM BSR_CATEGORY       │   └─────────────────────────┘
        │ WHERE BSRPROJECTID=?    │
        └─────────────────────────┘
                    │
                    ▼ (si count=0)
        ┌──────────────────────────────────────────────────────┐
        │         PROCESAR CATEGORY 1 (siempre)                 │
        │                                                        │
        │  1. _check_category_name_exists()                     │
        │     → Si existe: ValueError                           │
        │                                                        │
        │  2. _insert_category_with_elements()                  │
        │     ┌──────────────────────────────────────┐         │
        │     │ INSERT INTO BSR_CATEGORY             │         │
        │     │ (BSRPROJECTID, CATEGORY)             │         │
        │     │ VALUES (12345, 'Region')             │         │
        │     └──────────────────────────────────────┘         │
        │                    ↓                                   │
        │     ┌──────────────────────────────────────┐         │
        │     │ SELECT id FROM BSR_CATEGORY          │         │
        │     │ WHERE BSRPROJECTID=? AND CATEGORY=?  │         │
        │     │ → category_id = 501                  │         │
        │     └──────────────────────────────────────┘         │
        │                    ↓                                   │
        │     ┌──────────────────────────────────────┐         │
        │     │ INSERT INTO BSR_CATEGORY_ELEMENTS    │         │
        │     │ (CATEGORY_ID, CATEGORY_MEMBERS)      │         │
        │     │ VALUES (501, 'North America')        │         │
        │     │ VALUES (501, 'Europe')               │         │
        │     │ VALUES (501, 'Asia')                 │         │
        │     │ ...                                   │         │
        │     └──────────────────────────────────────┘         │
        │                    ↓                                   │
        │     conn.commit()                                     │
        │                                                        │
        │  Resultado: { category_id: 501, name: "Region",      │
        │              elements_count: 4 }                      │
        └──────────────────────────────────────────────────────┘
                                    │
                                    ▼
                    ┌───────────────────────────────┐
                    │ if mode == "both":            │
                    │   PROCESAR CATEGORY 2         │
                    │   (mismo proceso que cat1)    │
                    └───────────────────────────────┘
                                    │
                                    ▼
                ┌───────────────────────────────────────┐
                │ Retorna: categories_created = [       │
                │   {category_id: 501, name: "Region",  │
                │    elements_count: 4},                │
                │   {category_id: 502, name: "Channel", │
                │    elements_count: 3}                 │
                │ ]                                      │
                └───────────────────────────────────────┘
                                    │
╔═══════════════════════════════════════════════════════════════════╗
║                   FIN PROCESO DE CATEGORÍAS                        ║
╚═══════════════════════════════════════════════════════════════════╝
                                    │
                                    ▼
        ┌───────────────────────────────────────────────────────┐
        │ 8. Retornar al ROUTE                                   │
        │                                                         │
        │ return {                                                │
        │   "presentation_id": 12345,                            │
        │   "total_slides": 6,                                   │
        │   "categories_added": 2  ← NUEVO                      │
        │ }                                                       │
        └───────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────┐
│                  ROUTE: crear respuesta HTTP                         │
└─────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
        ┌───────────────────────────────────────────────────────┐
        │ BSRCreatePresentationResponse(                         │
        │   message="BSR Presentation created successfully",     │
        │   presentation_id=12345,                               │
        │   total_slides=6,                                      │
        │   processing_time_seconds=9.87,                        │
        │   categories_added=2  ← NUEVO                         │
        │ )                                                       │
        └───────────────────────────────────────────────────────┘
                                    │
                                    ▼
                           ┌─────────────────┐
                           │   HTTP 200 OK    │
                           └─────────────────┘
```

---

## Estado de Base de Datos Después del Proceso

### Tabla: `bsr_Master`
```
PresentationId | Project        | DisplayName     | SlideNumber | ...
12345          | BSR_Project    | Q1_Dashboard    | 3           | ...
```

### Tabla: `bsr_Details`
```
PresentationId | SlideNumber | SlideType    | SlideBGFileName              | SlideDescription
12345          | 1           | Image        | BRS_slides/Q1_Dashboard/001.jpg | Slide 1
12345          | 2           | Image        | BRS_slides/Q1_Dashboard/002.jpg | Slide 2
12345          | 3           | NameSummary  | BRS_slides/Q1_Dashboard/003.jpg | Summary
12345          | 4           | Image        | BRS_slides/Q1_Dashboard/004.jpg | Slide 3
12345          | 5           | Image        | BRS_slides/Q1_Dashboard/005.jpg | Slide 4
12345          | 6           | Image        | BRS_slides/Q1_Dashboard/006.jpg | Slide 5
```

### Tabla: `BSR_CATEGORY` (NUEVA)
```
id  | BSRPROJECTID | CATEGORY
501 | 12345        | Region
502 | 12345        | Channel
```

### Tabla: `BSR_CATEGORY_ELEMENTS` (NUEVA)
```
id | CATEGORY_ID | CATEGORY_MEMBERS
1  | 501         | North America
2  | 501         | Europe
3  | 501         | Asia
4  | 501         | Latin America
5  | 502         | Retail
6  | 502         | Online
7  | 502         | Corporate
```

---

## Casos de Uso

### Caso 1: Sin Categorías (Tradicional)
```
categories: null
  ↓
Step 7: SKIP (categories_created = [])
  ↓
Response: categories_added = 0
```

### Caso 2: Mode Single (1 Categoría)
```
categories: { add_categories: true, mode: "single", category1: {...} }
  ↓
Step 7: Procesar SOLO category1
  ↓
Response: categories_added = 1
```

### Caso 3: Mode Both (2 Categorías)
```
categories: { add_categories: true, mode: "both", category1: {...}, category2: {...} }
  ↓
Step 7: Procesar category1 Y category2
  ↓
Response: categories_added = 2
```

### Caso 4: Error en Categorías (No Bloqueante)
```
categories: { add_categories: true, ... }
  ↓
Step 7: Error al insertar (ej: nombre duplicado)
  ↓
Log error + categories_created = []
  ↓
Response: categories_added = 0
  ↓
Proyecto BSR SE CREÓ exitosamente ✓
```

---

## Ventanas de Transacción

```
┌────────────────────────────────────────────────────┐
│ Transacción 1: Proyecto BSR                        │
│                                                     │
│ BEGIN TRANSACTION                                  │
│   INSERT bsr_Master                                │
│   INSERT bsr_Details (múltiples)                   │
│ COMMIT                                             │
│                                                     │
│ Resultado: presentation_id = 12345                │
└────────────────────────────────────────────────────┘
                      │
                      ▼
┌────────────────────────────────────────────────────┐
│ Transacción 2: Categoría 1                         │
│                                                     │
│ BEGIN TRANSACTION                                  │
│   INSERT BSR_CATEGORY                              │
│   SELECT id → category_id                          │
│   INSERT BSR_CATEGORY_ELEMENTS (múltiples)         │
│ COMMIT                                             │
│                                                     │
│ Resultado: {category_id: 501, elements_count: 4}  │
└────────────────────────────────────────────────────┘
                      │
                      ▼
┌────────────────────────────────────────────────────┐
│ Transacción 3: Categoría 2 (si mode=both)          │
│                                                     │
│ BEGIN TRANSACTION                                  │
│   INSERT BSR_CATEGORY                              │
│   SELECT id → category_id                          │
│   INSERT BSR_CATEGORY_ELEMENTS (múltiples)         │
│ COMMIT                                             │
│                                                     │
│ Resultado: {category_id: 502, elements_count: 3}  │
└────────────────────────────────────────────────────┘

✓ Si Transacción 2 o 3 fallan → Proyecto BSR sigue existiendo
✓ Categorías son OPCIONALES e INDEPENDIENTES del proyecto
```

---

**Nota:** Este diagrama representa el flujo implementado en el branch AI-019.
