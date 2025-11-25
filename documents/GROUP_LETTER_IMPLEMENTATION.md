# Implementación de Letra de Grupo (Group Letter)

## 📋 Resumen

Este documento describe la implementación completa para almacenar y enviar al frontend la **letra del grupo** (A, B, C) de cada elemento en proyectos NW.

## 🎯 Objetivo

Permitir que el frontend identifique a qué tipo de grupo pertenece cada elemento para determinar el comportamiento de votación:

- **Grupo A**: Comportamiento actual (sin cambios)
- **Grupo B**: Permitirá votaciones (positivo, negativo, neutral)
- **Grupo C**: No permitirá votaciones

## ✅ Cambios Implementados en el Backend (Python)

### 1. Modelo `ProcessedExcelData` ([excel_models.py:23](app/models/excel_models.py#L23))

**Agregado:**
```python
lst_group_letters: List[str] = field(default_factory=list)  # Group letter (A, B, C) for each row
```

Este campo almacena la letra del grupo para cada fila del Excel procesado.

### 2. Servicio de Procesamiento Excel ([excel_service.py:147-277](app/services/excel_service.py#L147-L277))

**Lógica implementada:**
- Se extrae la letra del grupo de la **Columna A (Type)** usando el regex existente `_group_marker_regex`
- Se mantiene un `current_group_letter` que se actualiza cuando aparece un marcador de grupo
- Todas las filas subsiguientes heredan la letra del grupo actual hasta que aparezca un nuevo marcador
- Se popula `lst_group_letters` con la letra correspondiente para cada fila

**Ejemplo:**
```
Fila  | Type | Category          | Name      | Group Letter
------|------|-------------------|-----------|-------------
1     | A    | Prescreen Survivors| Akugiva  | A
2     |      | Prescreen Survivors| Atrekalyj | A
3     |      | Prescreen Survivors| Glaraonoki| A
4     | B    | Prescreen Failures | Adigira  | B
5     |      | Prescreen Failures | Adigira  | B
```

### 3. Modelo `DetailItem` ([presentation_models.py:364-367](app/models/presentation_models.py#L364-L367))

**Agregado:**
```python
group_letter: Optional[str] = Field(
    default="",
    description="Group letter (A, B, C) indicating the voting behavior for this slide's group."
)
```

### 4. Servicio de Presentación ([presentation_service.py](app/services/presentation_service.py))

**Funciones actualizadas:**
- `_build_slides_data_with_pptx` - Rastrea `last_group_letter`
- `_create_group_slide` - Incluye `group_letter` del Excel
- `_create_individual_slide` - Incluye `group_letter` del Excel
- `_create_summary_slide` - Incluye el `last_group_letter`

Todas las instancias de `DetailItem()` ahora incluyen el campo `group_letter`.

## 🔧 Cambios Requeridos en la Base de Datos

### Paso 1: Agregar columna a `nw_Details`

```sql
USE [BI_GUIDELINES]
GO

-- Agregar columna GroupLetter a nw_Details
ALTER TABLE [dbo].[nw_Details]
ADD [GroupLetter] NVARCHAR(10) NULL
GO

-- Opcional: Actualizar registros existentes con valor por defecto
UPDATE [dbo].[nw_Details]
SET [GroupLetter] = 'A'
WHERE [GroupLetter] IS NULL
GO
```

### Paso 2: Actualizar Stored Procedure `nw_InsertPresentationDetail_copy`

El stored procedure actual tiene 14 parámetros. Necesitamos agregar `@GroupLetter`:

```sql
USE [BI_GUIDELINES]
GO

-- Eliminar procedimiento existente
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[nw_InsertPresentationDetail_copy]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[nw_InsertPresentationDetail_copy]
GO

CREATE PROCEDURE [dbo].[nw_InsertPresentationDetail_copy]
    @PresentationId INT,
    @SlideNumber INT,
    @SlideType NVARCHAR(100),
    @SlideBGFileName NVARCHAR(500),
    @SlideDescription NVARCHAR(MAX),
    @NameGroup NVARCHAR(500),
    @NameCategory NVARCHAR(500),
    @Name NVARCHAR(MAX),
    @NameRationale NVARCHAR(MAX),
    @NameNotation NVARCHAR(500),
    @KanaNames NVARCHAR(MAX),
    @NameLogo NVARCHAR(500),
    @TemplateId INT,
    @NameSubGroup NVARCHAR(100),
    @GroupLetter NVARCHAR(10) = 'A'  -- NUEVO PARÁMETRO con valor por defecto
AS
BEGIN
    SET NOCOUNT ON;

    INSERT INTO [dbo].[nw_Details] (
        [PresentationId],
        [SlideNumber],
        [SlideType],
        [SlideBGFileName],
        [SlideDescription],
        [NameGroup],
        [NameCategory],
        [Name],
        [NameRationale],
        [NameNotation],
        [KanaNames],
        [NameLogo],
        [TemplateId],
        [NameSubGroup],
        [GroupLetter]  -- NUEVA COLUMNA
    )
    VALUES (
        @PresentationId,
        @SlideNumber,
        @SlideType,
        @SlideBGFileName,
        @SlideDescription,
        @NameGroup,
        @NameCategory,
        @Name,
        @NameRationale,
        @NameNotation,
        @KanaNames,
        @NameLogo,
        @TemplateId,
        @NameSubGroup,
        @GroupLetter  -- NUEVO VALOR
    );
END
GO
```

**Nota:** El parámetro `@GroupLetter` tiene un valor por defecto `'A'` para mantener compatibilidad con código que no envíe este parámetro.

## 🔄 Actualización del Código Python para Enviar `GroupLetter`

### Ubicación: [presentation_service.py:1673-1690](app/services/presentation_service.py#L1673-L1690)

**Cambio necesario:**

```python
# ANTES
detail_sql = """
    EXEC [dbo].[nw_InsertPresentationDetail_copy]
        @PresentationId=?, @SlideNumber=?, @SlideType=?, @SlideBGFileName=?,
        @SlideDescription=?, @NameGroup=?, @NameCategory=?, @Name=?,
        @NameRationale=?, @NameNotation=?, @KanaNames=?, @NameLogo=?,
        @TemplateId=?, @NameSubGroup=?;
"""

detail_params = (
    presentation_id,
    slide["slide_number"],
    slide_type,
    path_to_save,
    slide.get("slide_description") or "",
    slide.get("group_name") or "",
    slide.get("category") or "",
    slide.get("name") or "",
    slide.get("rationale") or "",
    slide.get("notation") or "",
    slide.get("kana") or "",
    slide.get("logo_filename") or "",
    template_id_to_save,
    slide.get("name_sub_group") or "",
)

# DESPUÉS
detail_sql = """
    EXEC [dbo].[nw_InsertPresentationDetail_copy]
        @PresentationId=?, @SlideNumber=?, @SlideType=?, @SlideBGFileName=?,
        @SlideDescription=?, @NameGroup=?, @NameCategory=?, @Name=?,
        @NameRationale=?, @NameNotation=?, @KanaNames=?, @NameLogo=?,
        @TemplateId=?, @NameSubGroup=?, @GroupLetter=?;
"""

detail_params = (
    presentation_id,
    slide["slide_number"],
    slide_type,
    path_to_save,
    slide.get("slide_description") or "",
    slide.get("group_name") or "",
    slide.get("category") or "",
    slide.get("name") or "",
    slide.get("rationale") or "",
    slide.get("notation") or "",
    slide.get("kana") or "",
    slide.get("logo_filename") or "",
    template_id_to_save,
    slide.get("name_sub_group") or "",
    slide.get("group_letter") or "A",  # NUEVO PARÁMETRO
)
```

## 📤 Respuesta del Frontend

### Endpoints Afectados

Cualquier endpoint que devuelva detalles de presentaciones ahora incluirá el campo `group_letter`:

#### Ejemplo de Respuesta:

```json
{
  "presentation_id": 12345,
  "details": [
    {
      "slide_number": 1,
      "slide_type": "NameGroup",
      "group_name": "Prescreen Survivors",
      "category": "Prescreen Survivors",
      "name": "Prescreen Survivors",
      "group_letter": "A"
    },
    {
      "slide_number": 2,
      "slide_type": "NameEvaluation",
      "group_name": "Prescreen Survivors",
      "category": "Prescreen Survivors",
      "name": "Akugiva",
      "rationale": "Atrasentan, Kidney",
      "group_letter": "A"
    },
    {
      "slide_number": 13,
      "slide_type": "NameGroup",
      "group_name": "Prescreen Failures",
      "category": "Prescreen Failures",
      "name": "Prescreen Failures",
      "group_letter": "B"
    },
    {
      "slide_number": 14,
      "slide_type": "NameEvaluation",
      "group_name": "Prescreen Failures",
      "category": "Prescreen Failures",
      "name": "Adigira",
      "rationale": "",
      "group_letter": "B"
    }
  ]
}
```

## 🎨 Uso en el Frontend

### Lógica de Votación

El frontend puede usar el campo `group_letter` para determinar qué elementos permiten votación:

```typescript
interface SlideDetail {
  slide_number: number;
  slide_type: string;
  group_name: string;
  name: string;
  group_letter: "A" | "B" | "C";
  // ... otros campos
}

function canVote(slide: SlideDetail): boolean {
  // Grupo A: Comportamiento actual (por ahora sin votación)
  if (slide.group_letter === "A") {
    return false;
  }

  // Grupo B: Permite votación
  if (slide.group_letter === "B") {
    return true;
  }

  // Grupo C: No permite votación
  if (slide.group_letter === "C") {
    return false;
  }

  // Por defecto, no permitir
  return false;
}

// Ejemplo de uso
function renderSlide(slide: SlideDetail) {
  return (
    <div className="slide">
      <h3>{slide.name}</h3>
      <p>Grupo: {slide.group_letter}</p>

      {canVote(slide) && (
        <div className="voting-buttons">
          <button onClick={() => vote(slide, 'positive')}>👍 Positivo</button>
          <button onClick={() => vote(slide, 'neutral')}>😐 Neutral</button>
          <button onClick={() => vote(slide, 'negative')}>👎 Negativo</button>
        </div>
      )}
    </div>
  );
}
```

## 📝 Orden de Implementación

1. ✅ **Backend Python** - COMPLETADO
   - Modelos actualizados
   - Lógica de extracción implementada
   - Propagación a través del flujo

2. ⏳ **Base de Datos** - PENDIENTE
   - Ejecutar script de migración para agregar columna `GroupLetter`
   - Actualizar stored procedure `nw_InsertPresentationDetail_copy`

3. ⏳ **Actualizar llamada al SP en Python** - PENDIENTE
   - Modificar `_insert_detail_records` para enviar el parámetro `@GroupLetter`

4. ⏳ **Frontend** - PENDIENTE
   - Actualizar modelos TypeScript para incluir `group_letter`
   - Implementar lógica de votación basada en `group_letter`

## 🧪 Testing

### Datos de Prueba

Crear un Excel con la siguiente estructura:

```
A  | Prescreen Survivors | Project Team Name Candidate Evaluation
   | Category           | Name              | Rationale
A  | Prescreen Survivors | Akugiva          | Atrasentan, Kidney
   | Prescreen Survivors | Atrekalyj        | Atrasentan, Kidney
   | Prescreen Survivors | Glaraonoki       | Glomerulus, Prevention, Atrasentan
B  | Prescreen Failures  | $Project Team Name Candidate Evaluation
   | Prescreen Failures  | Adigira          |
   | Prescreen Failures  | Akidira          |
C  | No Voting Group     | $Project Team Name Candidate Evaluation
   | No Voting Group     | TestName1        | Test Rationale
```

### Verificación

1. Subir el Excel a través del endpoint de creación de presentación
2. Verificar que `lst_group_letters` contiene: `["A", "A", "A", "B", "B", "B", "C", "C"]`
3. Verificar en la BD que la columna `GroupLetter` tiene los valores correctos
4. Consultar los detalles de la presentación y verificar que el JSON incluye `group_letter`

## 📚 Referencias

- [excel_models.py](app/models/excel_models.py) - Modelo de datos procesados del Excel
- [excel_service.py](app/services/excel_service.py) - Servicio de procesamiento del Excel
- [presentation_models.py](app/models/presentation_models.py) - Modelos de presentación
- [presentation_service.py](app/services/presentation_service.py) - Servicio de creación de presentaciones

## ❓ Preguntas Frecuentes

### ¿Qué pasa con proyectos antiguos que no tienen GroupLetter?

Los proyectos existentes en la BD tendrán `GroupLetter = NULL` o `'A'` (si se ejecuta el UPDATE). El frontend puede tratar `NULL` o `'A'` como el comportamiento por defecto.

### ¿Puedo cambiar el GroupLetter de un proyecto existente?

Actualmente no hay endpoint para modificar el `GroupLetter` de slides existentes. Se debería crear un endpoint PATCH si es necesario.

### ¿Qué pasa si el Excel no tiene marcadores de grupo?

Si el Excel no tiene letras en la columna A (Type), `current_group_letter` permanecerá como cadena vacía `""`, y todos los slides tendrán `group_letter = ""`. El frontend debe manejar esto como comportamiento por defecto (igual que `'A'`).

## ✅ Estado Actual

- ✅ Cambios en modelos Python completados
- ✅ Lógica de extracción del Excel implementada
- ✅ Propagación a través de todos los slides
- ⏳ Scripts SQL preparados (pendiente ejecución)
- ⏳ Actualización de llamada al SP (pendiente aplicar cambio)
- ⏳ Cambios en frontend (pendiente)
