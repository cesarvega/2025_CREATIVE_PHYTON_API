# Implementación de Letra de Grupo - Solución Alternativa (Sin cambios en BD)

## 📋 Resumen

Solución pragmática que **NO requiere cambios en la base de datos**. La letra del grupo se almacena dentro del campo existente `NameGroup` usando el formato: `{LETRA}|{NOMBRE_GRUPO}`.

## 🎯 Formato del Campo `NameGroup`

### Estructura:
```
NameGroup = "A|Prescreen Survivors"
NameGroup = "B|Prescreen Failures"
NameGroup = "C|Another Group"
```

**Delimitador:** `|` (pipe/barra vertical)

**Formato:** `{GROUP_LETTER}|{GROUP_NAME}`

### Ventajas de esta Solución

✅ **Sin cambios en BD** - No requiere ejecutar scripts SQL
✅ **Retrocompatible** - Registros antiguos sin `|` siguen funcionando
✅ **Fácil de parsear** - En frontend: `split('|')` o regex
✅ **No invasivo** - El campo `group_letter` sigue disponible en la respuesta API

## ✅ Cambios Implementados

### 1. Backend Python - Modelos (Completado)

**Archivos modificados:**
- ✅ [app/models/excel_models.py](app/models/excel_models.py) - Campo `lst_group_letters` agregado
- ✅ [app/services/excel_service.py](app/services/excel_service.py) - **Extracción de letra de Grupo1/Grupo2**
- ✅ [app/models/presentation_models.py](app/models/presentation_models.py) - Campo `group_letter` en `DetailItem`
- ✅ [app/services/presentation_service.py:1672-1702](app/services/presentation_service.py#L1672-L1702) - **Combinación de letra y nombre**

### 2. Lógica de Extracción de Letra

**Origen de la letra:** Columnas **Grupo1 (G)** o **Grupo2 (H)** del Excel

**Método:**
1. Se lee el valor de Grupo1 (ej: "a1", "a2", "b1", "b2")
2. Se toma el **primer carácter** del valor
3. Se convierte a mayúscula (ej: "a" → "A", "b" → "B")
4. Se almacena en `lst_group_letters`

**Ejemplos:**
- Grupo1 = "a1" → `group_letter = "A"`
- Grupo1 = "a2" → `group_letter = "A"`
- Grupo1 = "b1" → `group_letter = "B"`
- Grupo1 = "c5" → `group_letter = "C"`

### 3. Lógica de Almacenamiento ([presentation_service.py:1672-1702](app/services/presentation_service.py#L1672-L1702))

```python
# Combinar group_letter y group_name en formato: "A|Prescreen Survivors"
group_name_raw = slide.get("group_name") or ""
group_letter_raw = slide.get("group_letter") or ""

if group_letter_raw and group_name_raw:
    # Formato: "A|Prescreen Survivors"
    group_name_with_letter = f"{group_letter_raw}|{group_name_raw}"
elif group_name_raw:
    # No letter, just the name (retrocompatibilidad)
    group_name_with_letter = group_name_raw
else:
    group_name_with_letter = ""

detail_params = (
    ...,
    group_name_with_letter,  # Se guarda en NameGroup
    ...,
)
```

## 📤 Respuesta de la API

### Estructura JSON (Ejemplo)

La respuesta incluye **AMBOS** campos:
1. `group_name` - Con formato `"A|Prescreen Survivors"` (almacenado en BD)
2. `group_letter` - Con la letra extraída `"A"` (generado en Python)

```json
{
  "presentation_id": 12345,
  "details": [
    {
      "slide_number": 1,
      "slide_type": "NameGroup",
      "group_name": "A|Prescreen Survivors",  ← FORMATO COMBINADO
      "category": "Prescreen Survivors",
      "name": "Prescreen Survivors",
      "group_letter": "A"  ← LETRA EXTRAÍDA
    },
    {
      "slide_number": 2,
      "slide_type": "NameEvaluation",
      "group_name": "A|Prescreen Survivors",  ← FORMATO COMBINADO
      "category": "Prescreen Survivors",
      "name": "Akugiva",
      "rationale": "Atrasentan, Kidney",
      "group_letter": "A"  ← LETRA EXTRAÍDA
    },
    {
      "slide_number": 13,
      "slide_type": "NameGroup",
      "group_name": "B|Prescreen Failures",  ← FORMATO COMBINADO
      "category": "Prescreen Failures",
      "name": "Prescreen Failures",
      "group_letter": "B"  ← LETRA EXTRAÍDA
    }
  ]
}
```

## 🎨 Uso en el Frontend

### Opción 1: Usar el campo `group_letter` (Recomendado)

El campo `group_letter` ya viene parseado desde el backend:

```typescript
interface SlideDetail {
  slide_number: number;
  slide_type: string;
  group_name: string;  // "A|Prescreen Survivors"
  name: string;
  group_letter: string;  // "A" ← USAR ESTE
}

function canVote(slide: SlideDetail): boolean {
  const letter = slide.group_letter;

  if (letter === "A") return false;  // Comportamiento actual
  if (letter === "B") return true;   // Permite votación
  if (letter === "C") return false;  // No permite votación

  return false; // Por defecto
}
```

### Opción 2: Parsear `group_name` manualmente

Si prefieres extraer la letra del campo `group_name`:

```typescript
function extractGroupLetter(groupName: string): string {
  if (!groupName) return "";

  // Verificar si tiene el formato "A|Nombre"
  if (groupName.includes("|")) {
    const parts = groupName.split("|");
    return parts[0].trim();  // "A"
  }

  // Retrocompatibilidad: sin letra
  return "";
}

function getGroupName(groupName: string): string {
  if (!groupName) return "";

  // Verificar si tiene el formato "A|Nombre"
  if (groupName.includes("|")) {
    const parts = groupName.split("|");
    return parts[1].trim();  // "Prescreen Survivors"
  }

  // Retrocompatibilidad: devolver tal cual
  return groupName;
}

// Uso
const letter = extractGroupLetter(slide.group_name);  // "A"
const cleanName = getGroupName(slide.group_name);     // "Prescreen Survivors"
```

### Opción 3: Usar Regex (Más robusto)

```typescript
const GROUP_NAME_REGEX = /^([A-Z])\|(.+)$/;

function parseGroupName(groupName: string): { letter: string; name: string } {
  if (!groupName) return { letter: "", name: "" };

  const match = groupName.match(GROUP_NAME_REGEX);

  if (match) {
    return {
      letter: match[1],  // "A"
      name: match[2]     // "Prescreen Survivors"
    };
  }

  // Retrocompatibilidad: sin letra
  return { letter: "", name: groupName };
}

// Uso
const { letter, name } = parseGroupName(slide.group_name);
```

## 🔄 Comparación de Soluciones

| Aspecto | Solución Alternativa (Actual) | Solución con Columna BD |
|---------|------------------------------|-------------------------|
| Cambios en BD | ❌ No requiere | ✅ Requiere ALTER TABLE + SP |
| Complejidad | 🟢 Baja | 🟡 Media |
| Retrocompatibilidad | ✅ Total | ⚠️ Parcial |
| Parseo en Frontend | 🟡 Manual (1 línea) | 🟢 Automático |
| Mantenibilidad | 🟢 Simple | 🟢 Simple |
| Tiempo de implementación | ✅ Inmediato | ⏱️ Requiere deployment BD |

## 🧪 Casos de Prueba

### Caso 1: Proyecto con grupos A, B, C

**Excel:**
```
A | Prescreen Survivors | ...
  | Prescreen Survivors | Akugiva
B | Prescreen Failures  | ...
  | Prescreen Failures  | Adigira
C | No Voting Group     | ...
```

**Resultado en BD (`NameGroup`):**
```
"A|Prescreen Survivors"
"A|Prescreen Survivors"
"B|Prescreen Failures"
"B|Prescreen Failures"
"C|No Voting Group"
```

### Caso 2: Proyecto sin letra de grupo (retrocompatibilidad)

**Excel:**
```
  | Category | Name
  | General  | Name1
  | General  | Name2
```

**Resultado en BD (`NameGroup`):**
```
"General"
"General"
```

**En frontend:**
```typescript
slide.group_letter  // ""
slide.group_name    // "General"
```

### Caso 3: Proyecto antiguo (antes de esta implementación)

**BD existente (`NameGroup`):**
```
"Prescreen Survivors"
```

**Frontend debe manejar:**
```typescript
// No tiene "|", entonces:
extractGroupLetter("Prescreen Survivors")  // ""
getGroupName("Prescreen Survivors")        // "Prescreen Survivors"
```

## 📝 Recomendaciones para el Frontend

### 1. Usar el campo `group_letter` directamente

**Recomendado:** No parsear manualmente, usar el campo que ya viene del backend.

```typescript
// ✅ RECOMENDADO
const canVote = slide.group_letter === "B";

// ❌ NO NECESARIO (pero funciona)
const letter = slide.group_name.split("|")[0];
const canVote = letter === "B";
```

### 2. Mostrar el nombre limpio del grupo

Si quieres mostrar solo el nombre sin la letra:

```typescript
function getCleanGroupName(groupName: string): string {
  return groupName.includes("|")
    ? groupName.split("|")[1].trim()
    : groupName;
}

// Uso en componente
<h3>{getCleanGroupName(slide.group_name)}</h3>
```

### 3. Helper functions recomendadas

```typescript
// utils/groupHelpers.ts

export interface ParsedGroupName {
  letter: string;
  name: string;
}

export function parseGroupName(groupName: string): ParsedGroupName {
  if (!groupName) return { letter: "", name: "" };

  const parts = groupName.split("|");

  if (parts.length === 2) {
    return {
      letter: parts[0].trim(),
      name: parts[1].trim()
    };
  }

  return { letter: "", name: groupName };
}

export function canVoteByLetter(letter: string): boolean {
  return letter === "B";
}

export function getVotingBehavior(letter: string): "voting" | "no-voting" | "default" {
  if (letter === "B") return "voting";
  if (letter === "C") return "no-voting";
  return "default";  // A o sin letra
}
```

## 🔍 Verificación

### 1. Verificar en la BD

```sql
SELECT TOP 10
    SlideNumber,
    SlideType,
    NameGroup,  -- Debe tener formato "A|Nombre" o solo "Nombre"
    NameCategory,
    Name
FROM [BI_GUIDELINES].[dbo].[nw_Details]
WHERE PresentationId = 12345
ORDER BY SlideNumber
```

### 2. Verificar en la API

```bash
curl http://localhost:8000/api/presentations/12345/details
```

Verificar que:
- ✅ `group_name` tiene formato `"A|Nombre"` o solo `"Nombre"`
- ✅ `group_letter` contiene `"A"`, `"B"`, `"C"` o `""`

## ❓ Preguntas Frecuentes

### ¿Qué pasa si el nombre del grupo contiene el carácter `|`?

Es muy improbable, pero si ocurre, el parseo tomará todo después del primer `|` como nombre:
```
"A|Grupo|Con|Pipes" → letter="A", name="Grupo|Con|Pipes"
```

### ¿Puedo cambiar el delimitador?

Sí, pero debes actualizar en dos lugares:
1. [presentation_service.py:1679](app/services/presentation_service.py#L1679) - Cambiar el `|`
2. Frontend - Actualizar el `split()`

### ¿Funciona con proyectos existentes?

Sí. Proyectos antiguos:
- En BD tendrán `NameGroup = "Nombre"` (sin letra)
- El frontend debe manejar esto como comportamiento por defecto
- El campo `group_letter` vendrá vacío `""`

## ✅ Estado de Implementación

- ✅ Extracción de letra del Excel
- ✅ Almacenamiento de letra en `ProcessedExcelData`
- ✅ Propagación a `DetailItem`
- ✅ Combinación de letra y nombre en formato `"A|Nombre"`
- ✅ Almacenamiento en campo `NameGroup` existente
- ✅ Campo `group_letter` en respuesta API
- ⏳ Parseo en frontend (pendiente)
- ⏳ Lógica de votación en frontend (pendiente)

## 🎯 Próximos Pasos

1. **Backend:** ✅ Completado (no requiere más cambios)

2. **Frontend:**
   - Actualizar modelos TypeScript
   - Implementar helpers de parseo (opcional, ya viene el campo `group_letter`)
   - Implementar lógica de votación basada en `group_letter`
   - Actualizar UI para mostrar/ocultar botones de votación

3. **Testing:**
   - Crear presentación con grupos A, B, C
   - Verificar que `NameGroup` tiene formato correcto
   - Verificar que `group_letter` llega al frontend
   - Probar comportamiento de votación

## 📚 Archivos Relacionados

- [excel_models.py](app/models/excel_models.py) - Modelo con `lst_group_letters`
- [excel_service.py](app/services/excel_service.py) - Extracción de letra
- [presentation_models.py](app/models/presentation_models.py) - `DetailItem` con `group_letter`
- [presentation_service.py](app/services/presentation_service.py) - Combinación de letra y nombre

---

**Ventaja Principal:** Esta solución funciona **inmediatamente** sin necesidad de cambios en la base de datos ni despliegues complejos. 🚀
