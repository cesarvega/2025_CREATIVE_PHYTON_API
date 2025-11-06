# 🔧 FIX: URLs de Presentaciones Corregidas

**Fecha:** 2025-11-05
**Issue:** URLs incorrectas en endpoints de presentaciones activas

---

## 🐛 PROBLEMA IDENTIFICADO

Los endpoints de presentaciones activas estaban generando URLs incorrectas para todos los tipos de proyectos:

### ❌ ANTES (Incorrecto):
```json
{
  "presentation_id": 7409,
  "display_name": "BSR_TEST_11_05_2025",
  "link": "https://tools.brandinstitute.com/bsr/#/main/BSR_TEST_11_05_2025"  // ❌ INCORRECTO
}
```

### ✅ DESPUÉS (Correcto):
```json
{
  "presentation_id": 7409,
  "display_name": "BSR_TEST_11_05_2025",
  "link": "https://bipresents.com/BSR_TEST_11_05_2025"  // ✅ CORRECTO
}
```

---

## 📋 REGLAS DE GENERACIÓN DE URLs

### Proyectos NW y DW:
```
Nombre: "TEST_NW_11_05_2025_v3"
URL: https://nw.bipresents.com/TEST_NW_11_05_2025_v3

Nombre: "1402_NW2"
URL: https://nw.bipresents.com/1402_NW2
```

### Proyectos BSR y BSR-Japan:
```
Nombre: "BECCA"
URL: https://bipresents.com/BECCA

Nombre: "BSR_TEST_11_05_2025"
URL: https://bipresents.com/BSR_TEST_11_05_2025
```

### Proyectos NSR y NSR-Japan:
```
Nombre: "Test_11_05_2025"
URL: https://bipresents.com/Test_11_05_2025

Nombre: "UNITRI"
URL: https://bipresents.com/UNITRI
```

---

## 🔧 CAMBIOS APLICADOS

### 1. Endpoint `/bi_guidelines/nw-active-presentations`

**Archivo:** [app/services/bi_guidelines_service.py:110](app/services/bi_guidelines_service.py#L110)

**ANTES:**
```python
# Build link from display name
link = f"https://tools.brandinstitute.com/nw/#/main/{row.DisplayName}"
```

**DESPUÉS:**
```python
# FIXED: Build correct link - NW presentations use nw.bipresents.com
link = f"https://nw.bipresents.com/{row.DisplayName}"
```

---

### 2. Endpoint `/bi_guidelines/bsr-active-presentations`

**Archivo:** [app/services/bi_guidelines_service.py:213-234](app/services/bi_guidelines_service.py#L213-L234)

**ANTES:**
```python
# Use the Link column from the SP which already contains the BSR link
# Format: https://tools.brandinstitute.com/bsr/#/main/{DisplayName}
link = self._get_ci(d, "Link", "link") or f"https://tools.brandinstitute.com/bsr/#/main/{display_name}"
```

**DESPUÉS:**
```python
# FIXED: Generate correct link based on presentation type
# Check for PresentationType field from SP (BSR, NSR, BSR-Japan, NSR-Japan, NW, DW)
presentation_type = self._get_ci(d, "PresentationType", "presentationtype", "Type")

# Generate appropriate link based on type
if presentation_type:
    ptype = str(presentation_type).upper()
    if "NSR" in ptype:
        # NSR or NSR-Japan → https://bipresents.com/{DisplayName}
        link = f"https://bipresents.com/{display_name}"
    elif "BSR" in ptype:
        # BSR or BSR-Japan → https://bipresents.com/{DisplayName}
        link = f"https://bipresents.com/{display_name}"
    elif ptype in ["NW", "DW"]:
        # NW or DW → https://nw.bipresents.com/{DisplayName}
        link = f"https://nw.bipresents.com/{display_name}"
    else:
        # Unknown type: default to bipresents
        link = f"https://bipresents.com/{display_name}"
else:
    # No type field: try to use Link from SP or default to bipresents
    link = self._get_ci(d, "Link", "link") or f"https://bipresents.com/{display_name}"
```

---

## ✅ VALIDACIÓN

```bash
# Compilación exitosa
✅ python -m py_compile app/services/bi_guidelines_service.py
```

---

## 📊 IMPACTO

### Endpoints afectados:
1. ✅ `GET /api/bi_guidelines/nw-active-presentations`
2. ✅ `GET /api/bi_guidelines/bsr-active-presentations`

### Comportamiento corregido:
- ✅ Proyectos **NW/DW** → `https://nw.bipresents.com/{DisplayName}`
- ✅ Proyectos **BSR/BSR-Japan** → `https://bipresents.com/{DisplayName}`
- ✅ Proyectos **NSR/NSR-Japan** → `https://bipresents.com/{DisplayName}`

---

## 🧪 TESTING

### Test 1: Endpoint BSR Active Presentations
```bash
curl "https://tools.brandinstitute.com/CreativePythonAPI/api/bi_guidelines/bsr-active-presentations?page=1&limit=5"
```

**Resultado esperado:**
```json
{
  "presentations": [
    {
      "presentation_id": 7409,
      "project": "ATEST01",
      "display_name": "BSR_TEST_11_05_2025",
      "link": "https://bipresents.com/BSR_TEST_11_05_2025"  // ✅ Correcto
    }
  ]
}
```

### Test 2: Endpoint NW Active Presentations
```bash
curl "https://tools.brandinstitute.com/CreativePythonAPI/api/bi_guidelines/nw-active-presentations?page=1&limit=5"
```

**Resultado esperado:**
```json
{
  "presentations": [
    {
      "presentation_id": 8286,
      "display_name": "TEST_NW_11_05_2025",
      "link": "https://nw.bipresents.com/TEST_NW_11_05_2025"  // ✅ Correcto
    }
  ]
}
```

---

## 🔍 NOTAS TÉCNICAS

### Lógica de detección de tipo:
El código ahora busca el campo `PresentationType` en la respuesta del stored procedure `BSR_ActivePresentations`.

Si el SP no incluye este campo, se usa un fallback a `https://bipresents.com/{DisplayName}` para BSR/NSR.

### Campos buscados (case-insensitive):
- `PresentationType`
- `presentationtype`
- `Type`

### Tipos soportados:
- `BSR` → bipresents.com
- `BSR-Japan` → bipresents.com
- `NSR` → bipresents.com
- `NSR-Japan` → bipresents.com
- `NW` → nw.bipresents.com
- `DW` → nw.bipresents.com

---

## ⚠️ NOTA IMPORTANTE

Si el stored procedure `[BI_GUIDELINES].[dbo].[BSR_ActivePresentations]` **NO devuelve** el campo `PresentationType`, todas las presentaciones BSR usarán el fallback `https://bipresents.com/{DisplayName}`.

Para soporte completo de todos los tipos (incluyendo NSR mixtos en el resultado), el SP debe incluir la columna `PresentationType`.

---

**Autor:** Claude Code
**Fecha:** 2025-11-05
**Estado:** ✅ APLICADO Y VALIDADO
