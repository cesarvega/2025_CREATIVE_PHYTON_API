# Thumbnails para Background Templates

## 📋 Descripción

Sistema de thumbnails automáticos para optimizar la carga de imágenes de background templates en el frontend.

## 🎯 Problema Resuelto

Las imágenes de background templates son de alta resolución (1920x1080px, ~500KB-2MB cada una). Al cargar listados con muchas plantillas, esto causaba:
- ⏱️ Tiempos de carga lentos
- 📡 Alto consumo de ancho de banda
- 🐌 Experiencia de usuario lenta

## ✨ Solución Implementada

### Generación Automática de Thumbnails

Cuando se crea o actualiza un background template:
1. Se guarda la imagen original en su calidad completa
2. Automáticamente se genera un **thumbnail de 300x169px** (16:9)
3. El thumbnail se guarda como JPEG optimizado (~20-30KB)

### Estructura de Archivos

```
C:/inetpub/wwwroot/nw2/assets/images/BackGrounds/
├── Default.jpg                    ← Imagen original (1920x1080, ~500KB)
├── BMW_blue.jpg                   ← Imagen original (1920x1080, ~800KB)
└── thumbnails/
    ├── Default.jpg                ← Thumbnail (300x169, ~25KB)
    └── BMW_blue.jpg               ← Thumbnail (300x169, ~30KB)
```

### Respuesta de la API

Ambos endpoints ahora devuelven **dos URLs** por cada template:

```json
{
  "template_group_id": 123,
  "template_name": "BMW_blue",
  "category": "BMW",
  "template_file_name": "images/BackGrounds/BMW_blue.jpg",
  "thumbnail_url": "https://tools.brandinstitute.com/nw2/assets/images/BackGrounds/thumbnails/BMW_blue.jpg",
  "preview_url": "https://tools.brandinstitute.com/nw2/assets/images/BackGrounds/BMW_blue.jpg"
}
```

## 📖 Uso en el Frontend

### Para Listados (Rápido)
Usa `thumbnail_url` para mostrar previsualizaciones en grids/listas:

```jsx
// ✅ RECOMENDADO: Usa thumbnail para listados
<img 
  src={template.thumbnail_url} 
  alt={template.template_name}
  className="template-thumbnail"
/>
```

**Ventajas:**
- ⚡ Carga 20-30x más rápido
- 📉 Reduce ancho de banda en 95%
- 🚀 Mejora experiencia de usuario

### Para Detalles/Vista Completa (Calidad)
Usa `preview_url` cuando necesites la imagen en calidad completa:

```jsx
// ✅ Usa preview_url para modales/detalles
<img 
  src={template.preview_url} 
  alt={template.template_name}
  className="template-full-preview"
/>
```

### Carga Progresiva (Best Practice)

Implementa carga progresiva para mejor UX:

```jsx
const [imageLoaded, setImageLoaded] = useState(false);

<div className="template-preview">
  {/* Muestra thumbnail primero */}
  <img 
    src={template.thumbnail_url}
    alt={template.template_name}
    className={imageLoaded ? 'blur' : ''}
  />
  
  {/* Carga imagen completa en background */}
  <img 
    src={template.preview_url}
    alt={template.template_name}
    onLoad={() => setImageLoaded(true)}
    className={imageLoaded ? 'visible' : 'hidden'}
  />
</div>
```

## 🔧 Endpoints Afectados

### 1. GET `/api/bi_guidelines/template-groups`

**Antes:**
```json
{
  "template_group_id": 123,
  "template_name": "BMW_blue",
  "preview_url": "https://tools.brandinstitute.com/nw2/assets/images/BackGrounds/BMW_blue.jpg"
}
```

**Ahora:**
```json
{
  "template_group_id": 123,
  "template_name": "BMW_blue",
  "thumbnail_url": "https://tools.brandinstitute.com/nw2/assets/images/BackGrounds/thumbnails/BMW_blue.jpg",
  "preview_url": "https://tools.brandinstitute.com/nw2/assets/images/BackGrounds/BMW_blue.jpg"
}
```

### 2. GET `/api/bi_guidelines/background-templates`

Mismo formato que arriba. Incluye ambas URLs.

### 3. POST `/api/bi_guidelines/background-templates`

Ahora genera automáticamente el thumbnail al crear un template.

## 📊 Comparación de Tamaños

| Tipo | Dimensiones | Peso Promedio | Uso |
|------|-------------|---------------|-----|
| **Thumbnail** | 300x169px | ~25KB | Listados, grids, previews rápidas |
| **Original** | 1920x1080px | ~500KB | Vista completa, modales, detalles |

**Reducción:** ~95% menos datos al usar thumbnails

## 🛠️ Implementación Técnica

### Archivos Modificados

1. **`app/utils/image_processor.py`** (NUEVO)
   - Funciones para generar thumbnails
   - Maneja conversión PNG → JPEG
   - Optimización automática

2. **`app/models/response_models.py`**
   - Añadido campo `thumbnail_url` a `TemplateGroup`

3. **`app/api/routes/bi_guidelines.py`**
   - POST: Genera thumbnail al crear template
   - GET: Construye `thumbnail_url` automáticamente

4. **`app/services/bi_guidelines_service.py`**
   - Construye URLs de thumbnails en `get_template_groups()`

### Script de Migración

Para generar thumbnails de templates existentes:

```bash
python generate_all_thumbnails.py
```

**Resultado:** Genera thumbnails para las ~180 imágenes existentes en minutos.

## 🎨 Detalles Técnicos

### Dimensiones del Thumbnail
- **Width:** 300px
- **Height:** 169px
- **Aspect Ratio:** 16:9 (mantiene proporción original)

### Formato
- **Entrada:** JPG, PNG, o PPTX (primer slide)
- **Salida:** JPEG optimizado
- **Calidad:** 85 (balance entre calidad y tamaño)

### Conversión PNG → JPEG
- Transparencias se convierten a fondo blanco
- Optimización automática para reducir tamaño

### Path del Thumbnail
```
Original:   images/BackGrounds/BMW_blue.jpg
Thumbnail:  images/BackGrounds/thumbnails/BMW_blue.jpg
            └─────────────────┬────────────┘
                    Mismo path, subcarpeta "thumbnails"
```

## ⚡ Performance

### Antes (Sin Thumbnails)
```
50 templates × 500KB = 25MB de descarga
Tiempo estimado (3G): ~8-10 segundos
```

### Ahora (Con Thumbnails)
```
50 templates × 25KB = 1.25MB de descarga
Tiempo estimado (3G): ~400ms
Reducción: 95%
```

## 🔍 Verificación

Para verificar que los thumbnails se generaron correctamente:

```bash
# Verifica estructura de directorios
ls C:\inetpub\wwwroot\nw2\assets\images\BackGrounds\thumbnails\

# Debería mostrar ~180 archivos .jpg
```

## 🚀 Casos de Uso

### 1. Listado de Templates (Usa thumbnails)
```jsx
const TemplateGrid = ({ templates }) => (
  <div className="template-grid">
    {templates.map(t => (
      <div key={t.template_group_id} className="template-card">
        <img src={t.thumbnail_url} alt={t.template_name} />
        <span>{t.template_name}</span>
      </div>
    ))}
  </div>
);
```

### 2. Modal de Preview (Usa original)
```jsx
const TemplateModal = ({ template }) => (
  <Modal>
    <img 
      src={template.preview_url} 
      alt={template.template_name}
      className="full-resolution"
    />
  </Modal>
);
```

### 3. Selector con Preview (Progresivo)
```jsx
const TemplateSelector = ({ templates, onSelect }) => {
  const [selected, setSelected] = useState(null);
  
  return (
    <div>
      {/* Grid con thumbnails */}
      <div className="template-grid">
        {templates.map(t => (
          <img 
            key={t.template_group_id}
            src={t.thumbnail_url}
            onClick={() => setSelected(t)}
          />
        ))}
      </div>
      
      {/* Preview grande del seleccionado */}
      {selected && (
        <div className="preview">
          <img src={selected.preview_url} />
        </div>
      )}
    </div>
  );
};
```

## ✅ Checklist de Integración Frontend

- [ ] Actualizar modelos TypeScript para incluir `thumbnail_url`
- [ ] Modificar componentes de listado para usar `thumbnail_url`
- [ ] Mantener `preview_url` para modales/detalles
- [ ] Implementar carga progresiva (opcional pero recomendado)
- [ ] Probar en red lenta (throttling 3G)
- [ ] Verificar mejora en tiempos de carga

## 📝 Notas

- Los thumbnails se generan **automáticamente** al crear templates
- No requiere cambios en la base de datos (solo lectura)
- Compatible con templates antiguos (se generaron retrospectivamente)
- El sistema sigue funcionando si falta un thumbnail (fallback a original)
- Los thumbnails siempre son JPEG independientemente del formato original

## 🔗 Endpoints Relacionados

```
GET  /api/bi_guidelines/template-groups          → Incluye thumbnail_url
GET  /api/bi_guidelines/background-templates     → Incluye thumbnail_url
POST /api/bi_guidelines/background-templates     → Genera thumbnail automáticamente
```

## 🎯 Recomendación Final

**Para listados/grids:** Siempre usa `thumbnail_url`  
**Para detalles/modales:** Usa `preview_url`  
**Para mejor UX:** Implementa carga progresiva
