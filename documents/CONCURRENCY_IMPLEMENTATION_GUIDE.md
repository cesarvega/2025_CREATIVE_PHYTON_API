# Sistema de Control de Concurrencia con WebSockets - Guía de Implementación

## Resumen Ejecutivo

Se ha implementado un sistema completo de control de concurrencia con cola de tareas y notificaciones en tiempo real para la API de generación de presentaciones. El sistema permite gestionar eficientemente tareas pesadas de PowerPoint COM automation, evitando saturación del servidor.

## Arquitectura del Sistema

```
┌─────────────────────────────────────────────────────────────────┐
│                         FastAPI Application                      │
├─────────────────────────────────────────────────────────────────┤
│                                                                   │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────────┐      │
│  │   REST API   │    │  Socket.IO   │    │   Lifespan   │      │
│  │  Endpoints   │    │   Manager    │    │   Manager    │      │
│  └──────┬───────┘    └──────┬───────┘    └──────┬───────┘      │
│         │                    │                    │              │
│         └────────────────────┼────────────────────┘              │
│                              │                                   │
│  ┌───────────────────────────▼──────────────────────────────┐   │
│  │            ConcurrencyManager (Orchestrator)             │   │
│  ├──────────────────────────────────────────────────────────┤   │
│  │  • Coordinates all components                            │   │
│  │  • Submits tasks to queue                                │   │
│  │  • Manages worker pool                                   │   │
│  │  • Emits Socket.IO events                                │   │
│  │  • Tracks statistics                                     │   │
│  └──────┬───────────────────┬───────────────────┬───────────┘   │
│         │                   │                   │               │
│    ┌────▼────┐        ┌────▼────┐        ┌────▼────┐          │
│    │  Task   │        │ Worker  │        │ Socket  │          │
│    │  Queue  │        │  Pool   │        │   IO    │          │
│    └─────────┘        └─────────┘        └─────────┘          │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

## Componentes Implementados

### 1. **TaskQueue** (`app/utils/task_queue.py`)
Cola de prioridad thread-safe para gestionar tareas pendientes.

**Características:**
- Priorización de tareas (HIGH, NORMAL, LOW)
- Límite configurable de tamaño
- Lookup rápido por task_id
- Ordenamiento automático por prioridad

**Métodos principales:**
```python
# Agregar tarea a la cola
position = queue.enqueue(task_id, task_type, func, args, kwargs, priority)

# Obtener próxima tarea (por prioridad)
task = queue.dequeue()

# Cancelar tarea en cola
success = queue.cancel(task_id)

# Obtener posición en cola
position = queue.get_position(task_id)

# Estadísticas
stats = queue.get_stats()
```

### 2. **TaskWorker & WorkerPool** (`app/utils/task_worker.py`)
Pool de workers que consumen tareas de la cola y las ejecutan.

**Características:**
- Workers en threads separados
- Manejo de timeouts
- Callbacks para status y progress
- Auto-recuperación de errores

**Métodos principales:**
```python
# Inicializar y arrancar pool
pool = WorkerPool(num_workers=1, task_queue=queue)
pool.start()

# Detener pool
pool.stop(wait=True, timeout=10.0)

# Estadísticas
stats = pool.get_stats()
```

### 3. **SocketIOManager** (`app/utils/socketio_manager.py`)
Gestor de Socket.IO para notificaciones en tiempo real.

**Eventos emitidos:**

**Eventos de Tarea Individual:**
- `task:queued` → Tarea agregada a la cola
- `task:processing` → Tarea comenzó a ejecutarse
- `task:progress` → Actualización de progreso
- `task:completed` → Tarea completada
- `task:failed` → Tarea falló

**Eventos Globales:**
- `queue:updated` → Estado de la cola cambió
- `worker:status` → Estado de worker cambió

**Métodos principales:**
```python
# Emitir evento de tarea
await socketio_manager.emit_task_event(task_id, 'task:progress', {...})

# Emitir evento global
await socketio_manager.emit_global_event('queue:updated', {...})

# Métodos convenientes
await socketio_manager.emit_task_queued(task_id, position, estimated_wait)
await socketio_manager.emit_task_progress(task_id, progress, message)
await socketio_manager.emit_task_completed(task_id, result, total_seconds)
```

### 4. **ConcurrencyManager** (`app/utils/concurrency_manager_v2.py`)
Orquestador principal que coordina todos los componentes.

**Características:**
- Integra Queue + Workers + SocketIO
- Gestión de slots disponibles
- Estimación de tiempos de espera
- Estadísticas globales

**Métodos principales:**
```python
# Iniciar sistema
concurrency_manager.start()

# Enviar tarea
result = await concurrency_manager.submit_task(
    task_id, task_type, func, args, kwargs, priority, metadata
)

# Cancelar tarea
cancelled = await concurrency_manager.cancel_task(task_id)

# Obtener estado de cola
status = concurrency_manager.get_queue_status()

# Cleanup de tareas antiguas
cleaned = await concurrency_manager.cleanup_old_tasks(max_age_hours=24)
```

## Configuración

### Settings (`app/config/settings.py`)

```python
# Concurrency settings
concurrency_max_workers: int = 1  # PowerPoint COM: solo 1 simultánea
concurrency_queue_size: int = 100
task_timeout_seconds: int = 600  # 10 minutos máximo
progress_update_interval: int = 5  # segundos

# Socket.IO settings
socketio_cors_origins: str = "http://localhost:4200,https://tools.brandinstitute.com"
socketio_async_mode: str = "asgi"
```

Puedes sobrescribir estos valores usando variables de entorno en `.env`:

```env
CONCURRENCY_MAX_WORKERS=1
CONCURRENCY_QUEUE_SIZE=100
TASK_TIMEOUT_SECONDS=600
SOCKETIO_CORS_ORIGINS=http://localhost:4200,https://tools.brandinstitute.com
```

## API REST Endpoints

### Gestión de Tareas

#### `GET /api/tasks/{task_id}`
Obtener estado detallado de una tarea.

**Response:**
```json
{
  "task_id": "abc123-def456-ghi789",
  "task_type": "create_presentation",
  "status": "processing",
  "progress": 45,
  "description": "Creating presentation...",
  "result": null,
  "error": null,
  "created_at": "2025-01-15T10:30:00",
  "started_at": "2025-01-15T10:30:05",
  "completed_at": null
}
```

#### `DELETE /api/tasks/{task_id}`
Cancelar una tarea en cola.

**Response:**
```json
{
  "success": true,
  "message": "Task abc123 cancelled successfully",
  "task_id": "abc123"
}
```

#### `GET /api/tasks/queue/status`
Obtener estado global de la cola.

**Response:**
```json
{
  "active_count": 1,
  "queued_count": 3,
  "available_slots": 0,
  "max_workers": 1,
  "queue_capacity": 100,
  "statistics": {
    "total_submitted": 150,
    "total_completed": 145,
    "total_failed": 2
  }
}
```

#### `GET /api/tasks/queue/position/{task_id}`
Obtener posición de una tarea en la cola.

**Response:**
```json
{
  "task_id": "abc123",
  "in_queue": true,
  "position": 2,
  "estimated_wait_seconds": 180,
  "estimated_wait_minutes": 3.0
}
```

#### `GET /api/tasks/queue/history`
Obtener historial de tareas.

**Query Parameters:**
- `limit` (int): Máximo número de tareas (default: 100, max: 500)
- `task_type` (str): Filtrar por tipo de tarea
- `status` (str): Filtrar por estado (pending, processing, completed, failed)

**Response:**
```json
{
  "total": 50,
  "limit": 100,
  "filters": {
    "task_type": null,
    "status": "completed"
  },
  "tasks": [...]
}
```

#### `POST /api/tasks/queue/cleanup`
Limpiar tareas antiguas.

**Query Parameters:**
- `max_age_hours` (int): Edad máxima en horas (default: 24, max: 168)

**Response:**
```json
{
  "success": true,
  "cleaned_count": 15,
  "max_age_hours": 24,
  "message": "Cleaned up 15 tasks older than 24 hours"
}
```

## Integración con Socket.IO

### Cliente HTML/JavaScript

Ver archivo: `documents/socketio_client_example.html`

Ejemplo básico:
```javascript
const socket = io('http://localhost:50100', {
    transports: ['websocket', 'polling'],
    path: '/socket.io'
});

// Conexión
socket.on('connect', () => {
    console.log('Connected!');
    // Suscribirse a actualizaciones globales
    socket.emit('subscribe_global');
});

// Unirse a room de tarea específica
socket.emit('join_task', { task_id: 'abc123' });

// Escuchar eventos de tarea
socket.on('task:progress', (data) => {
    console.log(`Progress: ${data.progress}%`);
});

socket.on('task:completed', (data) => {
    console.log('Task completed!', data.result);
});

// Escuchar actualizaciones de cola
socket.on('queue:updated', (data) => {
    console.log('Queue:', data.active_count, data.queued_count);
});
```

### Cliente JavaScript Clase

Ver archivo: `documents/socketio_client_example.js`

```javascript
import { TaskMonitor } from './socketio_client_example.js';

const monitor = new TaskMonitor('http://localhost:50100');

// Conectar
monitor.connect();

// Escuchar eventos
monitor.on('taskProgress', (data) => {
    console.log(`Progress: ${data.progress}%`);
});

// Trackear tarea específica
monitor.trackTask('abc123');
```

## Flujo de Ejecución Completo

### Escenario 1: Worker disponible (ejecución inmediata)

```
1. Cliente → POST /api/presentations/create
2. ConcurrencyManager verifica slots disponibles
3. Hay worker idle → Agregar a queue
4. Worker toma tarea inmediatamente
5. Socket.IO → 'task:processing'
6. Worker ejecuta tarea
7. Socket.IO → 'task:progress' (cada 5s)
8. Worker completa
9. Socket.IO → 'task:completed'
10. Cliente recibe resultado
```

### Escenario 2: Sin workers disponibles (espera en cola)

```
1. Cliente → POST /api/presentations/create
2. ConcurrencyManager verifica slots disponibles
3. No hay workers idle → Agregar a queue
4. Socket.IO → 'task:queued' (position: 3, estimated_wait: 270s)
5. Cliente espera en cola
6. Worker anterior termina
7. Queue.dequeue() obtiene siguiente tarea
8. Socket.IO → 'task:processing'
9. [Resto igual a Escenario 1]
```

## Migración desde Sistema Anterior

El nuevo sistema es **compatible hacia atrás** con el código existente:

1. El `TaskManager` existente se mantiene sin cambios
2. Los background tasks existentes funcionan sin modificaciones
3. Solo necesitas actualizar el código de submission para usar el nuevo sistema:

### Antes:
```python
task_id = task_manager.create_task("create_presentation")
background_tasks.add_task(_create_presentation_background, task_id, ...)
return {"task_id": task_id, "status": "pending"}
```

### Ahora:
```python
task_id = task_manager.create_task("create_presentation")
result = await concurrency_manager.submit_task(
    task_id=task_id,
    task_type="create_presentation",
    func=_create_presentation_background,
    args=(task_id, metadata, excel_content, ...),
    priority=TaskPriority.NORMAL
)
return {
    "task_id": task_id,
    "status": result["status"],
    "position": result.get("position"),
    "estimated_wait_seconds": result.get("estimated_wait_seconds")
}
```

## Instalación de Dependencias

```bash
# Instalar nuevas dependencias
pip install -r requirements.txt

# Las nuevas dependencias son:
# - python-socketio==5.11.4
# - aiohttp==3.11.11
```

## Pruebas

### 1. Verificar que el servidor inicia correctamente

```bash
# En el directorio del proyecto
python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 50100
```

Deberías ver en los logs:
```
INFO: Starting CreativePythonAPI...
INFO: Initializing concurrency manager...
INFO: ConcurrencyManager initialized: max_workers=1, queue_size=100, timeout=600s
INFO: Starting worker pool with 1 workers
INFO: TaskWorker 0 initialized
INFO: Worker 0 started
INFO: Concurrency manager started with 1 workers
```

### 2. Probar conexión Socket.IO

Abre `documents/socketio_client_example.html` en un navegador.

Deberías ver:
- Status: Connected (verde)
- Queue statistics: 0 active, 0 queued, 1 available

### 3. Probar endpoints REST

```bash
# Obtener estado de cola
curl http://localhost:50100/api/tasks/queue/status

# Obtener historial
curl http://localhost:50100/api/tasks/queue/history?limit=10
```

### 4. Probar creación de tarea

Usa Postman o curl para enviar una tarea:
```bash
curl -X POST http://localhost:50100/api/presentations/create \
  -F "project=TestProject" \
  -F "display_name=Test" \
  # ... otros campos
```

Deberías ver en el cliente Socket.IO:
- Evento `task:queued` o `task:processing`
- Eventos `task:progress` periódicamente
- Evento `task:completed` al finalizar

## Monitoreo y Debugging

### Logs

Los logs del sistema incluyen información detallada:

```
INFO: Task abc123 enqueued at position 0 (priority=NORMAL, type=create_presentation)
INFO: Worker 0 executing task abc123 (type=create_presentation, priority=2)
INFO: Task abc123 updated: status=TaskStatus.PROCESSING, progress=20
INFO: Worker 0 completed task abc123 in 62.5s
```

### Métricas

Obtener estadísticas en tiempo real:

```python
# En el código
stats = concurrency_manager.get_queue_status()
print(f"Active: {stats['active_count']}, Queued: {stats['queued_count']}")
```

## Troubleshooting

### Problema: Socket.IO no conecta

**Solución:**
1. Verificar CORS en settings.py: `socketio_cors_origins`
2. Verificar que el puerto está abierto
3. Verificar que la app Socket.IO está montada en `/socket.io`

### Problema: Tareas no se ejecutan

**Solución:**
1. Verificar que el worker pool está iniciado:
   ```python
   concurrency_manager.start()
   ```
2. Verificar logs para errores en workers
3. Verificar que la tarea está en la cola:
   ```bash
   curl http://localhost:50100/api/tasks/queue/status
   ```

### Problema: Tareas tardan mucho

**Solución:**
1. Verificar el límite de workers (default: 1 para PowerPoint COM)
2. Aumentar `concurrency_max_workers` si es seguro para tu caso
3. Verificar logs para cuellos de botella
4. Monitorear progreso con Socket.IO

## Mejoras Futuras

1. **Persistencia en Base de Datos:**
   - Guardar tareas en SQL Server para recuperación tras reinicio
   - Implementar tabla `task_queue` con estado y prioridad

2. **Métricas Avanzadas:**
   - Prometheus/Grafana para monitoreo
   - Alertas cuando cola crece demasiado
   - Análisis de tiempos de ejecución

3. **Escalabilidad:**
   - Redis como cola compartida para múltiples workers/servidores
   - Load balancing de Socket.IO con Redis adapter

4. **Seguridad:**
   - Autenticación en Socket.IO
   - Rate limiting por usuario
   - Validación de permisos para cancelar tareas

## Soporte

Para preguntas o problemas, contactar:
- Email: creative@brandinstitute.com
- Logs del servidor: `logs/app.log`
- Documentación FastAPI: http://localhost:50100/docs

---

**Autor:** Claude Code
**Fecha:** 2025-01-15
**Versión:** 1.0.0
