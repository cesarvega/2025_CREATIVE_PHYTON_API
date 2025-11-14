# Sistema de Posición en Cola - Actualización Dinámica

## Problema Resuelto

Cuando un usuario cancela una tarea que está en cola, los demás usuarios que están esperando verán su posición actualizada automáticamente en el siguiente polling.

## Cómo Funciona

### Escenario Ejemplo:

```
Usuario A: Tarea 1 → Procesándose
Usuario B: Tarea 2 → Posición 1 en cola
Usuario C: Tarea 3 → Posición 2 en cola
```

**Usuario B cancela su tarea:**

```
Usuario A: Tarea 1 → Procesándose
Usuario C: Tarea 3 → Posición 1 en cola (actualizado automáticamente)
```

## Implementación

### 1. Cálculo en Tiempo Real

La posición **NO se guarda** como un valor estático. Se calcula dinámicamente cada vez que:
- El frontend hace polling al endpoint `/api/presentations/tasks/{task_id}`
- El frontend consulta `/api/tasks/queue/position/{task_id}`

### 2. Fórmula de Cálculo

```python
actual_position = processing_workers + queue_position
```

Donde:
- `processing_workers`: Número de workers ocupados procesando tareas
- `queue_position`: Posición de la tarea dentro de la cola (0-indexed)

### 3. Endpoints Actualizados

#### GET `/api/presentations/tasks/{task_id}` (Endpoint principal de polling)

**Respuesta cuando está en cola:**
```json
{
    "task_id": "...",
    "status": "pending",
    "progress": 0,
    "position": 1,                    // Calculado dinámicamente
    "estimated_wait_seconds": 90.0,   // Calculado dinámicamente
    ...
}
```

#### GET `/api/tasks/queue/position/{task_id}` (Endpoint específico de posición)

**Respuesta:**
```json
{
    "task_id": "...",
    "in_queue": true,
    "position": 1,                  // Total de tareas adelante
    "queue_position": 0,            // Posición dentro de la cola
    "processing_tasks": 1,          // Tareas procesándose
    "estimated_wait_seconds": 90.0,
    "estimated_wait_minutes": 1.5
}
```

## Flujo de Actualización en Frontend

```typescript
// Polling cada 2 segundos
setInterval(async () => {
  const status = await this.http.get(`/api/presentations/tasks/${taskId}`).toPromise();

  if (status.status === 'pending' && status.position !== null) {
    // Mostrar posición actualizada
    console.log(`Posición en cola: ${status.position}`);
    console.log(`Tiempo estimado: ${status.estimated_wait_seconds}s`);
  }
}, 2000);
```

## Ventajas

✅ **Siempre actualizado**: La posición refleja el estado real de la cola en cada consulta

✅ **Sin datos obsoletos**: No hay valores "cacheados" que puedan quedar desactualizados

✅ **Cancelaciones reflejadas inmediatamente**: Cuando alguien cancela, la siguiente consulta de los demás usuarios ya muestra la posición correcta

✅ **Múltiples usuarios**: Funciona correctamente sin importar cuántos usuarios cancelen o agreguen tareas

## Comportamiento

| Escenario | Posición Inicial | Después de Cancelación | Próximo Polling |
|-----------|-----------------|------------------------|-----------------|
| Usuario A procesa tarea | N/A | N/A | N/A |
| Usuario B espera (pos 1) | 1 | Cancela | N/A |
| Usuario C espera (pos 2) | 2 | Pasa a pos 1 | **Ve posición 1** ✅ |

## Protección Contra Race Conditions

El sistema también protege contra cancelaciones en el momento exacto en que una tarea va a empezar:

- Si el usuario cancela **antes** de que el worker tome la tarea → Cancelación exitosa
- Si el worker toma la tarea **antes** de que llegue la cancelación → Error 409 "Task is already processing"

Esto se garantiza mediante el uso de locks en la cola (ver `task_queue.py`).
