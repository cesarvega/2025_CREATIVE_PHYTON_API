"""
Diagnostic script for position=2 issue after cancellation.

SCENARIO REPORTED:
------------------
1. Task 1 created
2. Task 2 created (queued)
3. Task 2 cancelled
4. Task 3 created -> Shows position=2 (WRONG)

EXPECTED BEHAVIOR:
------------------
At Task 3 creation:
- If Task 1 finished: processing_workers=0, queue_pos=0 -> position=0
- If Task 1 running: processing_workers=1, queue_pos=0 -> position=1

POSSIBLE CAUSES:
----------------
1. Ghost worker: Worker._current_task_id not cleared after task completes
2. Phantom queue: Queue not properly cleared after cancellation
3. Stats caching: get_stats() returning stale data
4. Race condition: Task cancelled but worker already dequeued it

DIAGNOSIS STEPS:
----------------
1. Check logs for "Pool stats before enqueue" to see processing_workers count
2. Check if processing_workers=2 or queue_pos=2
3. Review worker lifecycle to ensure _current_task_id is cleared

To diagnose, run the test scenario and check the logs for:
- "Pool stats before enqueue: idle=X, processing=Y, total=Z"
- If processing=2 -> Ghost worker issue (worker not clearing task ID)
- If queue_pos=2 -> Phantom queue issue (cancelled tasks not removed)
"""
print(__doc__)
