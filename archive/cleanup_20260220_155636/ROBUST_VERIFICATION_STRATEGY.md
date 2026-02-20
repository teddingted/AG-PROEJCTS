# Robust Verification Strategy - Lessons Learned

## Key Success Patterns from `hysys_optimizer_final.py`

### 1. **Solver Control**
```python
self.solver.CanSolve = False  # Pause solver
# Make changes...
self.solver.CanSolve = True   # Resume solver
```
- **Why**: Prevents solver from interfering during state changes
- **Application**: Use in recovery and hard reset

### 2. **Health Checks**
```python
def is_healthy(self):
    t = self.s1.Temperature.Value
    if t < -200: return False  # Cryogenic failure
    if self.comp_k100.EnergyValue < 0.1: return False  # Compressor trip
    return True
```
- **Why**: Early detection of divergence
- **Application**: Check before every operation

### 3. **Hard Reset Pattern**
```python
def recover_state(self, flow, p_bar):
    self.solver.CanSolve = False
    # Set safe warm state (40°C, -90°C target)
    # Reset all adjusts
    self.solver.CanSolve = True
    self.wait_stable(20, 1.0)
    # **Double Reset** to clear logical latches
    for adj in self.adjs: adj.Reset()
    self.wait_stable(10, 1.0)
```
- **Why**: Multi-stage reset ensures complete recovery
- **Application**: Call before each verification point

### 4. **Gradual Pressure Changes**
```python
if abs(current_P - target_P) > 100:  # > 1 bar
    # Apply change gradually
    self.wait_stable(5, 0.5)
```
- **Why**: Large pressure jumps cause instability
- **Application**: Step pressure changes in verification

### 5. **Anti-Lag Logic**
```python
if app < 0.5:  # Potentially transient
    time.sleep(1.5)  # Wait for settle
    app = self.get_metrics()['app']  # Re-check
```
- **Why**: Adjust blocks can lag behind solver
- **Application**: Double-check critical metrics

### 6. **Settle Buffers**
```python
self.wait_stable(20, 1.0)
time.sleep(1.5)  # CRITICAL: Settle buffer for laggy adjusts
```
- **Why**: Spreadsheet/Adjust convergence lags solver
- **Application**: Add buffers after wait_stable

### 7. **Temperature Stability Check**
```python
while timeout:
    curr = self.s1.Temperature.Value
    if abs(curr - last_val) < 0.01:
        if stable for > 1.0s: return True
```
- **Why**: More reliable than IsSolving flag
- **Application**: Use for convergence verification

## Proposed Robust Verification Scheme

### Phase 1: Initialize from Known Good State
1. Start at 500 kg/h (lowest, most stable)
2. Use `recover_state()` to establish baseline
3. Verify health before proceeding

### Phase 2: Sequential Upward Verification (500→1500)
```python
for flow in [500, 600, ..., 1500]:
    1. recover_state(flow, PRESET_P[flow])  # Hard reset
    2. wait_stable(20, 1.0)
    3. Double reset adjusts
    4. set_inputs_safe(flow, P, T)  # With gradual P change
    5. wait_stable(20, 1.0)
    6. time.sleep(1.5)  # Settle buffer
    7. is_healthy() check
    8. Anti-lag double-check on metrics
    9. Collect data only if healthy
```

### Phase 3: Critical Failure Handling
- If 3 consecutive failures → full solver restart recommended
- Log failures but continue (collect what we can)

### Phase 4: Data Validation
- MA: must be > 0 (not -32767)
- Power: must be > 10 kW
- S6_Pres: must be < 40 bar
- Flag extrapolated/invalid points clearly

## Implementation Priority

**HIGH:**
- Solver control (CanSolve)
- Health checks
- Double reset pattern
- Settle buffers

**MEDIUM:**
- Gradual pressure changes
- Anti-lag logic

**LOW:**
- Enhanced logging
- Retry counters
