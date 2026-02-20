# HYSYS Optimization Process Report & High-Flow Strategy

## 1. Optimization Process Logic (Based on `hysys_optimizer_final.py`)

Our optimization engine uses a multi-layered robust approach to handle HYSYS instability:

### **Phase 1: Safe Initialization (Recovery Mode)**
- **Concept**: Before attempting any optimization, the system performs a "Hard Reset" to a known stable state.
- **Action**:
  - `CanSolve = False`: Suspend solver to prevent interference.
  - Set N2 Pressure to a "Safe Warm Start" (e.g., 40°C, low target -90°C).
  - **Reset Adjust Blocks**: Force all controllers to reset integral terms.
  - `CanSolve = True`: Resume solver and wait for stability (20s).
- **Benefit**: Ensures every optimization run starts from a clean slate, eliminating "memory" of previous crashes.

### **Phase 2: Coarse Grid Search (The "Scan")**
- **Trigger**: Once stable.
- **Action**:
  - **Pressure Scan**: Iterate P in steps of 0.1 bar around a center point (e.g., 4.0 ± 0.2 bar).
  - **Temperature Scan**: For each P, sweep Target Temp from -90°C down to -124°C in steps of -2°C.
- **Logic**:
  - If `MA > 3.5`: Too wide -> Continue lowering T ("W").
  - If `0.5 <= MA <= 3.5`: **HIT** -> Feasible region found.
  - If `MA < 0.5` or Crash: Too tight -> Stop scanning this P (".").

### **Phase 3: Fine Tuning & Anti-Lag**
- **Trigger**: Valid range identified in Phase 2.
- **Action**:
  - Re-scan the valid temperature range with finer resolution (1°C steps).
  - **Anti-Lag Check**: If a result looks suspicious (e.g., sudden MA drop), wait 1.5s and re-read. HYSYS Adjust blocks often lag behind the steady-state solver.
- **Selection**: Choose the point with **Lowest Power** that maintains constraints (MA > 0.5°C, P7 < 36.5 bar).

---

## 2. Process Logs & Evidence

The `robust_verification_report.csv` demonstrates the effectiveness of this process. Note the clean progression:

| Flow | P (bar) | T (°C) | Power (kW) | MA (°C) | Result |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **500** | 3.1 | -116 | 694.8 | 2.45 | **Stable** (Wide margin) |
| **900** | 4.8 | -110 | 987.4 | 0.54 | **Critical** (Tight MA, but solved) |
| **1300** | 6.3 | -100 | 1238.3 | 0.54 | **Limit** (Very tight, requires robust handling) |

*Key Insight*: The optimization successfully pushed the system to the edge (MA ~0.54°C) without crashing, maximizing efficiency.

---

## 3. Strategy for 1400 & 1500 kg/h (High-Flow Optimization)

The 1400/1500 kg/h points failed due to solver instability (popups/divergence). Here is the proposed strategy to conquer them:

### **Strategy A: "Creeping" Convergence (Incremental Step)**
Instead of jumping from 1300 to 1400 (a 100 kg/h shock), we should approach it incrementally.
1.  **Start at 1300 kg/h** (Trusted State).
2.  **Increase to 1350 kg/h**: Allow solver to adjust.
3.  **Increase to 1400 kg/h**:
    - **Pre-set Pressure**: Set P directly to **6.7 bar** (Predicted by linear formula).
    - **Relax MA**: Set Target T higher initially (-95°C) to ensure convergence, then lower it.

### **Strategy B: Manual "Kick" & Hold**
If automatic creeping fails, use manual intervention for the first convergence.
1.  **Manual Set**: Set Flow=1400, P=6.7 bar in HYSYS manually.
2.  **Solver Hold**: If it oscillates, put Solver on "Hold", manually type in estimated stream values (e.g., Stream 1 Flow ~28,000 kg/h), then releasing Hold.
3.  **Lock**: Once converged manually, run the **Verification Script** to "lock in" the data point for the report.

### **Strategy C: Relaxed Constraints**
- The MA limit of 0.5°C might be physically impossible at 1500 kg/h with the current equipment sizing.
- **Action**: Accept MA < 0.5°C (e.g., 0.1°C) if stable, OR acknowledge that 1500 kg/h exceeds the equipment's efficient operating envelope (hardware limit).

### **Recommended Next Step**
Try **Strategy A (Creeping)** using a modified script or manual operation, as the linear pressure prediction ($P \approx 3.1 + 0.004 \Delta F$) is proving highly accurate.
