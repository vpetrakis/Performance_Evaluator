from models import EngineSnapshot, EvaluationResult

# Stage 1 Fleet Rules
MAX_EXHAUST_DEV = 15.0  # Celsius
MAX_PMAX_DEV = 3.0      # Bar

def evaluate_engine(snapshot: EngineSnapshot) -> EvaluationResult:
    alerts = []
    
    # Calculate Pmax average
    avg_pmax = sum(c.p_max for c in snapshot.cylinders) / len(snapshot.cylinders)
    
    for cyl in snapshot.cylinders:
        # 1. Thermal Imbalance Check
        if abs(cyl.exhaust_temp - snapshot.average_exhaust) > MAX_EXHAUST_DEV:
            alerts.append(
                f"Cylinder {cyl.id} exhaust ({cyl.exhaust_temp}°C) deviates from average ({snapshot.average_exhaust}°C) by > {MAX_EXHAUST_DEV}°C."
            )
            
        # 2. Pressure Imbalance Check
        if abs(cyl.p_max - avg_pmax) > MAX_PMAX_DEV:
            alerts.append(
                f"Cylinder {cyl.id} Pmax ({cyl.p_max} bar) deviates from average ({avg_pmax:.1f} bar)."
            )
            
    # 3. Scavenge Air Check
    if snapshot.scavenge_temp > 50.0:
        alerts.append(f"Scavenge air temp high ({snapshot.scavenge_temp}°C). Check cooler/fouling.")

    # Determine Executive Status
    if len(alerts) >= 3:
        status = "RED"
    elif len(alerts) > 0:
        status = "YELLOW"
    else:
        status = "GREEN"

    return EvaluationResult(status=status, alerts=alerts, snapshot=snapshot)
