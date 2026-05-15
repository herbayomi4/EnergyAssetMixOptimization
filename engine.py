import xlwings as xw
from collections import deque
import time

def run_optimization():
    print("=" * 60)
    print("HYPERSCALER CAPACITY OPTIMIZATION")
    print("Cost-First Search: Guaranteed Cheapest Feasible")
    print("=" * 60)
    
    start_time = time.time()
    
    wb = xw.Book('Optimization & Financial Model.xlsm')
    sht = wb.sheets['Optimization Dashboard']
    sht.range('M3').value = "Running..."

    # Load inputs
    demand_mw = float(sht.range('C5').value)
    target_uptime = float(sht.range('C6').value) / 100
    project_years = float(sht.range('C7').value)
    solar_life = float(sht.range('C8').value)
    wind_life = float(sht.range('C9').value)
    solar_ilr = float(sht.range('C10').value)
    batt_eff_rt = float(sht.range('C11').value) / 100
    batt_duration = float(sht.range('C12').value)
    max_cycles_per_day = float(sht.range('C13').value or 2)

    # Wind toggle
    raw_wind = sht.range('C14').value
    wind_enabled = int(raw_wind) == 1 if raw_wind is not None else True
    
    ref_solar = float(sht.range('C15').value)
    ref_wind = float(sht.range('C16').value)
    
    
    # Derived constraint - Energy expires after batt_duration hours
    energy_shelf_life = int(batt_duration)

    # Cost inputs
    capex_linear_s = float(sht.range('F5').value or 0)
    capex_linear_w = float(sht.range('F6').value or 0)
    capex_linear_b = float(sht.range('F7').value or 0)
    capex_constant_s = float(sht.range('F10').value or 0)
    capex_constant_w = float(sht.range('F11').value or 0)
    capex_constant_b = float(sht.range('F12').value or 0)

    opex_lin_recur_s = float(sht.range('I5').value or 0)
    opex_lin_recur_w = float(sht.range('I6').value or 0)
    opex_lin_recur_b = float(sht.range('I7').value or 0)
    opex_con_recur_s = float(sht.range('I10').value or 0)
    opex_con_recur_w = float(sht.range('I11').value or 0)
    opex_con_recur_b = float(sht.range('I12').value or 0)
    opex_con_one_s = float(sht.range('I15').value or 0)
    opex_con_one_w = float(sht.range('I16').value or 0)
    opex_con_one_b = float(sht.range('I17').value or 0)

    max_fail_hours = int(8760 * (1 - target_uptime))
    
    # Tolerance for "full" capacity (99.9%)
    CAPACITY_TOLERANCE = 0.999

    print(f"\nDemand: {demand_mw} MW")
    print(f"Target Uptime: {target_uptime*100:.3f}%")
    print(f"Max Failed Hours: {max_fail_hours}")
    print(f"Battery Duration: {batt_duration} hours")
    print(f"Energy Shelf Life: {energy_shelf_life} hours")
    print(f"Max Cycles/Day: {max_cycles_per_day}")
    print(f"  -> 1 cycle = SOC: 0 → FULL capacity → back to 0")
    print(f"Wind Enabled: {'Yes' if wind_enabled else 'No'}")

    # Build profiles
    sht_solar = wb.sheets['EPE_Solar']
    sht_wind = wb.sheets['EPE_Wind']
    real_solar_24x12 = sht_solar.range('B2:M25').value
    real_wind_24x12 = sht_wind.range('B2:M25').value

    days_in_month = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    solar_cf = []
    wind_cf = []
    
    for month_idx, days in enumerate(days_in_month):
        for day in range(days):
            for hour in range(24):
                s_raw = real_solar_24x12[hour][month_idx]
                w_raw = real_wind_24x12[hour][month_idx]
                s_cf = s_raw / ref_solar if ref_solar > 0 else 0
                w_cf = w_raw / ref_wind if ref_wind > 0 else 0
                solar_cf.append(max(0.0, min(1.0, s_cf)))
                wind_cf.append(max(0.0, min(1.0, w_cf)))

    ilr_cap = 1.0 / solar_ilr if solar_ilr > 0 else 1.0
    solar_cf_capped = [min(cf, ilr_cap) for cf in solar_cf]

    # =========================================================
    # FAST SIMULATION
    # =========================================================
    
    def simulate_fast(S_mw, W_mw, B_mwh):
        B_mw = B_mwh / batt_duration if batt_duration > 0 else 0
        full_threshold = B_mwh * CAPACITY_TOLERANCE
        
        failed = 0
        soc = 0.0
        
        cycle_completions = deque()
        in_cycle = False
        reached_full = False
        
        charge_history_shelf = deque([0.0] * energy_shelf_life, maxlen=energy_shelf_life)
        
        for t in range(8760):
            while cycle_completions and cycle_completions[0] <= t - 24:
                cycle_completions.popleft()
            
            cycles_in_window = len(cycle_completions)
            expiring = min(charge_history_shelf[0], soc)
            gen = S_mw * solar_cf_capped[t] + W_mw * wind_cf[t]
            
            if gen >= demand_mw:
                if expiring > 0:
                    soc -= expiring
                    if soc < 0.01 and in_cycle and reached_full:
                        cycle_completions.append(t)
                        in_cycle = False
                        reached_full = False
                    elif soc < 0.01:
                        in_cycle = False
                        reached_full = False
                    charge = 0.0
                else:
                    if soc < 0.01:
                        if cycles_in_window >= max_cycles_per_day:
                            charge = 0.0
                        else:
                            excess = gen - demand_mw
                            charge = min(excess, B_mw, B_mwh - soc)
                            if charge > 0.01:
                                soc += charge
                                in_cycle = True
                                if soc >= full_threshold:
                                    reached_full = True
                    else:
                        excess = gen - demand_mw
                        charge = min(excess, B_mw, B_mwh - soc)
                        soc += charge
                        if soc >= full_threshold:
                            reached_full = True
                
                charge_history_shelf.append(charge)
                
            else:
                deficit = demand_mw - gen
                needed = deficit / batt_eff_rt
                
                if expiring > 0:
                    soc -= expiring
                    if soc < 0.01 and in_cycle and reached_full:
                        cycle_completions.append(t)
                        in_cycle = False
                        reached_full = False
                    elif soc < 0.01:
                        in_cycle = False
                        reached_full = False
                    
                    if expiring >= needed:
                        pass
                    else:
                        if soc > 0.01 and in_cycle:
                            additional_needed = needed - expiring
                            limit_power = max(0.0, B_mw - expiring)
                            additional = min(additional_needed, limit_power, soc)
                            soc -= additional
                            
                            if soc < 0.01 and reached_full:
                                cycle_completions.append(t)
                                in_cycle = False
                                reached_full = False
                            elif soc < 0.01:
                                in_cycle = False
                                reached_full = False
                            
                            delivered = gen + (expiring + additional) * batt_eff_rt
                        else:
                            delivered = gen + expiring * batt_eff_rt
                        
                        if demand_mw - delivered > 0.01:
                            failed += 1
                else:
                    if soc > 0.01 and in_cycle:
                        discharge = min(needed, B_mw, soc)
                        soc -= discharge
                        
                        if soc < 0.01 and reached_full:
                            cycle_completions.append(t)
                            in_cycle = False
                            reached_full = False
                        elif soc < 0.01:
                            in_cycle = False
                            reached_full = False
                        
                        delivered = gen + discharge * batt_eff_rt
                        if demand_mw - delivered > 0.01:
                            failed += 1
                    else:
                        failed += 1
                
                charge_history_shelf.append(0.0)
            
        return failed

    # =========================================================
    # ACCURATE SIMULATION (Full FIFO)
    # =========================================================
    
    def simulate(S_mw, W_mw, B_mwh):
        B_mw = B_mwh / batt_duration if batt_duration > 0 else 0
        full_threshold = B_mwh * CAPACITY_TOLERANCE
        
        failed = 0
        hourly = []
        
        energy_queue = deque()
        current_soc = 0.0
        
        cycle_completions = deque()
        in_cycle = False
        reached_full = False
        
        def check_expiring_energy(current_hour):
            expiring = 0.0
            for charge_hour, energy in energy_queue:
                if current_hour - charge_hour >= energy_shelf_life:
                    expiring += energy
                else:
                    break
            return expiring
        
        def remove_expired_energy(current_hour):
            nonlocal current_soc
            expired = 0.0
            while energy_queue:
                charge_hour, energy = energy_queue[0]
                if current_hour - charge_hour >= energy_shelf_life:
                    energy_queue.popleft()
                    expired += energy
                else:
                    break
            current_soc -= expired
            return expired
        
        def discharge_fifo(amount_needed):
            nonlocal current_soc
            discharged = 0.0
            remaining = amount_needed
            while remaining > 0.001 and energy_queue:
                charge_hour, energy = energy_queue[0]
                if energy <= remaining:
                    energy_queue.popleft()
                    discharged += energy
                    remaining -= energy
                else:
                    energy_queue[0] = (charge_hour, energy - remaining)
                    discharged += remaining
                    remaining = 0
            current_soc -= discharged
            return discharged
        
        for t in range(8760):
            while cycle_completions and cycle_completions[0] <= t - 24:
                cycle_completions.popleft()
            
            cycles_in_window = len(cycle_completions)
            
            s_gen = S_mw * solar_cf_capped[t]
            w_gen = W_mw * wind_cf[t]
            gen = s_gen + w_gen
            
            charge, discharge_to_demand, curtail, shortfall = 0.0, 0.0, 0.0, 0.0
            expired_this_hour = 0.0
            
            expiring = check_expiring_energy(t)
            
            if gen >= demand_mw:
                delivered = demand_mw
                excess = gen - demand_mw
                
                if expiring > 0:
                    expired_this_hour = remove_expired_energy(t)
                    
                    if current_soc < 0.01 and in_cycle and reached_full:
                        cycle_completions.append(t)
                        in_cycle = False
                        reached_full = False
                    elif current_soc < 0.01:
                        in_cycle = False
                        reached_full = False
                    
                    curtail = excess
                    charge = 0.0
                else:
                    if current_soc < 0.01:
                        if cycles_in_window >= max_cycles_per_day:
                            charge = 0.0
                            curtail = excess
                        else:
                            space_available = B_mwh - current_soc
                            charge = min(excess, B_mw, max(0.0, space_available))
                            
                            if charge > 0.01:
                                energy_queue.append((t, charge))
                                current_soc += charge
                                in_cycle = True
                                if current_soc >= full_threshold:
                                    reached_full = True
                            
                            curtail = excess - charge
                    else:
                        space_available = B_mwh - current_soc
                        charge = min(excess, B_mw, max(0.0, space_available))
                        
                        if charge > 0.01:
                            energy_queue.append((t, charge))
                            current_soc += charge
                            if current_soc >= full_threshold:
                                reached_full = True
                        
                        curtail = excess - charge
                
            else:
                deficit = demand_mw - gen
                needed = deficit / batt_eff_rt
                
                if expiring > 0:
                    expiring_amount = remove_expired_energy(t)
                    
                    if current_soc < 0.01 and in_cycle and reached_full:
                        cycle_completions.append(t)
                        in_cycle = False
                        reached_full = False
                    elif current_soc < 0.01:
                        in_cycle = False
                        reached_full = False
                    
                    if expiring_amount >= needed:
                        discharge_to_demand = needed
                        expired_this_hour = expiring_amount - needed
                        delivered = demand_mw
                        shortfall = 0.0
                    else:
                        expired_this_hour = 0.0
                        
                        if current_soc > 0.01 and in_cycle:
                            additional_needed = needed - expiring_amount
                            limit_power = max(0.0, B_mw - expiring_amount)
                            
                            max_possible = min(additional_needed, limit_power, current_soc)
                            additional_discharge = discharge_fifo(max_possible)
                            
                            if current_soc < 0.01 and reached_full:
                                cycle_completions.append(t)
                                in_cycle = False
                                reached_full = False
                            elif current_soc < 0.01:
                                in_cycle = False
                                reached_full = False
                            
                            discharge_to_demand = expiring_amount + additional_discharge
                            delivered = gen + (discharge_to_demand * batt_eff_rt)
                            shortfall = max(0.0, demand_mw - delivered)
                        else:
                            discharge_to_demand = expiring_amount
                            delivered = gen + (expiring_amount * batt_eff_rt)
                            shortfall = max(0.0, demand_mw - delivered)
                        
                        if shortfall > 0.01:
                            failed += 1
                else:
                    if current_soc > 0.01 and in_cycle:
                        max_possible = min(needed, B_mw, current_soc)
                        discharge_to_demand = discharge_fifo(max_possible)
                        
                        if current_soc < 0.01 and reached_full:
                            cycle_completions.append(t)
                            in_cycle = False
                            reached_full = False
                        elif current_soc < 0.01:
                            in_cycle = False
                            reached_full = False
                        
                        delivered = gen + (discharge_to_demand * batt_eff_rt)
                        shortfall = max(0.0, demand_mw - delivered)
                        
                        if shortfall > 0.01:
                            failed += 1
                    else:
                        delivered = gen
                        shortfall = demand_mw - gen
                        failed += 1
            
            hourly.append({
                'hour': t,
                'solar_cf': solar_cf[t],
                'wind_cf': wind_cf[t],
                'solar_gen': s_gen,
                'wind_gen': w_gen,
                'total_gen': gen,
                'charge': charge,
                'discharge': discharge_to_demand * batt_eff_rt,
                'soc': current_soc,
                'curtail': curtail,
                'shortfall': shortfall,
                'delivered': delivered,
                'served': 1 if shortfall <= 0.01 else 0,
                'expired': expired_this_hour
            })
        
        return failed, hourly, sum(h['expired'] for h in hourly)

    # =========================================================
    # COST FUNCTION
    # =========================================================
    
    def cost(S, W, B):
        y_s = 1 if S > 0 else 0
        y_w = 1 if W > 0 else 0
        y_b = 1 if B > 0 else 0
        
        c_s = (S * capex_linear_s + y_s * capex_constant_s +
               S * opex_lin_recur_s * solar_life + y_s * opex_con_recur_s * solar_life +
               y_s * opex_con_one_s)
        
        c_w = (W * capex_linear_w + y_w * capex_constant_w +
               W * opex_lin_recur_w * wind_life + y_w * opex_con_recur_w * wind_life +
               y_w * opex_con_one_w)
        
        c_b = (B * capex_linear_b + y_b * capex_constant_b +
               B * opex_lin_recur_b * project_years + y_b * opex_con_recur_b * project_years +
               y_b * opex_con_one_b)
        
        return c_s + c_w + c_b

    # =========================================================
    # STAGE 1: GENERATE ALL CONFIGS, SORT BY COST
    # =========================================================
    
    print("\n" + "-" * 40)
    print("STAGE 1: Generate & Sort by Cost")
    print("-" * 40)
    
    max_solar = int(demand_mw * 5)
    max_wind = int(demand_mw * 5)
    max_battery = int(demand_mw * 12)
    
    step_solar = max(50, int(max_solar / 40))
    step_wind = max(50, int(max_wind / 40))
    step_battery = max(100, int(max_battery / 30))
    
    print(f"\nSearch Range:")
    print(f"  Solar:   0 to {max_solar:,} MW (step {step_solar})")
    if wind_enabled:
        print(f"  Wind:    0 to {max_wind:,} MW (step {step_wind})")
    else:
        print(f"  Wind:    DISABLED")
    print(f"  Battery: 0 to {max_battery:,} MWh (step {step_battery})")
    
    all_configs = []
    
    wind_range = range(0, max_wind + 1, step_wind) if wind_enabled else [0]
    
    for S in range(0, max_solar + 1, step_solar):
        for W in wind_range:
            for B in range(0, max_battery + 1, step_battery):
                if S == 0 and W == 0:
                    continue
                c = cost(S, W, B)
                all_configs.append((c, S, W, B))
    
    all_configs.sort()
    
    print(f"\nTotal configurations: {len(all_configs):,}")
    
    stage1_time = time.time() - start_time
    print(f"Time to generate & sort: {stage1_time:.1f}s")

    # =========================================================
    # STAGE 2: FIND CHEAPEST FEASIBLE (FAST SIMULATION)
    # =========================================================
    
    print("\n" + "-" * 40)
    print("STAGE 2: Find Cheapest Feasible (Fast)")
    print("-" * 40)
    
    coarse_best = None
    configs_checked = 0
    
    for c, S, W, B in all_configs:
        configs_checked += 1
        
        fails = simulate_fast(S, W, B)
        
        if fails <= max_fail_hours:
            coarse_best = (c, S, W, B, fails)
            print(f"\nFirst feasible found at config #{configs_checked:,}")
            print(f"  Solar: {S}, Wind: {W}, Battery: {B}")
            print(f"  Cost: ${c/1e6:,.1f}M, Failed Hours: {fails}")
            break
    
    stage2_time = time.time() - start_time - stage1_time
    print(f"\nConfigs checked: {configs_checked:,}")
    print(f"Time: {stage2_time:.1f}s")
    
    if not coarse_best:
        print("\nNO FEASIBLE SOLUTION FOUND")
        if not wind_enabled:
            print("Note: Wind is disabled. Consider enabling wind or lowering uptime target.")
        sht.range('M3').value = "No Solution"        
        sht.range('M5').value = 0
        sht.range('M6').value = 0
        sht.range('M7').value = 0
        return

    # =========================================================
    # STAGE 3: FINE SEARCH (ACCURATE SIMULATION)
    # =========================================================
    
    print("\n" + "-" * 40)
    print("STAGE 3: Fine Search (Accurate Simulation)")
    print("-" * 40)
    
    _, S0, W0, B0, _ = coarse_best
    
    fine_step_solar = 10
    fine_step_wind = 10
    fine_step_battery = 20
    
    fine_configs = []
    
    if wind_enabled:
        fine_wind_range = range(max(0, W0 - step_wind), W0 + step_wind + 1, fine_step_wind)
    else:
        fine_wind_range = [0]
    
    for S in range(max(0, S0 - step_solar), S0 + step_solar + 1, fine_step_solar):
        for W in fine_wind_range:
            for B in range(max(0, B0 - step_battery), B0 + step_battery + 1, fine_step_battery):
                if S == 0 and W == 0:
                    continue
                c = cost(S, W, B)
                fine_configs.append((c, S, W, B))
    
    fine_configs.sort()
    
    print(f"Fine configs to check: {len(fine_configs):,}")
    
    best_cost = float('inf')
    best = None
    best_hourly = None
    best_expired = 0
    fine_checked = 0
    
    for c, S, W, B in fine_configs:
        if c >= best_cost:
            continue
        
        fine_checked += 1
        fails, hourly, expired = simulate(S, W, B)
        
        if fails <= max_fail_hours:
            best_cost = c
            best = (S, W, B, fails)
            best_hourly = hourly
            best_expired = expired
    
    stage3_time = time.time() - start_time - stage1_time - stage2_time
    total_time = time.time() - start_time
    
    print(f"Fine configs checked (accurate): {fine_checked:,}")
    print(f"Time: {stage3_time:.1f}s")

    # =========================================================
    # RESULTS
    # =========================================================
    
    if best:
        S_r, W_r, B_r, fails = best
        B_mw = B_r / batt_duration if batt_duration > 0 else 0
        
        total_gen = sum(h['total_gen'] for h in best_hourly)
        total_del = sum(h['delivered'] for h in best_hourly)
        total_curt = sum(h['curtail'] for h in best_hourly)
        total_short = sum(h['shortfall'] for h in best_hourly)
        total_charge = sum(h['charge'] for h in best_hourly)
        total_discharge = sum(h['discharge'] for h in best_hourly)
        total_expired = sum(h['expired'] for h in best_hourly)
        
        # Calculate total solar and wind generation
        total_solar_gen = sum(h['solar_gen'] for h in best_hourly)
        total_wind_gen = sum(h['wind_gen'] for h in best_hourly)
        
        utilization = total_discharge / total_charge * 100 if total_charge > 0 else 0
        
        max_soc = max(h['soc'] for h in best_hourly)
        
        print("\n" + "=" * 60)
        print("OPTIMAL SOLUTION")
        print("=" * 60)
        
        print(f"\nCapacity:")
        print(f"  Solar:   {S_r:>8,} MWdc")
        if wind_enabled:
            print(f"  Wind:    {W_r:>8,} MWac")
        else:
            print(f"  Wind:    {'DISABLED':>8}")
        print(f"  Battery: {B_r:>8,} MWh ({B_mw:,.0f} MW)")
        
        print(f"\nPerformance:")
        print(f"  Failed Hours: {fails:,} (max: {max_fail_hours:,})")
        print(f"  Uptime:       {(8760-fails)/8760*100:.3f}%")
        
        print(f"\nCost:")
        print(f"  Total: ${best_cost/1e6:,.1f}M")
        
        print(f"\nEnergy (Annual):")
        print(f"  Generated:  {total_gen:>12,.0f} MWh")
        print(f"  Solar:      {total_solar_gen:>12,.0f} MWh")
        print(f"  Wind:       {total_wind_gen:>12,.0f} MWh")
        print(f"  Delivered:  {total_del:>12,.0f} MWh")
        print(f"  Curtailed:  {total_curt:>12,.0f} MWh")
        print(f"  Shortfall:  {total_short:>12,.0f} MWh")
        
        if B_r > 0:
            print(f"\nBattery (Annual):")
            print(f"  Charged:              {total_charge:>12,.0f} MWh")
            print(f"  Discharged:           {total_discharge:>12,.0f} MWh")
            print(f"  Expired (wasted):     {total_expired:>12,.0f} MWh")
            print(f"  Utilization:          {utilization:>11.1f}%")
            print(f"  Max SOC:              {max_soc:>10,.1f} MWh ({max_soc/B_r*100:.1f}% of capacity)")
        else:
            print(f"\nBattery: None (solution uses generation only)")
        
        print(f"\nTotal Time: {total_time:.1f}s")

        # Write hourly data - CHANGED: "Hourly" to "Hourly Output"
        try:
            sht_h = wb.sheets['Hourly Output']
            sht_h.clear()
        except:
            sht_h = wb.sheets.add('Hourly Output')
        
        headers = ['Hour', 'Solar_CF', 'Wind_CF', 'Solar_MW', 'Wind_MW',
                   'Total_Gen', 'Discharge', 'Charge', 'SOC',
                   'Curtail', 'Shortfall', 'Delivered', 'Served', 'Demand', 'Expired']
        
        rows = [[h['hour'], round(h['solar_cf'],4), round(h['wind_cf'],4),
                 round(h['solar_gen'],2), round(h['wind_gen'],2),
                 round(h['total_gen'],2), round(h['discharge'],2),
                 round(h['charge'],2), round(h['soc'],2),
                 round(h['curtail'],2), round(h['shortfall'],2),
                 round(h['delivered'],2), h['served'], demand_mw,
                 round(h['expired'],2)]
                for h in best_hourly]
        
        sht_h.range('A1').value = headers
        sht_h.range('A2').value = rows
        
        sht.range('M3').value = "Optimal"
        sht.range('M5').value = S_r
        sht.range('M6').value = W_r if wind_enabled else 0
        sht.range('M7').value = B_r
        
        # ADDED: Write total solar and wind energy to M10 and M11
        sht.range('M10').value = total_solar_gen
        sht.range('M11').value = total_wind_gen
        
        print("\nDone!")
        print("=" * 60)
        
    else:
        print("\nNO FEASIBLE SOLUTION")
        print("Coarse solution did not pass accurate simulation.")
        if not wind_enabled:
            print("Note: Wind is disabled. Consider enabling wind or lowering uptime target.")
        sht.range('M3').value = "No Solution"
        sht.range('M5').value = 0
        sht.range('M6').value = 0
        sht.range('M7').value = 0


if __name__ == "__main__":
    try:
        run_optimization()
    except Exception as e:
        import traceback
        import sys
        import os

        if getattr(sys, 'frozen', False):
            current_folder = os.path.dirname(sys.executable)
        else:
            current_folder = os.path.dirname(os.path.abspath(__file__))
            
        error_file = os.path.join(current_folder, "error_log.txt")

        with open(error_file, "w") as f:
            f.write("THE ENGINE CRASHED. HERE IS THE REASON:\n\n")
            f.write(traceback.format_exc())
        sys.exit(1) 