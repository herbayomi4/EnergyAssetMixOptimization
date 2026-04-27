import xlwings as xw

def run_optimization():
    print("=" * 60)
    print("HYPERSCALER CAPACITY OPTIMIZATION")
    print("=" * 60)
    
    wb = xw.Book('Optimization & Financial Model.xlsm')
    sht = wb.sheets['Optimization Dashboard']
    sht.range('F1').value = "Running..."

    # Load inputs
    demand_mw = float(sht.range('C3').value)
    project_years = float(sht.range('C4').value)
    solar_life = float(sht.range('C5').value)
    wind_life = float(sht.range('C6').value)
    target_uptime = float(sht.range('C7').value) / 100
    batt_duration = float(sht.range('C8').value)
    ref_solar = float(sht.range('C9').value)
    ref_wind = float(sht.range('C10').value)
    solar_ilr = float(sht.range('C11').value)
    batt_eff_rt = float(sht.range('C12').value) / 100

    capex_linear_s = float(sht.range('C15').value or 0)
    capex_linear_w = float(sht.range('C16').value or 0)
    capex_linear_b = float(sht.range('C17').value or 0)
    capex_constant_s = float(sht.range('C20').value or 0)
    capex_constant_w = float(sht.range('C21').value or 0)
    capex_constant_b = float(sht.range('C22').value or 0)
    opex_lin_recur_s = float(sht.range('C26').value or 0)
    opex_lin_recur_w = float(sht.range('C27').value or 0)
    opex_lin_recur_b = float(sht.range('C28').value or 0)
    opex_con_recur_s = float(sht.range('C31').value or 0)
    opex_con_recur_w = float(sht.range('C32').value or 0)
    opex_con_recur_b = float(sht.range('C33').value or 0)
    opex_lin_one_s = float(sht.range('C36').value or 0)
    opex_lin_one_w = float(sht.range('C37').value or 0)
    opex_lin_one_b = float(sht.range('C38').value or 0)
    opex_con_one_s = float(sht.range('C41').value or 0)
    opex_con_one_w = float(sht.range('C42').value or 0)
    opex_con_one_b = float(sht.range('C43').value or 0)

    max_fail_hours = int(8760 * (1 - target_uptime))

    print(f"\nDemand: {demand_mw} MW")
    print(f"Target Uptime: {target_uptime*100:.3f}%")
    print(f"Max Failed Hours: {max_fail_hours}")

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

    # Simulation
    def simulate(S_mw, W_mw, B_mwh):
        B_mw = B_mwh / batt_duration if batt_duration > 0 else 0
        soc = 0.0
        failed = 0
        hourly = []
        
        for t in range(8760):
            s_gen = S_mw * solar_cf_capped[t]
            w_gen = W_mw * wind_cf[t]
            gen = s_gen + w_gen
            
            charge, discharge, curtail, shortfall = 0.0, 0.0, 0.0, 0.0
            
            if gen >= demand_mw:
                delivered = demand_mw
                excess = gen - demand_mw
                charge = min(excess, B_mw, B_mwh - soc)
                curtail = excess - charge
                soc += charge
            else:
                deficit = demand_mw - gen
                max_dis = min(B_mw, soc)
                needed = deficit / batt_eff_rt
                discharge = min(max_dis, needed)
                soc -= discharge
                delivered = gen + discharge * batt_eff_rt
                shortfall = demand_mw - delivered
                if shortfall > 0.01:
                    failed += 1
            
            hourly.append({
                'hour': t, 'solar_cf': solar_cf[t], 'wind_cf': wind_cf[t],
                'solar_gen': s_gen, 'wind_gen': w_gen, 'total_gen': gen,
                'charge': charge, 'discharge': discharge * batt_eff_rt,
                'soc': soc, 'curtail': curtail, 'shortfall': shortfall,
                'delivered': delivered, 'served': 1 if shortfall <= 0.01 else 0
            })
        
        return failed, hourly

    # Cost
    def cost(S, W, B):
        y_s, y_w, y_b = (1 if S > 0 else 0), (1 if W > 0 else 0), (1 if B > 0 else 0)
        c_s = (S * capex_linear_s + y_s * capex_constant_s +
               S * opex_lin_recur_s * solar_life + y_s * opex_con_recur_s * solar_life +
               S * opex_lin_one_s + y_s * opex_con_one_s)
        c_w = (W * capex_linear_w + y_w * capex_constant_w +
               W * opex_lin_recur_w * wind_life + y_w * opex_con_recur_w * wind_life +
               W * opex_lin_one_w + y_w * opex_con_one_w)
        c_b = (B * capex_linear_b + y_b * capex_constant_b +
               B * opex_lin_recur_b * project_years + y_b * opex_con_recur_b * project_years +
               B * opex_lin_one_b + y_b * opex_con_one_b)
        return c_s + c_w + c_b

    # Search
    print("\nSearching...")
    
    best_cost = float('inf')
    best = None
    best_hourly = None

    for S in range(0, 1501, 50):
        for W in range(0, 1501, 50):
            for B in range(0, 2001, 100):
                if S == 0 and W == 0:
                    continue
                c = cost(S, W, B)
                if c >= best_cost:
                    continue
                fails, hourly = simulate(S, W, B)
                if fails <= max_fail_hours:
                    best_cost = c
                    best = (S, W, B, fails)
                    best_hourly = hourly

    if best:
        S0, W0, B0, _ = best
        for S in range(max(0, S0-50), S0+60, 10):
            for W in range(max(0, W0-50), W0+60, 10):
                for B in range(max(0, B0-100), B0+110, 10):
                    if S == 0 and W == 0:
                        continue
                    c = cost(S, W, B)
                    if c >= best_cost:
                        continue
                    fails, hourly = simulate(S, W, B)
                    if fails <= max_fail_hours:
                        best_cost = c
                        best = (S, W, B, fails)
                        best_hourly = hourly

    # Results
    if best:
        S_r, W_r, B_r, fails = best
        B_mw = B_r / batt_duration if batt_duration > 0 else 0
        
        total_gen = sum(h['total_gen'] for h in best_hourly)
        total_del = sum(h['delivered'] for h in best_hourly)
        total_curt = sum(h['curtail'] for h in best_hourly)
        total_short = sum(h['shortfall'] for h in best_hourly)
        total_demand = demand_mw * 8760
        
        print("\n" + "=" * 60)
        print("OPTIMAL SOLUTION")
        print("=" * 60)
        
        print(f"\nCapacity:")
        print(f"  Solar:   {S_r:>8,} MWdc")
        print(f"  Wind:    {W_r:>8,} MWac")
        print(f"  Battery: {B_r:>8,} MWh ({B_mw:,.0f} MW)")
        
        print(f"\nPerformance:")
        print(f"  Failed Hours: {fails:,} (max: {max_fail_hours:,})")
        print(f"  Uptime:       {(8760-fails)/8760*100:.3f}%")
        
        print(f"\nCost:")
        print(f"  Total: ${best_cost/1e6:,.1f}M")
        
        print(f"\nEnergy (Annual):")
        print(f"  Delivered:  {total_del:>12,.0f} MWh")
        print(f"  Curtailed:  {total_curt:>12,.0f} MWh")
        print(f"  Shortfall:  {total_short:>12,.0f} MWh")

        # Write hourly
        try:
            sht_h = wb.sheets['Hourly']
            sht_h.clear()
        except:
            sht_h = wb.sheets.add('Hourly')
        
        headers = ['Hour', 'Solar_CF', 'Wind_CF', 'Solar_MW', 'Wind_MW',
                   'Total_Gen', 'Discharge', 'Charge', 'SOC',
                   'Curtail', 'Shortfall', 'Delivered', 'Served', 'Demand']
        
        rows = [[h['hour'], round(h['solar_cf'],4), round(h['wind_cf'],4),
                 round(h['solar_gen'],2), round(h['wind_gen'],2),
                 round(h['total_gen'],2), round(h['discharge'],2),
                 round(h['charge'],2), round(h['soc'],2),
                 round(h['curtail'],2), round(h['shortfall'],2),
                 round(h['delivered'],2), h['served'], demand_mw]
                for h in best_hourly]
        
        sht_h.range('A1').value = headers
        sht_h.range('A2').value = rows
        
        sht.range('F1').value = "Optimal"
        sht.range('F3').value = S_r
        sht.range('F4').value = W_r
        sht.range('F5').value = B_r
        
        print("\nDone!")
        print("=" * 60)
        
    else:
        print("\nNO FEASIBLE SOLUTION")
        sht.range('F1').value = "No Solution"


if __name__ == "__main__":
    try:
        run_optimization()
    except Exception as e:
        import traceback

        # 1. Figure out exactly where the .exe is currently sitting
        if getattr(sys, 'frozen', False):
            current_folder = os.path.dirname(sys.executable)
        else:
            current_folder = os.path.dirname(os.path.abspath(__file__))
            
        # 2. Build the exact file path for the error log
        error_file = os.path.join(current_folder, "error_log.txt")

        with open(error_file, "w") as f:
            f.write("THE ENGINE CRASHED. HERE IS THE REASON:\n\n")
            f.write(traceback.format_exc())
        sys.exit(1)