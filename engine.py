# import pandas as pd
# import pulp
# import xlwings as xw

# def run_optimization():
#     print("=" * 60)
#     print("HYPERSCALER ENERGY OPTIMIZATION")
#     print("=" * 60)
    
#     # ---------------------------------------------------------
#     # STEP 1: CONNECT TO EXCEL AND GRAB INPUTS
#     # ---------------------------------------------------------
#     wb = xw.Book('Hyperscaler_Model.xlsm')
#     sht = wb.sheets['Dashboard']
#     sht.range('F1').value = "Calculating..."

#     # Core inputs
#     demand_mw = float(sht.range('C3').value)
#     project_years = float(sht.range('C4').value)
#     solar_life = float(sht.range('C5').value)
#     wind_life = float(sht.range('C6').value)
#     target_uptime = float(sht.range('C7').value) / 100
#     batt_duration = float(sht.range('C8').value)
#     ref_solar = float(sht.range('C9').value)
#     ref_wind = float(sht.range('C10').value)
#     solar_ilr = float(sht.range('C11').value)
    
#     # Battery efficiency
#     batt_eff = 0.90

#     # Cost inputs
#     capex_linear_s = float(sht.range('C14').value or 0)
#     capex_linear_w = float(sht.range('C15').value or 0)
#     capex_linear_b = float(sht.range('C16').value or 0)

#     capex_constant_s = float(sht.range('C19').value or 0)
#     capex_constant_w = float(sht.range('C20').value or 0)
#     capex_constant_b = float(sht.range('C21').value or 0)

#     opex_lin_recur_s = float(sht.range('C25').value or 0)
#     opex_lin_recur_w = float(sht.range('C26').value or 0)
#     opex_lin_recur_b = float(sht.range('C27').value or 0)

#     opex_con_recur_s = float(sht.range('C30').value or 0)
#     opex_con_recur_w = float(sht.range('C31').value or 0)
#     opex_con_recur_b = float(sht.range('C32').value or 0)

#     opex_lin_one_s = float(sht.range('C35').value or 0)
#     opex_lin_one_w = float(sht.range('C36').value or 0)
#     opex_lin_one_b = float(sht.range('C37').value or 0)

#     opex_con_one_s = float(sht.range('C40').value or 0)
#     opex_con_one_w = float(sht.range('C41').value or 0)
#     opex_con_one_b = float(sht.range('C42').value or 0)

#     print(f"\nInputs loaded:")
#     print(f"  Demand: {demand_mw} MW")
#     print(f"  Target Uptime: {target_uptime*100}%")
#     print(f"  Battery Duration: {batt_duration} hours")
#     print(f"  Solar ILR: {solar_ilr}")

#     # ---------------------------------------------------------
#     # STEP 2: BUILD 8760 CAPACITY FACTOR PROFILES
#     # ---------------------------------------------------------
#     sht_solar = wb.sheets['EPE_Solar']
#     sht_wind = wb.sheets['EPE_Wind']
    
#     real_solar_24x12 = sht_solar.range('B2:M25').value
#     real_wind_24x12 = sht_wind.range('B2:M25').value

#     days_in_month = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
#     solar_cf = []
#     wind_cf = []
    
#     for month_idx, days in enumerate(days_in_month):
#         for day in range(days):
#             for hour in range(24):
#                 s_cf = real_solar_24x12[hour][month_idx] / ref_solar
#                 w_cf = real_wind_24x12[hour][month_idx] / ref_wind
                
#                 solar_cf.append(max(0.0, min(1.0, s_cf)))
#                 wind_cf.append(max(0.0, min(1.0, w_cf)))

#     print(f"\n8760 profile built:")
#     print(f"  Avg Solar CF: {sum(solar_cf)/len(solar_cf):.3f}")
#     print(f"  Avg Wind CF: {sum(wind_cf)/len(wind_cf):.3f}")

#     # ---------------------------------------------------------
#     # STEP 3: RANK HOURS BY DIFFICULTY
#     # ---------------------------------------------------------
#     # Combined renewable score: higher = easier to serve
#     # We weight solar and wind equally (can adjust based on costs)
    
#     hour_scores = []
#     for t in range(8760):
#         # Combined CF - hours with high combined renewables are "easy"
#         combined_cf = solar_cf[t] + wind_cf[t]
#         hour_scores.append((t, combined_cf))
    
#     # Sort by score ascending (lowest score = hardest hour)
#     hour_scores_sorted = sorted(hour_scores, key=lambda x: x[1])
    
#     # Determine how many hours we're allowed to fail
#     max_fail_hours = int(8760 * (1 - target_uptime))
    
#     # The hardest hours (that we're allowed to skip)
#     hours_allowed_to_fail = set(h[0] for h in hour_scores_sorted[:max_fail_hours])
    
#     # The hours we MUST serve
#     hours_must_serve = [t for t in range(8760) if t not in hours_allowed_to_fail]
    
#     print(f"\nUptime calculation:")
#     print(f"  Target uptime: {target_uptime*100}%")
#     print(f"  Max fail hours: {max_fail_hours}")
#     print(f"  Hours must serve: {len(hours_must_serve)}")
    
#     # Show some example hard hours
#     print(f"\n  Hardest hours (will skip if allowed):")
#     for i in range(min(5, max_fail_hours)):
#         h, score = hour_scores_sorted[i]
#         print(f"    Hour {h}: Solar CF={solar_cf[h]:.3f}, Wind CF={wind_cf[h]:.3f}, Score={score:.3f}")

#     # ---------------------------------------------------------
#     # STEP 4: BUILD OPTIMIZATION MODEL
#     # ---------------------------------------------------------
#     model = pulp.LpProblem("Hyperscaler", pulp.LpMinimize)
    
#     # We only model the hours we must serve
#     T = hours_must_serve
#     T_set = set(T)
    
#     print(f"\nOptimizing over {len(T)} hours (skipping {max_fail_hours} hardest hours)")

#     # ---------------------------------------------------------
#     # DECISION VARIABLES
#     # ---------------------------------------------------------
    
#     # Capacity to build
#     S = pulp.LpVariable("Solar_DC_MW", lowBound=0)
#     W = pulp.LpVariable("Wind_AC_MW", lowBound=0)
#     B_mwh = pulp.LpVariable("Battery_MWh", lowBound=0)
#     B_mw = pulp.LpVariable("Battery_MW", lowBound=0)
    
#     # Binary: do we build this asset at all?
#     y_s = pulp.LpVariable("Build_Solar", cat='Binary')
#     y_w = pulp.LpVariable("Build_Wind", cat='Binary')
#     y_b = pulp.LpVariable("Build_Battery", cat='Binary')
    
#     # Hourly operations (only for hours we must serve)
#     charge = pulp.LpVariable.dicts("Chg", T, lowBound=0)
#     discharge = pulp.LpVariable.dicts("Dis", T, lowBound=0)
#     soc = pulp.LpVariable.dicts("SOC", range(8760), lowBound=0)  # Need SOC for all hours for continuity
#     curtail = pulp.LpVariable.dicts("Curt", T, lowBound=0)

#     # ---------------------------------------------------------
#     # LINK CAPACITY VARIABLES
#     # ---------------------------------------------------------
    
#     BIG_M = 50000
    
#     model += S <= BIG_M * y_s, "Solar_Link"
#     model += W <= BIG_M * y_w, "Wind_Link"
#     model += B_mwh <= BIG_M * y_b, "Batt_MWh_Link"
#     model += B_mw <= BIG_M * y_b, "Batt_MW_Link"
    
#     model += B_mwh == B_mw * batt_duration, "Batt_Duration"

#     # ---------------------------------------------------------
#     # OBJECTIVE: MINIMIZE TOTAL COST
#     # ---------------------------------------------------------
    
#     solar_cost = (
#         S * capex_linear_s +
#         y_s * capex_constant_s +
#         S * opex_lin_recur_s * solar_life +
#         y_s * opex_con_recur_s * solar_life +
#         S * opex_lin_one_s +
#         y_s * opex_con_one_s
#     )
    
#     wind_cost = (
#         W * capex_linear_w +
#         y_w * capex_constant_w +
#         W * opex_lin_recur_w * wind_life +
#         y_w * opex_con_recur_w * wind_life +
#         W * opex_lin_one_w +
#         y_w * opex_con_one_w
#     )
    
#     batt_cost = (
#         B_mwh * capex_linear_b +
#         y_b * capex_constant_b +
#         B_mwh * opex_lin_recur_b * project_years +
#         y_b * opex_con_recur_b * project_years +
#         B_mwh * opex_lin_one_b +
#         y_b * opex_con_one_b
#     )
    
#     model += solar_cost + wind_cost + batt_cost, "Total_Cost"

#     # ---------------------------------------------------------
#     # CONSTRAINTS
#     # ---------------------------------------------------------
    
#     for t in range(8760):
#         # Calculate generation this hour
#         effective_solar_cf = min(solar_cf[t] * solar_ilr, 1.0) / solar_ilr
#         solar_gen = S * effective_solar_cf
#         wind_gen = W * wind_cf[t]
        
#         if t in T_set:
#             # THIS HOUR MUST BE SERVED
            
#             # Power balance: Gen + Discharge = Demand + Charge + Curtail
#             model += (
#                 solar_gen +
#                 wind_gen +
#                 discharge[t] * batt_eff
#                 ==
#                 demand_mw +
#                 charge[t] +
#                 curtail[t]
#             ), f"Balance_{t}"
            
#             # Charge/discharge limits
#             model += charge[t] <= B_mw, f"Chg_Limit_{t}"
#             model += discharge[t] <= B_mw, f"Dis_Limit_{t}"
            
#             # SOC dynamics
#             if t == 0:
#                 model += soc[t] == 0.5 * B_mwh + charge[t] * batt_eff - discharge[t], f"SOC_{t}"
#                 model += discharge[t] <= 0.5 * B_mwh, f"Dis_SOC_{t}"
#             else:
#                 model += soc[t] == soc[t-1] + charge[t] * batt_eff - discharge[t], f"SOC_{t}"
#                 model += discharge[t] <= soc[t-1], f"Dis_SOC_{t}"
        
#         else:
#             # THIS HOUR CAN BE SKIPPED
#             # Battery can still charge if there's excess generation, but no requirement to meet demand
            
#             # Excess generation can charge battery (simplified: assume we capture some)
#             # For skipped hours, we just track SOC passively
#             if t == 0:
#                 # First hour skipped - just set initial SOC
#                 model += soc[t] == 0.5 * B_mwh, f"SOC_Skip_{t}"
#             else:
#                 # During skipped hours, assume battery holds steady (no charge/discharge requirement)
#                 # But allow opportunistic charging if generation > 0
#                 # Simplified: SOC stays same as previous
#                 model += soc[t] == soc[t-1], f"SOC_Skip_{t}"
        
#         # SOC bounds (all hours)
#         model += soc[t] <= B_mwh, f"SOC_Max_{t}"
#         model += soc[t] >= 0, f"SOC_Min_{t}"

#     # End SOC constraint
#     model += soc[8759] >= 0.4 * B_mwh, "SOC_End_Min"
#     model += soc[8759] <= 0.6 * B_mwh, "SOC_End_Max"

#     # ---------------------------------------------------------
#     # SOLVE
#     # ---------------------------------------------------------
#     print("\nSolving optimization...")
    
#     solver = pulp.PULP_CBC_CMD(
#         timeLimit=120,
#         gapRel=0.01,
#         msg=True,
#         threads=4
#     )
    
#     status = model.solve(solver)
#     status_str = pulp.LpStatus[status]
    
#     print(f"\nSolver Status: {status_str}")

#     # ---------------------------------------------------------
#     # EXTRACT AND DISPLAY RESULTS
#     # ---------------------------------------------------------
    
#     if status in [pulp.LpStatusOptimal, pulp.LpStatusNotSolved]:
#         solar_result = S.varValue or 0
#         wind_result = W.varValue or 0
#         batt_mwh_result = B_mwh.varValue or 0
#         batt_mw_result = B_mw.varValue or 0
        
#         # Calculate total curtailment
#         total_curtail = sum(curtail[t].varValue or 0 for t in T)
        
#         print("\n" + "=" * 60)
#         print("RESULTS")
#         print("=" * 60)
#         print(f"\n🔧 CAPACITY TO BUILD:")
#         print(f"   Solar:   {solar_result:>10,.2f} MWdc")
#         print(f"   Wind:    {wind_result:>10,.2f} MWac")
#         print(f"   Battery: {batt_mwh_result:>10,.2f} MWh ({batt_mw_result:,.2f} MW)")
        
#         print(f"\n📊 PERFORMANCE:")
#         print(f"   Target Uptime:   {target_uptime*100:>10.2f}%")
#         print(f"   Hours Served:    {len(T):>10,}")
#         print(f"   Hours Skipped:   {max_fail_hours:>10,}")
#         print(f"   Curtailed:       {total_curtail:>10,.0f} MWh")
        
#         # Sanity checks
#         print(f"\n🔍 SANITY CHECK:")
#         print(f"   Demand:          {demand_mw:>10,.0f} MW")
#         solar_ac = solar_result / solar_ilr
#         total_firm = solar_ac + wind_result + batt_mw_result
#         print(f"   Solar AC:        {solar_ac:>10,.0f} MW")
#         print(f"   Max Firm Power:  {total_firm:>10,.0f} MW (Solar AC + Wind + Batt)")
        
#         # Check if we have enough to meet demand at worst served hour
#         worst_served_hour = min(T, key=lambda t: solar_cf[t] + wind_cf[t])
#         worst_solar = solar_result * min(solar_cf[worst_served_hour] * solar_ilr, 1.0) / solar_ilr
#         worst_wind = wind_result * wind_cf[worst_served_hour]
#         worst_gen = worst_solar + worst_wind
#         print(f"\n   Worst served hour: {worst_served_hour}")
#         print(f"   Solar CF: {solar_cf[worst_served_hour]:.3f}, Wind CF: {wind_cf[worst_served_hour]:.3f}")
#         print(f"   Generation: {worst_gen:,.0f} MW, Gap filled by battery: {demand_mw - worst_gen:,.0f} MW")
        
#         print("=" * 60)
        
#         # Write to Excel
#         sht.range('F1').value = status_str
#         sht.range('F3').value = round(solar_result, 2)
#         sht.range('F4').value = round(wind_result, 2)
#         sht.range('F5').value = round(batt_mwh_result, 2)
        
#         # ---------------------------------------------------------
#         # HOURLY OUTPUT
#         # ---------------------------------------------------------
#         print("\nExporting hourly data...")
        
#         hourly_data = []
#         for t in range(8760):
#             effective_solar_cf = min(solar_cf[t] * solar_ilr, 1.0) / solar_ilr
#             is_served = t in T_set
            
#             row = {
#                 'Hour': t,
#                 'Demand_MW': demand_mw,
#                 'Solar_CF': solar_cf[t],
#                 'Wind_CF': wind_cf[t],
#                 'Solar_Gen_MW': solar_result * effective_solar_cf,
#                 'Wind_Gen_MW': wind_result * wind_cf[t],
#                 'SOC_MWh': soc[t].varValue or 0,
#                 'Hour_Served': 1 if is_served else 0
#             }
            
#             if is_served:
#                 row['Charge_MW'] = charge[t].varValue or 0
#                 row['Discharge_MW'] = discharge[t].varValue or 0
#                 row['Curtail_MW'] = curtail[t].varValue or 0
#             else:
#                 row['Charge_MW'] = 0
#                 row['Discharge_MW'] = 0
#                 row['Curtail_MW'] = 0
            
#             hourly_data.append(row)
        
#         df_out = pd.DataFrame(hourly_data)
        
#         try:
#             sht_out = wb.sheets['Hourly_Output']
#             sht_out.clear_contents()
#         except:
#             sht_out = wb.sheets.add('Hourly_Output')
        
#         sht_out.range('A1').value = df_out
#         print("Hourly output exported.")
        
#     else:
#         print(f"\n❌ Optimization FAILED: {status_str}")
#         sht.range('F1').value = status_str
#         sht.range('F3').value = "ERROR"
#         sht.range('F4').value = "ERROR"
#         sht.range('F5').value = "ERROR"

#     print("\nDone!")
#     return model


# if __name__ == "__main__":
#     run_optimization()

import xlwings as xw
from pathlib import Path

def run_optimization():
    print("=" * 60)
    print("HYPERSCALER CAPACITY OPTIMIZATION")
    print("=" * 60)
    
    current_dir = Path(__file__).parent
    outer_file_path = current_dir.parent / "Hyperscaler_Model.xlsm"

    # wb = xw.Book('../Hyperscaler_Model.xlsm')
    wb = xw.Book(outer_file_path)
    sht = wb.sheets['Dashboard']
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
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
    input("\nPress ENTER...")