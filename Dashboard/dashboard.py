import pandas as pd
import streamlit as st
import plotly.express as px
from streamlit_option_menu import option_menu
import calendar
from datetime import datetime, timedelta
import plotly.graph_objects as go
from plotly.subplots import make_subplots

today = datetime.today()
yesterday = today - timedelta(days=1)
day_before_yesterday = today - timedelta(days=2)


def get_filtered_table(xls, active_sheet, df, start_date, end_date, selected_cluster):
    result_df = df.copy()
    
    start_year = start_date.year
    end_year = end_date.year
    selected_years = list(range(start_year, end_year + 1))
    
    result_df["kWh"] = None
    result_df["Specific Production"] = None
    result_df["PPR"] = None
    
    all_kwh_sheets = [str(sheet) for sheet in xls.sheet_names if 'kWh' in str(sheet) and 'Target' not in str(sheet)]
    
    sheet_years = {}
    for sheet in all_kwh_sheets:
        digits = ''.join(filter(str.isdigit, sheet))
        if len(digits) >= 4:
            for i in range(len(digits) - 3):
                potential_year = int(digits[i:i+4])
                if 2000 <= potential_year <= 3000: 
                    sheet_years[sheet] = potential_year
                    break
    
    kwh_sheets = []
    for sheet in all_kwh_sheets:
        if sheet in sheet_years:
            if sheet_years[sheet] in selected_years:
                kwh_sheets.append(sheet)
        else:
            kwh_sheets.append(sheet)
    
    if kwh_sheets:
        site_kwh_totals = {}
        
        for sheet in kwh_sheets:
            try:
                raw_kwh_df = pd.read_excel(xls, sheet_name=sheet, header=0)
                raw_kwh_df.columns = raw_kwh_df.columns.str.strip()
                
                smip_rows = raw_kwh_df.apply(lambda row: any('SMIP' in str(cell) and 'OUTSIDE' not in str(cell) 
                                                        for cell in row), axis=1)
                outside_smip_rows = raw_kwh_df.apply(lambda row: any('OUTSIDE SMIP' in str(cell) 
                                                                for cell in row), axis=1)
                
                if any(smip_rows) and any(outside_smip_rows):
                    smip_indices = raw_kwh_df[smip_rows].index
                    outside_smip_indices = raw_kwh_df[outside_smip_rows].index
                    
                    if active_sheet == "SMIP Database":
                        start_idx = min(smip_indices)
                        end_idx = min(outside_smip_indices)
                        raw_kwh_df = raw_kwh_df.iloc[start_idx:end_idx].reset_index(drop=True)
                    elif active_sheet == "Outside-SMIP Database":
                        start_idx = min(outside_smip_indices)
                        next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
                        end_idx = min(next_smip_indices) if next_smip_indices else len(raw_kwh_df)
                        raw_kwh_df = raw_kwh_df.iloc[start_idx:end_idx].reset_index(drop=True)
                
                sheet_year = sheet_years.get(sheet, None)
                
                kwh_date_columns = []
                for date in pd.date_range(start=start_date, end=end_date, freq='D'):
                    if sheet_year is not None and date.year != sheet_year:
                        continue
                        
                    date_col = date.strftime('%b-%d')
                    matching_columns = [col for col in raw_kwh_df.columns.astype(str) if date_col in col]
                    kwh_date_columns.extend(matching_columns)
                
                if kwh_date_columns and "Site Name" in raw_kwh_df.columns:
                    kwh_site_indices = {}
                    for idx, site_name in enumerate(raw_kwh_df["Site Name"]):
                        if pd.notna(site_name):
                            if site_name not in kwh_site_indices:
                                kwh_site_indices[site_name] = []
                            kwh_site_indices[site_name].append(idx)
                    
                    for idx, row in result_df.iterrows():
                        site_name = row["Site Name"]
                        if site_name in kwh_site_indices:
                            matched = False
                            for kwh_idx in kwh_site_indices[site_name]:
                                match_found = True
                                for col in ["Inverter Name", "Inverter SN"]:
                                    if col in result_df.columns and col in raw_kwh_df.columns:
                                        if row[col] != raw_kwh_df.iloc[kwh_idx][col]:
                                            match_found = False
                                            break
                                
                                if match_found:
                                    kwh_sum = 0
                                    for date_col in kwh_date_columns:
                                        if date_col in raw_kwh_df.columns:
                                            val = raw_kwh_df.iloc[kwh_idx][date_col]
                                            if pd.notna(val):
                                                try:
                                                    kwh_sum += float(val)
                                                except (ValueError, TypeError):
                                                    pass
                                    
                                    site_key = (site_name, row.get("Inverter Name", ""), row.get("Inverter SN", ""))
                                    if site_key not in site_kwh_totals:
                                        site_kwh_totals[site_key] = 0
                                    site_kwh_totals[site_key] += kwh_sum
                                    
                                    matched = True
                                    break
                            
                            if not matched and kwh_site_indices[site_name]:
                                kwh_idx = kwh_site_indices[site_name][0]
                                kwh_sum = 0
                                for date_col in kwh_date_columns:
                                    if date_col in raw_kwh_df.columns:
                                        val = raw_kwh_df.iloc[kwh_idx][date_col]
                                        if pd.notna(val):
                                            try:
                                                kwh_sum += float(val)
                                            except (ValueError, TypeError):
                                                pass
                                
                                site_key = (site_name, row.get("Inverter Name", ""), row.get("Inverter SN", ""))
                                if site_key not in site_kwh_totals:
                                    site_kwh_totals[site_key] = 0
                                site_kwh_totals[site_key] += kwh_sum
            
            except Exception as e:
                print(f"Error processing kWh sheet '{sheet}': {str(e)}")
        
        for idx, row in result_df.iterrows():
            site_key = (row["Site Name"], row.get("Inverter Name", ""), row.get("Inverter SN", ""))
            if site_key in site_kwh_totals:
                result_df.at[idx, "kWh"] = round(site_kwh_totals[site_key], 2)
    
    all_sp_sheets = [str(sheet) for sheet in xls.sheet_names if 'SP' in str(sheet) and 'Target' not in str(sheet)]
    
    sheet_years = {}
    for sheet in all_sp_sheets:
        digits = ''.join(filter(str.isdigit, sheet))
        if len(digits) >= 4:
            for i in range(len(digits) - 3):
                potential_year = int(digits[i:i+4])
                if 2000 <= potential_year <= 3000: 
                    sheet_years[sheet] = potential_year
                    break
    
    sp_sheets = []
    for sheet in all_sp_sheets:
        if sheet in sheet_years:
            if sheet_years[sheet] in selected_years:
                sp_sheets.append(sheet)
        else:
            sp_sheets.append(sheet) 
    
    if sp_sheets:
        site_sp_values = {}
        site_sp_counts = {}
        
        for sheet in sp_sheets:
            try:
                raw_sp_df = pd.read_excel(xls, sheet_name=sheet, header=0)
                raw_sp_df.columns = raw_sp_df.columns.str.strip()
                
                smip_rows = raw_sp_df.apply(lambda row: any('SMIP' in str(cell) and 'OUTSIDE' not in str(cell) 
                                                        for cell in row), axis=1)
                outside_smip_rows = raw_sp_df.apply(lambda row: any('OUTSIDE SMIP' in str(cell) 
                                                                for cell in row), axis=1)
                
                if any(smip_rows) and any(outside_smip_rows):
                    smip_indices = raw_sp_df[smip_rows].index
                    outside_smip_indices = raw_sp_df[outside_smip_rows].index
                    
                    if active_sheet == "SMIP Database":
                        start_idx = min(smip_indices)
                        end_idx = min(outside_smip_indices)
                        raw_sp_df = raw_sp_df.iloc[start_idx:end_idx].reset_index(drop=True)
                    elif active_sheet == "Outside-SMIP Database":
                        start_idx = min(outside_smip_indices)
                        next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
                        end_idx = min(next_smip_indices) if next_smip_indices else len(raw_sp_df)
                        raw_sp_df = raw_sp_df.iloc[start_idx:end_idx].reset_index(drop=True)
                
                sheet_year = sheet_years.get(sheet, None)
                
                sp_date_columns = []
                for date in pd.date_range(start=start_date, end=end_date, freq='D'):
                    if sheet_year is not None and date.year != sheet_year:
                        continue
                        
                    date_col = date.strftime('%b-%d')
                    matching_columns = [col for col in raw_sp_df.columns.astype(str) if date_col in col]
                    sp_date_columns.extend(matching_columns)
                
                if sp_date_columns and "Site Name" in raw_sp_df.columns:
                    sp_site_indices = {}
                    for idx, site_name in enumerate(raw_sp_df["Site Name"]):
                        if pd.notna(site_name):
                            if site_name not in sp_site_indices:
                                sp_site_indices[site_name] = []
                            sp_site_indices[site_name].append(idx)
                    
                    for idx, row in result_df.iterrows():
                        site_name = row["Site Name"]
                        if site_name in sp_site_indices:
                            matched = False
                            for sp_idx in sp_site_indices[site_name]:
                                match_found = True
                                for col in ["Inverter Name", "Inverter SN"]:
                                    if col in result_df.columns and col in raw_sp_df.columns:
                                        if row[col] != raw_sp_df.iloc[sp_idx][col]:
                                            match_found = False
                                            break
                                
                                if match_found:
                                    valid_values = []
                                    for date_col in sp_date_columns:
                                        if date_col in raw_sp_df.columns:
                                            val = raw_sp_df.iloc[sp_idx][date_col]
                                            if pd.notna(val):
                                                try:
                                                    valid_values.append(float(val))
                                                except (ValueError, TypeError):
                                                    pass
                                    
                                    site_key = (site_name, row.get("Inverter Name", ""), row.get("Inverter SN", ""))
                                    if site_key not in site_sp_values:
                                        site_sp_values[site_key] = 0
                                        site_sp_counts[site_key] = 0
                                    
                                    if valid_values:
                                        site_sp_values[site_key] += sum(valid_values)
                                        site_sp_counts[site_key] += len(valid_values)
                                    
                                    matched = True
                                    break
                            
                            if not matched and sp_site_indices[site_name]:
                                sp_idx = sp_site_indices[site_name][0]
                                valid_values = []
                                for date_col in sp_date_columns:
                                    if date_col in raw_sp_df.columns:
                                        val = raw_sp_df.iloc[sp_idx][date_col]
                                        if pd.notna(val):
                                            try:
                                                valid_values.append(float(val))
                                            except (ValueError, TypeError):
                                                pass
                                
                                site_key = (site_name, row.get("Inverter Name", ""), row.get("Inverter SN", ""))
                                if site_key not in site_sp_values:
                                    site_sp_values[site_key] = 0
                                    site_sp_counts[site_key] = 0
                                
                                if valid_values:
                                    site_sp_values[site_key] += sum(valid_values)
                                    site_sp_counts[site_key] += len(valid_values)
            
            except Exception as e:
                print(f"Error processing SP sheet '{sheet}': {str(e)}")
        
        for idx, row in result_df.iterrows():
            site_key = (row["Site Name"], row.get("Inverter Name", ""), row.get("Inverter SN", ""))
            if site_key in site_sp_values and site_sp_counts[site_key] > 0:
                result_df.at[idx, "Specific Production"] = round(site_sp_values[site_key] / site_sp_counts[site_key], 2)
    
    all_ppr_sheets = [str(sheet) for sheet in xls.sheet_names if 'PPR' in str(sheet) and 'Target' not in str(sheet)]
    
    sheet_years = {}
    for sheet in all_ppr_sheets:
        digits = ''.join(filter(str.isdigit, sheet))
        if len(digits) >= 4:
            for i in range(len(digits) - 3):
                potential_year = int(digits[i:i+4])
                if 2000 <= potential_year <= 3000: 
                    sheet_years[sheet] = potential_year
                    break
    
    ppr_sheets = []
    for sheet in all_ppr_sheets:
        if sheet in sheet_years:
            if sheet_years[sheet] in selected_years:
                ppr_sheets.append(sheet)
        else:
            ppr_sheets.append(sheet) 
    
    if ppr_sheets:
        site_ppr_values = {}
        site_ppr_counts = {}
        
        for sheet in ppr_sheets:
            try:
                raw_ppr_df = pd.read_excel(xls, sheet_name=sheet, header=0)
                raw_ppr_df.columns = raw_ppr_df.columns.str.strip()
                
                smip_rows = raw_ppr_df.apply(lambda row: any('SMIP per SITE' in str(cell) and 'OUTSIDE' not in str(cell) 
                                                        for cell in row), axis=1)
                outside_smip_rows = raw_ppr_df.apply(lambda row: any('OUTSIDE SMIP per SITE' in str(cell) 
                                                                for cell in row), axis=1)
                
                if any(smip_rows) and any(outside_smip_rows):
                    smip_indices = raw_ppr_df[smip_rows].index
                    outside_smip_indices = raw_ppr_df[outside_smip_rows].index
                    
                    if active_sheet == "SMIP Database":
                        start_idx = min(smip_indices)
                        end_idx = min(outside_smip_indices)
                        raw_ppr_df = raw_ppr_df.iloc[start_idx:end_idx].reset_index(drop=True)
                    elif active_sheet == "Outside-SMIP Database":
                        start_idx = min(outside_smip_indices)
                        next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
                        end_idx = min(next_smip_indices) if next_smip_indices else len(raw_ppr_df)
                        raw_ppr_df = raw_ppr_df.iloc[start_idx:end_idx].reset_index(drop=True)
                
                sheet_year = sheet_years.get(sheet, None)
                
                ppr_date_columns = []
                for date in pd.date_range(start=start_date, end=end_date, freq='D'):
                    if sheet_year is not None and date.year != sheet_year:
                        continue
                        
                    date_col = date.strftime('%b-%d')
                    matching_columns = [col for col in raw_ppr_df.columns.astype(str) if date_col in col]
                    ppr_date_columns.extend(matching_columns)
                
                if ppr_date_columns and "Site Name" in raw_ppr_df.columns:
                    ppr_site_indices = {}
                    for idx, site_name in enumerate(raw_ppr_df["Site Name"]):
                        if pd.notna(site_name):
                            if site_name not in ppr_site_indices:
                                ppr_site_indices[site_name] = []
                            ppr_site_indices[site_name].append(idx)
                    
                    for idx, row in result_df.iterrows():
                        site_name = row["Site Name"]
                        if site_name in ppr_site_indices:
                            matched = False
                            for ppr_idx in ppr_site_indices[site_name]:
                                match_found = True
                                for col in ["Inverter Name", "Inverter SN"]:
                                    if col in result_df.columns and col in raw_ppr_df.columns:
                                        if row[col] != raw_ppr_df.iloc[ppr_idx][col]:
                                            match_found = False
                                            break
                                
                                if match_found:
                                    valid_values = []
                                    for date_col in ppr_date_columns:
                                        if date_col in raw_ppr_df.columns:
                                            val = raw_ppr_df.iloc[ppr_idx][date_col]
                                            if pd.notna(val):
                                                try:
                                                    valid_values.append(float(val))
                                                except (ValueError, TypeError):
                                                    pass
                                    
                                    site_key = (site_name, row.get("Inverter Name", ""), row.get("Inverter SN", ""))
                                    if site_key not in site_ppr_values:
                                        site_ppr_values[site_key] = 0
                                        site_ppr_counts[site_key] = 0
                                    
                                    if valid_values:
                                        site_ppr_values[site_key] += sum(valid_values)
                                        site_ppr_counts[site_key] += len(valid_values)
                                    
                                    matched = True
                                    break
                            
                            if not matched and ppr_site_indices[site_name]:
                                ppr_idx = ppr_site_indices[site_name][0]
                                valid_values = []
                                for date_col in ppr_date_columns:
                                    if date_col in raw_ppr_df.columns:
                                        val = raw_ppr_df.iloc[ppr_idx][date_col]
                                        if pd.notna(val):
                                            try:
                                                valid_values.append(float(val))
                                            except (ValueError, TypeError):
                                                pass
                                
                                site_key = (site_name, row.get("Inverter Name", ""), row.get("Inverter SN", ""))
                                if site_key not in site_ppr_values:
                                    site_ppr_values[site_key] = 0
                                    site_ppr_counts[site_key] = 0
                                
                                if valid_values:
                                    site_ppr_values[site_key] += sum(valid_values)
                                    site_ppr_counts[site_key] += len(valid_values)
            
            except Exception as e:
                print(f"Error processing PPR sheet '{sheet}': {str(e)}")
        
        for idx, row in result_df.iterrows():
            site_key = (row["Site Name"], row.get("Inverter Name", ""), row.get("Inverter SN", ""))
            if site_key in site_ppr_values and site_ppr_counts[site_key] > 0:
                result_df.at[idx, "PPR"] = round(site_ppr_values[site_key] / site_ppr_counts[site_key] * 100, 2)
    
    if "kWp" not in result_df.columns and "kWp (DC)" in result_df.columns:
        result_df.rename(columns={"kWp (DC)": "kWp"}, inplace=True)
    
    if selected_cluster and "Cluster" in result_df.columns:
        result_df = result_df[result_df["Cluster"] == selected_cluster]
    
    required_columns = ["Site Name", "kWh", "Specific Production", "PPR"]
    for col in required_columns:
        if col not in result_df.columns:
            result_df[col] = None
    
    desired_order = ["Site Name", "kWp", "Inverter Name", "Inverter SN", "kWh", "Specific Production", "PPR"]
    available_columns = [col for col in desired_order if col in result_df.columns]
    result_df = result_df[available_columns].copy()
    
    if "PPR" in result_df.columns and not result_df["PPR"].isna().all():
        result_df = result_df.sort_values(by="PPR", ascending=False)
    else:
        result_df = result_df.sort_values(by="Site Name")
    
    result_df.reset_index(drop=True, inplace=True)
    return result_df

def get_daily_kwh(xls, active_sheet, df, start_date, end_date):
    cluster_kwh = []
    
    start_year = start_date.year
    end_year = end_date.year
    selected_years = list(range(start_year, end_year + 1))
    
    all_kwh_sheets = [str(sheet) for sheet in xls.sheet_names if 'kWh' in str(sheet) and 'Target' not in str(sheet)]
    
    sheet_years = {}
    for sheet in all_kwh_sheets:
        digits = ''.join(filter(str.isdigit, sheet))
        if len(digits) >= 4:
            for i in range(len(digits) - 3):
                potential_year = int(digits[i:i+4])
                if 2000 <= potential_year <= 3000: 
                    sheet_years[sheet] = potential_year
                    break
    
    kwh_sheets = []
    for sheet in all_kwh_sheets:
        if sheet in sheet_years:
            if sheet_years[sheet] in selected_years:
                kwh_sheets.append(sheet)
        else:
            kwh_sheets.append(sheet)
    
    if not kwh_sheets:
        st.warning(f"No kWh sheets found for years {selected_years}.")
        return pd.DataFrame(columns=["Date", "kWh", "Cluster"])
    
    for sheet in kwh_sheets:
        try:
            df_kwh = pd.read_excel(xls, sheet_name=sheet, header=0)
            df_kwh.columns = df_kwh.columns.str.strip()
            
            smip_rows = df_kwh.apply(lambda row: any('SMIP' in str(cell) and 'OUTSIDE' not in str(cell) for cell in row), axis=1)
            outside_smip_rows = df_kwh.apply(lambda row: any('OUTSIDE SMIP' in str(cell) for cell in row), axis=1)
            
            if any(smip_rows) and any(outside_smip_rows):
                smip_indices = df_kwh[smip_rows].index
                outside_smip_indices = df_kwh[outside_smip_rows].index
                
                if active_sheet == "SMIP Database":
                    start_idx = min(smip_indices)
                    end_idx = min(outside_smip_indices)
                else:
                    start_idx = min(outside_smip_indices)
                    next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
                    end_idx = min(next_smip_indices) if next_smip_indices else len(df_kwh)
                
                df_kwh = df_kwh.iloc[start_idx:end_idx].reset_index(drop=True)
            
            if "Cluster" in df.columns and "Cluster" in df_kwh.columns:
                df_kwh = df_kwh[df_kwh["Cluster"].isin(df["Cluster"].unique())]
            
            daily_columns = [col for col in df_kwh.columns if '-' in str(col)]
            if not daily_columns:
                continue
            
            df_kwh[daily_columns] = df_kwh[daily_columns].apply(pd.to_numeric, errors='coerce')
            
            sheet_year = sheet_years.get(sheet, None)
            
            for date in pd.date_range(start=start_date, end=end_date, freq='D'):
                date_str = date.strftime('%b-%d')
                
                if sheet_year is not None and date.year != sheet_year:
                    continue
                    
                if date_str in daily_columns:
                    grouped = df_kwh.groupby("Cluster")[date_str].sum().reset_index()
                    for _, row in grouped.iterrows():
                        cluster_kwh.append({
                            "Date": date,
                            "kWh": round(row[date_str], 2),
                            "Cluster": row["Cluster"]
                        })
        except Exception as e:
            st.warning(f"Error processing sheet '{sheet}': {str(e)}")
    
    if cluster_kwh:
        daily_kwh_df = pd.DataFrame(cluster_kwh)
        daily_kwh_df = daily_kwh_df[(daily_kwh_df["Date"] >= start_date) & (daily_kwh_df["Date"] <= end_date)].round(2)
        return daily_kwh_df
    else:
        st.warning("No matching kWh data found in any sheet.")
        return pd.DataFrame(columns=["Date", "kWh", "Cluster"])

def get_target_kwh(xls, active_sheet, df, start_date, end_date):
    cluster_target = []
    
    start_year = start_date.year
    end_year = end_date.year
    selected_years = list(range(start_year, end_year + 1))
    
    all_target_sheets = [str(sheet) for sheet in xls.sheet_names if 'Target kWh' in str(sheet)]
    
    sheet_years = {}
    for sheet in all_target_sheets:
        digits = ''.join(filter(str.isdigit, sheet))
        if len(digits) >= 4:
            for i in range(len(digits) - 3):
                potential_year = int(digits[i:i+4])
                if 2000 <= potential_year <= 2100:
                    sheet_years[sheet] = potential_year
                    break
    
    target_sheets = []
    for sheet in all_target_sheets:
        if sheet in sheet_years:
            if sheet_years[sheet] in selected_years:
                target_sheets.append(sheet)
        else:
            target_sheets.append(sheet)
    
    if not target_sheets:
        st.warning(f"No Target kWh sheets found for years {selected_years}.")
        return pd.DataFrame(columns=["Date", "Target kWh", "Cluster"])
    
    
    for sheet in target_sheets:
        try:
            df_target = pd.read_excel(xls, sheet_name=sheet, header=0)
            df_target.columns = df_target.columns.str.strip()
            
            smip_rows = df_target.apply(lambda row: any('SMIP per CLUSTER' in str(cell) and 'OUTSIDE SMIP per CLUSTER' not in str(cell) for cell in row), axis=1)
            outside_smip_rows = df_target.apply(lambda row: any('OUTSIDE SMIP per CLUSTER' in str(cell) for cell in row), axis=1)
            
            if any(smip_rows) and any(outside_smip_rows):
                smip_indices = df_target[smip_rows].index
                outside_smip_indices = df_target[outside_smip_rows].index
                
                if active_sheet == "SMIP Database":
                    start_idx = min(smip_indices)
                    end_idx = min(outside_smip_indices)
                else:
                    start_idx = min(outside_smip_indices)
                    next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
                    end_idx = min(next_smip_indices) if next_smip_indices else len(df_target)
                
                df_target = df_target.iloc[start_idx:end_idx].reset_index(drop=True)
            
            if "Cluster" in df.columns and "Cluster" in df_target.columns:
                df_target = df_target[df_target["Cluster"].isin(df["Cluster"].unique())]
            
            target_columns = [col for col in df_target.columns if '-' in str(col)]
            if not target_columns:
                continue
            
            df_target[target_columns] = df_target[target_columns].apply(pd.to_numeric, errors='coerce')
            
            sheet_year = sheet_years.get(sheet, None)
            
            for date in pd.date_range(start=start_date, end=end_date, freq='D'):
                date_str = date.strftime('%b-%d')
                
                if sheet_year is not None and date.year != sheet_year:
                    continue
                    
                if date_str in target_columns:
                    grouped = df_target.groupby("Cluster")[date_str].sum().reset_index()
                    for _, row in grouped.iterrows():
                        cluster_target.append({
                            "Date": date,
                            "Target kWh": round(row[date_str], 2),
                            "Cluster": row["Cluster"]
                        })
        except Exception as e:
            st.warning(f"Error processing sheet '{sheet}': {str(e)}")
    
    if cluster_target:
        target_kwh_df = pd.DataFrame(cluster_target)
        target_kwh_df = target_kwh_df[(target_kwh_df["Date"] >= start_date) & (target_kwh_df["Date"] <= end_date)].round(2)
        return target_kwh_df
    else:
        st.warning("No matching Target kWh data found in any sheet.")
        return pd.DataFrame(columns=["Date", "Target kWh", "Cluster"])
        
def get_yesterday_kwh(xls, active_sheet, df):
    yesterday_abbr = yesterday.strftime('%b-%d')

    kwh_sheets = [sheet for sheet in xls.sheet_names if 'kWh' in sheet and 'Target' not in sheet]
    if not kwh_sheets:
        st.warning("No kWh sheets found.")
        return pd.DataFrame(columns=["Site Name", "kWh"]) 

    latest_kwh_sheet = max(kwh_sheets, key=lambda x: int(''.join(filter(str.isdigit, x))))

    df_kwh = pd.read_excel(xls, sheet_name=latest_kwh_sheet, header=0) 
    df_kwh.columns = df_kwh.columns.str.strip()
    
    smip_rows = df_kwh.apply(lambda row: any('SMIP' in str(cell) and 'OUTSIDE' not in str(cell) 
                                             for cell in row), axis=1)
    outside_smip_rows = df_kwh.apply(lambda row: any('OUTSIDE SMIP' in str(cell) 
                                                    for cell in row), axis=1)
    
    if any(smip_rows) and any(outside_smip_rows):
        smip_indices = df_kwh[smip_rows].index
        outside_smip_indices = df_kwh[outside_smip_rows].index
        
        if active_sheet == "SMIP Database":
            start_idx = min(smip_indices)
            end_idx = min(outside_smip_indices)
            df_kwh = df_kwh.iloc[start_idx:end_idx].reset_index(drop=True)
        elif active_sheet == "Outside-SMIP Database":
            start_idx = min(outside_smip_indices)
            next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
            end_idx = min(next_smip_indices) if next_smip_indices else len(df_kwh)
            df_kwh = df_kwh.iloc[start_idx:end_idx].reset_index(drop=True)

    if "Site Name" in df_kwh.columns and "Cluster" in df.columns:
        df_kwh = df_kwh[df_kwh["Site Name"].isin(df["Site Name"].unique())]
        df_kwh = df_kwh[df_kwh["Cluster"].isin(df["Cluster"].unique())]

    matching_columns = [col for col in df_kwh.columns.astype(str) if yesterday_abbr in col]

    if matching_columns:
        df_kwh["kWh"] = pd.to_numeric(df_kwh[matching_columns].sum(axis=1), errors='coerce')
        result_df = df_kwh.groupby("Site Name", as_index=False)["kWh"].sum()
        
        result_df["kWh"] = result_df["kWh"].round(2)
        
        return result_df 
    else:
        return pd.DataFrame(columns=["Site Name", "kWh"])
    
def get_yesterday_sp(xls, active_sheet, df):
    yesterday_abbr = yesterday.strftime('%b-%d')

    sp_sheets = [sheet for sheet in xls.sheet_names if 'SP' in sheet and 'Target' not in sheet]
    if not sp_sheets:
        st.warning("No Specific Production sheets found.")
        return pd.DataFrame(columns=["Site Name", "Specific Production"]) 

    latest_sp_sheet = max(sp_sheets, key=lambda x: int(''.join(filter(str.isdigit, x))))
    
    df_sp = pd.read_excel(xls, sheet_name=latest_sp_sheet, header=0) 
    df_sp.columns = df_sp.columns.str.strip()
    
    smip_rows = df_sp.apply(lambda row: any('SMIP' in str(cell) and 'OUTSIDE' not in str(cell) 
                                         for cell in row), axis=1)
    outside_smip_rows = df_sp.apply(lambda row: any('OUTSIDE SMIP' in str(cell) 
                                                for cell in row), axis=1)
    
    if any(smip_rows) and any(outside_smip_rows):
        smip_indices = df_sp[smip_rows].index
        outside_smip_indices = df_sp[outside_smip_rows].index
        
        if active_sheet == "SMIP Database":
            start_idx = min(smip_indices)
            end_idx = min(outside_smip_indices)
            df_sp = df_sp.iloc[start_idx:end_idx].reset_index(drop=True)
        elif active_sheet == "Outside-SMIP Database":
            start_idx = min(outside_smip_indices)
            next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
            end_idx = min(next_smip_indices) if next_smip_indices else len(df_sp)
            df_sp = df_sp.iloc[start_idx:end_idx].reset_index(drop=True)

    if "Site Name" in df_sp.columns:
        df_sp = df_sp[df_sp["Site Name"].isin(df["Site Name"].unique())]
    
    matching_columns = [col for col in df_sp.columns.astype(str) if yesterday_abbr in col]

    if matching_columns:
        df_sp["Specific Production"] = pd.to_numeric(df_sp[matching_columns].mean(axis=1), errors='coerce')
        result_df = df_sp.groupby("Site Name", as_index=False)["Specific Production"].mean()
        return result_df 
    else:
        return pd.DataFrame(columns=["Site Name", "Specific Production"])
    
def get_yesterday_ppr(xls, active_sheet, df):
    yesterday_abbr = yesterday.strftime('%b-%d')

    ppr_sheets = [sheet for sheet in xls.sheet_names if 'PPR' in sheet and 'Target' not in sheet]
    if not ppr_sheets:
        st.warning("No PPR sheets found.")
        return pd.DataFrame(columns=["Site Name", "PPR"]) 

    latest_ppr_sheet = max(ppr_sheets, key=lambda x: int(''.join(filter(str.isdigit, x))))
    
    df_ppr = pd.read_excel(xls, sheet_name=latest_ppr_sheet, header=0) 
    df_ppr.columns = df_ppr.columns.str.strip()
    
    smip_rows = df_ppr.apply(lambda row: any('SMIP per SITE' in str(cell) and 'OUTSIDE' not in str(cell) 
                                         for cell in row), axis=1)
    outside_smip_rows = df_ppr.apply(lambda row: any('OUTSIDE SMIP per SITE' in str(cell) 
                                                for cell in row), axis=1)
    
    if any(smip_rows) and any(outside_smip_rows):
        smip_indices = df_ppr[smip_rows].index
        outside_smip_indices = df_ppr[outside_smip_rows].index
        
        if active_sheet == "SMIP Database":
            start_idx = min(smip_indices)
            end_idx = min(outside_smip_indices)
            df_ppr = df_ppr.iloc[start_idx:end_idx].reset_index(drop=True)
        elif active_sheet == "Outside-SMIP Database":
            start_idx = min(outside_smip_indices)
            next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
            end_idx = min(next_smip_indices) if next_smip_indices else len(df_ppr)
            df_ppr = df_ppr.iloc[start_idx:end_idx].reset_index(drop=True)

    if "Site Name" in df_ppr.columns:
        df_ppr = df_ppr[df_ppr["Site Name"].isin(df["Site Name"].unique())]

    matching_columns = [col for col in df_ppr.columns.astype(str) if yesterday_abbr in col]

    if matching_columns:
        df_ppr["PPR"] = pd.to_numeric(df_ppr[matching_columns].mean(axis=1), errors='coerce') * 100
        result_df = df_ppr.groupby("Site Name", as_index=False)["PPR"].mean()
        return result_df 
    else:
        return pd.DataFrame(columns=["Site Name", "PPR"])

def get_last_7_days_table(xls, active_sheet, df):
    yesterday_abbrs = [(yesterday - timedelta(days=i)).strftime('%b-%d') for i in range(7)]
    
    last_7_days_kwh_df = pd.DataFrame(columns=["Site Name", "kWh"])
    last_7_days_sp_df = pd.DataFrame(columns=["Site Name", "Specific Production"])
    last_7_days_ppr_df = pd.DataFrame(columns=["Site Name", "PPR"])

    kwh_sheets = [sheet for sheet in xls.sheet_names if 'kWh' in sheet and 'Target' not in sheet]
    if not kwh_sheets:
        st.warning("No kWh sheets found.")
        return pd.DataFrame(columns=["Site Name", "kWh"]) 

    latest_kwh_sheet = max(kwh_sheets, key=lambda x: int(''.join(filter(str.isdigit, x))))

    df_kwh = pd.read_excel(xls, sheet_name=latest_kwh_sheet, header=0) 
    df_kwh.columns = df_kwh.columns.str.strip()

    smip_rows = df_kwh.apply(lambda row: any('SMIP' in str(cell) and 'OUTSIDE' not in str(cell) 
                                         for cell in row), axis=1)
    outside_smip_rows = df_kwh.apply(lambda row: any('OUTSIDE SMIP' in str(cell) 
                                                for cell in row), axis=1)
    
    if any(smip_rows) and any(outside_smip_rows):
        smip_indices = df_kwh[smip_rows].index
        outside_smip_indices = df_kwh[outside_smip_rows].index
        
        if active_sheet == "SMIP Database":
            start_idx = min(smip_indices)
            end_idx = min(outside_smip_indices)
            df_kwh = df_kwh.iloc[start_idx:end_idx].reset_index(drop=True)
        elif active_sheet == "Outside-SMIP Database":
            start_idx = min(outside_smip_indices)
            next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
            end_idx = min(next_smip_indices) if next_smip_indices else len(df_kwh)
            df_kwh = df_kwh.iloc[start_idx:end_idx].reset_index(drop=True)

    if "Site Name" in df_kwh.columns and "Cluster" in df.columns:
        df_kwh = df_kwh[df_kwh["Site Name"].isin(df["Site Name"].unique())]
        df_kwh = df_kwh[df_kwh["Cluster"].isin(df["Cluster"].unique())]

    for date_abbr in yesterday_abbrs:
        matching_columns = [col for col in df_kwh.columns.astype(str) if date_abbr in col]

        if matching_columns:
            df_kwh["kWh"] = pd.to_numeric(df_kwh[matching_columns].sum(axis=1), errors='coerce')
            daily_result_kwh_df = df_kwh.groupby("Site Name", as_index=False)["kWh"].sum()
            last_7_days_kwh_df = pd.concat([last_7_days_kwh_df, daily_result_kwh_df], ignore_index=True)

    last_7_days_kwh_df = last_7_days_kwh_df.groupby("Site Name", as_index=False).agg({"kWh": "sum"})
    last_7_days_kwh_df["kWh"] = (pd.to_numeric(last_7_days_kwh_df["kWh"], errors="coerce").round(2).fillna(0))


    sp_sheets = [sheet for sheet in xls.sheet_names if 'SP' in sheet and 'Target' not in sheet]
    if not sp_sheets:
        st.warning("No Specific Production sheets found.")
        return pd.DataFrame(columns=["Site Name", "Specific Production"]) 

    latest_sp_sheet = max(sp_sheets, key=lambda x: int(''.join(filter(str.isdigit, x))))
    
    df_sp = pd.read_excel(xls, sheet_name=latest_sp_sheet, header=0) 
    df_sp.columns = df_sp.columns.str.strip()
    
    smip_rows = df_sp.apply(lambda row: any('SMIP' in str(cell) and 'OUTSIDE' not in str(cell) 
                                         for cell in row), axis=1)
    outside_smip_rows = df_sp.apply(lambda row: any('OUTSIDE SMIP' in str(cell) 
                                                for cell in row), axis=1)
    
    if any(smip_rows) and any(outside_smip_rows):
        smip_indices = df_sp[smip_rows].index
        outside_smip_indices = df_sp[outside_smip_rows].index
        
        if active_sheet == "SMIP Database":
            start_idx = min(smip_indices)
            end_idx = min(outside_smip_indices)
            df_sp = df_sp.iloc[start_idx:end_idx].reset_index(drop=True)
        elif active_sheet == "Outside-SMIP Database":
            start_idx = min(outside_smip_indices)
            next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
            end_idx = min(next_smip_indices) if next_smip_indices else len(df_sp)
            df_sp = df_sp.iloc[start_idx:end_idx].reset_index(drop=True)

    if "Site Name" in df_sp.columns:
        df_sp = df_sp[df_sp["Site Name"].isin(df["Site Name"].unique())]

    for date_abbr in yesterday_abbrs:
        matching_columns = [col for col in df_sp.columns.astype(str) if date_abbr in col]

        if matching_columns:
            df_sp["Specific Production"] = pd.to_numeric(df_sp[matching_columns].mean(axis=1), errors='coerce')
            daily_result_sp_df = df_sp.groupby("Site Name", as_index=False)["Specific Production"].mean()
            last_7_days_sp_df = pd.concat([last_7_days_sp_df, daily_result_sp_df], ignore_index=True)

    last_7_days_sp_df = last_7_days_sp_df.groupby("Site Name", as_index=False).agg({"Specific Production": "mean"})
    last_7_days_sp_df["Specific Production"] = (pd.to_numeric(last_7_days_sp_df["Specific Production"], errors="coerce").round(2).fillna(0))


    ppr_sheets = [sheet for sheet in xls.sheet_names if 'PPR' in sheet and 'Target' not in sheet]
    if not ppr_sheets:
        st.warning("No PPR sheets found.")
        return pd.DataFrame(columns=["Site Name", "PPR"]) 

    latest_ppr_sheet = max(ppr_sheets, key=lambda x: int(''.join(filter(str.isdigit, x))))
    
    df_ppr = pd.read_excel(xls, sheet_name=latest_ppr_sheet, header=0) 
    df_ppr.columns = df_ppr.columns.str.strip()
    
    smip_rows = df_ppr.apply(lambda row: any('SMIP per SITE' in str(cell) and 'OUTSIDE' not in str(cell) 
                                         for cell in row), axis=1)
    outside_smip_rows = df_ppr.apply(lambda row: any('OUTSIDE SMIP per SITE' in str(cell) 
                                                for cell in row), axis=1)
    
    if any(smip_rows) and any(outside_smip_rows):
        smip_indices = df_ppr[smip_rows].index
        outside_smip_indices = df_ppr[outside_smip_rows].index
        
        if active_sheet == "SMIP Database":
            start_idx = min(smip_indices)
            end_idx = min(outside_smip_indices)
            df_ppr = df_ppr.iloc[start_idx:end_idx].reset_index(drop=True)
        elif active_sheet == "Outside-SMIP Database":
            start_idx = min(outside_smip_indices)
            next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
            end_idx = min(next_smip_indices) if next_smip_indices else len(df_ppr)
            df_ppr = df_ppr.iloc[start_idx:end_idx].reset_index(drop=True)

    if "Site Name" in df_ppr.columns:
        df_ppr = df_ppr[df_ppr["Site Name"].isin(df["Site Name"].unique())]

    for date_abbr in yesterday_abbrs:
        matching_columns = [col for col in df_ppr.columns.astype(str) if date_abbr in col]

        if matching_columns:
            df_ppr["PPR"] = df_ppr[matching_columns].mean(axis=1) * 100
            daily_result_ppr_df = df_ppr.groupby("Site Name", as_index=False)["PPR"].mean()
            last_7_days_ppr_df = pd.concat([last_7_days_ppr_df, daily_result_ppr_df], ignore_index=True)
        else:
            st.warning(f"No matching columns found for date abbreviation: {date_abbr}")

    last_7_days_ppr_df = last_7_days_ppr_df.groupby("Site Name", as_index=False).agg({"PPR": "mean"})
    last_7_days_ppr_df["PPR"] = (pd.to_numeric(last_7_days_ppr_df["PPR"], errors="coerce").round(2).fillna(0))


    last_7_days_df = pd.merge(last_7_days_kwh_df, last_7_days_sp_df, on="Site Name", how="outer")
    last_7_days_df = pd.merge(last_7_days_df, last_7_days_ppr_df, on="Site Name", how="outer")

    last_7_days_df = last_7_days_df.sort_values(by="PPR", ascending=False)

    return last_7_days_df

def get_last_30_days_table(xls, active_sheet, df):
    last_30_days_kwh_df = pd.DataFrame(columns=["Site Name", "kWh"])
    last_30_days_sp_df = pd.DataFrame(columns=["Site Name", "Specific Production"])
    last_30_days_ppr_df = pd.DataFrame(columns=["Site Name", "PPR"])

    kwh_sheets = [sheet for sheet in xls.sheet_names if 'kWh' in sheet and 'Target' not in sheet]
    if not kwh_sheets:
        st.warning("No kWh sheets found.")
        return pd.DataFrame(columns=["Site Name", "kWh"]) 

    latest_kwh_sheet = max(kwh_sheets, key=lambda x: int(''.join(filter(str.isdigit, x))))
    df_kwh = pd.read_excel(xls, sheet_name=latest_kwh_sheet, header=0) 
    df_kwh.columns = df_kwh.columns.str.strip()

    smip_rows = df_kwh.apply(lambda row: any('SMIP' in str(cell) and 'OUTSIDE' not in str(cell) 
                                         for cell in row), axis=1)
    outside_smip_rows = df_kwh.apply(lambda row: any('OUTSIDE SMIP' in str(cell) 
                                                for cell in row), axis=1)
    
    if any(smip_rows) and any(outside_smip_rows):
        smip_indices = df_kwh[smip_rows].index
        outside_smip_indices = df_kwh[outside_smip_rows].index
        
        if active_sheet == "SMIP Database":
            start_idx = min(smip_indices)
            end_idx = min(outside_smip_indices)
            df_kwh = df_kwh.iloc[start_idx:end_idx].reset_index(drop=True)
        elif active_sheet == "Outside-SMIP Database":
            start_idx = min(outside_smip_indices)
            next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
            end_idx = min(next_smip_indices) if next_smip_indices else len(df_kwh)
            df_kwh = df_kwh.iloc[start_idx:end_idx].reset_index(drop=True)

    if "Site Name" in df_kwh.columns and "Cluster" in df.columns:
        df_kwh = df_kwh[df_kwh["Site Name"].isin(df["Site Name"].unique())]
        df_kwh = df_kwh[df_kwh["Cluster"].isin(df["Cluster"].unique())]

    for i in range(30):
        date_abbr = (yesterday - timedelta(days=i)).strftime('%b-%d')
        matching_columns = [col for col in df_kwh.columns.astype(str) if date_abbr in col]

        if matching_columns:
            df_kwh["kWh"] = pd.to_numeric(df_kwh[matching_columns].sum(axis=1), errors='coerce')
            daily_result_kwh_df = df_kwh.groupby("Site Name", as_index=False)["kWh"].sum()
            last_30_days_kwh_df = pd.concat([last_30_days_kwh_df, daily_result_kwh_df], ignore_index=True)

    last_30_days_kwh_df = last_30_days_kwh_df.groupby("Site Name", as_index=False).agg({"kWh": "sum"})
    last_30_days_kwh_df["kWh"] = last_30_days_kwh_df["kWh"].round(2)

    sp_sheets = [sheet for sheet in xls.sheet_names if 'SP' in sheet and 'Target' not in sheet]
    if not sp_sheets:
        st.warning("No Specific Production sheets found.")
        return pd.DataFrame(columns=["Site Name", "Specific Production"]) 

    latest_sp_sheet = max(sp_sheets, key=lambda x: int(''.join(filter(str.isdigit, x))))
    df_sp = pd.read_excel(xls, sheet_name=latest_sp_sheet, header=0) 
    df_sp.columns = df_sp.columns.str.strip()

    smip_rows = df_sp.apply(lambda row: any('SMIP' in str(cell) and 'OUTSIDE' not in str(cell) 
                                         for cell in row), axis=1)
    outside_smip_rows = df_sp.apply(lambda row: any('OUTSIDE SMIP' in str(cell) 
                                                for cell in row), axis=1)
    
    if any(smip_rows) and any(outside_smip_rows):
        smip_indices = df_sp[smip_rows].index
        outside_smip_indices = df_sp[outside_smip_rows].index
        
        if active_sheet == "SMIP Database":
            start_idx = min(smip_indices)
            end_idx = min(outside_smip_indices)
            df_sp = df_sp.iloc[start_idx:end_idx].reset_index(drop=True)
        elif active_sheet == "Outside-SMIP Database":
            start_idx = min(outside_smip_indices)
            next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
            end_idx = min(next_smip_indices) if next_smip_indices else len(df_sp)
            df_sp = df_sp.iloc[start_idx:end_idx].reset_index(drop=True)

    if "Site Name" in df_sp.columns:
        df_sp = df_sp[df_sp["Site Name"].isin(df["Site Name"].unique())]

    for i in range(30):
        date_abbr = (yesterday - timedelta(days=i)).strftime('%b-%d')
        matching_columns = [col for col in df_sp.columns.astype(str) if date_abbr in col]

        if matching_columns:
            df_sp["Specific Production"] = pd.to_numeric(df_sp[matching_columns].mean(axis=1), errors='coerce')
            daily_result_sp_df = df_sp.groupby("Site Name", as_index=False)["Specific Production"].mean()
            last_30_days_sp_df = pd.concat([last_30_days_sp_df, daily_result_sp_df], ignore_index=True)

    last_30_days_sp_df = last_30_days_sp_df.groupby("Site Name", as_index=False).agg({"Specific Production": "mean"})
    last_30_days_sp_df["Specific Production"] = last_30_days_sp_df["Specific Production"].round(2)

    ppr_sheets = [sheet for sheet in xls.sheet_names if 'PPR' in sheet and 'Target' not in sheet]
    if not ppr_sheets:
        st.warning("No PPR sheets found.")
        return pd.DataFrame(columns=["Site Name", "PPR"]) 

    latest_ppr_sheet = max(ppr_sheets, key=lambda x: int(''.join(filter(str.isdigit, x))))
    df_ppr = pd.read_excel(xls, sheet_name=latest_ppr_sheet, header=0) 
    df_ppr.columns = df_ppr.columns.str.strip()

    smip_rows = df_ppr.apply(lambda row: any('SMIP per SITE' in str(cell) and 'OUTSIDE' not in str(cell) 
                                         for cell in row), axis=1)
    outside_smip_rows = df_ppr.apply(lambda row: any('OUTSIDE SMIP per SITE' in str(cell) 
                                                for cell in row), axis=1)
    
    if any(smip_rows) and any(outside_smip_rows):
        smip_indices = df_ppr[smip_rows].index
        outside_smip_indices = df_ppr[outside_smip_rows].index
        
        if active_sheet == "SMIP Database":
            start_idx = min(smip_indices)
            end_idx = min(outside_smip_indices)
            df_ppr = df_ppr.iloc[start_idx:end_idx].reset_index(drop=True)
        elif active_sheet == "Outside-SMIP Database":
            start_idx = min(outside_smip_indices)
            next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
            end_idx = min(next_smip_indices) if next_smip_indices else len(df_ppr)
            df_ppr = df_ppr.iloc[start_idx:end_idx].reset_index(drop=True)

    if "Site Name" in df_ppr.columns:
        df_ppr = df_ppr[df_ppr["Site Name"].isin(df["Site Name"].unique())]

    for i in range(30):
        date_abbr = (yesterday - timedelta(days=i)).strftime('%b-%d')
        matching_columns = [col for col in df_ppr.columns.astype(str) if date_abbr in col]

        if matching_columns:
            df_ppr["PPR"] = df_ppr[matching_columns].mean(axis=1) * 100
            daily_result_ppr_df = df_ppr.groupby("Site Name", as_index=False)["PPR"].mean()
            last_30_days_ppr_df = pd.concat([last_30_days_ppr_df, daily_result_ppr_df], ignore_index=True)

    last_30_days_ppr_df = last_30_days_ppr_df.groupby("Site Name", as_index=False).agg({"PPR": "mean"})
    last_30_days_ppr_df["PPR"] = last_30_days_ppr_df["PPR"].round(2)

    last_30_days_df = pd.merge(last_30_days_kwh_df, last_30_days_sp_df, on="Site Name", how="outer")
    last_30_days_df = pd.merge(last_30_days_df, last_30_days_ppr_df, on="Site Name", how="outer")

    last_30_days_df = last_30_days_df.sort_values(by="PPR", ascending=False)

    return last_30_days_df

def get_this_year_table(xls, active_sheet, df):
    today = datetime.today()
    current_year = today.year
    this_year_kwh_df = pd.DataFrame(columns=["Site Name", "kWh"])
    this_year_sp_df = pd.DataFrame(columns=["Site Name", "Specific Production"])
    this_year_ppr_df = pd.DataFrame(columns=["Site Name", "PPR"])

    kwh_sheet_name = f"{current_year} kWh"
    if kwh_sheet_name not in xls.sheet_names:
        st.warning(f"No sheet found for {kwh_sheet_name}.")
        return pd.DataFrame(columns=["Site Name", "kWh"])

    df_kwh = pd.read_excel(xls, sheet_name=kwh_sheet_name, header=0)
    df_kwh.columns = df_kwh.columns.str.strip()

    smip_rows = df_kwh.apply(lambda row: any('SMIP' in str(cell) and 'OUTSIDE' not in str(cell) 
                                         for cell in row), axis=1)
    outside_smip_rows = df_kwh.apply(lambda row: any('OUTSIDE SMIP' in str(cell) 
                                                for cell in row), axis=1)
    
    if any(smip_rows) and any(outside_smip_rows):
        smip_indices = df_kwh[smip_rows].index
        outside_smip_indices = df_kwh[outside_smip_rows].index
        
        if active_sheet == "SMIP Database":
            start_idx = min(smip_indices)
            end_idx = min(outside_smip_indices)
            df_kwh = df_kwh.iloc[start_idx:end_idx].reset_index(drop=True)
        elif active_sheet == "Outside-SMIP Database":
            start_idx = min(outside_smip_indices)
            next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
            end_idx = min(next_smip_indices) if next_smip_indices else len(df_kwh)
            df_kwh = df_kwh.iloc[start_idx:end_idx].reset_index(drop=True)

    if "Site Name" in df_kwh.columns:
        df_kwh = df_kwh[df_kwh["Site Name"].isin(df["Site Name"].unique())]

    daily_columns = [col for col in df_kwh.columns if '-' in str(col)]

    if daily_columns:
        df_kwh[daily_columns] = df_kwh[daily_columns].apply(pd.to_numeric, errors='coerce')
        df_kwh["kWh"] = df_kwh[daily_columns].sum(axis=1)
        this_year_kwh_df = df_kwh.groupby("Site Name", as_index=False)["kWh"].sum()
    else:
        st.warning("No daily columns found in the filtered data.")

    this_year_kwh_df["kWh"] = this_year_kwh_df["kWh"].round(2)

    sp_sheet_name = f"{current_year} SP"
    if sp_sheet_name not in xls.sheet_names:
        st.warning(f"No sheet found for {sp_sheet_name}.")
        return pd.DataFrame(columns=["Site Name", "Specific Production"])

    df_sp = pd.read_excel(xls, sheet_name=sp_sheet_name, header=0)
    df_sp.columns = df_sp.columns.str.strip()

    smip_rows = df_sp.apply(lambda row: any('SMIP' in str(cell) and 'OUTSIDE' not in str(cell) 
                                         for cell in row), axis=1)
    outside_smip_rows = df_sp.apply(lambda row: any('OUTSIDE SMIP' in str(cell) 
                                                for cell in row), axis=1)
    
    if any(smip_rows) and any(outside_smip_rows):
        smip_indices = df_sp[smip_rows].index
        outside_smip_indices = df_sp[outside_smip_rows].index
        
        if active_sheet == "SMIP Database":
            start_idx = min(smip_indices)
            end_idx = min(outside_smip_indices)
            df_sp = df_sp.iloc[start_idx:end_idx].reset_index(drop=True)
        elif active_sheet == "Outside-SMIP Database":
            start_idx = min (outside_smip_indices)
            next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
            end_idx = min(next_smip_indices) if next_smip_indices else len(df_sp)
            df_sp = df_sp.iloc[start_idx:end_idx].reset_index(drop=True)

    if "Site Name" in df_sp.columns:
        df_sp = df_sp[df_sp["Site Name"].isin(df["Site Name"].unique())]

    sp_columns = [col for col in df_sp.columns if '-' in str(col)]
    if sp_columns:
        df_sp[sp_columns] = df_sp[sp_columns].apply(pd.to_numeric, errors='coerce')
        df_sp["Specific Production"] = df_sp[sp_columns].mean(axis=1)
        this_year_sp_df = df_sp.groupby("Site Name", as_index=False)["Specific Production"].mean()
    else:
        st.warning("No valid 'SP' columns found in the filtered data.")

    this_year_sp_df["Specific Production"] = this_year_sp_df["Specific Production"].round(2)

    ppr_sheet_name = f"{current_year} PPR"
    if ppr_sheet_name not in xls.sheet_names:
        st.warning(f"No sheet found for {ppr_sheet_name}.")
        return pd.DataFrame(columns=["Site Name", "PPR"])

    df_ppr = pd.read_excel(xls, sheet_name=ppr_sheet_name, header=0)
    df_ppr.columns = df_ppr.columns.str.strip()

    smip_rows = df_ppr.apply(lambda row: any('SMIP per SITE' in str(cell) and 'OUTSIDE' not in str(cell) 
                                         for cell in row), axis=1)
    outside_smip_rows = df_ppr.apply(lambda row: any('OUTSIDE SMIP per SITE' in str(cell) 
                                                for cell in row), axis=1)
    
    if any(smip_rows) and any(outside_smip_rows):
        smip_indices = df_ppr[smip_rows].index
        outside_smip_indices = df_ppr[outside_smip_rows].index
        
        if active_sheet == "SMIP Database":
            start_idx = min(smip_indices)
            end_idx = min(outside_smip_indices)
            df_ppr = df_ppr.iloc[start_idx:end_idx].reset_index(drop=True)
        elif active_sheet == "Outside-SMIP Database":
            start_idx = min(outside_smip_indices)
            next_smip_indices = [idx for idx in smip_indices if idx > start_idx]
            end_idx = min(next_smip_indices) if next_smip_indices else len(df_ppr)
            df_ppr = df_ppr.iloc[start_idx:end_idx].reset_index(drop=True)

    if "Site Name" in df_ppr.columns:
        df_ppr = df_ppr[df_ppr["Site Name"].isin(df["Site Name"].unique())]

    ppr_columns = [col for col in df_ppr.columns if '-' in str(col)]
    if ppr_columns:
        df_ppr[ppr_columns] = df_ppr[ppr_columns].apply(pd.to_numeric, errors='coerce')
        df_ppr["PPR"] = df_ppr[ppr_columns].mean(axis=1) * 100
        this_year_ppr_df = df_ppr.groupby("Site Name", as_index=False)["PPR"].mean()
    else:
        st.warning("No valid 'PPR' columns found in the filtered data.")

    this_year_ppr_df["PPR"] = this_year_ppr_df["PPR"].round(2)

    this_year_df = pd.merge(this_year_kwh_df, this_year_sp_df, on="Site Name", how="outer")
    this_year_df = pd.merge(this_year_df, this_year_ppr_df, on="Site Name", how="outer")
    this_year_df = this_year_df.sort_values(by="PPR", ascending=False)

    return this_year_df

def filter_by_cluster(df):
        clusters = df["Cluster"].dropna().unique().tolist()
        cluster_options = ["Select All"] + clusters
        selected_clusters = st.sidebar.multiselect("Select Cluster", cluster_options)
        if selected_clusters and "Select All" not in selected_clusters:
            return df[df["Cluster"].isin(selected_clusters)]
        return df

def filter_by_site(df):
    site_names = df["Site Name"].dropna().unique().tolist()
    site_options = ["Select All"] + site_names
    selected_sites = st.sidebar.multiselect("Select Site Name", site_options)
    if selected_sites and "Select All" not in selected_sites:
        return df[df["Site Name"].isin(selected_sites)]
    return df

def create_date_column_and_filter(df, time_filter):
    if 'Year' in df.columns and 'Month' in df.columns:
        if time_filter == "Daily":
            if 'Day' in df.columns:
                df["Date"] = pd.to_datetime(df[["Year", "Month", "Day"]])
            else:
                st.warning("Day column is missing for daily data.")
                return df
        else:
            df["Date"] = pd.to_datetime(df[["Year", "Month"]].assign(Day=1))

    if not df.empty and 'Date' in df.columns:
        df = df[df['Date'] <= today]
    return df

def kWh(show_subheader=True):
        if combined_df.empty:
            return None, None, None, 0, 0

        if time_filter == "Yearly":
            grouped_df = combined_df.groupby(["Year", "Site Name"], as_index=False)["kWh"].sum()
            grouped_df["Year"] = grouped_df["Year"].astype(str)
            if show_subheader:
                st.subheader("📈Yearly kWh Trend per Site")
            x_col = "Year"

        elif time_filter == "Cumulative":
            combined_df["Date"] = pd.to_datetime(combined_df[["Year", "Month"]].assign(Day=1))
            grouped_df = combined_df.groupby(["Date", "Site Name"], as_index=False)["kWh"].sum()
            grouped_df.sort_values(["Date", "Site Name"], inplace=True)
            grouped_df["Cumulative kWh"] = grouped_df.groupby("Site Name")["kWh"].cumsum().round(2)
            if show_subheader:
                st.subheader("📈Cumulative kWh Trend from Start to Present")
            x_col = "Date"

        elif time_filter == "Monthly":
            grouped_df = combined_df.groupby(["Year", "Month", "Site Name"], as_index=False)["kWh"].sum()
            grouped_df["Date"] = pd.to_datetime(
                grouped_df["Year"].astype(str) + '-' + grouped_df["Month"].astype(str) + '-01'
            )
            if show_subheader:
                st.subheader("📈Monthly kWh Trend per Site")
            x_col = "Date"

        else:
            grouped_df = combined_df.groupby(["Date", "Site Name"], as_index=False)["kWh"].sum()
            if show_subheader:
                st.subheader(f"📈Daily kWh Trend for {selected_month} {selected_year}")
            x_col = "Date"

        hover_data = {"kWh": ":,.2f", x_col: True}
        total_kwh = grouped_df["kWh"].sum().round(2)

        if time_filter == "Cumulative":
            hover_data["Cumulative kWh"] = ":,.2f"

        return grouped_df, x_col, hover_data, total_kwh

def SP(show_subheader=True):
    if combined_sp_df.empty:
        return None, None, None, 0

    if time_filter == "Yearly":
        grouped_sp_df = combined_sp_df.groupby(["Year", "Site Name"], as_index=False)["Specific Production"].sum()
        grouped_sp_df["Year"] = grouped_sp_df["Year"].astype(str)
        if show_subheader:
            st.subheader("📈Yearly Specific Production Trend per Site")
        x_col_sp = "Year"

    elif time_filter == "Cumulative":
        combined_sp_df["Date"] = pd.to_datetime(combined_sp_df[["Year", "Month"]].assign(Day=1))
        grouped_sp_df = combined_sp_df.groupby(["Date", "Site Name"], as_index=False)["Specific Production"].sum()
        grouped_sp_df.sort_values(["Date", "Site Name"], inplace=True)
        grouped_sp_df["Specific Production"] = pd.to_numeric(grouped_sp_df["Specific Production"], errors='coerce').fillna(0)
        grouped_sp_df["Cumulative Specific Production"] = grouped_sp_df.groupby("Site Name")["Specific Production"].cumsum().round(2)
        if show_subheader:
            st.subheader("📈Cumulative Specific Production Trend from Start to Present")
        x_col_sp = "Date"

    elif time_filter == "Monthly":
        grouped_sp_df = combined_sp_df.groupby(["Year", "Month", "Site Name"], as_index=False)["Specific Production"].mean()
        grouped_sp_df["Date"] = pd.to_datetime(
            grouped_sp_df["Year"].astype(str) + '-' + grouped_sp_df["Month"].astype(str) + '-01')
        if show_subheader:
            st.subheader("📈Monthly Specific Production Trend per Site")
        x_col_sp = "Date"

    else:
        grouped_sp_df = combined_sp_df.groupby(["Date", "Site Name"], as_index=False)["Specific Production"].sum()
        if show_subheader:
            st.subheader(f"📈Daily Specific Production Trend for {selected_month} {selected_year}")
        x_col_sp = "Date"

    hover_data_sp = {"Specific Production": ":,.2f", x_col_sp: True}
    total_sp = round(combined_sp_df["Specific Production"].mean(), 2)

    if time_filter == "Cumulative":
        hover_data_sp["Cumulative Specific Production"] = ":,.2f"

    return grouped_sp_df, x_col_sp, hover_data_sp, total_sp

def AF(show_subheader=True):
    if combined_af_df.empty:
        return None, None, None, 0, 0
    if time_filter == "Yearly":
        grouped_af_df = combined_af_df.groupby(["Year", "Site Name"], as_index=False)["Availability Factor"].mean()
        grouped_af_df["Year"] = grouped_af_df["Year"].astype(str)
        if show_subheader:
            st.subheader("📈Yearly Average Availability Factor Trend per Site")
        x_col_af = "Year"

    elif time_filter == "Cumulative":
        combined_af_df["Date"] = pd.to_datetime(combined_af_df[["Year", "Month"]].assign(Day=1))
        grouped_af_df = combined_af_df.groupby(["Date", "Site Name"], as_index=False)["Availability Factor"].mean()
        grouped_af_df.sort_values(["Date", "Site Name"], inplace=True)
        grouped_af_df["Cumulative AF"] = grouped_af_df.groupby("Site Name")["Availability Factor"].cumsum().round(2)
        grouped_af_df["Cumulative AF"] *= 100
        if show_subheader:
            st.subheader("📈Cumulative Availability Factor Trend from Start to Present")
        x_col_af = "Date"

    elif time_filter == "Monthly":
        grouped_af_df = combined_af_df.groupby(["Year", "Month", "Site Name"], as_index=False)["Availability Factor"].mean()
        grouped_af_df["Date"] = pd.to_datetime(
            grouped_af_df["Year"].astype(str) + '-' + grouped_af_df["Month"].astype(str) + '-01'
        )
        if show_subheader:
            st.subheader("📈Monthly Average Availability Factor Trend per Site")
        x_col_af = "Date"

    else:
        grouped_af_df = combined_af_df.groupby(["Date", "Site Name"], as_index=False)["Availability Factor"].sum()
        if show_subheader:
            st.subheader(f"📈Daily Availability Factor Trend for {selected_month} {selected_year}")
        x_col_af = "Date"

    hover_data_af = {"Availability Factor": ":,.2f", x_col_af: True}
    if time_filter == "Cumulative":
        hover_data_af["Cumulative AF"] = ":,.2f"

    grouped_af_df["Availability Factor"] = grouped_af_df["Availability Factor"] * 100

    total_af = (
        combined_af_df["Availability Factor"].mean().round(2)
        if time_filter != "Cumulative"
        else grouped_af_df["Cumulative AF"].max().round(2)
    )

    avg_af = grouped_af_df["Availability Factor"].mean()
    return grouped_af_df, x_col_af, hover_data_af, total_af, avg_af

st.set_page_config(page_title="Dashboard", page_icon="Dashboard/GEMI_logo.png", layout="wide")
st.subheader("Analysis")

st.markdown(
    """
    <style>
    .stMarkdown, .stText, .stWrite {
        color: #4CAF50; 
    }
    </style>
    """,
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file, engine='openpyxl')

    if "active_sheet" not in st.session_state:
        st.session_state["active_sheet"] = "SMIP Database"

    st.sidebar.image("Dashboard/GEMI_logo.png")
    
    selected_page = option_menu(
        menu_title="",
        options=["Dashboard", "Analysis"],
        icons=["bar-chart-line", "graph-up"],
        menu_icon="cast",
        default_index=0,
        orientation="horizontal",
        styles={
            "container": {"padding": "5px"},
            "icon": {"color": "white", "font-size": "20px"},
            "nav-link": {
                "font-size": "16px",
                "text-align": "center",
                "transition": "0.3s",
                "padding": "10px 20px",
                "border-radius": "5px"
            },
            "nav-link-selected": {
                "background-color": "#4CAF50",
                "color": "white"
            }
        }
    )
    with st.sidebar:
        selected_option = option_menu(
            "Please select",
            ["SMIP", "Outside SMIP"],
            icons=["geo-alt", "geo-alt-fill"],
            menu_icon="cast",
            default_index=0 if st.session_state["active_sheet"] == "SMIP Database" else 1,
            styles={
                "container": {"padding": "5px"},
                "icon": {"color": "white", "font-size": "20px"},
                "nav-link": {
                    "font-size": "16px",
                    "text-align": "left",
                    "margin": "5px",
                    "transition": "0.3s",
                    "padding": "10px 20px",
                    "border-radius": "5px"
                },
                "nav-link-selected": {
                    "background-color": "#4CAF50",
                    "color": "white"
                }
            }
        )
    st.session_state["active_sheet"] = "SMIP Database" if selected_option == "SMIP" else "Outside-SMIP Database"
    sheet_name = st.session_state["active_sheet"]
    if sheet_name not in xls.sheet_names:
        st.error(f"The '{sheet_name}' sheet was not found.")
        st.stop()

    df = pd.read_excel(xls, sheet_name=sheet_name, header=8)
    df.columns = df.columns.str.strip()

    if selected_page == "Analysis":

        if "Cluster" in df.columns and "Site Name" in df.columns:
            df = filter_by_cluster(df)
            df = filter_by_site(df)
        else:
            st.error("Required columns ('Cluster' or 'Site Name') missing.")
            st.stop()

        time_filter = st.sidebar.selectbox("Select Timeframe", ["Daily", "Monthly", "Yearly", "Cumulative"])

        # if time_filter == "Cumulative":
        #    show_prediction = st.sidebar.checkbox("Show Prediction", value=False)
        #else:
        #    show_prediction = False

        kwh_sheets = [str(sheet) for sheet in xls.sheet_names if 'kWh' in str(sheet) and 'Target' not in str(sheet)]
        sp_sheets = [str(sheet) for sheet in xls.sheet_names if 'SP' in str(sheet) and 'Target' not in str(sheet)]
        af_sheets = [str(sheet) for sheet in xls.sheet_names if 'AF' in str(sheet) and 'Target' not in str(sheet)]

        if not kwh_sheets:
            st.error("No valid 'kWh' sheets found.")
            st.stop()

        years = sorted(set(''.join(filter(str.isdigit, sheet)) for sheet in kwh_sheets))

        if time_filter in ["Daily", "Monthly"] and years:
            selected_year = st.sidebar.selectbox("Select Year", years)

        if time_filter == "Daily":
            months = list(calendar.month_name[1:])
            selected_month = st.sidebar.selectbox("Select Month", months)
            month_number = months.index(selected_month) + 1

        combined_df = pd.DataFrame()
        combined_sp_df = pd.DataFrame()
        combined_af_df = pd.DataFrame()

        for sheet in kwh_sheets:
            try:
                year = ''.join(filter(str.isdigit, str(sheet)))

                if time_filter in ["Monthly", "Daily"] and year != selected_year:
                    continue

                yearly_df = pd.read_excel(xls, sheet_name=sheet, header=0)
                yearly_df.columns = yearly_df.columns.str.strip().astype(str)

                month_names = [month[:3] for month in calendar.month_name[1:]]
                month_cols = [str(col) for col in yearly_df.columns if any(month in str(col) for month in month_names)]

                if not month_cols:
                    st.warning(f"No valid month columns in '{sheet}'. Skipping.")
                    continue

                yearly_df = yearly_df.iloc[453:].reset_index(drop=True)

                if "Site Name" in yearly_df.columns:
                    yearly_df = yearly_df[yearly_df["Site Name"].astype(str).isin(df["Site Name"].astype(str))]

                if year:
                    yearly_df["Year"] = int(year)

                if time_filter == "Daily":
                    month_cols = [col for col in month_cols if col.startswith(selected_month[:3])]

                    daily_df = yearly_df.melt(
                        id_vars=["Cluster", "Site Name", "Year"],
                        value_vars=month_cols,
                        var_name="Day",
                        value_name="kWh"
                    )

                    daily_df["Day"] = daily_df["Day"].str.extract(r'(\d+)')[0].astype(int)
                    daily_df["Date"] = pd.to_datetime(
                        f"{selected_year}-{month_number}-" + daily_df["Day"].astype(str),
                        errors="coerce"
                    )

                    daily_df["kWh"] = pd.to_numeric(daily_df["kWh"], errors='coerce').fillna(0).round(2)
                    combined_df = pd.concat([combined_df, daily_df], ignore_index=True)

                else:
                    melted_df = yearly_df.melt(
                        id_vars=["Cluster", "Site Name", "Year"],
                        value_vars=month_cols,
                        var_name="Month",
                        value_name="kWh"
                    )

                    melted_df["Month"] = pd.to_datetime(melted_df["Month"].str[:3], format='%b').dt.month
                    melted_df["kWh"] = pd.to_numeric(melted_df["kWh"], errors='coerce').fillna(0).round(2)
                    combined_df = pd.concat([combined_df, melted_df], ignore_index=True)

            except Exception as e:
                st.error(f"Failed to load {sheet}: {str(e)}")

        for sheet in sp_sheets:
            try:
                year = ''.join(filter(str.isdigit, str(sheet)))

                if time_filter in ["Monthly", "Daily"] and year != selected_year:
                    continue

                yearly_sp_df = pd.read_excel(xls, sheet_name=sheet, header=0)
                yearly_sp_df.columns = yearly_sp_df.columns.str.strip().astype(str)

                month_names = [month[:3] for month in calendar.month_name[1:]]
                month_cols = [str(col) for col in yearly_sp_df.columns if any(month in str(col) for month in month_names)]

                if not month_cols:
                    st.warning(f"No valid month columns in '{sheet}'. Skipping.")
                    continue

                yearly_sp_df = yearly_sp_df.iloc[453:].reset_index(drop=True)

                if "Site Name" in yearly_sp_df.columns:
                    yearly_sp_df = yearly_sp_df[yearly_sp_df["Site Name"].astype(str).isin(df["Site Name"].astype(str))]

                if year:
                    yearly_sp_df["Year"] = int(year)

                if time_filter == "Daily":
                    month_cols = [col for col in month_cols if col.startswith(selected_month[:3])]

                    daily_sp_df = yearly_sp_df.melt(
                        id_vars=["Cluster", "Site Name", "Year"],
                        value_vars=month_cols,
                        var_name="Day",
                        value_name="Specific Production"
                    )

                    daily_sp_df["Day"] = daily_sp_df["Day"].str.extract(r'(\d+)')[0].astype(int)
                    daily_sp_df["Date"] = pd.to_datetime(
                        f"{selected_year}-{month_number}-" + daily_sp_df["Day"].astype(str),
                        errors="coerce"
                    )

                    daily_sp_df["Specific Production"] = pd.to_numeric(
                        daily_sp_df["Specific Production"], errors='coerce'
                    ).fillna(0).round(2)

                    combined_sp_df = pd.concat([combined_sp_df, daily_sp_df], ignore_index=True)

                else:
                    melted_sp_df = yearly_sp_df.melt(
                        id_vars=["Cluster", "Site Name", "Year"],
                        value_vars=month_cols,
                        var_name="Month",
                        value_name="Specific Production"
                    )

                    melted_sp_df["Month"] = pd.to_datetime(melted_sp_df["Month"].str[:3], format='%b').dt.month
                    melted_sp_df["Date"] = pd.to_datetime(
                        melted_sp_df[["Year", "Month"]].assign(Day=1)
                    )

                    combined_sp_df = pd.concat([combined_sp_df, melted_sp_df], ignore_index=True)

            except Exception as e:
                st.error(f"Failed to load {sheet}: {str(e)}")



        for sheet in af_sheets:
            try:
                year = ''.join(filter(str.isdigit, str(sheet)))

                if time_filter in ["Monthly", "Daily"] and year != selected_year:
                    continue

                yearly_af_df = pd.read_excel(xls, sheet_name=sheet, header=0)
                yearly_af_df.columns = yearly_af_df.columns.str.strip().astype(str)

                month_names = [month[:3] for month in calendar.month_name[1:]]
                month_cols = [str(col) for col in yearly_af_df.columns if any(month in str(col) for month in month_names)]

                if not month_cols:
                    st.warning(f"No valid month columns in '{sheet}'. Skipping.")
                    continue

                yearly_af_df = yearly_af_df.iloc[453:].reset_index(drop=True)

                if "Site Name" in yearly_af_df.columns:
                    yearly_af_df = yearly_af_df[yearly_af_df["Site Name"].astype(str).isin(df["Site Name"].astype(str))]

                if year:
                    yearly_af_df["Year"] = int(year)

                if time_filter == "Daily":
                    month_cols = [col for col in month_cols if col.startswith(selected_month[:3])]

                    daily_af_df = yearly_af_df.melt(
                        id_vars=["Cluster", "Site Name", "Year"],
                        value_vars=month_cols,
                        var_name="Day",
                        value_name="Availability Factor"
                    )

                    daily_af_df["Day"] = daily_af_df["Day"].str.extract(r'(\d+)')[0].astype(int)
                    daily_af_df["Date"] = pd.to_datetime(
                        f"{selected_year}-{month_number}-" + daily_af_df["Day"].astype(str),
                        errors="coerce"
                    )

                    daily_af_df["Availability Factor"] = pd.to_numeric(
                        daily_af_df["Availability Factor"], errors='coerce'
                    ).fillna(0).round(2)

                    combined_af_df = pd.concat([combined_af_df, daily_af_df], ignore_index=True)

                else:
                    melted_af_df = yearly_af_df.melt(
                        id_vars=["Cluster", "Site Name", "Year"],
                        value_vars=month_cols,
                        var_name="Month",
                        value_name="Availability Factor"
                    )

                    melted_af_df["Month"] = pd.to_datetime(melted_af_df["Month"].str[:3], format='%b').dt.month
                    melted_af_df["Availability Factor"] = pd.to_numeric(melted_af_df["Availability Factor"], errors='coerce').fillna(0).round(2)

                    combined_af_df = pd.concat([combined_af_df, melted_af_df], ignore_index=True)

            except Exception as e:
                st.error(f"Failed to load {sheet}: {str(e)}")
            combined_df = create_date_column_and_filter(combined_df, time_filter)
            combined_sp_df = create_date_column_and_filter(combined_sp_df, time_filter)
            combined_af_df = create_date_column_and_filter(combined_af_df, time_filter)

            color_palette = px.colors.qualitative.Bold

            grouped_df, x_col, hover_data, total_kwh = kWh(show_subheader=True)
            fig = px.line(
                grouped_df,
                x=x_col,
                y="kWh" if time_filter != "Cumulative" else "Cumulative kWh",
                color="Site Name",
                markers=True,
                labels={"kWh": "Total kWh", "Cumulative kWh": "Cumulative kWh", x_col: "Time"},
                hover_data=hover_data,
                height=300,
                color_discrete_sequence=color_palette
            )
            st.plotly_chart(fig, key="kWh_chart")

            grouped_sp_df, x_col_sp, hover_data_sp, total_sp = SP(show_subheader=True)
            if grouped_sp_df is None or grouped_sp_df.empty:
                st.warning("No data available for Specific Production.")
            else:
                fig_sp = px.line(
                    grouped_sp_df,
                    x=x_col_sp,
                    y="Specific Production" if time_filter != "Cumulative" else "Cumulative Specific Production",
                    color="Site Name",
                    markers=True,
                    labels={"Specific Production": "Total SP", "Cumulative Specific Production": "Cumulative SP", x_col_sp: "Time"},
                    hover_data=hover_data_sp,
                    height=300,
                    color_discrete_sequence=px.colors.qualitative.Dark24
                )
                st.plotly_chart(fig_sp, key="SP_chart")

            grouped_af_df, x_col_af, hover_data_af, total_af, avg_af = AF(show_subheader=True)
            if grouped_af_df is None or grouped_af_df.empty:
                st.warning("No data available for Availability Factor.")

            else:
                y_col = "Cumulative AF" if time_filter == "Cumulative" else "Availability Factor"
                fig_af = px.line(
                    grouped_af_df,
                    x=x_col_af,
                    y=y_col, 
                    color="Site Name",
                    markers=True,
                    labels={
                        "Cumulative AF": "Cumulative AF (%)",
                        "Availability Factor": "Average AF (%)",
                        x_col_af: "Time"
                    },
                    hover_data=hover_data_af,
                    height=300,
                    color_discrete_sequence=px.colors.qualitative.Prism
                )
                st.plotly_chart(fig_af, key="AF_chart")

            combined_df["kWh"] = combined_df["kWh"].fillna(0)
            site_production_df = combined_df.groupby("Site Name", as_index=False)["kWh"].sum()
            site_production_df = site_production_df.sort_values("kWh", ascending=False)

            top_producers = site_production_df.nlargest(5, "kWh")
            lowest_producers = site_production_df.nsmallest(5, "kWh")

            st.subheader("📊 Highest & Lowest kWh Production by Site")

            site_production_df["kWh_hover"] = site_production_df["kWh"].map("{:,.2f}".format)

            fig_bar = px.bar(
                site_production_df,
                x="kWh",
                y="Site Name",
                color="kWh",
                orientation="h",
                labels={"kWh": "Total kWh", "Site Name": "Site"},
                color_continuous_scale="viridis",
                height=650,
                hover_data={"kWh_hover": True, "kWh": False}
            )

            fig_bar.update_traces(hovertemplate='<b>Site:</b> %{y}<br><b>Total kWh:</b> %{customdata[0]}')

            fig_bar.add_annotation(
                x=top_producers["kWh"].iloc[0],
                y=top_producers["Site Name"].iloc[0],
                text="🏅 Highest",
                showarrow=True,
                arrowhead=2,
                ax=50,
                ay=0,
                font=dict(color="green", size=14, family="Arial, sans-serif")
            )

            fig_bar.add_annotation(
                x=lowest_producers["kWh"].iloc[0],
                y=lowest_producers["Site Name"].iloc[0],
                text="🔻 Lowest",
                showarrow=True,
                arrowhead=2,
                ax=-50,
                ay=0,
                font=dict(color="red", size=14, family="Arial, sans-serif")
            )

            fig_bar.update_layout(
                yaxis=dict(categoryorder="total ascending")
            )

            st.plotly_chart(fig_bar)

    if selected_page == "Dashboard":
        if 'tables_loaded' not in st.session_state:
            st.session_state.tables_loaded = False
            
        if not st.session_state.tables_loaded:
            yesterday_kwh_df = get_yesterday_kwh(xls, st.session_state["active_sheet"], df)
            yesterday_sp_df = get_yesterday_sp(xls, st.session_state["active_sheet"], df)
            yesterday_af_df = get_yesterday_ppr(xls, st.session_state["active_sheet"], df)

            results_df = pd.merge(yesterday_kwh_df, yesterday_sp_df, on="Site Name", how="outer")
            results_df = pd.merge(results_df, yesterday_af_df, on="Site Name", how="outer")
            results_df = results_df.sort_values(by=["PPR", "Specific Production", "kWh"], ascending=False)
            results_df.reset_index(drop=True, inplace=True)

            if len(results_df) < 15:
                results_df = pd.concat([results_df, pd.DataFrame(columns=results_df.columns)], ignore_index=True)
                results_df = results_df.iloc[:15]

            least_performers = results_df.nsmallest(5, "PPR")
            
            if len(results_df) > 5:
                results_df = pd.concat([results_df.iloc[:-5], least_performers], ignore_index=True)
            else:
                results_df = least_performers  

            if "Availability Factor" in results_df.columns:
                results_df["Availability Factor"] = (results_df["Availability Factor"]).round(2)

            results_df.index = range(1, len(results_df) + 1)
            
            last_7_days_df = get_last_7_days_table(xls, st.session_state["active_sheet"], df)
            last_7_days_df.index = range(1, len(last_7_days_df) + 1)
            
            last_30_days_df = get_last_30_days_table(xls, st.session_state["active_sheet"], df)
            last_30_days_df.index = range(1, len(last_30_days_df) + 1)
            
            this_year_df = get_this_year_table(xls, st.session_state["active_sheet"], df)
            this_year_df.index = range(1, len(this_year_df) + 1)
            
            st.session_state.results_df = results_df
            st.session_state.last_7_days_df = last_7_days_df
            st.session_state.last_30_days_df = last_30_days_df
            st.session_state.this_year_df = this_year_df
            
            st.session_state.tables_loaded = True
        
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Yesterday's Ranking per Site")
            st.write(st.session_state.results_df.style.format({
                "kWh": "{:.2f}",
                "Specific Production": "{:.1f}",
                "PPR": "{:.2f}%" 
            }).set_properties(**{'border-color': 'black', 'border-width': '1px'})
            .set_table_styles([dict(selector='th', props=[('font-size', '12pt'), ('text-align', 'center'), ('position', 'sticky'), ('top', '0')])]))

        with col2:
            st.subheader("Last 7 Days' Ranking per Site")
            
            styled_df = st.session_state.last_7_days_df.style.format({
                "kWh": "{:.2f}",
                "Specific Production": "{:.1f}",
                "PPR": "{:.2f}%" 
            }).set_properties(**{
                'border-color': 'black', 
                'border-width': '1px',
                'font-size': '14pt',
                'text-align': 'center'
            }).set_table_styles([
                dict(selector='th', props=[('font-size', '14pt'), ('text-align', 'center'), ('background-color', '#4CAF50'), ('color', ' white')]),
                dict(selector='td', props=[('border', '1px solid black')]), 
                dict(selector='tr:nth-child(even)', props=[('background-color', '#f2f2f2')]),
                dict(selector='tr:nth-child(odd)', props=[('background-color', 'white')]) 
            ])
            
            st.write(styled_df)

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Last 30 Days' Ranking per Site")
            
            styled_30_days_df = st.session_state.last_30_days_df.style.format({
                "kWh": "{:.2f}",
                "Specific Production": "{:.1f}",
                "PPR": "{:.2f}%" 
            }).set_properties(**{
                'border-color': 'black', 
                'border-width': '1px',
                'font-size': '14pt',
                'text-align': 'center' 
            }).set_table_styles([
                dict(selector='th', props=[('font-size', '14pt'), ('text-align', 'center'), ('background-color', '#4CAF50'), ('color', 'white')]),
                dict(selector='td', props=[('border', '1px solid black')]),  
                dict(selector='tr:nth-child(even)', props=[('background-color', '#f2f2f2')]),
                dict(selector='tr:nth-child(odd)', props=[('background-color', 'white')])
            ])
            
            st.write(styled_30_days_df)

        with col2:
            st.subheader("This Year's Ranking per Site")
            
            styled_this_year_df = st.session_state.this_year_df.style.format({
                "kWh": "{:.2f}",
                "Specific Production": "{:.1f}",
                "PPR": "{:.2f}%" 
            }).set_properties(**{
                'border-color': 'black', 
                'border-width': '1px',
                'font-size': '14pt',
                'text-align': 'center' 
            }).set_table_styles([
                dict(selector='th', props=[('font-size', '14pt'), ('text-align', 'center'), ('background-color', '#4CAF50'), ('color', 'white')]),
                dict(selector='td', props=[('border', '1px solid black')]),  
                dict(selector='tr:nth-child(even)', props=[('background-color', '#f2f2f2')]),
                dict(selector='tr:nth-child(odd)', props=[('background-color', 'white')])
            ])
            
            st.write(styled_this_year_df) 
                    
        if st.button("Refresh Tables"):
            st.session_state.tables_loaded = False
            st.rerun()
        st.markdown("#")

        st.subheader("kWh/Target kWh Chart")
        filter_container = st.container()
        
        with filter_container:
            col1, col2, col3 = st.columns(3)
            
            with col1:
                clusters = df["Cluster"].dropna().unique().tolist()
                selected_cluster = st.selectbox("Select Cluster", clusters, key="graph_cluster")

            with col2:
                start_date = st.date_input("Select Start Date", value=datetime.today() - timedelta(days=7), key="graph_start_date")

            with col3:
                end_date = st.date_input("Select End Date", value=datetime.today(), key="graph_end_date")

        start_date = pd.to_datetime(start_date)
        end_date = pd.to_datetime(end_date)

        daily_kwh_df = get_daily_kwh(xls, st.session_state["active_sheet"], df, start_date, end_date)
        target_kwh_df = get_target_kwh(xls, st.session_state["active_sheet"], df, start_date, end_date)

        if selected_cluster:
            daily_kwh_df = daily_kwh_df[daily_kwh_df["Cluster"] == selected_cluster]
            target_kwh_df = target_kwh_df[target_kwh_df["Cluster"] == selected_cluster]

        if not daily_kwh_df.empty and not target_kwh_df.empty:
            daily_kwh_sum = daily_kwh_df.groupby("Date")["kWh"].sum().reset_index().round(2)
            target_kwh_sum = target_kwh_df.groupby("Date")["Target kWh"].sum().reset_index().round(2)

            plot_data = pd.merge(daily_kwh_sum, target_kwh_sum, on="Date", how="outer")

            fig = make_subplots(specs=[[{"secondary_y": True}]])

            fig.add_trace(
                go.Bar(x=plot_data["Date"], y=plot_data["kWh"], name="Daily kWh", marker_color='blue'),
                secondary_y=False,
            )

            fig.add_trace(
                go.Scatter(x=plot_data["Date"], y=plot_data["Target kWh"], name="Target kWh", mode='lines+markers', line=dict(color='red')),
                secondary_y=False,
            )

            fig.update_layout(
                title_text=f"Daily kWh Production vs Target kWh for {selected_cluster}",
                xaxis_title="Date",
                legend=dict(x=0.1, y=1.1, orientation="h"),
            )

            fig.update_yaxes(title_text="Daily kWh", secondary_y=False, tickformat=".2f")
            fig.update_yaxes(title_text="Target kWh", secondary_y=True, tickformat=".2f")

            st.plotly_chart(fig)    
        else:
            st.warning("No matching dates found for daily kWh and target kWh data.")
        filtered_table_df = get_filtered_table(xls, st.session_state["active_sheet"], df, start_date, end_date, selected_cluster)
        filtered_table_df = filtered_table_df.fillna(0)
        if not filtered_table_df.empty:
            st.dataframe(filtered_table_df.style.format({
                "kWh": "{:.2f}",
                "Specific Production": "{:.2f}",
                "PPR": "{:.2f}%",
                "kWp": "{:.2f}"
            }).set_properties(**{
                'border-color': 'black', 
                'border-width': '1px',
                'font-size': '14pt',
                'text-align': 'center' 
            }).set_table_styles([
                dict(selector='th', props=[('font-size', '14pt'), ('text-align', 'center'), ('background-color', '#4CAF50'), ('color', 'white')]),
                dict(selector='td', props=[('border', '1px solid black')]),  
                dict(selector='tr:nth-child(even)', props=[('background-color', '#f2f2f2')]),
                dict(selector='tr:nth-child(odd)', props=[('background-color', 'white')])
            ]))
        else:
            st.warning("No data available for the selected filters.")