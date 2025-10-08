"""
Utility functions for DSS Template processing
"""

import pandas as pd
import re

def find_column(df, possible_names):
    """Find column by checking multiple possible names (case-insensitive, space-insensitive)"""
    if df is None or df.empty:
        return None
    
    df_columns_lower = {col.lower().strip().replace(' ', '').replace('_', ''): col for col in df.columns}
    
    for name in possible_names:
        name_normalized = name.lower().strip().replace(' ', '').replace('_', '')
        if name_normalized in df_columns_lower:
            return df_columns_lower[name_normalized]
    return None

def safe_get_value(df, column_name, row_idx=0, default=''):
    """Safely get value from dataframe with multiple fallback mechanisms"""
    if df is None or column_name is None or column_name not in df.columns:
        return default
    try:
        if len(df) <= row_idx:
            return default
        val = df[column_name].iloc[row_idx]
        if pd.isna(val):
            return default
        return val
    except:
        return default

def safe_load_sheet(excel_file, sheet_name, alternative_names=None):
    """Safely load Excel sheet with multiple possible names"""
    try:
        excel_data = pd.ExcelFile(excel_file)
        available_sheets = excel_data.sheet_names
        
        if sheet_name in available_sheets:
            return pd.read_excel(excel_file, sheet_name=sheet_name)
        
        sheet_lower = sheet_name.lower().strip()
        for avail_sheet in available_sheets:
            if avail_sheet.lower().strip() == sheet_lower:
                return pd.read_excel(excel_file, sheet_name=avail_sheet)
        
        if alternative_names:
            for alt_name in alternative_names:
                alt_lower = alt_name.lower().strip()
                for avail_sheet in available_sheets:
                    if avail_sheet.lower().strip() == alt_lower:
                        return pd.read_excel(excel_file, sheet_name=avail_sheet)
        
        return None
    except:
        return None

def process_template(excel_path, template_path, log_callback=None):
    """
    Main processing function
    Returns: (filled_content, replacements_dict, warnings_list)
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
    
    warnings = []
    replacements = {}
    
    # Load Excel
    log("\n📂 Loading Excel file...")
    try:
        excel_data = pd.ExcelFile(excel_path)
        log(f"✓ Found {len(excel_data.sheet_names)} worksheets")
    except Exception as e:
        raise Exception(f"Cannot open Excel file: {e}")
    
    # Load sheets
    mixed_mode_df = safe_load_sheet(excel_path, 'Mixed Mode Info', ['MixedModeInfo', 'Mixed_Mode_Info'])
    five_g_df = safe_load_sheet(excel_path, '5G Info', ['5GInfo', '5G_Info'])
    eutran_df = safe_load_sheet(excel_path, 'eUtran Parameters', ['EUtranParameters', 'eUtran_Parameters'])
    
    sheets_loaded = []
    if mixed_mode_df is not None: sheets_loaded.append("Mixed Mode Info")
    if five_g_df is not None: sheets_loaded.append("5G Info")
    if eutran_df is not None: sheets_loaded.append("eUtran Parameters")
    
    log(f"✓ Loaded: {', '.join(sheets_loaded)}")
    
    # Load template
    log("\n📄 Loading template...")
    with open(template_path, 'r', encoding='utf-8') as f:
        template_content = f.read()
    
    placeholder_pattern = r'xx[A-Za-z0-9_]+xx'
    all_placeholders = re.findall(placeholder_pattern, template_content)
    log(f"✓ Found {len(set(all_placeholders))} unique placeholders")
    
    # Extract Primary Node
    log("\n🔍 Extracting primary node...")
    if mixed_mode_df is not None:
        cabinet_col = find_column(mixed_mode_df, ['Cabinet Controlling DUL', 'CabinetControllingDUL'])
        enb_name_col = find_column(mixed_mode_df, ['eNodeB Name', 'eNodeBName'])
        enb_id_col = find_column(mixed_mode_df, ['eNBId', 'eNBID'])
        gnb_name_col = find_column(mixed_mode_df, ['gNodeB Name', 'gNodeBName'])
        gnb_id_col = find_column(mixed_mode_df, ['gNBId', 'gNBID'])
        
        primary_node = None
        
        # Try finding primary node
        if cabinet_col:
            try:
                primary_rows = mixed_mode_df[mixed_mode_df[cabinet_col] == True]
                if len(primary_rows) > 0:
                    primary_node = primary_rows.iloc[0]
            except:
                pass
        
        if primary_node is None and enb_name_col:
            for idx, row in mixed_mode_df.iterrows():
                enb_val = safe_get_value(pd.DataFrame([row]), enb_name_col, 0)
                if enb_val and str(enb_val).strip():
                    primary_node = row
                    break
        
        if primary_node is not None:
            site_name = str(safe_get_value(pd.DataFrame([primary_node]), enb_name_col, 0)).strip()
            if site_name:
                replacements['xxMMBB_Primary_Node_Namexx'] = site_name
                replacements['xxLTE_Site_IDxx'] = site_name
                log(f"✓ Primary Node: {site_name}")
            
            enb_id_val = safe_get_value(pd.DataFrame([primary_node]), enb_id_col, 0)
            if pd.notna(enb_id_val):
                replacements['xxLTE_eNBIDxx'] = str(int(float(enb_id_val))).strip()
            
            gnb_name = str(safe_get_value(pd.DataFrame([primary_node]), gnb_name_col, 0)).strip()
            if gnb_name:
                replacements['xx5G_NR_Node_Namexx'] = gnb_name
            
            gnb_id_val = safe_get_value(pd.DataFrame([primary_node]), gnb_id_col, 0)
            if pd.notna(gnb_id_val):
                replacements['xx5G_NR_gNBIDxx'] = str(int(float(gnb_id_val))).strip()
    
    # Extract DSS Cells
    log("\n🔍 Extracting DSS cells...")
    if five_g_df is not None and 'xx5G_NR_Node_Namexx' in replacements:
        try:
            gnb_name = replacements['xx5G_NR_Node_Namexx']
            gnb_name_5g_col = find_column(five_g_df, ['gNB Name', 'gNodeB Name'])
            dss_col = find_column(five_g_df, ['DSS', 'dss'])
            nr_cell_col = find_column(five_g_df, ['NRCellDU', 'NRCellDu'])
            cell_local_id_col = find_column(five_g_df, ['cellLocalId', 'celllocalid'])
            nr_sector_col = find_column(five_g_df, ['NRSectorCarrier', 'nrsectorcarrier'])
            
            if dss_col and gnb_name_5g_col and nr_cell_col:
                nr_dss_cells = five_g_df[
                    (five_g_df[gnb_name_5g_col] == gnb_name) &
                    (five_g_df[dss_col].notna()) &
                    (five_g_df[dss_col] != 'NO')
                ].sort_values(nr_cell_col)
                
                if len(nr_dss_cells) >= 3:
                    alpha_nr = nr_dss_cells.iloc[0]
                    beta_nr = nr_dss_cells.iloc[1]
                    gamma_nr = nr_dss_cells.iloc[2]
                    
                    lte_alpha = str(alpha_nr[dss_col]).strip()
                    lte_beta = str(beta_nr[dss_col]).strip()
                    lte_gamma = str(gamma_nr[dss_col]).strip()
                    
                    nr_alpha = str(alpha_nr[nr_cell_col]).strip()
                    nr_beta = str(beta_nr[nr_cell_col]).strip()
                    nr_gamma = str(gamma_nr[nr_cell_col]).strip()
                    
                    lte_band_match = re.search(r'_(\d+)[ABC]_', lte_alpha)
                    nr_band_match = re.search(r'_(N\d{3})[ABC]_', nr_alpha)
                    
                    if lte_band_match and nr_band_match:
                        lte_band = lte_band_match.group(1)
                        nr_band = nr_band_match.group(1)
                        
                        if cell_local_id_col:
                            replacements['xx5G_celllocalidAxx'] = str(int(alpha_nr[cell_local_id_col]))
                            replacements['xx5G_celllocalidBxx'] = str(int(beta_nr[cell_local_id_col]))
                            replacements['xx5G_celllocalidCxx'] = str(int(gamma_nr[cell_local_id_col]))
                        
                        if nr_sector_col:
                            replacements['xx5G_NRSectorCarrier_Alphaxx'] = str(alpha_nr[nr_sector_col])
                            replacements['xx5G_NRSectorCarrier_Betaxx'] = str(beta_nr[nr_sector_col])
                            replacements['xx5G_NRSectorCarrier_Gammaxx'] = str(gamma_nr[nr_sector_col])
                        
                        replacements['xxLTE_Site_IDxx_XA_1'] = lte_alpha
                        replacements['xxLTE_Site_IDxx_XB_1'] = lte_beta
                        replacements['xxLTE_Site_IDxx_XC_1'] = lte_gamma
                        
                        replacements['xx5G_NR_Node_Namexx_N00XA_1'] = nr_alpha
                        replacements['xx5G_NR_Node_Namexx_N00XB_1'] = nr_beta
                        replacements['xx5G_NR_Node_Namexx_N00XC_1'] = nr_gamma
                        
                        if 'xxLTE_Site_IDxx' in replacements:
                            site_id = replacements['xxLTE_Site_IDxx']
                            replacements['xxMMBB_Primary_Node_Namexx_N00XA_1'] = f"{site_id}_{nr_band}A_1"
                            replacements['xxMMBB_Primary_Node_Namexx_N00XB_1'] = f"{site_id}_{nr_band}B_1"
                            replacements['xxMMBB_Primary_Node_Namexx_N00XC_1'] = f"{site_id}_{nr_band}C_1"
                            replacements['xxLTE_Site_IDxx_X*'] = f"{site_id}_X*"
                        
                        replacements['N00XA'] = f"{nr_band}A"
                        replacements['N00XB'] = f"{nr_band}B"
                        replacements['N00XC'] = f"{nr_band}C"
                        
                        log(f"✓ Found {len(nr_dss_cells)} DSS cells")
                        log(f"✓ Bands: LTE {lte_band}, 5G {nr_band}")
                        
                        # Get LTE cell details
                        if eutran_df is not None:
                            eutran_cell_col = find_column(eutran_df, ['EutranCellFDDId', 'EUtranCellFDDId'])
                            cell_id_col = find_column(eutran_df, ['cellId', 'cellid'])
                            sector_id_col = find_column(eutran_df, ['sectorId', 'sectorid'])
                            
                            if eutran_cell_col:
                                lte_alpha_row = eutran_df[eutran_df[eutran_cell_col] == lte_alpha]
                                lte_beta_row = eutran_df[eutran_df[eutran_cell_col] == lte_beta]
                                lte_gamma_row = eutran_df[eutran_df[eutran_cell_col] == lte_gamma]
                                
                                if len(lte_alpha_row) > 0 and cell_id_col and sector_id_col:
                                    replacements['LTE_cellidA'] = str(int(lte_alpha_row.iloc[0][cell_id_col]))
                                    replacements['LTE_cellidB'] = str(int(lte_beta_row.iloc[0][cell_id_col]))
                                    replacements['LTE_cellidC'] = str(int(lte_gamma_row.iloc[0][cell_id_col]))
                                    
                                    replacements['xxLTE_SectorCarrier_No_Alphaxx'] = str(int(lte_alpha_row.iloc[0][sector_id_col]))
                                    replacements['xxLTE_SectorCarrier_No_Betaxx'] = str(int(lte_beta_row.iloc[0][sector_id_col]))
                                    replacements['xxLTE_SectorCarrier_No_Gammaxx'] = str(int(lte_gamma_row.iloc[0][sector_id_col]))
                                    
                                    log(f"✓ Extracted LTE cell details")
        except Exception as e:
            warnings.append(f"Error extracting DSS cells: {e}")
    
    # Replace placeholders
    log("\n🔄 Replacing placeholders...")
    filled_content = template_content
    replacement_count = {}
    
    sorted_replacements = sorted(replacements.items(), key=lambda x: len(x[0]), reverse=True)
    
    for placeholder, value in sorted_replacements:
        count = filled_content.count(placeholder)
        if count > 0:
            filled_content = filled_content.replace(placeholder, value)
            replacement_count[placeholder] = count
    
    log(f"✓ Replaced {len(replacement_count)} placeholders")
    
    return filled_content, replacement_count, warnings
