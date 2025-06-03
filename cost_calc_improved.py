import pandas as pd
import math
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Constants
BUFFER_MULTIPLIER = 1.3
MINIMUM_QUANTITY = 210
PRICE_RANGES = [
    (0, 200, '<200'),
    (201, 500, '201-500'),
    (501, 1000, '501-1000'),
    (1001, 2000, '1001 - 2000'),
    (2001, 3000, '2001 - 3000'),
    (3001, 4000, '3001 - 4000'),
    (4001, 5000, '4001 - 5000'),
    (5001, float('inf'), '>5000')
]

def apply_quantity_rules(qty, priority):
    """
    Apply different quantity rules based on priority:
    - Priority 1: Apply buffer and enforce minimum quantity
    - Priority 2: Only round up to next multiple of 5
    """
    if priority == 1:
        # Apply buffer and enforce minimum quantity
        return max(
            # Add buffer and round up to next 10
            math.ceil((qty * BUFFER_MULTIPLIER) / 10) * 10,
            # Ensure minimum quantity
            MINIMUM_QUANTITY
        )
    else:  # priority == 2
        # Only round up to next multiple of 5
        return math.ceil(qty / 5) * 5

def get_price_range(quantity):
    """Determine the price range for a given quantity."""
    for min_val, max_val, range_name in PRICE_RANGES:
        if min_val <= quantity <= max_val:
            return range_name
    return PRICE_RANGES[-1][2]  # Return the highest range if no match

def calculate_price(posm_type, quantity, price_df):
    """Calculate price for a specific POSM type and quantity."""
    range_col = get_price_range(quantity)
    
    # Filter price table by POSM type and range
    price_table = price_df[price_df['posm'] == posm_type]
    price_row = price_table[price_table['range'] == range_col]
    
    if not price_row.empty:
        unit_price = price_row['price'].values[0]
        return unit_price * quantity, unit_price
    else:
        logging.warning(f"No price found for POSM {posm_type} with quantity {quantity}")
        return 0, 0

def calculate_posm_report(fact_display_df, dim_storelist_df, dim_model_df, dim_posm_df, price_posm_df):
    """Calculate POSM report with multiple summaries including Care and SDA categories."""
    try:
        logging.info("Starting POSM calculation...")
        
        # Merge all data sources
        merged_df = fact_display_df.merge(
            dim_storelist_df[['shop', 'Address']], 
            on='shop', 
            how='left'
        ).merge(
            dim_model_df[['model', 'priority', 'subcategory']], 
            on='model', 
            how='left'
        ).merge(
            dim_posm_df[['model', 'posm']], 
            on='model', 
            how='left'
        )

        # ================== POSM Summary ==================
        # Bước 1: Group by posm và priority để áp dụng logic buffer
        posm_priority_grouped = merged_df.groupby(['posm', 'priority'])['quantity'].sum().reset_index()
        posm_priority_grouped['adjusted_quantity'] = posm_priority_grouped.apply(
            lambda row: apply_quantity_rules(row['quantity'], row['priority']), 
            axis=1
        )
        
        # Bước 2: Group by chỉ theo posm (bỏ priority) để tính tổng cuối cùng
        posm_final_grouped = posm_priority_grouped.groupby(['posm']).agg({
            'quantity': 'sum',
            'adjusted_quantity': 'sum'
        }).reset_index()
        
        # Calculate costs for POSM Summary
        posm_results = []
        for _, row in posm_final_grouped.iterrows():
            original_cost, original_unit_price = calculate_price(row['posm'], row['quantity'], price_posm_df)
            adjusted_cost, adjusted_unit_price = calculate_price(row['posm'], row['adjusted_quantity'], price_posm_df)
            
            posm_results.append({
                'posm': row['posm'],
                'original_quantity': row['quantity'],
                'adjusted_quantity': row['adjusted_quantity'],
                'original_price': original_unit_price,
                'original_cost': original_cost,
                'unit_price': adjusted_unit_price,
                'total_cost': adjusted_cost
            })
        posm_summary_df = pd.DataFrame(posm_results)

        # ================== Model-POSM Summary ==================
        # SỬA: Sử dụng cùng logic group như POSM Summary để đảm bảo consistency
        model_posm_grouped = merged_df.groupby(['model', 'posm', 'priority'])['quantity'].sum().reset_index()
        model_posm_grouped['adjusted_quantity'] = model_posm_grouped.apply(
            lambda row: apply_quantity_rules(row['quantity'], row['priority']), 
            axis=1
        )
        
        model_posm_results = []
        for _, row in model_posm_grouped.iterrows():
            original_cost, original_unit_price = calculate_price(row['posm'], row['quantity'], price_posm_df)
            adjusted_cost, adjusted_unit_price = calculate_price(row['posm'], row['adjusted_quantity'], price_posm_df)
            
            model_posm_results.append({
                'model': row['model'],
                'posm': row['posm'],
                'priority': row['priority'],
                'original_quantity': row['quantity'],
                'adjusted_quantity': row['adjusted_quantity'],
                'original_price': original_unit_price,
                'original_cost': original_cost,
                'unit_price': adjusted_unit_price,
                'total_cost': adjusted_cost
            })
        model_posm_summary_df = pd.DataFrame(model_posm_results)

        # ================== Address Summary ==================
        # SỬA: Sử dụng cùng base data để đảm bảo consistency
        address_summary = merged_df.groupby(['Address', 'posm'])['quantity'].sum().reset_index()
        address_summary = address_summary.rename(columns={
            'Address': 'address',
            'quantity': 'total_quantity'
        })

        # ================== POSM by Care by Address ==================
        care_subcat = ['Front Load', 'Dryer']
        care_data = merged_df[merged_df['subcategory'].isin(care_subcat)]
        
        care_summary = care_data.groupby(['Address', 'posm'])['quantity'].sum().reset_index()
        care_summary = care_summary.rename(columns={
            'Address': 'address',
            'quantity': 'total_quantity'
        })
        care_summary['category'] = 'Care'

        # ================== POSM by SDA by Address ==================
        sda_data = merged_df[~merged_df['subcategory'].isin(care_subcat)]
        
        sda_summary = sda_data.groupby(['Address', 'posm'])['quantity'].sum().reset_index()
        sda_summary = sda_summary.rename(columns={
            'Address': 'address',
            'quantity': 'total_quantity'
        })
        sda_summary['category'] = 'SDA'

        logging.info("POSM calculation completed successfully")
        return posm_summary_df, model_posm_summary_df, address_summary, care_summary, sda_summary
        
    except Exception as e:
        logging.error(f"Error in calculate_posm_report: {e}")
        return None, None, None, None, None

