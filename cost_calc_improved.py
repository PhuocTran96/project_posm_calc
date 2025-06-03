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
    """Calculate POSM report with POSM summary and address-based summary."""
    try:
        logging.info("Starting POSM calculation...")
        
        # Merge fact_display with storelist to get address information
        merged_df = fact_display_df.merge(
            dim_storelist_df[['shop', 'Address']], 
            on='shop', 
            how='left'
        )
        
        # Merge with model data to get priority
        merged_df = merged_df.merge(
            dim_model_df[['model', 'priority']], 
            on='model', 
            how='left'
        )
        
        # Merge with POSM data
        merged_df = merged_df.merge(
            dim_posm_df[['model', 'posm']], 
            on='model', 
            how='left'
        )
        
        # Group by POSM and priority first
        posm_grouped = merged_df.groupby(['posm', 'priority'])['quantity'].sum().reset_index()
        
        # Apply quantity rules cho tổng quantity của từng POSM
        posm_grouped['adjusted_quantity'] = posm_grouped.apply(
            lambda row: apply_quantity_rules(row['quantity'], row['priority']), 
            axis=1
        )
        
        # Calculate costs for POSM Summary
        posm_results = []
        for _, row in posm_grouped.iterrows():
            # Calculate original cost (based on original quantity)
            original_cost, original_unit_price = calculate_price(row['posm'], row['quantity'], price_posm_df)
            
            # Calculate adjusted cost (based on adjusted quantity)
            adjusted_cost, adjusted_unit_price = calculate_price(row['posm'], row['adjusted_quantity'], price_posm_df)
            
            posm_results.append({
                'posm': row['posm'],
                'priority': row['priority'],
                'original_quantity': row['quantity'],
                'adjusted_quantity': row['adjusted_quantity'],
                'original_price': original_unit_price,
                'original_cost': original_cost,
                'unit_price': adjusted_unit_price,
                'total_cost': adjusted_cost
            })
        
        posm_summary_df = pd.DataFrame(posm_results)
        
        # Create address-based summary by POSM (chỉ dựa trên original quantity)
        # Group by shop, posm và address để có thông tin chi tiết theo shop
        shop_posm_grouped = merged_df.groupby(['shop', 'posm', 'Address'])['quantity'].sum().reset_index()
        
        # Group by address và posm để tạo address summary (chỉ tính tổng original quantity)
        address_posm_summary = shop_posm_grouped.groupby(['Address', 'posm']).agg({
            'quantity': 'sum'
        }).reset_index()
        
        # Rename columns for clarity
        address_posm_summary = address_posm_summary.rename(columns={
            'Address': 'address',
            'quantity': 'total_quantity'
        })
        
        logging.info("POSM calculation completed successfully")
        return posm_summary_df, address_posm_summary
        
    except Exception as e:
        logging.error(f"Error in calculate_posm_report: {e}")
        return None, None
