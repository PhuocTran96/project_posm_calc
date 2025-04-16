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

def load_data():
    """Load all required data from Excel files."""
    try:
        fact_display = pd.read_excel('fact_display.xlsx')
        dim_storelist = pd.read_excel('dim_storelist.xlsx').rename(columns={"Store name": "shop"})
        dim_model = pd.read_excel('dim_model.xlsx')
        dim_posm = pd.read_excel('dim_posm.xlsx', sheet_name='posm')
        price_posm = pd.read_excel('dim_posm.xlsx', sheet_name='price')
        
        return fact_display, dim_storelist, dim_model, dim_posm, price_posm
    except Exception as e:
        logging.error(f"Error loading data: {e}")
        raise

def apply_buffer_and_minimum(qty):
    """Apply buffer and enforce minimum quantity."""
    return max(
        # Add buffer and round up to next 10
        math.ceil((qty * BUFFER_MULTIPLIER) / 10) * 10,
        # Ensure minimum quantity
        MINIMUM_QUANTITY
    )

def get_price_range(quantity):
    """Determine the price range for a given quantity."""
    for min_val, max_val, range_name in PRICE_RANGES:
        if quantity <= max_val:
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

def calculate_send_quantity(quantity, original_quantity):
    """
    Calculate the 'send' quantity which is:
    - 70% of the adjusted quantity
    - But must be at least the original quantity
    - And must be divisible by 10
    """
    # Calculate 70% of quantity
    send = int(quantity * 0.7)
    
    # Ensure it's at least the original quantity
    send = max(send, original_quantity)
    
    # Make it divisible by 10 (round up to next multiple of 10)
    if send % 10 != 0:
        send = math.ceil(send / 10) * 10
        
    # Ensure send doesn't exceed quantity
    send = min(send, quantity)
    
    return send

def calculate_posm_report(fact_display, dim_storelist, dim_model, dim_posm, price_posm):
    """
    Calculates POSM costs and province allocations based on input dataframes.

    Args:
        fact_display (pd.DataFrame): DataFrame from fact_display.xlsx
        dim_storelist (pd.DataFrame): DataFrame from dim_storelist.xlsx
        dim_model (pd.DataFrame): DataFrame from dim_model.xlsx
        dim_posm (pd.DataFrame): DataFrame from dim_posm.xlsx (posm sheet)
        price_posm (pd.DataFrame): DataFrame from dim_posm.xlsx (price sheet)

    Returns:
        tuple: (final_df, province_summary_df) containing the results.
               Returns (None, None) if an error occurs during calculation.
    """
    try:
        # --- Start of calculation logic ---
    
        # Calculate total quantity by model
        model_quantity = fact_display.groupby('model')['quantity'].sum().reset_index()
        
        # Merge model quantities with POSM requirements
        model_posm = pd.merge(
            model_quantity,
            dim_posm,
            on='model',
            how='left'
        )
        
        # Calculate total quantity needed for each POSM type
        posm_by_model = model_posm.copy()
        posm_quantity = posm_by_model.groupby('posm')['quantity'].sum().reset_index()
        
        # Get POSM names
        dim_posm_unique = price_posm.drop_duplicates(subset=['posm'])
        posm_quantity = pd.merge(
            posm_quantity,
            dim_posm_unique,
            on='posm',
            how='left'
        )
        posm_quantity = posm_quantity[['posm', 'name', 'quantity']]
        
        # Save original quantities and apply buffer
        posm_quantity['original_quantity'] = posm_quantity['quantity']
        posm_quantity['quantity'] = posm_quantity['quantity'].apply(apply_buffer_and_minimum)
        
        # Get model information
        model_info = pd.merge(
            model_quantity,
            dim_model,
            on='model',
            how='left'
        )
        
        # Calculate province-level missing POSM data
        # First, merge fact_display with dim_storelist to get province information
        display_with_store = pd.merge(
            fact_display,
            dim_storelist,
            on='shop',
            how='left'
        )
        
        # Calculate total POSM needed by province and model
        province_model_qty = display_with_store.groupby(['Province', 'model'])['quantity'].sum().reset_index()
        
        # Merge with POSM requirements to get POSM needed by province
        province_posm = pd.merge(
            province_model_qty,
            dim_posm,
            on='model',
            how='left'
        )
        
        # Calculate total POSM needed by province
        province_posm_qty = province_posm.groupby(['Province', 'posm'])['quantity'].sum().reset_index()
        
        # First, calculate province allocations based on need
        province_summary_data = []
        
        # Dictionary to store total send quantities by POSM type
        posm_send_quantities = {}
        
        # Process each POSM type and calculate province allocations
        for _, row in posm_quantity.iterrows():
            posm_type = row['posm']
            adjusted_qty = row['quantity']
            original_qty = row['original_quantity']
            
            # Initialize total send for this POSM type
            posm_send_quantities[posm_type] = 0
            
            # Calculate province distribution for this POSM type
            if posm_type in province_posm_qty['posm'].values:
                # Get province data for this POSM type
                provinces_for_posm = province_posm_qty[province_posm_qty['posm'] == posm_type]
                
                # Calculate total quantity needed across all provinces
                total_province_need = provinces_for_posm['quantity'].sum()
                
                # Calculate percentage distribution by province
                for _, prov_row in provinces_for_posm.iterrows():
                    province = prov_row['Province']
                    province_qty = prov_row['quantity']
                    
                    # Calculate percentage of total need
                    if total_province_need > 0:
                        province_pct = province_qty / total_province_need
                    else:
                        province_pct = 0
                    
                    # Calculate initial allocation based on need percentage
                    # Make sure it's rounded to the nearest multiple of 5
                    province_allocation = math.ceil(province_qty / 5) * 5
                    
                    # Add to the total send quantity for this POSM type
                    posm_send_quantities[posm_type] += province_allocation
                    
                    province_summary_data.append({
                        'Province': province,
                        'posm': posm_type,
                        'posm_name': row['name'],
                        'needed_quantity': province_qty,
                        'percentage_of_total': province_pct * 100,
                        'allocated_quantity': province_allocation
                    })
        
        # Now calculate the final results with the new send quantities
        results = []
        for _, row in posm_quantity.iterrows():
            posm_type = row['posm']
            adjusted_qty = row['quantity']
            original_qty = row['original_quantity']
            
            # Get the send quantity from our calculated province allocations
            # If no provinces need this POSM, use the original calculation
            if posm_type in posm_send_quantities and posm_send_quantities[posm_type] > 0:
                send_qty = posm_send_quantities[posm_type]
                
                # Ensure send quantity is at least the original quantity
                send_qty = max(send_qty, original_qty)
                
                # Ensure send quantity doesn't exceed adjusted quantity
                send_qty = min(send_qty, adjusted_qty)
            else:
                # Fallback to original calculation if no province data
                send_qty = calculate_send_quantity(adjusted_qty, original_qty)
            
            # Calculate backup quantity
            backup_qty = adjusted_qty - send_qty
            
            # Calculate costs for adjusted quantity
            total_cost, unit_price = calculate_price(posm_type, adjusted_qty, price_posm)
            
            # Calculate costs for original quantity
            original_total_cost, original_unit_price = calculate_price(posm_type, original_qty, price_posm)
            
            results.append({
                'posm': posm_type,
                'name': row['name'],
                'original_quantity': original_qty,
                'quantity': adjusted_qty,
                'send': send_qty,
                'backup': backup_qty,
                'original_unit_price': original_unit_price,
                'unit_price': unit_price,
                'original_total_cost': original_total_cost,
                'total_cost': total_cost
            })
        
        # Create final dataframes
        final_df = pd.DataFrame(results)
        province_summary_df = pd.DataFrame(province_summary_data)
        # --- End of calculation logic ---

        return final_df, province_summary_df

    except Exception as e:
        logging.error(f"Error during calculation: {e}")
        return None, None


# Keep the main execution block for standalone script usage if needed
def main():
    # Load data
    try:
        fact_display, dim_storelist, dim_model, dim_posm, price_posm = load_data()
    except Exception:
        return # Error already logged in load_data

    # Perform calculation
    final_df, province_summary_df = calculate_posm_report(
        fact_display, dim_storelist, dim_model, dim_posm, price_posm
    )

    if final_df is not None and province_summary_df is not None:
        # Export to Excel
        output_filename = 'posm_cost_report_new.xlsx'
        try:
            with pd.ExcelWriter(output_filename) as writer:
                final_df.to_excel(writer, sheet_name='POSM Summary', index=False)
                province_summary_df.to_excel(writer, sheet_name='Province Summary', index=False)
            logging.info(f"Report successfully generated: {output_filename}")
        except Exception as e:
            logging.error(f"Error writing Excel report: {e}")
            # Try with a timestamp in the filename if the first attempt failed
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f'posm_cost_report_{timestamp}.xlsx'
            try:
                with pd.ExcelWriter(output_filename) as writer:
                    final_df.to_excel(writer, sheet_name='POSM Summary', index=False)
                    province_summary_df.to_excel(writer, sheet_name='Province Summary', index=False)
                logging.info(f"Report successfully generated with timestamp: {output_filename}")
            except Exception as e2:
                logging.error(f"Second attempt also failed: {e2}")

if __name__ == "__main__":
    main()
