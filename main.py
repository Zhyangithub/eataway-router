import pandas as pd
import requests
import urllib.parse
import os

# ================= Configuration =================
API_KEY = "AIzaSyDzAxmeKWeWQGK3G5VVnEDLM0IY-RTjzrw"  # Your real API key
WAREHOUSE_COORD = "59.8542194,17.6650221"

# ================= 1. Data Cleaning and Merging =================
def load_and_merge_data(driver_name="Abbe"):
    print("Starting Excel reader (ignoring file extension disguise)...")
    
    # --- A. Process coordinate table ---
    coords_file = 'coords.xlsx' if os.path.exists('coords.xlsx') else 'coords.csv'
    try:
        # Force using openpyxl engine to read file as Excel
        df_coords = pd.read_excel(coords_file, engine='openpyxl')
    except Exception as e:
        print(f"Failed to read coordinate file: {e}")
        return []
        
    # Dynamically find the real header row containing 'Namn'
    header_row_idx = 0
    for i in range(min(15, len(df_coords))):
        row_vals = [str(v).lower() for v in df_coords.iloc[i].tolist()]
        if any('namn' in str(v) for v in row_vals):
            header_row_idx = i
            break
            
    # Reset header
    df_coords.columns = df_coords.iloc[header_row_idx]
    df_coords = df_coords.iloc[header_row_idx+1:].reset_index(drop=True)
    
    # Keep only first three columns
    df_coords = df_coords.iloc[:, :3]
    df_coords.columns = ['Namn', 'Latitude', 'Longitude']
    
    # Remove duplicates and create dictionary
    df_coords['match_name'] = df_coords['Namn'].astype(str).str.strip().str.lower()
    df_coords = df_coords.drop_duplicates(subset=['match_name'], keep='first')
    coord_dict = df_coords.set_index('match_name')[['Latitude', 'Longitude']].to_dict('index')

    # --- B. Process route table ---
    routes_file = 'routes.xlsx' if os.path.exists('routes.xlsx') else 'routes.csv'
    try:
        df_routes = pd.read_excel(routes_file, engine='openpyxl')
    except Exception as e:
        print(f"Failed to read route file: {e}")
        return []
        
    # Find driver column and row
    driver_col_name = None
    name_row_idx = -1
    
    for col in df_routes.columns:
        # Check column header
        if driver_name.lower() in str(col).lower():
            driver_col_name = col
            break
        # Check cells within column
        for r_idx in range(min(15, len(df_routes))):
            cell_val = str(df_routes[col].iloc[r_idx]).lower()
            if driver_name.lower() in cell_val:
                driver_col_name = col
                name_row_idx = r_idx
                break
        if driver_col_name is not None:
            break
            
    if driver_col_name is None:
        print(f"Driver {driver_name} not found.")
        return []
        
    # Extract store list (skip region title row)
    if name_row_idx == -1:
        raw_stores = df_routes[driver_col_name].iloc[1:].tolist() 
    else:
        raw_stores = df_routes[driver_col_name].iloc[name_row_idx+2:].tolist()
        
    matched_stores = []
    unmatched_stores = []
    
    for store in raw_stores:
        if pd.isna(store) or str(store).strip() == "" or str(store).lower() == 'nan':
            continue
            
        clean_name = str(store).strip().lower()
        if clean_name in coord_dict:
            matched_stores.append({
                "name": str(store).strip(),
                "lat": str(coord_dict[clean_name]['Latitude']),
                "lng": str(coord_dict[clean_name]['Longitude'])
            })
        else:
            unmatched_stores.append(str(store).strip())
            
    print(f"\n[{driver_name}] Total stores found: {len(matched_stores) + len(unmatched_stores)}")
    print(f"Successfully matched coordinates: {len(matched_stores)}")
    if unmatched_stores:
        print(f"Unmatched stores: {len(unmatched_stores)} -> {unmatched_stores}")
        
    return matched_stores


# ================= 2. Request API for Route Optimization =================
def optimize_route(stores):
    print("\nRequesting Google Maps API for optimized route...")
    waypoints_list = [f"{s['lat']},{s['lng']}" for s in stores]
    waypoints_str = "optimize:true|" + "|".join(waypoints_list)

    url = "https://maps.googleapis.com/maps/api/directions/json"
    params = {
        "origin": WAREHOUSE_COORD,
        "destination": WAREHOUSE_COORD, 
        "waypoints": waypoints_str,
        "key": API_KEY
    }

    response = requests.get(url, params=params)
    data = response.json()

    if data['status'] == 'OK':
        optimized_order = data['routes'][0]['waypoint_order']
        optimized_stores = [stores[i] for i in optimized_order]
        return optimized_stores
    else:
        print(f"API error: {data['status']}")
        if 'error_message' in data:
            print(data['error_message'])
        return None


# ================= 3. Generate Cross-Platform Navigation Links =================
def generate_google_maps_urls(optimized_stores):
    urls = []
    base_url = "https://www.google.com/maps/dir/?api=1"
    
    # Convert warehouse into node format
    wh_lat, wh_lng = WAREHOUSE_COORD.split(',')
    warehouse_node = {"name": "Lager (Uppsala)", "lat": wh_lat, "lng": wh_lng}
    
    # Build full path: warehouse -> stores -> warehouse
    full_path = [warehouse_node] + optimized_stores + [warehouse_node]
    
    # Each link can contain at most 11 points
    step = 10 
    
    for i in range(0, len(full_path) - 1, step):
        chunk = full_path[i : i + step + 1]
        
        origin_node = chunk[0]
        dest_node = chunk[-1]
        waypoints = chunk[1:-1]
        
        origin_param = f"&origin={origin_node['lat']},{origin_node['lng']}"
        destination_param = f"&destination={dest_node['lat']},{dest_node['lng']}"
        
        if waypoints:
            waypoints_str = "|".join([f"{s['lat']},{s['lng']}" for s in waypoints])
            waypoints_param = "&waypoints=" + urllib.parse.quote(waypoints_str)
        else:
            waypoints_param = ""
            
        final_url = base_url + origin_param + destination_param + waypoints_param
        urls.append(final_url)
        
    return urls


# ================= Main Execution =================
if __name__ == "__main__":
    drivers = ["Abbe", "Saman", "Sarkis", "Cornelia", "Pawlos"]

    for DRIVER in drivers:
        try:
            stores = load_and_merge_data(DRIVER)
            
            if len(stores) > 0:
                optimized = optimize_route(stores)
                
                if optimized:
                    print("\nFull optimization completed. Final delivery order:")
                    for idx, store in enumerate(optimized, 1):
                        print(f"{idx}. {store['name']}")
                    
                    nav_links = generate_google_maps_urls(optimized)
                    
                    print(f"\nNavigation links for {DRIVER}:")
                    for j, link in enumerate(nav_links, 1):
                        print(f"\nSegment {j}:")
                        print(link)
            else:
                print("\nNo valid store data extracted.")
                
        except Exception as e:
            print(f"\nUnexpected error occurred:\n{e}")