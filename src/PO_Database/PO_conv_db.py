"""
Process: Automated Data Archival and Incremental Loading
Description: Scans a designated directory for new Excel reports, cleans the data, 
             and performs a delta-load into a centralized DuckDB warehouse.
"""

import polars as pl
from pathlib import Path
from datetime import datetime 
import duckdb

# --- 1. Configuration & Path Setup ---
# Professional standardized paths for PDD documentation
base_dir = Path(__file__).parent.resolve()
staging_folder = base_dir / "Data_Staging_Area"
db_path = base_dir / "Central_Archive_Repository.db"
target_table = "historical_transaction_logs"

# Ensure staging area exists
staging_folder.mkdir(parents=True, exist_ok=True)
input_files = list(staging_folder.glob("*.xlsx"))

cleaned_buffer = []

# --- 2. Data Ingestion & Sanitization Loop ---
for file_path in input_files:
    try:
        # Metadata Extraction: Capture file modification timestamp for audit trails
        file_stats = file_path.stat()
        mod_timestamp = datetime.fromtimestamp(file_stats.st_mtime)

        # High-speed ingestion using the Calamine engine
        df = pl.read_excel(file_path, engine="calamine")
        
        # Data Sanitization: 
        # 1. Remove systemic artifacts (e.g., 'Textbox' columns from SAP exports)
        # 2. Exclude non-business metadata columns
        clean_cols = [c for c in df.columns if not c.lower().startswith('textbox') and c != 'Source']
        
        # Feature Engineering: Inject audit columns (source origin and modification date)
        df = df.select(clean_cols).with_columns(
            pl.lit(f"{file_path.parent.name}\\{file_path.name}").alias("origin_metadata"),
            pl.lit(mod_timestamp).alias("load_timestamp") 
        )
        
        cleaned_buffer.append(df)
        print(f"Validated: {file_path.name} | Timestamp: {mod_timestamp}")
        
    except Exception as e:
        print(f"System Alert: Error processing {file_path.name}. Exception: {e}")

# --- 3. Consolidation Logic ---
if cleaned_buffer:
    # Use vertical_relaxed to handle slight schema variations across files
    final_batch_df = pl.concat(cleaned_buffer, how="vertical_relaxed")
    print(f"\nBatch Consolidation Complete. Total Records: {len(final_batch_df)}")
else:
    print("Execution Bypassed: No new data found in staging area.")
    final_batch_df = None

# --- 4. Incremental Delta-Load (DuckDB) ---
if final_batch_df is not None:
    conn = duckdb.connect(str(db_path))

    try:
        # Check for existing records based on unique origin metadata
        existing_sources = conn.sql("SELECT DISTINCT origin_metadata FROM historical_transaction_logs").pl()
        
        # ANTI-JOIN LOGIC: Only select data from files that do not exist in the database
        delta_payload = final_batch_df.join(existing_sources, on=['origin_metadata'], how='anti')
        
        if delta_payload.height > 0:
            print(f"Delta-Load Initiated: Syncing {delta_payload.height} new records...")
            conn.sql("INSERT INTO historical_transaction_logs SELECT * FROM delta_payload")
            print("Database synchronization successful.")
        else:
            print("Synchronization Bypassed: All files already exist in archive.")
            
    except duckdb.CatalogException:
        # Repository Initialization: Create table if it does not exist
        print(f"Repository Initialization: Creating {target_table}...")
        conn.sql("CREATE TABLE historical_transaction_logs AS SELECT * FROM final_batch_df")
        print("Master table successfully initialized.")

    finally:
        conn.close()
