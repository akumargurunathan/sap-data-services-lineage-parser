#!/usr/bin/env python3
"""Runner script to execute DS XML extraction in a separate process.
Usage: python ds_engine_runner.py <input_path> <output_path>
Prints simple prefixed messages to stdout for the parent process to parse.
"""
import sys
import json
import traceback
import os

def status_callback(message, level="info"):
    """Callback that prints status messages with prefixes for parent process to parse"""
    if level == 'progress':
        print(f"PROGRESS: {message}", flush=True)
    elif level == 'error':
        print(f"ERROR: {message}", flush=True)
    elif level == 'warning':
        print(f"WARNING: {message}", flush=True)
    elif level == 'status':
        print(f"STATUS: {message}", flush=True)
    else:
        print(f"INFO: {message}", flush=True)

def main():
    # Immediately confirm runner is being called
    print(f"STATUS: [RUNNER] Process started with args: input={sys.argv[1] if len(sys.argv) > 1 else '?'}, output={sys.argv[2] if len(sys.argv) > 2 else '?'}", flush=True)
    
    if len(sys.argv) < 3:
        print("ERROR: Missing arguments. Usage: ds_engine_runner.py <input_path> <output_path>", flush=True)
        sys.exit(2)

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    try:
        status_callback(f"Starting DS XML extraction for {input_path}")
        print("STATUS: Importing ds_xml_engine...", flush=True)
        import ds_xml_engine
        print("STATUS: Imported ds_xml_engine", flush=True)

        status_callback(f"Initializing extraction engine...")
        print("STATUS: Calling extract_ds_lineage_from_path...", flush=True)
        rows, lineage_graph, table_lineage_df, processed_files = ds_xml_engine.extract_ds_lineage_from_path(
            input_path,
            status_callback=status_callback
        )
        print("STATUS: extract_ds_lineage_from_path returned", flush=True)

        if not processed_files:
            status_callback(f"No XML files found in {input_path}", "error")
            sys.exit(3)

        status_callback(f"Found and processed {len(processed_files)} XML file(s)")

        if rows:
            status_callback(f"Extracted {len(rows)} total lineage rows")
            
            status_callback(f"Building table-level lineage...")
            table_lineage_df = ds_xml_engine.extract_table_lineage(rows)
            
            status_callback(f"Exporting results to {output_path}...")
            export_paths = ds_xml_engine.export_ds_results(rows, table_lineage_df, output_path)

            status_callback(f"Export complete")
            print(json.dumps({"processed_files": len(processed_files), "rows": len(rows)}), flush=True)

            for key, path in export_paths.items():
                if path:
                    status_callback(f"{key}: {path}", "info")

            sys.exit(0)
        else:
            status_callback(f"No data extracted from files in {input_path}", "warning")
            sys.exit(4)

    except Exception as e:
        status_callback(f"Extraction failed: {str(e)}", "error")
        traceback.print_exc(file=sys.stdout)
        sys.exit(1)

if __name__ == '__main__':
    main()
