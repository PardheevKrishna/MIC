import os
import glob
import base64
import zipfile

def reassemble_chunks(chunk_pattern, combined_filename):
    """
    Reassembles chunk files matching the given glob pattern into one file.
    
    Args:
        chunk_pattern (str): Glob pattern for the chunk files.
        combined_filename (str): Output file name for the reassembled file.
    Returns:
        bool: True if reassembly succeeded, False otherwise.
    """
    chunk_files = sorted(glob.glob(chunk_pattern))
    if not chunk_files:
        print(f"[DEBUG] No chunk files found with pattern: {chunk_pattern}")
        return False
    
    print(f"[DEBUG] Found {len(chunk_files)} chunk files. Reassembling into '{combined_filename}'...")
    
    with open(combined_filename, "wb") as outfile:
        for chunk_file in chunk_files:
            print(f"[DEBUG] Reading {chunk_file}...")
            with open(chunk_file, "rb") as infile:
                data = infile.read()
                outfile.write(data)
            print(f"[DEBUG] Appended {chunk_file} ({len(data)} bytes)")
    
    print("[DEBUG] Reassembly of chunks complete.")
    return True

def decode_and_extract(base64_file, zip_filename, extract_dir):
    """
    Decodes a Base64-encoded file into a ZIP file and extracts its contents.
    
    Args:
        base64_file (str): The combined Base64 text file.
        zip_filename (str): The output ZIP file name.
        extract_dir (str): Directory where the ZIP contents will be extracted.
    """
    print(f"[DEBUG] Reading combined Base64 file: {base64_file}...")
    with open(base64_file, "rb") as f:
        b64_data = f.read()
    print(f"[DEBUG] Combined Base64 file has {len(b64_data)} bytes.")
    
    print("[DEBUG] Decoding Base64 data to binary ZIP format...")
    try:
        zip_data = base64.b64decode(b64_data)
    except Exception as e:
        print(f"[ERROR] Decoding failed: {e}")
        return
    print(f"[DEBUG] Decoded ZIP data size: {len(zip_data)} bytes.")
    
    print(f"[DEBUG] Writing ZIP data to file: {zip_filename}...")
    with open(zip_filename, "wb") as f:
        f.write(zip_data)
    print(f"[DEBUG] ZIP file '{zip_filename}' written successfully.")
    
    print(f"[DEBUG] Creating output directory: {extract_dir} (if not exists)...")
    os.makedirs(extract_dir, exist_ok=True)
    
    print(f"[DEBUG] Extracting '{zip_filename}' to directory '{extract_dir}'...")
    with zipfile.ZipFile(zip_filename, "r") as zip_ref:
        zip_ref.extractall(extract_dir)
    print(f"[DEBUG] Extraction complete. Model restored in '{extract_dir}'.")

if __name__ == "__main__":
    # Step 1: Reassemble chunk files into one combined Base64 file.
    # Adjust the chunk_pattern if your chunk file naming convention is different.
    chunk_pattern = "model_base64.txt_chunk_*.txt"  # Example pattern for 12 chunks.
    combined_base64_file = "model_base64_combined.txt"
    
    reassemble_success = reassemble_chunks(chunk_pattern, combined_base64_file)
    if not reassemble_success:
        print("[ERROR] Reassembly failed. Exiting.")
        exit(1)
    
    # Step 2: Decode the combined Base64 file back into a ZIP file and extract it.
    zip_file = "cssupport_t5_small_awesome_text_to_sql.zip"
    output_directory = "cssupport_t5_small_awesome_text_to_sql"
    decode_and_extract(combined_base64_file, zip_file, output_directory)