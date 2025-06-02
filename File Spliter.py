import os
import zipfile
import tempfile
from pathlib import Path

# Adjust these constants as needed:
SOURCE_ZIP = "Items.1.001.zip"
OUTPUT_PREFIX = "part"       # will produce part1.zip, part2.zip, …
MAX_SIZE_BYTES = 190 * 1024 * 1024  # 190 MB per chunk

def split_msg_zip(source_zip, output_prefix, max_size):
    # Create a temp directory to extract everything
    with tempfile.TemporaryDirectory() as tmpdir:
        # 1) Extract entire source ZIP into tmpdir
        with zipfile.ZipFile(source_zip, "r") as z:
            z.extractall(tmpdir)

        # 2) Collect all .msg file paths
        msg_paths = []
        for root, _, files in os.walk(tmpdir):
            for name in files:
                if name.lower().endswith(".msg"):
                    msg_paths.append(os.path.join(root, name))

        # 3) Sort or shuffle if you prefer a different order
        msg_paths.sort()

        part_index = 1
        current_zip_path = f"{output_prefix}{part_index}.zip"
        current_zip = zipfile.ZipFile(current_zip_path, mode="w", compression=zipfile.ZIP_DEFLATED)
        current_size = 0

        for path in msg_paths:
            fname = Path(path).name
            file_size = os.path.getsize(path)

            # If adding this file would exceed max_size AND the current ZIP is not empty, finalize and start a new one
            if current_size + file_size > max_size and current_size > 0:
                current_zip.close()
                part_index += 1
                current_zip_path = f"{output_prefix}{part_index}.zip"
                current_zip = zipfile.ZipFile(current_zip_path, mode="w", compression=zipfile.ZIP_DEFLATED)
                current_size = 0

            # Add the .msg into current ZIP
            current_zip.write(path, arcname=fname)
            current_size += file_size

        # Close the final chunk
        current_zip.close()
        print(f"Created {part_index} chunks: {output_prefix}1.zip through {output_prefix}{part_index}.zip")

if __name__ == "__main__":
    split_msg_zip(SOURCE_ZIP, OUTPUT_PREFIX, MAX_SIZE_BYTES)
