import sys
import os
import subprocess

def convert_image_to_csharp_hex(file_path, bytes_per_line=64, indent_spaces=8):
    """
    Reads a file and returns it as a C# formatted byte array hex dump string.
    """
    if not os.path.exists(file_path):
        error_msg = f"Error: File not found at '{file_path}'"
        print(error_msg)
        return error_msg

    try:
        with open(file_path, 'rb') as f:
            byte_data = f.read()
    except Exception as e:
        error_msg = f"Error reading file: {e}"
        print(error_msg)
        return error_msg

    # Create a list of hex strings (e.g., "0x1b")
    hex_values = [f"0x{b:02x}" for b in byte_data]
    
    total_bytes = len(hex_values)
    indentation = " " * indent_spaces
    
    lines = []
    # Process data in chunks to create lines
    for i in range(0, total_bytes, bytes_per_line):
        # Get the current chunk of bytes
        chunk = hex_values[i:i + bytes_per_line]
        
        # Join them with commas
        line_content = ", ".join(chunk)
        
        # Determine if we need a trailing comma (all lines except the absolute last one)
        if i + bytes_per_line < total_bytes:
            lines.append(f"{indentation}{line_content},")
        else:
            # Last line, no trailing comma usually required for the array end
            lines.append(f"{indentation}{line_content}")
    
    return "\n".join(lines)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python pngjpgconvert.py <path_to_image.jpg>")
    else:
        input_image = sys.argv[1]
        output = convert_image_to_csharp_hex(input_image)
        if output.startswith("Error"):
            # Error already printed
            pass
        else:
            print(output)  # Print the hex dump to console
            # Copy to clipboard using Windows clip.exe
            try:
                subprocess.run(['clip'], input=output, text=True, check=True)
                print("Hex dump copied to clipboard.")
            except subprocess.CalledProcessError as e:
                print(f"Failed to copy to clipboard: {e}")