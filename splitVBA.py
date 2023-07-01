import os

def split_vba_macros(input_file):
    with open(input_file, 'r') as f:
        content = f.read()

    blocks = content.split('VBA MACRO ')
    for block in blocks[1:]:  # skip first block as it is empty
        lines = block.split('\n')
        module_name = lines[0].strip()  # Get the module name
        module_content = '\n'.join(lines[1:])  # All other lines
        with open(f'{module_name}.bas', 'w') as f:
            f.write(module_content)

# The path to your file
file_path = r"C:\Users\david\Desktop\Temp\060623 vba\perviousver.txt"
split_vba_macros(file_path)
