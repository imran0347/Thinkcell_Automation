import os

def to_camel_case(snake_str):
    components = snake_str.split('_')
    return ''.join(x.title() for x in components)

def format_key_path_value(table_name, format_type):
    if format_type == '1':
        return table_name.upper()
    elif format_type == '2':
        return table_name.lower()
    elif format_type == '3':
        return to_camel_case(table_name)
    else:
        raise ValueError("Invalid format type. Please enter 1 for uppercase, 2 for lowercase, or 3 for camelcase.")

def generate_output(table_names, source_name, key_path_format, key_path_name):
    table_list = table_names.split(',')
    outputs = []
    
    for table_name in table_list:
        table_name = table_name.strip()  # Remove any leading/trailing whitespace
        camel_case_table_name = to_camel_case(table_name)
        formatted_key_path_value = format_key_path_value(table_name, key_path_format)
        output = f"""
        - data_contract: app.src.data_structures.data_contracts.source.{source_name.lower()}.v0.{source_name.lower()}_{table_name}_source_data_contract
          class_name: {source_name}{camel_case_table_name}Model
          data_contract_version: 0.0.1
          key_path:
            - name: {key_path_name}
              value: {formatted_key_path_value}
          router_schema: *schema
          router_table: {table_name.upper()}
"""
        outputs.append(output.strip())
    
    return '\n'.join(outputs)

def save_to_yaml(data, filename):
    with open(filename, 'w') as file:
        file.write(data)

# Example usage:
table_names = input("Enter table names (comma-separated): ")
source_name = input("Enter source name: ")
key_path_format = input("Enter key path format (1 for uppercase, 2 for lowercase, 3 for camelcase): ")
key_path_name = input("Enter key path name: ")

output_data = generate_output(table_names, source_name, key_path_format, key_path_name)

# Save to the specified folder
output_folder = r"C:\Users\nitin.singh\Downloads"
os.makedirs(output_folder, exist_ok=True)
output_file_path = os.path.join(output_folder, 'output.yaml')

save_to_yaml(output_data, output_file_path)

print(f"Output saved to {output_file_path}")
