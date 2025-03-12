import onnxruntime
from transformers import T5Tokenizer

# Initialize the tokenizer from Hugging Face Transformers
tokenizer = T5Tokenizer.from_pretrained('t5-small')

# Load the ONNX model (ensure the file is available in your working directory)
session = onnxruntime.InferenceSession("t5-small-awesome-text-to-sql.onnx",
                                       providers=['CPUExecutionProvider'])

def generate_sql_onnx(table_name: str, columns: list, prompt: str) -> str:
    """
    Constructs the input prompt using the table name, list of column definitions,
    and the natural language query, then runs the ONNX model to generate the SQL query.
    
    Args:
        table_name: Name of the table (e.g., "bank_accounts")
        columns: List of column definitions (e.g., ["account_id VARCHAR", "account_holder VARCHAR",
                                                    "balance FLOAT", "account_type VARCHAR"])
        prompt: Natural language query (e.g., "List the account holder names with balance greater than 1000")
    
    Returns:
        Generated SQL query as a string.
    """
    # Convert the list of columns into a comma-separated string.
    columns_str = ", ".join(columns)
    
    # Construct the model input.
    # The prompt format here is:
    # "tables:" followed by the table name and its columns, then "query for:" followed by the natural language query.
    model_input = f"tables:\n{table_name} ({columns_str})\nquery for: {prompt}"
    
    # Tokenize the input prompt; return tensors as NumPy arrays for ONNX Runtime.
    inputs = tokenizer(model_input, return_tensors="np", padding=True, truncation=True)
    
    # Prepare the inputs dictionary for the ONNX model.
    ort_inputs = {
        session.get_inputs()[0].name: inputs['input_ids'],
        session.get_inputs()[1].name: inputs['attention_mask']
    }
    
    # Run the ONNX inference session.
    ort_outs = session.run(None, ort_inputs)
    
    # Decode the output token IDs to obtain the SQL query.
    output_ids = ort_outs[0]
    generated_sql = tokenizer.decode(output_ids[0], skip_special_tokens=True)
    
    return generated_sql

# Example usage with dummy bank data
if __name__ == "__main__":
    # Define the inputs separately
    table_name = "bank_accounts"
    columns = [
        "account_id VARCHAR",
        "account_holder VARCHAR",
        "balance FLOAT",
        "account_type VARCHAR"
    ]
    prompt = "List the account holder names with balance greater than 1000"
    
    # Generate the SQL query using ONNX Runtime for fast CPU inference
    sql_query = generate_sql_onnx(table_name, columns, prompt)
    
    print("Generated SQL Query:")
    print(sql_query)