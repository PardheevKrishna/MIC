import torch
from transformers import T5ForConditionalGeneration, T5Tokenizer

# --- Debug: Check device ---
device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
print(f"[DEBUG] Using device: {device}")

# --- Load the Restored Model and Tokenizer ---
model_dir = "cssupport_t5_small_awesome_text_to_sql"  # Directory from the extraction step
print(f"[DEBUG] Loading model and tokenizer from: {model_dir}")

model = T5ForConditionalGeneration.from_pretrained(model_dir)
tokenizer = T5Tokenizer.from_pretrained(model_dir)
model = model.to(device)
model.eval()
print("[DEBUG] Model and tokenizer loaded successfully.")

def generate_sql(table_name: str, columns: list, prompt: str) -> str:
    """
    Generates an SQL query based on the table name, list of columns, and natural language prompt.
    
    Args:
        table_name (str): The name of the table.
        columns (list): A list of column definitions (e.g., ["id VARCHAR", "name VARCHAR"]).
        prompt (str): The natural language query.
    
    Returns:
        str: The generated SQL query.
    """
    # Create a comma-separated string from the list of columns.
    columns_str = ", ".join(columns)
    
    # Construct the complete model input using the expected format.
    model_input = f"tables:\n{table_name} ({columns_str})\nquery for: {prompt}"
    print(f"[DEBUG] Constructed model input:\n{model_input}")
    
    # Tokenize the input string.
    inputs = tokenizer(model_input, return_tensors="pt", padding=True, truncation=True)
    inputs = {k: v.to(device) for k, v in inputs.items()}
    print("[DEBUG] Tokenization complete.")
    
    # Generate output tokens using the model.
    with torch.no_grad():
        output_ids = model.generate(**inputs, max_length=512)
    print("[DEBUG] Model inference complete.")
    
    # Decode the tokens to form the final SQL query.
    generated_sql = tokenizer.decode(output_ids[0], skip_special_tokens=True)
    return generated_sql

if __name__ == "__main__":
    # --- Sample Usage ---
    table_name = "bank_accounts"
    columns = [
        "account_id VARCHAR",
        "account_holder VARCHAR",
        "balance FLOAT",
        "account_type VARCHAR"
    ]
    prompt_text = "List the account holder names with balance greater than 1000"
    
    print("[DEBUG] Generating SQL query...")
    sql_query = generate_sql(table_name, columns, prompt_text)
    print("[DEBUG] Generated SQL Query:")
    print(sql_query)