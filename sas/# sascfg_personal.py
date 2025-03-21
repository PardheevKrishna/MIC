SAS_config_names = ['winiom']

winiom = {
    # Full path to your Java executable (make sure Java is installed)
    'java': r'C:\Program Files\Java\jre1.8.0_251\bin\java.exe',
    # Hostname where your SAS Workspace Server is running (often 'localhost')
    'iomhost': 'localhost',
    # Port number for the SAS Integration Technologies server (default is usually 8591)
    'iomport': 8591,
    # Authentication key, if required by your setup (adjust as needed)
    'authkey': 'your_auth_key',  # Replace with your actual key or remove if not needed
    # Encoding setting; adjust according to your environment
    'encoding': 'windows-1252',
}