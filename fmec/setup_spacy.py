"""
Setup script to download spaCy model from GitHub releases.
Use this if 'python -m spacy download' is blocked.
"""

import subprocess
import sys

def install_spacy_model():
    """Download and install spaCy model from GitHub releases."""
    print("📦 Downloading spaCy model from GitHub releases...")
    
    model_url = "https://github.com/explosion/spacy-models/releases/download/en_core_web_sm-3.7.0/en_core_web_sm-3.7.0-py3-none-any.whl"
    
    try:
        # Install directly from GitHub release
        subprocess.check_call([
            sys.executable, 
            "-m", 
            "pip", 
            "install", 
            model_url,
            "--quiet"
        ])
        print("✅ spaCy model installed successfully!")
        
        # Verify installation
        import spacy
        nlp = spacy.load("en_core_web_sm")
        print(f"✅ Model loaded successfully: {nlp.meta['name']} v{nlp.meta['version']}")
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Error installing spaCy model: {e}")
        print("\n💡 Alternative: Try downloading manually from:")
        print(f"   {model_url}")
        print("   Then run: pip install en_core_web_sm-3.7.0-py3-none-any.whl")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Error verifying model: {e}")
        sys.exit(1)

if __name__ == "__main__":
    install_spacy_model()
