import sys
import os

# Add the parent directory to Python path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Now import and run the UI
from outlook_extractor.ui.main_window import main

if __name__ == "__main__":
    main()
