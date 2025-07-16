import sys
from pathlib import Path

# Add the current directory to Python path
sys.path.insert(0, str(Path(__file__).parent))

# Now import and run the UI
from outlook_extractor.ui.main_window import main

if __name__ == "__main__":
    main()
