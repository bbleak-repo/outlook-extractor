import shutil
from pathlib import Path

def clean_build_artifacts():
    build_dir = Path(__file__).parent.parent / "dist"
    if build_dir.exists():
        shutil.rmtree(build_dir)
    build_dir = Path(__file__).parent.parent / "build" / "outlook_extractor"
    if build_dir.exists():
        shutil.rmtree(build_dir)

if __name__ == "__main__":
    clean_build_artifacts()
