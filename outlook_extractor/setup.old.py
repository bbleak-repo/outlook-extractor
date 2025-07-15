# From the project root directory
cat > setup.py << 'EOL'
from setuptools import setup, find_packages

setup(
    name="outlook_extractor",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        'PySimpleGUI',
        'pywin32; sys_platform == "win32"',
    ],
    entry_points={
        'console_scripts': [
            'outlook-extractor=outlook_extractor.ui.main_window:main',
        ],
    },
)
EOL
