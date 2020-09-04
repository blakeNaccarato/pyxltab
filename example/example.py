"""
Example script.
"""

import pyxlrand

IN_DIR = r"example"

file_structure = pyxlrand.read(in_dir=IN_DIR)
pyxlrand.write(file_structure)
