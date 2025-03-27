import os

# from sympy import false
# from decouple import config
from unipath import Path

# Build paths inside the project like this: os.path.join(BASE_DIR, ...)
# BASE_DIR = Path(__file__).parent
CORE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ROOT_DIR = os.path.join(CORE_DIR, 'ScriptPY')
print("CORE_DIR-Config.py: ",CORE_DIR)
print("ROOT_DIR-Config.py: ",ROOT_DIR)
