import sys
from pathlib import Path

# Add the src directory to sys.path
sys.path.append(str(Path(__file__).resolve().parent.parent / 'src'))

import lambda_function as pdf

basic_event = {
    "key1": "value1",
    "key2": "value2",
    "key3": "value3"
}

pdf.lambda_handler(basic_event, None)
