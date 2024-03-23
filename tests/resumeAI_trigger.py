import sys
from pathlib import Path

# Add the src directory to sys.path
sys.path.append(str(Path(__file__).resolve().parent.parent / "src"))

import lambda_function as pdf

local_event = {
    "body": {
        "filename": "staging/mem_sb_clskhxlvc004r0ss07v1kd35k_1711227652153_resume.pdf",
    }
}

pdf.lambda_handler(local_event, None)
