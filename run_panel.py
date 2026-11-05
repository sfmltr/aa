"""Double-click friendly launcher (rename to run_panel.pyw on Windows to hide the console)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from ncr_panel.main import main  # noqa: E402

sys.exit(main())
