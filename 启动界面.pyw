"""Open the existing UI launcher from the suite root."""
from pathlib import Path
import runpy
runpy.run_path(str(Path(__file__).resolve().parent/'WPS_Forms_UI/启动界面.pyw'),run_name='__main__')
