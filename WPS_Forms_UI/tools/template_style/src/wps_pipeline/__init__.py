"""Local document organization runtime."""
__version__ = "0.5.7"


def resource_root():
    """The same config/template files in source and the standalone template tool."""
    from pathlib import Path
    import sys
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS) / 'wps_tool'
    return Path(__file__).resolve().parents[2]
