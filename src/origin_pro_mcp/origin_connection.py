_origin = None
_MAINWND_SHOW = 1

def get_origin():
    global _origin
    if _origin is None:
        try:
            import win32com.client
            import pywintypes
        except ModuleNotFoundError as exc:
            msg = (
                "Origin Pro COM automation requires Windows Python with pywin32. "
                "Run this MCP server on Windows, not WSL/Linux."
            )
            raise RuntimeError(msg) from exc
        _origin = win32com.client.Dispatch("Origin.ApplicationSI")
        try:
            _origin.Visible = _MAINWND_SHOW
        except (AttributeError, pywintypes.com_error):
            pass
    return _origin

def execute_labtalk(script: str) -> bool:
    o = get_origin()
    return o.Execute(script)

def get_lt_var(name: str) -> float:
    return get_origin().LTVar(name)

def get_lt_str(name: str) -> str:
    return get_origin().LTStr(name)
