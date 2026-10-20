"""
fabric_python_helper has been renamed alt_fabric_m365.

This package only forwards to alt_fabric_m365, so notebooks that still import fabric_python_helper keep
working. Every module is the same object under both names: fabric_python_helper.auth is alt_fabric_m365.auth.
"""
import sys
import warnings

import alt_fabric_m365

__version__ = alt_fabric_m365.__version__
__all__ = alt_fabric_m365.__all__
globals().update({name: getattr(alt_fabric_m365, name) for name in __all__})

warnings.warn(
    "fabric_python_helper has been renamed alt_fabric_m365. Install alt-fabric-m365 and change "
    "'fabric_python_helper' to 'alt_fabric_m365' in your imports.",
    FutureWarning,
    stacklevel=2,
)

# Alias every submodule (auth, graph_api, sharepoint and its private modules), so
# 'import fabric_python_helper.auth' finds the module that's already loaded instead of loading a second copy
for _name, _module in list(sys.modules.items()):
    if _name.startswith("alt_fabric_m365."):
        sys.modules[__name__ + _name[len("alt_fabric_m365"):]] = _module
        if "." not in _name[len("alt_fabric_m365."):]:
            globals()[_name.rsplit(".", 1)[1]] = _module
