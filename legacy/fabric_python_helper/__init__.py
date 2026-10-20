"""
fabric_python_helper has been renamed archwaytrust_fabric.

This package only forwards to archwaytrust_fabric, so notebooks that still import fabric_python_helper keep
working. Every module is the same object under both names: fabric_python_helper.auth is archwaytrust_fabric.auth.
"""
import sys
import warnings

import archwaytrust_fabric

__version__ = archwaytrust_fabric.__version__
__all__ = archwaytrust_fabric.__all__
globals().update({name: getattr(archwaytrust_fabric, name) for name in __all__})

warnings.warn(
    "fabric_python_helper has been renamed archwaytrust_fabric. Install archwaytrust-fabric and change "
    "'fabric_python_helper' to 'archwaytrust_fabric' in your imports.",
    FutureWarning,
    stacklevel=2,
)

# Alias every submodule (auth, graph_api, sharepoint and its private modules), so
# 'import fabric_python_helper.auth' finds the module that's already loaded instead of loading a second copy
for _name, _module in list(sys.modules.items()):
    if _name.startswith("archwaytrust_fabric."):
        sys.modules[__name__ + _name[len("archwaytrust_fabric"):]] = _module
        if "." not in _name[len("archwaytrust_fabric."):]:
            globals()[_name.rsplit(".", 1)[1]] = _module
