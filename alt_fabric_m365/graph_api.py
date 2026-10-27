"""
Deprecated: moved to alt_fabric_m365.exchange.

Kept so code written for 0.6.1 and earlier, and the fabric_python_helper alias, keep working. Using a name
from here shows a FutureWarning. It will be removed in a future release.
"""
import warnings

from . import exchange

_MOVED = ("Emails", "MessageNotFoundError")


def __getattr__(name: str) -> object:
    if name in _MOVED:
        warnings.warn(
            "alt_fabric_m365.graph_api has moved to alt_fabric_m365.exchange. "
            "Change 'graph_api' to 'exchange' in your imports.",
            FutureWarning,
            stacklevel=2,
        )
        return getattr(exchange, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
