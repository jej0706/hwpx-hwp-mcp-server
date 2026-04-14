"""Minimal pandas stub injected into ``sys.modules`` before pyhwpx import.

Why this exists
---------------

``pyhwpx`` imports ``pandas as pd`` at the top of ``core.py``. On 32-bit
Windows Python 3.10+ there are **no prebuilt pandas wheels** (pandas dropped
win32 wheels after 1.5.3/cp39), and building from source needs MSBuild +
Cython + Meson which most users do not have. Since this MCP server must run
under 32-bit Python to talk to the 32-bit Hancom HWP COM server, we cannot
install real pandas.

We don't actually *need* pandas for the tool surface we expose:

- ``put_field_text(dict, "")`` — hits the ``isinstance(field, dict)`` branch
  before any ``pd.*`` reference is evaluated. Safe under stub.
- ``find_replace_all``, ``insert_text``, ``create_table``, ``save_as``,
  ``open``, … — no pandas.
- The only pyhwpx methods that genuinely need pandas are
  ``table_from_data`` / ``table_to_df`` / ``table_to_df_q`` / ``table_to_csv``.
  We implement our table read/write paths without calling those.

So we register a stub that satisfies the import statement and the
``type(x) in [pd.Series]`` checks scattered through pyhwpx (those just
compare class identity — any sentinel class works).

Call :func:`install` **before** the first ``import pyhwpx`` anywhere in the
process.
"""

from __future__ import annotations

import sys
import types


def _make_stub() -> types.ModuleType:
    stub = types.ModuleType("pandas")
    stub.__version__ = "0.0.0-hwpx-hwp-mcp-stub"
    stub.__doc__ = (
        "Stub module supplied by hwpx_hwp_mcp so pyhwpx can import without "
        "real pandas. Any attempt to actually use a DataFrame/Series raises."
    )

    class _StubCallableClass:
        """Sentinel used for ``type(x) in [pd.DataFrame]`` comparisons.

        Instantiating raises — we never want real objects. But the class
        identity is stable, which is all the isinstance/type checks need.
        """

        def __init__(self, *args, **kwargs):  # noqa: D401 - sentinel
            raise RuntimeError(
                "pandas is stubbed in this 32-bit build of hwpx-hwp-mcp. "
                "Use the dedicated MCP tools instead of raw DataFrame/Series "
                "operations."
            )

        @classmethod
        def __class_getitem__(cls, _item):  # support pd.DataFrame[...] annotations
            return cls

    stub.DataFrame = _StubCallableClass  # type: ignore[attr-defined]
    stub.Series = _StubCallableClass  # type: ignore[attr-defined]
    stub.Index = _StubCallableClass  # type: ignore[attr-defined]

    def _unsupported(*_args, **_kwargs):
        raise RuntimeError(
            "pandas is stubbed in this 32-bit build of hwpx-hwp-mcp. "
            "The underlying pyhwpx method you called requires real pandas."
        )

    stub.concat = _unsupported  # type: ignore[attr-defined]
    stub.read_csv = _unsupported  # type: ignore[attr-defined]
    stub.read_excel = _unsupported  # type: ignore[attr-defined]
    stub.read_json = _unsupported  # type: ignore[attr-defined]
    stub.read_parquet = _unsupported  # type: ignore[attr-defined]
    stub.isna = lambda x: x is None  # type: ignore[attr-defined]
    stub.notna = lambda x: x is not None  # type: ignore[attr-defined]

    return stub


def install() -> None:
    """Inject the stub into ``sys.modules`` unless real pandas is present."""
    if "pandas" in sys.modules:
        return
    try:
        import pandas  # noqa: F401  # real pandas wins if the user has it
        return
    except ImportError:
        pass
    sys.modules["pandas"] = _make_stub()
