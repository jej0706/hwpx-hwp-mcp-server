"""FastMCP server wiring.

Creates the single ``FastMCP`` instance, registers tools from each category
module, and exposes lifespan hooks that tear down the Hancom COM session on
shutdown.
"""

from __future__ import annotations

from contextlib import asynccontextmanager
from typing import AsyncIterator

from mcp.server.fastmcp import FastMCP

from .backend.hancom_com import session
from .tools import batch, create, read, session as session_tools, template

SERVER_NAME = "hwpx-hwp-mcp"

INSTRUCTIONS = (
    "COM-backed editor for HWP and HWPX documents. Requires a local Hancom HWP "
    "installation; uses the real HWP engine for rendering fidelity. Prefer this "
    "server when faithful save/reload through 한/글 is important."
)


@asynccontextmanager
async def _lifespan(_mcp: FastMCP) -> AsyncIterator[None]:
    try:
        yield
    finally:
        await session.shutdown()


mcp = FastMCP(SERVER_NAME, instructions=INSTRUCTIONS, lifespan=_lifespan)

# Register tools grouped by category. Each module exposes ``register(mcp)``.
session_tools.register(mcp)
read.register(mcp)
template.register(mcp)
create.register(mcp)
batch.register(mcp)
