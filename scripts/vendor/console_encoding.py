"""Standalone UTF-8 console setup for the vendored PPT Master core.

The upstream module also performs repository-specific integrity and transcript
work. This distribution keeps only the stream-encoding behavior; the complete
MIT license and upstream attribution are shipped beside the vendored core.
"""

from __future__ import annotations

import io
import sys
from typing import TextIO


def _reconfigure_stream(stream: TextIO) -> TextIO:
    try:
        stream.reconfigure(encoding="utf-8", errors="replace")
        return stream
    except AttributeError:
        buffer = getattr(stream, "buffer", None)
        return stream if buffer is None else io.TextIOWrapper(
            buffer, encoding="utf-8", errors="replace"
        )
    except (OSError, ValueError):
        return stream


def configure_utf8_stdio() -> None:
    sys.stdout = _reconfigure_stream(sys.stdout)
    sys.stderr = _reconfigure_stream(sys.stderr)
