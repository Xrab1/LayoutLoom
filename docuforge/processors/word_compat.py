"""Public entry point for the independent region-level Word rebuild.

The PDF-to-Word converter remains untouched.  This task consumes its
high-fidelity positioned DOCX and creates a separate region-copyable Word
through :mod:`docuforge.processors.word_region`.
"""

from __future__ import annotations

from pathlib import Path


def optimize_word_full_compatibility(
    input_docx: str | Path,
    output_path: str | Path,
    *,
    verification_engine: str = "auto",
    timeout: float = 300,
    overwrite: bool = False,
) -> list[Path]:
    """Create a region-level editable desktop/mobile-compatible Word copy."""

    from docuforge.processors.word_region import build_region_compatible_word

    return build_region_compatible_word(
        input_docx,
        output_path,
        verification_engine=verification_engine,
        timeout=timeout,
        overwrite=overwrite,
    )


__all__ = ["optimize_word_full_compatibility"]
