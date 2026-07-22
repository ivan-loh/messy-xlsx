from __future__ import annotations

import warnings


class LegacyAPIWarning(DeprecationWarning):
    """A materialized compatibility API has a preferred bounded alternative."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)


def warn_legacy(api_name: str) -> None:
    warnings.warn(
        f"{api_name} is a legacy materialized API retained through messy-xlsx v1.x",
        LegacyAPIWarning,
        stacklevel=3,
    )
