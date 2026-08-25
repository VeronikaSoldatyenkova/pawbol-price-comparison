"""Production ASGI entrypoint."""

from powerbi_purchase_fields import apply

apply()

from runtime_main import app  # noqa: E402,F401
