# Compatibility entrypoint. Production is served by Uvicorn from runtime_main:app.
from runtime_main import app  # noqa: F401
