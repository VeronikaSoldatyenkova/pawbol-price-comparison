# Compatibility entrypoint. Production is served by Uvicorn from main:app.
from main import app  # noqa: F401
