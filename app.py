# Compatibility entrypoint. Production is served by Uvicorn from production_app:app.
from production_app import app  # noqa: F401
