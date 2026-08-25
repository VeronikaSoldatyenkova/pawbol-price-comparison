"""Production ASGI entrypoint."""

from powerbi_purchase_fields import apply

apply()

from runtime_main import app  # noqa: E402,F401
from performance_optimization import install  # noqa: E402

install(app)

# Keep the existing configure.html form action (/compare/{workspace_id}) while
# replacing its old long-running request handler with the optimized background
# comparison endpoint. This avoids proxy timeouts / Bad Gateway responses.
_fast_compare_route = next(
    (
        route
        for route in app.router.routes
        if getattr(route, "path", None) == "/compare-fast/{workspace_id}"
        and "POST" in (getattr(route, "methods", set()) or set())
    ),
    None,
)
if _fast_compare_route is not None:
    app.router.routes[:] = [
        route
        for route in app.router.routes
        if not (
            getattr(route, "path", None) == "/compare/{workspace_id}"
            and "POST" in (getattr(route, "methods", set()) or set())
        )
    ]
    app.add_api_route(
        "/compare/{workspace_id}",
        _fast_compare_route.endpoint,
        methods=["POST"],
        include_in_schema=False,
    )
