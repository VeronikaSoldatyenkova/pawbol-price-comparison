"""Production ASGI entrypoint."""

from powerbi_purchase_fields import apply

apply()

from runtime_main import app  # noqa: E402,F401
from performance_optimization import install  # noqa: E402

install(app)


def _route(path, method):
    return next(
        (
            route
            for route in app.router.routes
            if getattr(route, "path", None) == path
            and method in (getattr(route, "methods", set()) or set())
        ),
        None,
    )


# Keep the existing templates/URLs while replacing long-running handlers with
# the optimized implementations installed above.
_fast_compare_route = _route("/compare-fast/{workspace_id}", "POST")
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

_lazy_download_route = _route("/download-full/{workspace_id}", "GET")
if _lazy_download_route is not None:
    app.router.routes[:] = [
        route
        for route in app.router.routes
        if not (
            getattr(route, "path", None) == "/download/{workspace_id}"
            and "GET" in (getattr(route, "methods", set()) or set())
        )
    ]
    app.add_api_route(
        "/download/{workspace_id}",
        _lazy_download_route.endpoint,
        methods=["GET"],
        include_in_schema=False,
    )
