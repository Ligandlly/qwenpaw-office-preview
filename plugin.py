# -*- coding: utf-8 -*-
"""QwenPaw Office Document Preview plugin — backend entry point."""

from __future__ import annotations

import logging
import os
import sys
from pathlib import Path

_plugin_dir = Path(__file__).resolve().parent
_plugin_dir_str = str(_plugin_dir)
if _plugin_dir_str not in sys.path:
    sys.path.insert(0, _plugin_dir_str)

# QwenPaw execs this file as a plain module.  On plugin reload the previous
# ``converter`` package can still be cached in ``sys.modules``; evict only the
# modules that belong to this plugin directory so the new source is used.
for _module_name in list(sys.modules):
    if _module_name != "converter" and not _module_name.startswith("converter."):
        continue
    _module = sys.modules.get(_module_name)
    _module_file = getattr(_module, "__file__", None) or ""
    if _module_file and os.path.realpath(_module_file).startswith(
        os.path.realpath(_plugin_dir_str) + os.sep,
    ):
        del sys.modules[_module_name]

from qwenpaw.plugins.api import PluginApi  # noqa: E402

from converter.api import router as office_router  # noqa: E402

logger = logging.getLogger("qwenpaw.office_preview")


class OfficePreviewPlugin:
    """Serve DOCX/PPTX/XLSX (and legacy Office) previews to the Console."""

    def register(self, api: PluginApi) -> None:
        """Expose the converter endpoints under ``/api/office-preview``."""

        api.register_http_router(
            office_router,
            prefix="/office-preview",
            tags=["office-preview"],
        )
        logger.info(
            "Office preview plugin registered: GET /api/office-preview/health",
        )


plugin = OfficePreviewPlugin()
