import { installPreviewInterceptor } from "./intercept";
import { localizedLabel, t } from "./i18n";
import { OfficePreviewPage } from "./OfficePreviewPage";
import { installStyles } from "./styles";

const PLUGIN_ID = "qwenpaw-office-preview";

installStyles();
installPreviewInterceptor();

try {
  window.QwenPaw.registerRoutes?.(PLUGIN_ID, [
    {
      path: `/plugin/${PLUGIN_ID}`,
      component: OfficePreviewPage,
      label: localizedLabel("Office 预览", "Office Preview"),
      title: localizedLabel("Office 文档预览", "Office Document Preview"),
      icon: "📄",
      priority: 46,
    },
  ]);
} catch (error) {
  console.error(`[${PLUGIN_ID}] failed to register route`, error);
}

// Expose a tiny manual hook for other plugins/console scripts and for
// debugging: QwenPawOfficePreview.open({ path, root }).
try {
  (window as unknown as Record<string, unknown>).QwenPawOfficePreview = {
    open: (options: { path?: string; artifact?: string; name?: string; root?: string }) => {
      const target = {
        source: options.artifact ? "attachment" : "workspace",
        path: options.path || "",
        artifactUrl: options.artifact,
        root: options.root || "project",
      };
      window.dispatchEvent(
        new CustomEvent("qwenpaw:open-file-preview", {
          detail: { target },
        }),
      );
    },
    title: t("routeLabel"),
  };
} catch {
  /* ignore */
}
