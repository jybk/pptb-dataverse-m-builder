import react from "@vitejs/plugin-react";
import type { Plugin } from "vite";
import { defineConfig } from "vite";

/**
 * Makes the built HTML loadable by Power Platform ToolBox:
 * - drops type="module" / crossorigin (the bundle is IIFE, loaded from a custom protocol)
 * - moves script tags to the end of <body> so the DOM exists when the IIFE runs
 */
function fixHtmlForPPTB(): Plugin {
    return {
        name: "fix-html-for-pptb",
        enforce: "post",
        transformIndexHtml(html) {
            html = html.replace(/\s*type="module"/g, "");
            html = html.replace(/\s*crossorigin/g, "");
            html = html.replace(/\s+>/g, ">");
            const scripts: string[] = [];
            html = html.replace(/(<script[^>]*src="[^"]*"[^>]*><\/script>)/g, (match) => {
                scripts.push(match);
                return "";
            });
            if (scripts.length > 0) {
                html = html.replace("</body>", "\n  " + scripts.join("\n  ") + "\n</body>");
            }
            return html;
        },
    };
}

export default defineConfig((configEnv) => ({
    plugins: [react(), fixHtmlForPPTB()],
    base: "./",
    build: {
        outDir: "dist",
        assetsDir: "assets",
        sourcemap: configEnv.mode === "development",
        chunkSizeWarningLimit: 2000,
        rollupOptions: {
            output: {
                format: "iife",
                inlineDynamicImports: true,
                manualChunks: undefined,
            },
        },
    },
    test: {
        environment: "jsdom",
        include: ["tests/**/*.test.ts"],
    },
}));
