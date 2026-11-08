import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import App from "./App";
import "./index.css";

const rootElement = document.getElementById("root");
if (rootElement && !rootElement.hasAttribute("data-reactroot-initialized")) {
    rootElement.setAttribute("data-reactroot-initialized", "true");
    createRoot(rootElement).render(
        <StrictMode>
            <App />
        </StrictMode>,
    );
}
