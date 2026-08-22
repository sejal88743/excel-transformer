import { StrictMode } from "react";
import ReactDOM from "react-dom/client";
import { RouterProvider } from "@tanstack/react-router";
import { getRouter } from "./router";

// Suppress benign SheetJS internal decompression warnings from polluting error loggers
if (typeof window !== "undefined") {
  const originalError = console.error;
  console.error = function (...args: unknown[]) {
    const firstArg = typeof args[0] === "string" ? args[0] : "";
    if (
      firstArg.startsWith("Bad uncompressed size:") ||
      firstArg.startsWith("Bad compressed size:") ||
      firstArg.startsWith("Bad CRC32 checksum:")
    ) {
      console.debug(...args);
      return;
    }
    originalError.apply(console, args);
  };
}

const router = getRouter();

const rootElement = document.getElementById("root")!;
ReactDOM.createRoot(rootElement).render(
  <StrictMode>
    <RouterProvider router={router} />
  </StrictMode>
);

