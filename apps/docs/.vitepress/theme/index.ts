import type { Theme } from "vitepress";
import DefaultTheme from "vitepress/theme";
import type { Component } from "vue";
import "./styles/index.css";

// Auto-register every SFC in components/, so adding a demo/card component does
// not require touching this file.
const components = import.meta.glob<{ default: Component }>(
  "./components/*.vue",
  { eager: true },
);

interface ViewTransitionLike {
  ready: Promise<void>;
}

/**
 * Circular clip reveal for the light/dark toggle (View Transitions API).
 *
 * The default theme flips `.dark` on <html> in its own click handler. To run
 * that flip inside a transition we intercept the appearance-button click in
 * the capture phase (before VitePress' listener on the button runs), then
 * replay the same click from within document.startViewTransition(). When the
 * API is unavailable or the user prefers reduced motion, the interception
 * stands down and VitePress toggles plainly.
 */
function setupAppearanceTransition(): void {
  const doc = document as Document & {
    startViewTransition?: (callback: () => void) => ViewTransitionLike;
  };
  if (typeof doc.startViewTransition !== "function") return;

  let replaying = false;
  document.addEventListener(
    "click",
    (event) => {
      if (replaying) return;
      const target = event.target as HTMLElement | null;
      const button = target?.closest?.("button");
      // Only the appearance toggle in the top bar; doc content is untouched.
      if (!button || !button.closest(".VPNavBarAppearance")) return;

      const reduced =
        typeof window.matchMedia === "function" &&
        window.matchMedia("(prefers-reduced-motion: reduce)").matches;
      if (reduced) return;

      event.preventDefault();
      event.stopImmediatePropagation();

      // Expand from the button center to the farthest screen corner.
      const rect = button.getBoundingClientRect();
      const x = rect.left + rect.width / 2;
      const y = rect.top + rect.height / 2;
      const endRadius = Math.hypot(
        Math.max(x, window.innerWidth - x),
        Math.max(y, window.innerHeight - y),
      );

      const replayToggle = (): void => {
        replaying = true;
        try {
          button.click();
        } finally {
          replaying = false;
        }
      };

      const transition = doc.startViewTransition(replayToggle);
      void transition.ready.then(() => {
        document.documentElement.animate(
          {
            clipPath: [
              `circle(0px at ${x}px ${y}px)`,
              `circle(${endRadius}px at ${x}px ${y}px)`,
            ],
          },
          {
            duration: 450,
            easing: "ease-in-out",
            pseudoElement: "::view-transition-new(root)",
          },
        );
      });
    },
    true, // capture: win the race against VitePress' own listener
  );
}

export default {
  extends: DefaultTheme,
  enhanceApp({ app }) {
    if (typeof window !== "undefined") setupAppearanceTransition();
    for (const [filePath, mod] of Object.entries(components)) {
      const name = filePath
        .split("/")
        .pop()
        ?.replace(/\.vue$/, "");
      if (name) app.component(name, mod.default);
    }
  },
} satisfies Theme;
