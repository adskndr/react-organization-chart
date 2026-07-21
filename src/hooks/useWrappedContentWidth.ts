import * as React from "react";

/**
 * Measures the actual rendered width of a flex-wrap container's content —
 * i.e. the width of its widest wrapped row — and returns it so it can be
 * applied as an explicit width on the container itself.
 *
 * Why this is needed: a flex (or inline-flex) container with an automatic
 * width either sizes to fit all its children on one line, or — once that
 * doesn't fit the available space — falls back to the full available
 * width. There's no CSS-only way to make it shrink to "the width of the
 * widest row it actually wrapped into", which is exactly what makes a grey
 * box around 1-2 cards (single row, no wrapping) hug them tightly, while
 * the same box around 3+ cards (wraps to multiple rows) ends up far wider
 * than the cards it contains.
 *
 * This measures the real rendered position of each direct child and
 * derives the box's own needed width from that, re-measuring whenever the
 * content or the available space changes.
 */
export const useWrappedContentWidth = <T extends HTMLElement>(
  deps: React.DependencyList
): [React.RefObject<T>, number | undefined] => {
  const ref = React.useRef<T>(null);
  const [width, setWidth] = React.useState<number>();

  React.useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;

    const measure = (): void => {
      const children = Array.from(el.children) as HTMLElement[];
      // TEMPORARY diagnostic — remove once we've confirmed what's happening.
      // eslint-disable-next-line no-console
      console.log("[useWrappedContentWidth] measuring", {
        childCount: children.length,
        elRect: el.getBoundingClientRect(),
      });
      if (children.length === 0) {
        setWidth(undefined);
        return;
      }
      const rects = children.map((child) => child.getBoundingClientRect());
      const minLeft = Math.min(...rects.map((r) => r.left));
      const maxRight = Math.max(...rects.map((r) => r.right));
      // eslint-disable-next-line no-console
      console.log("[useWrappedContentWidth] result", {
        minLeft,
        maxRight,
        width: maxRight - minLeft,
      });
      setWidth(maxRight - minLeft);
    };

    // Measure once we have real layout, then keep re-measuring if the
    // available width changes (e.g. the web part is resized) — a size
    // change can shift how many cards fit per row, which changes the
    // widest-row width even though nothing else about the content changed.
    measure();
    const resizeObserver = new ResizeObserver(measure);
    resizeObserver.observe(el);

    return () => {
      resizeObserver.disconnect();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);

  return [ref, width];
};