import * as React from "react";

/**
 * Measures the actual width of the widest rendered row
 * inside a flex-wrap container.
 *
 * Example:
 *
 *   Card   Card   Card
 *   Card
 *
 * The returned width is the width of:
 *
 *   Card   Card   Card
 *
 * and NOT the complete container width.
 *
 * This allows the surrounding box to shrink-wrap around
 * the widest actually rendered row.
 */
export const useWrappedContentWidth = <T extends HTMLElement>(
  deps: React.DependencyList
): [React.RefObject<T>, number | undefined] => {
  const ref = React.useRef<T>(null);
  const [width, setWidth] = React.useState<number>();

  React.useLayoutEffect(() => {
    const el = ref.current;

    if (!el) {
      return;
    }

    const measure = (): void => {
      const children = Array.from(el.children) as HTMLElement[];

      if (children.length === 0) {
        setWidth(undefined);
        return;
      }

      const rects = children
        .map((child) => child.getBoundingClientRect())
        .filter((rect) => rect.width > 0 && rect.height > 0);

      if (rects.length === 0) {
        setWidth(undefined);
        return;
      }

      /*
       * Group cards by their vertical position.
       *
       * Cards in the same flex row have approximately
       * the same top position.
       */
      const rows: DOMRect[][] = [];

      const ROW_TOLERANCE = 2;

      rects
        .sort((a, b) => a.top - b.top || a.left - b.left)
        .forEach((rect) => {
          const existingRow = rows.find(
            (row) => Math.abs(row[0].top - rect.top) <= ROW_TOLERANCE
          );

          if (existingRow) {
            existingRow.push(rect);
          } else {
            rows.push([rect]);
          }
        });

      /*
       * Find the widest actual row.
       */
      let widestRowWidth = 0;

      rows.forEach((row) => {
        const minLeft = Math.min(...row.map((rect) => rect.left));
        const maxRight = Math.max(...row.map((rect) => rect.right));

        const rowWidth = maxRight - minLeft;

        if (rowWidth > widestRowWidth) {
          widestRowWidth = rowWidth;
        }
      });

      if (widestRowWidth > 0) {
        /*
         * Add the container's horizontal padding/border.
         *
         * offsetWidth includes padding + border.
         * clientWidth includes padding but not the border.
         *
         * The difference gives us the border width.
         */
        const computedStyle = window.getComputedStyle(el);

        const paddingLeft = parseFloat(computedStyle.paddingLeft) || 0;
        const paddingRight = parseFloat(computedStyle.paddingRight) || 0;

        const borderLeft =
          parseFloat(computedStyle.borderLeftWidth) || 0;

        const borderRight =
          parseFloat(computedStyle.borderRightWidth) || 0;

        const finalWidth =
          widestRowWidth +
          paddingLeft +
          paddingRight +
          borderLeft +
          borderRight;

        setWidth((previousWidth) => {
          /*
           * Avoid unnecessary renders when the measured width
           * has not actually changed.
           */
          if (
            previousWidth !== undefined &&
            Math.abs(previousWidth - finalWidth) < 1
          ) {
            return previousWidth;
          }

          return finalWidth;
        });
      }
    };

    /*
     * Measure after the browser has completed the current layout.
     */
    const measureAsync = (): void => {
      window.requestAnimationFrame(measure);
    };

    measureAsync();

    /*
     * Re-measure when the available size changes.
     */
    const resizeObserver = new ResizeObserver(() => {
      measureAsync();
    });

    resizeObserver.observe(el);

    /*
     * Also observe the parent because the parent can change width
     * without the element itself changing its intrinsic dimensions.
     */
    if (el.parentElement) {
      resizeObserver.observe(el.parentElement);
    }

    return () => {
      resizeObserver.disconnect();
    };

    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);

  return [ref, width];
};