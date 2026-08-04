function createJellyAnimator(selector) {
  let frame = null;
  const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  function clear(root = document) {
    if (frame) {
      cancelAnimationFrame(frame);
      frame = null;
    }

    root.querySelectorAll(selector).forEach((element) => {
      element.style.opacity = "";
      element.style.transform = "";
      element.style.willChange = "";
    });
  }

  function animate(root = document) {
    if (reducedMotion) return;
    if (frame) cancelAnimationFrame(frame);

    const elements = Array.from(root.querySelectorAll(selector));
    const states = elements.map((element, order) => {
      const fromLeft = element.dataset.from === "left";
      const fromY = element.dataset.from === "top" ? -170 : 95;
      const fromX = fromLeft ? -150 : 0;

      element.style.opacity = "0";
      element.style.willChange = "transform, opacity";
      element.style.transform = `translate3d(${fromX}px, ${fromY}px, 0) scaleX(0.82) scaleY(1.18) rotate(-4deg)`;

      return {
        element,
        delay: Math.min(order * 75, 520),
        x: fromX,
        y: fromY,
        vx: fromLeft ? 24 : 0,
        vy: -8,
        rotation: fromLeft ? -4 : -2,
        vRotation: fromLeft ? 1.3 : 0.7,
        opacity: 0
      };
    });

    const stiffness = 0.105;
    const damping = 0.765;
    const rotationStiffness = 0.085;
    const rotationDamping = 0.72;
    let start = null;

    function tick(now) {
      if (start === null) start = now;
      let activeCount = 0;

      states.forEach((state) => {
        const elapsed = now - start - state.delay;
        if (elapsed < 0) {
          activeCount++;
          return;
        }

        state.vx += (0 - state.x) * stiffness;
        state.vy += (0 - state.y) * stiffness;
        state.vx *= damping;
        state.vy *= damping;
        state.x += state.vx;
        state.y += state.vy;

        state.vRotation += (0 - state.rotation) * rotationStiffness;
        state.vRotation *= rotationDamping;
        state.rotation += state.vRotation;

        const speed = Math.min(Math.hypot(state.vx, state.vy), 42);
        const squash = speed / 420;
        state.opacity = Math.min(1, state.opacity + 0.075);

        state.element.style.opacity = String(state.opacity);
        state.element.style.transform =
          `translate3d(${state.x.toFixed(2)}px, ${state.y.toFixed(2)}px, 0) ` +
          `rotate(${state.rotation.toFixed(2)}deg) ` +
          `scaleX(${(1 + squash).toFixed(3)}) scaleY(${(1 - squash * 0.75).toFixed(3)})`;

        const settled = Math.abs(state.x) < 0.12 &&
          Math.abs(state.y) < 0.12 &&
          Math.abs(state.vx) < 0.12 &&
          Math.abs(state.vy) < 0.12 &&
          Math.abs(state.rotation) < 0.08 &&
          Math.abs(state.vRotation) < 0.08 &&
          state.opacity >= 1;

        if (!settled) {
          activeCount++;
        } else {
          state.element.style.opacity = "1";
          state.element.style.transform = "translate3d(0, 0, 0) rotate(0) scaleX(1) scaleY(1)";
          state.element.style.willChange = "";
        }
      });

      frame = activeCount > 0 ? requestAnimationFrame(tick) : null;
    }

    frame = requestAnimationFrame(tick);
  }

  return { animate, clear };
}
