// Flutter 3.44 reports a benign
//   "LateInitializationError: Field '_handledContextLostEvent' has not been initialized."
// on every hot restart. It is an engine bug: CkSurface.onContextLost()
// (engine/canvaskit/surface.dart) reads that field first thing, but only
// triggerContextLoss() — an engine test hook — ever assigns it, so a REAL
// webglcontextlost event throws. A hot restart drops the old WebGL context,
// which is exactly when it fires. Nothing in the app can set the field, and
// the restart builds a fresh surface anyway, so there is no UI effect.
//
// It reaches the console by three different routes and each needs its own
// gag: the handler runs as a DOM listener (uncaught 'error' event), some
// paths run it from a .then() (unhandledrejection), and the framework may
// also relay it through console.error. Every filter matches only this one
// message — anything else passes through untouched.
//
// This lives in its own file rather than inline in index.html because the
// production build gets a Content-Security-Policy with `script-src 'self'`
// (see tool\inject_csp.ps1), which blocks inline scripts outright.
(function () {
  var MARKER = "_handledContextLostEvent";
  function isEngineContextLostNoise(value) {
    return !!value && ("" + value).indexOf(MARKER) !== -1;
  }
  var origError = console.error.bind(console);
  console.error = function () {
    for (var i = 0; i < arguments.length; i++) {
      if (isEngineContextLostNoise(arguments[i])) return;
    }
    return origError.apply(null, arguments);
  };
  // Capture phase so this runs before the browser's default reporting.
  window.addEventListener("error", function (event) {
    if (isEngineContextLostNoise(event.message) || isEngineContextLostNoise(event.error)) {
      event.preventDefault();
    }
  }, true);
  window.addEventListener("unhandledrejection", function (event) {
    if (isEngineContextLostNoise(event.reason)) event.preventDefault();
  });
})();
