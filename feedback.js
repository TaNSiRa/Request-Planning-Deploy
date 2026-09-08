// Sound and haptics for the app's small confirmations — the "ting" when a
// notification lands, the click under a button, the chord when something is
// approved.
//
// Everything here is SYNTHESISED, not loaded. Two reasons:
//   * the app has to work on machines with no internet and no shared drive, and
//     a folder of .mp3 files is one more thing that can fail to ship (the same
//     reason CanvasKit and the Thai font are bundled);
//   * a 40 ms decayed sine is a few hundred bytes of code and sounds cleaner
//     than a compressed sample of one, at any volume.
//
// Everything is also DELIBERATELY QUIET. These play in an office where several
// people sit within earshot of each other; the loudest cue here is about a
// third of a Windows notification. If it is ever the thing you notice, it is
// wrong.
(function () {
  var ctx = null;
  var master = null;

  // Browsers refuse to start an AudioContext until the user has interacted with
  // the page, and throw if you build one earlier. So it is built on the first
  // sound asked for (which is always after a click or a keypress) and resumed
  // every time after, because a background tab suspends it again.
  function audio() {
    if (ctx === null) {
      var Ctor = window.AudioContext || window.webkitAudioContext;
      if (!Ctor) return null;
      try {
        ctx = new Ctor();
      } catch (e) {
        ctx = false;
        return null;
      }
      master = ctx.createGain();
      // ── DIAL: how loud any of this is. ──
      // Started at 0.5 and was raised because "quiet" turned into "inaudible"
      // on an office PC at half volume — a cue nobody hears is the same as no
      // cue, and the point of these is to reach someone reading the middle of a
      // table. Still roughly a third of a Windows notification. Turn it DOWN
      // here, in one place, if a room ever complains.
      master.gain.value = 0.9;
      master.connect(ctx.destination);
    }
    if (ctx === false) return null;
    if (ctx.state === 'suspended') { try { ctx.resume(); } catch (e) {} }
    return ctx;
  }

  // One note. `at` is an offset in seconds from now, so a chord is written as
  // three calls with staggered offsets instead of a scheduler.
  //
  // The envelope matters more than the waveform: a note that starts instantly
  // clicks, and one that stops instantly clicks again. 8 ms in, exponential out
  // — that shape is what makes a bare sine read as a chime rather than a beep.
  function note(freq, at, dur, gain, type) {
    var c = audio();
    if (!c) return;
    var t0 = c.currentTime + at;
    var osc = c.createOscillator();
    var env = c.createGain();
    osc.type = type || 'sine';
    osc.frequency.setValueAtTime(freq, t0);
    env.gain.setValueAtTime(0.0001, t0);
    env.gain.exponentialRampToValueAtTime(gain, t0 + 0.008);
    env.gain.exponentialRampToValueAtTime(0.0001, t0 + dur);
    osc.connect(env);
    env.connect(master);
    osc.start(t0);
    osc.stop(t0 + dur + 0.02);
  }

  // A note that slides — used only for the error cue, where the fall is the
  // whole message.
  function slide(from, to, at, dur, gain) {
    var c = audio();
    if (!c) return;
    var t0 = c.currentTime + at;
    var osc = c.createOscillator();
    var env = c.createGain();
    osc.type = 'triangle';
    osc.frequency.setValueAtTime(from, t0);
    osc.frequency.exponentialRampToValueAtTime(to, t0 + dur);
    env.gain.setValueAtTime(0.0001, t0);
    env.gain.exponentialRampToValueAtTime(gain, t0 + 0.01);
    env.gain.exponentialRampToValueAtTime(0.0001, t0 + dur);
    osc.connect(env);
    env.connect(master);
    osc.start(t0);
    osc.stop(t0 + dur + 0.02);
  }

  // The vocabulary. Pitches are from a C-major pentatonic so that two cues
  // landing on top of each other — a save finishing as a notification arrives —
  // still agree with each other instead of beating.
  var cues = {
    // Under a click. Barely there on purpose: it is felt more than heard.
    tick: function () { note(1174.7, 0, 0.035, 0.030, 'sine'); },
    // Something saved. One note, up.
    ok: function () { note(987.8, 0, 0.09, 0.055, 'sine'); },
    // Something completed — a rising third.
    success: function () {
      note(783.99, 0, 0.10, 0.055);
      note(1046.5, 0.075, 0.16, 0.050);
    },
    // A notification arrived. The classic two-note "ting", high and short so it
    // carries over a room without being loud.
    notify: function () {
      note(1318.5, 0, 0.09, 0.050);
      note(1975.5, 0.085, 0.22, 0.038);
    },
    // Something went wrong. Falls, and is the only cue with any edge to it.
    error: function () { slide(415.3, 233.1, 0, 0.26, 0.055); },
    // The one cue allowed to be a flourish: an approval going through, a board
    // being cleared. Four notes up, then the octave ringing under them.
    celebrate: function () {
      note(523.25, 0.00, 0.12, 0.048);
      note(659.25, 0.075, 0.12, 0.048);
      note(783.99, 0.150, 0.14, 0.048);
      note(1046.5, 0.225, 0.45, 0.055);
      note(1567.98, 0.240, 0.40, 0.022);
    }
  };

  // Called from Dart (AppFeedback). Unknown names are ignored rather than
  // thrown — a cue added on the Dart side before this file catches up should
  // fall silent, not take the button down with it.
  window.rapSound = function (name) {
    try {
      var cue = cues[name];
      if (cue) cue();
    } catch (e) { /* audio is a nicety; never let it break an action */ }
  };

  // Haptics. Real on Android; Safari and desktop Chrome have no vibrator and
  // return false, which is fine — this is additive everywhere.
  window.rapBuzz = function (ms) {
    try {
      if (navigator.vibrate) navigator.vibrate(ms);
    } catch (e) {}
  };
})();
