"use strict";

/**
 * Built-in pictograms: the "visual built in" for slides that need a picture
 * but must never invent one (megaprompt section 18 / 29).
 *
 * A pictogram is a simple, flat glyph (Phosphor Icons, MIT licence) rendered
 * as a white glyph on a coloured circle or tile, or as a flat glyph in a theme
 * colour. It is deliberately not an illustration: it names a thing (frog,
 * cloud, book, happy face) so a Foundation student can read the slide, and it
 * gives a vocabulary card, a science stage or a wellbeing feeling a picture
 * without anyone hunting for a photo.
 *
 * Rendering is SYNCHRONOUS (resvg-js) so every builder stays synchronous and
 * a build script can never forget an `await` and ship a blank space.
 *
 * Unknown names emit a `WARN [pictogram]` line, which fails the build gate on
 * purpose: a missing picture is invisible in the rendered file, so the build
 * has to be the thing that shouts.
 */

const React = require("react");
const ReactDOMServer = require("react-dom/server");
const Pi = require("react-icons/pi");
const { DEFAULT_SIZES, byBand } = require("./gradeBand");

let Resvg = null;
function getResvg() {
  if (Resvg) return Resvg;
  // Lazy: keep theme load fast for scripts that never draw a pictogram.
  ({ Resvg } = require("@resvg/resvg-js"));
  return Resvg;
}

/**
 * Friendly name -> Phosphor icon component name. Names are the words a
 * teacher would type. Add here, then rebuild the visual catalogue.
 */
const PICTOGRAMS = {
  // Animals and living things
  butterfly: "PiButterflyFill", bird: "PiBirdFill", cat: "PiCatFill", dog: "PiDogFill",
  fish: "PiFishFill", horse: "PiHorseFill", cow: "PiCowFill", rabbit: "PiRabbitFill",
  bug: "PiBugFill", beetle: "PiBugBeetleFill", paw: "PiPawPrintFill", feather: "PiFeatherFill",
  shrimp: "PiShrimpFill", egg: "PiEggFill",
  // Plants and nature
  tree: "PiTreeFill", pine: "PiTreeEvergreenFill", leaf: "PiLeafFill", flower: "PiFlowerFill",
  lotus: "PiFlowerLotusFill", plant: "PiPlantFill", cactus: "PiCactusFill", mountains: "PiMountainsFill",
  island: "PiIslandFill", park: "PiParkFill",
  // Weather, sky, water
  sun: "PiSunFill", moon: "PiMoonFill", stars: "PiMoonStarsFill", star: "PiStarFill",
  cloud: "PiCloudFill", rain: "PiCloudRainFill", snow: "PiCloudSnowFill", storm: "PiCloudLightningFill",
  partlyCloudy: "PiCloudSunFill", fog: "PiCloudFogFill", drop: "PiDropFill", snowflake: "PiSnowflakeFill",
  wind: "PiWindFill", lightning: "PiLightningFill", rainbow: "PiRainbowFill", fire: "PiFireFill",
  waves: "PiWavesFill", umbrella: "PiUmbrellaFill", thermometer: "PiThermometerFill",
  hot: "PiThermometerHotFill", cold: "PiThermometerColdFill",
  // Earth and space
  planet: "PiPlanetFill", rocket: "PiRocketLaunchFill", globe: "PiGlobeHemisphereWestFill",
  earth: "PiGlobeFill", compass: "PiCompassFill", map: "PiMapTrifoldFill", pin: "PiMapPinFill",
  // Science and tools
  flask: "PiFlaskFill", testTube: "PiTestTubeFill", atom: "PiAtomFill", magnet: "PiMagnetFill",
  lightbulb: "PiLightbulbFill", gear: "PiGearFill", binoculars: "PiBinocularsFill", dna: "PiDnaFill",
  virus: "PiVirusFill", magnifier: "PiMagnifyingGlassFill", battery: "PiBatteryFullFill", plug: "PiPlugFill",
  recycle: "PiRecycleFill", solar: "PiSolarPanelFill", windmill: "PiWindmillFill", factory: "PiFactoryFill",
  hammer: "PiHammerFill", wrench: "PiWrenchFill", shovel: "PiShovelFill", key: "PiKeyFill",
  lock: "PiLockFill", shield: "PiShieldFill", anchor: "PiAnchorFill", target: "PiTargetFill",
  flashlight: "PiFlashlightFill", scales: "PiScalesFill",
  // Body and health
  brain: "PiBrainFill", heart: "PiHeartFill", heartbeat: "PiHeartbeatFill", tooth: "PiToothFill",
  bone: "PiBoneFill", eye: "PiEyeFill", ear: "PiEarFill", hand: "PiHandFill", wave: "PiHandWavingFill",
  clap: "PiHandsClappingFill", handHeart: "PiHandHeartFill", footprints: "PiFootprintsFill",
  firstAid: "PiFirstAidFill", bed: "PiBedFill", bath: "PiBathtubFill",
  // People
  person: "PiPersonFill", run: "PiPersonSimpleRunFill", walk: "PiPersonSimpleWalkFill",
  people: "PiUsersFill", group: "PiUsersThreeFill", student: "PiStudentFill",
  teacher: "PiChalkboardTeacherFill", family: "PiHouseLineFill", baby: "PiBabyFill",
  // Feelings
  happy: "PiSmileyFill", sad: "PiSmileySadFill", calm: "PiSmileyMehFill", angry: "PiSmileyAngryFill",
  worried: "PiSmileyNervousFill", wink: "PiSmileyWinkFill", shocked: "PiSmileyXEyesFill",
  blank: "PiSmileyBlankFill", sticker: "PiSmileyStickerFill",
  // School and making
  book: "PiBookFill", openBook: "PiBookOpenTextFill", books: "PiBooksFill", notebook: "PiNotebookFill",
  pencil: "PiPencilFill", pen: "PiPenFill", brush: "PiPaintBrushFill", palette: "PiPaletteFill",
  scissors: "PiScissorsFill", ruler: "PiRulerFill", calculator: "PiCalculatorFill",
  maths: "PiMathOperationsFill", chalkboard: "PiChalkboardFill", backpack: "PiBackpackFill",
  eraser: "PiEraserFill", clipboard: "PiClipboardTextFill", newspaper: "PiNewspaperFill",
  article: "PiArticleFill", envelope: "PiEnvelopeFill", chat: "PiChatCircleFill",
  chats: "PiChatsFill", megaphone: "PiMegaphoneFill", speaker: "PiSpeakerHighFill",
  microphone: "PiMicrophoneFill", music: "PiMusicNotesFill", guitar: "PiGuitarFill",
  camera: "PiCameraFill", film: "PiFilmSlateFill", tv: "PiTelevisionFill", game: "PiGameControllerFill",
  puzzle: "PiPuzzlePieceFill", cube: "PiCubeFill", shapes: "PiShapesFill",
  // Time and money
  clock: "PiClockFill", alarm: "PiAlarmFill", timer: "PiTimerFill", hourglass: "PiHourglassFill",
  calendar: "PiCalendarFill", coins: "PiCoinsFill", coin: "PiCoinFill", money: "PiMoneyFill",
  // Places and transport
  house: "PiHouseFill", city: "PiBuildingsFill", castle: "PiCastleTurretFill", tent: "PiTentFill",
  car: "PiCarFill", bus: "PiBusFill", bike: "PiBicycleFill", train: "PiTrainFill",
  plane: "PiAirplaneFill", boat: "PiBoatFill", sailboat: "PiSailboatFill", truck: "PiTruckFill",
  shop: "PiStorefrontFill", hospital: "PiHospitalFill",
  // Food
  carrot: "PiCarrotFill", bread: "PiBreadFill", pizza: "PiPizzaFill", cookie: "PiCookieFill",
  iceCream: "PiIceCreamFill", burger: "PiHamburgerFill", cup: "PiCoffeeFill", cutlery: "PiForkKnifeFill",
  orange: "PiOrangeSliceFill", cherries: "PiCherriesFill", avocado: "PiAvocadoFill",
  // Symbols, signals and routines
  tick: "PiCheckCircleFill", cross: "PiXCircleFill", question: "PiQuestionFill", warning: "PiWarningFill",
  info: "PiInfoFill", flag: "PiFlagFill", trophy: "PiTrophyFill", medal: "PiMedalFill",
  crown: "PiCrownFill", gift: "PiGiftFill", thumbsUp: "PiThumbsUpFill", thumbsDown: "PiThumbsDownFill",
  stop: "PiHandPalmFill", point: "PiHandPointingFill", exit: "PiDoorOpenFill", talk: "PiChatCircleDotsFill",
  write: "PiNotePencilFill", read: "PiBookOpenTextFill", listen: "PiEarFill", look: "PiEyeFill",
  partners: "PiUsersFill", whiteboard: "PiChalkboardSimpleFill", think: "PiHeadCircuitFill",
  draw: "PiPencilLineFill", build: "PiCubeFill", act: "PiMaskHappyFill", sort: "PiFunnelFill",
  plus: "PiPlusFill", minus: "PiMinusFill", times: "PiXFill", divide: "PiDivideFill",
  equals: "PiEqualsFill", percent: "PiPercentFill", arrowRight: "PiArrowRightFill",
  arrowDown: "PiArrowDownFill", cycle: "PiArrowsClockwiseFill", dice: "PiDiceFiveFill",
  // Signing-room cues (see CUES below)
  voicesOff: "PiMicrophoneSlashFill", eyesUp: "PiEyeFill", watchCopy: "PiCopyFill",
};

/**
 * The fixed cue set for a signing room, from the Auslan slides prompt section 6.
 *
 * Six cues and only six. The same image and the same words mean the same thing
 * on every slide of every Auslan deck, so a student with a working-memory load
 * reads the cue by shape before they read anything else. Adding a seventh is a
 * change to the prompt, not a change to a build script.
 */
const CUES = {
  voicesOff: { pictogram: "voicesOff", label: "Voices off", tone: "ALERT" },
  eyesUp: { pictogram: "eyesUp", label: "Eyes up", tone: "PRIMARY" },
  watchCopy: { pictogram: "watchCopy", label: "Watch, then copy", tone: "ASSESS" },
  partner: { pictogram: "partners", label: "Partner", tone: "SUCCESS" },
  whiteboards: { pictogram: "whiteboard", label: "Whiteboards", tone: "SECONDARY" },
  timer: { pictogram: "timer", label: "Timer", tone: "ACCENT" },
};

/** Subject glyphs for title and closing slides. */
const SUBJECT_PICTOGRAMS = {
  numeracy: "maths",
  literacy: "openBook",
  science: "flask",
  inquiry: "compass",
  wellbeing: "handHeart",
};

function listPictograms() {
  return Object.keys(PICTOGRAMS).sort();
}

function resolveIconComponent(name) {
  const iconName = PICTOGRAMS[name];
  if (!iconName) return null;
  return Pi[iconName] || null;
}

const pngCache = new Map();

/**
 * Render a pictogram to a base64 PNG data string (PptxGenJS `data` format).
 * @param {string} name      catalogue key
 * @param {string} colorHex  glyph colour, 6-char hex
 * @param {number} [px]      raster size in pixels (default 320)
 * @returns {string|null}    "image/png;base64,..." or null for unknown names
 */
function renderPictogramPng(name, colorHex, px) {
  const Comp = resolveIconComponent(name);
  if (!Comp) return null;
  const color = String(colorHex || "FFFFFF").replace("#", "").toUpperCase();
  const size = Math.max(64, Math.min(1024, Math.round(px || 320)));
  const key = `${name}|${color}|${size}`;
  if (pngCache.has(key)) return pngCache.get(key);

  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: `#${color}`, size: String(size) })
  );
  const R = getResvg();
  // loadSystemFonts:false matters: icons carry no text, and the font scan
  // costs ~5 seconds per render on macOS.
  const png = new R(svg, { fitTo: { mode: "width", value: size }, font: { loadSystemFonts: false } }).render().asPng();
  const data = "image/png;base64," + Buffer.from(png).toString("base64");
  pngCache.set(key, data);
  return data;
}

/**
 * Factory bound to a palette. Returns synchronous helpers.
 */
function createPictogramHelpers(C, FONT_B, el, S) {
  const sz = S || DEFAULT_SIZES;

  function warnUnknown(name) {
    console.warn(
      `WARN [pictogram] unknown pictogram "${name}". Use one of: ${listPictograms().join(", ")}`
    );
  }

  /**
   * Draw one pictogram.
   *
   * @param {object} slide
   * @param {string} name   catalogue key (see PICTOGRAMS)
   * @param {number} x      left edge (inches)
   * @param {number} y      top edge (inches)
   * @param {number} size   circle/tile diameter (inches); the glyph is ~58% of it
   * @param {object} [opts] {
   *   style: "circle" | "tile" | "flat"   (default circle)
   *   color: fill colour for circle/tile, or glyph colour for flat (default C.PRIMARY)
   *   glyphColor: override glyph colour (default WHITE on circle/tile)
   *   label: short caption under the pictogram
   *   labelColor, labelFontSize, labelW
   * }
   * @returns {{x,y,w,h,labelBottom}} drawn bounds
   */
  function addPictogram(slide, name, x, y, size, opts) {
    const o = opts || {};
    const style = o.style || "circle";
    const color = o.color || C.PRIMARY;
    const glyphColor = o.glyphColor || (style === "flat" ? color : C.WHITE);
    const d = Math.max(0.3, Number(size) || 0.8);

    const data = renderPictogramPng(name, glyphColor, Math.round(d * 220));
    if (!data) {
      warnUnknown(name);
      slide.addShape("roundRect", {
        x, y, w: d, h: d, rectRadius: d / 2,
        fill: { color: C.MUTED },
      });
    } else {
      if (style === "circle") {
        slide.addShape("roundRect", {
          x, y, w: d, h: d, rectRadius: d / 2,
          fill: { color },
        });
      } else if (style === "tile") {
        slide.addShape("roundRect", {
          x, y, w: d, h: d, rectRadius: Math.min(0.18, d * 0.2),
          fill: { color },
        });
      }
      const glyph = style === "flat" ? d : d * 0.58;
      slide.addImage({
        data,
        x: x + (d - glyph) / 2,
        y: y + (d - glyph) / 2,
        w: glyph,
        h: glyph,
      });
    }

    let labelBottom = y + d;
    if (o.label) {
      const labelFontSize = o.labelFontSize || byBand(sz, 22, 19, 15);
      const labelH = Math.max(0.3, labelFontSize * 0.022 + 0.12);
      const labelW = o.labelW || Math.max(d + 1.2, 1.6);
      slide.addText(String(o.label), {
        x: x + d / 2 - labelW / 2, y: y + d + 0.06,
        w: labelW, h: labelH,
        fontSize: labelFontSize, fontFace: FONT_B,
        color: o.labelColor || C.CHARCOAL, bold: true,
        align: "center", valign: "top", margin: 0,
        fit: "shrink", shrinkText: true,
      });
      labelBottom = y + d + 0.06 + labelH;
    }
    return { x, y, w: d, h: d, labelBottom };
  }

  /**
   * Evenly spaced row of pictograms with labels, sized to the width given.
   * items: ["happy", "sad"] or [{ name, label, color }]
   */
  function addPictogramRow(slide, x, y, w, items, opts) {
    const o = opts || {};
    const list = (Array.isArray(items) ? items : [items])
      .map((it) => (typeof it === "string" ? { name: it } : (it || {})))
      .filter((it) => it.name);
    if (!list.length) return { size: 0, bottom: y };
    const gap = o.gap != null ? o.gap : 0.3;
    const maxSize = o.size || byBand(sz, 1.6, 1.4, 1.15);
    const size = Math.min(maxSize, (w - gap * (list.length - 1)) / list.length);
    const totalW = size * list.length + gap * (list.length - 1);
    const startX = x + (w - totalW) / 2;
    const palette = o.colors || [C.PRIMARY, C.SECONDARY, C.ACCENT, C.ALERT, C.SUCCESS, C.ASSESS];
    let bottom = y + size;
    list.forEach((it, i) => {
      const px = startX + i * (size + gap);
      const geo = addPictogram(slide, it.name, px, y, size, {
        style: o.style,
        color: it.color || (o.color || palette[i % palette.length]),
        glyphColor: o.glyphColor,
        label: it.label != null ? it.label : (o.labels === false ? "" : it.name),
        labelFontSize: o.labelFontSize,
        labelW: size + gap - 0.04,
      });
      bottom = Math.max(bottom, geo.labelBottom);
    });
    return { size, bottom, startX, totalW };
  }

  /**
   * The signing-room cue strip: small labelled chips, always the same image and
   * words for the same cue, in the same place on every slide that uses them.
   *
   * @param {object} slide
   * @param {string[]} names  keys of CUES, in the order they should read
   * @param {object} [opts] {
   *   x, y      top-left (inches). Default: right-aligned to 9.5, y 0.16
   *   align     "right" (default) | "left", used when x is not given
   *   right     right edge to align to (default 9.5)
   *   fontSize  label size; defaults by band
   *   fill      chip fill (default C.BG_LIGHT)
   *   lineColor chip outline (default C.MUTED), so the chip reads on any background
   *   color     glyph colour (default C.PRIMARY)
   * }
   * @returns {{x,y,w,h,chips:number}} drawn bounds
   */
  function addCueStrip(slide, names, opts) {
    const o = opts || {};
    const list = (Array.isArray(names) ? names : [names]).filter(Boolean);
    const unknown = list.filter((n) => !CUES[n]);
    if (unknown.length) {
      console.warn(
        `WARN [cue] unknown cue(s) ${unknown.join(", ")}. The six cues are: ${Object.keys(CUES).join(", ")}`
      );
    }
    const cues = list.filter((n) => CUES[n]).map((n) => CUES[n]);
    if (!cues.length) return { x: 0, y: 0, w: 0, h: 0, chips: 0 };

    // Each cue keeps its own colour on every slide of every deck, so a student
    // reads it by colour and shape before they read the words. Bigger than the
    // first build: these are for students, not for the teacher.
    const fontSize = o.fontSize || byBand(sz, 14, 13, 12);
    const h = Math.max(0.36, fontSize * 0.0235 + 0.13);
    const glyph = h * 0.66;
    const padX = 0.11;
    const gap = 0.12;
    // Width per chip: padding, glyph, a small gap, then the label. Calibri-ish
    // at this size is about 0.0078in per character per point.
    const widths = cues.map(
      (c) => padX * 2 + glyph + 0.07 + c.label.length * fontSize * 0.0078
    );
    const totalW = widths.reduce((a, b) => a + b, 0) + gap * (cues.length - 1);

    const y = o.y != null ? o.y : 0.16;
    let x = o.x;
    if (x == null) {
      const right = o.right != null ? o.right : 9.5;
      x = o.align === "left" ? 0.5 : right - totalW;
    }
    const startX = x;

    cues.forEach((c, i) => {
      const w = widths[i];
      const tone = o.color || C[c.tone] || C.PRIMARY;
      // Outlined as well as filled: a deck's background is sometimes the same
      // colour as the chip fill, and an invisible cue is worse than none.
      slide.addShape("roundRect", {
        x, y, w, h, rectRadius: h / 2,
        fill: { color: o.fill || C.BG_CARD || C.WHITE },
        line: { color: tone, width: 1.25 },
      });
      const data = renderPictogramPng(c.pictogram, tone, Math.round(glyph * 220));
      if (data) {
        slide.addImage({ data, x: x + padX, y: y + (h - glyph) / 2, w: glyph, h: glyph });
      }
      slide.addText(c.label, {
        x: x + padX + glyph + 0.04, y,
        w: w - padX * 2 - glyph - 0.04, h,
        fontSize, fontFace: FONT_B, bold: true,
        color: o.labelColor || C.CHARCOAL,
        align: "left", valign: "middle", margin: 0,
      });
      x += w + gap;
    });

    return { x: startX, y, w: totalW, h, chips: cues.length };
  }

  return {
    addPictogram,
    addPictogramRow,
    addCueStrip,
    renderPictogramPng,
    hasPictogram: (name) => Boolean(PICTOGRAMS[name]),
    listCues: () => Object.keys(CUES),
  };
}

module.exports = {
  PICTOGRAMS,
  CUES,
  SUBJECT_PICTOGRAMS,
  listPictograms,
  renderPictogramPng,
  createPictogramHelpers,
};
