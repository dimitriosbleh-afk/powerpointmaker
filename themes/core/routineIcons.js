"use strict";

const {
  MdEditNote,
  MdGroups,
  MdHearing,
  MdSchool,
  MdRecordVoiceOver,
  MdMenuBook,
  MdContentCut,
  MdExtension,
  MdDraw,
  MdAutoStories,
  MdEdit,
  MdTheaterComedy,
  MdTouchApp,
  MdThumbUp,
  MdLogout,
  MdQuestionAnswer,
} = require("react-icons/md");

const { iconToBase64Png } = require("./icons");

/**
 * Megaprompt §18a classroom routine icon set.
 *
 * Maps each named routine to a Material Design icon component. Used to give
 * Foundation to Year 2 students a consistent visual signal for classroom
 * actions ("show me on your whiteboard", "turn and tell your partner",
 * etc.). The icons must SUPPORT classroom action; they do not replace the
 * mathematical, literacy or concept visual on the slide.
 *
 * Keys match the megaprompt routine list verbatim; aliases are provided for
 * common shorthand callers might use.
 */
const ROUTINES = {
  miniWhiteboard: MdEditNote,
  partnerTalk:    MdGroups,
  listen:         MdHearing,
  teacherModel:   MdSchool,
  turnAndTell:    MdRecordVoiceOver,
  workbook:       MdMenuBook,
  scissorsAndGlue: MdContentCut,
  manipulatives: MdExtension,
  draw:          MdDraw,
  read:          MdAutoStories,
  write:         MdEdit,
  actItOut:      MdTheaterComedy,
  pointTo:       MdTouchApp,
  thumbsUp:      MdThumbUp,
  exitTicket:    MdLogout,
  check:         MdQuestionAnswer,
};

const ROUTINE_LABELS = {
  miniWhiteboard: "Mini-whiteboard",
  partnerTalk:    "Partner talk",
  listen:         "Listen",
  teacherModel:   "Teacher model",
  turnAndTell:    "Turn & tell",
  workbook:       "Workbook",
  scissorsAndGlue: "Cut & paste",
  manipulatives: "Manipulatives",
  draw:          "Draw",
  read:          "Read",
  write:         "Write",
  actItOut:      "Act it out",
  pointTo:       "Point to",
  thumbsUp:      "Thumbs up",
  exitTicket:    "Exit ticket",
  check:         "Check",
};

/**
 * Resolve the icon component for a routine key. Returns null for unknown
 * keys after logging a warning so the build doesn't throw, but the missing
 * icon is loud in build output.
 */
function getRoutineIcon(key) {
  if (!ROUTINES[key]) {
    console.warn(`[routineIcons] Unknown routine "${key}". Valid: ${Object.keys(ROUTINES).join(", ")}`);
    return null;
  }
  return ROUTINES[key];
}

/**
 * Factory that returns routine-icon helpers bound to a palette, body font,
 * and the shared element helpers (uses el.addIconCircle for the badge).
 *
 * @param {object} C       Palette colours
 * @param {string} FONT_B  Body font name
 * @param {object} el      Element helpers from createElements (addIconCircle)
 */
function createRoutineHelpers(C, FONT_B, el) {

  async function addRoutineBadge(slide, routineKey, x, y, opts) {
    const o = opts || {};
    const Icon = getRoutineIcon(routineKey);
    if (!Icon) return;

    const size = o.size || 0.55;
    const circleColor = o.color || C.PRIMARY;
    const iconColor = o.iconColor || C.WHITE;
    const showLabel = o.showLabel !== false;
    const label = o.label != null ? o.label : ROUTINE_LABELS[routineKey];

    const iconData = await iconToBase64Png(Icon, iconColor, 256);
    el.addIconCircle(slide, iconData, x + size / 2, y + size / 2, size / 2, circleColor);

    if (showLabel && label) {
      slide.addText(String(label), {
        x: x - 0.4, y: y + size + 0.02,
        w: size + 0.8, h: 0.22,
        fontSize: o.labelFontSize || 10,
        fontFace: FONT_B,
        color: o.labelColor || C.CHARCOAL,
        bold: true, align: "center", valign: "top", margin: 0,
      });
    }
  }

  return { addRoutineBadge, ROUTINES, ROUTINE_LABELS, getRoutineIcon };
}

module.exports = { createRoutineHelpers, ROUTINES, ROUTINE_LABELS, getRoutineIcon };
