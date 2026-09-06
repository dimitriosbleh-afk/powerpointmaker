"use strict";

// War Horse Unit -- Lesson 23: Chapter 18 -- Joey's Illness
// Week 5, Session 3, Grade 5/6 Literacy
// Chapter 18 reading + Topic Sentences (sentence types + vivid words)

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  addTextOnShape, withReveal,
  titleSlide, liSlide, contentSlide,
  cfuSlide, taskSlide, closingSlide,
  vocabSlide, quoteSlide, modellingSlide,
} = require("../themes/wh5_helpers");

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 3;
const FOOTER = "War Horse | Lesson 23 of 25 | Week 5 | Year 5/6 Literacy";
const OUT_DIR = "output/WH5_Lesson23_Joeys_Illness";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const WORKSHEET_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Topic Sentences Worksheet",
  "Student worksheet: improve three boring topic sentences using sentence types and vivid words, then write a paragraph."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Answer Key",
  "Teacher reference: model improved topic sentences with alternative approaches and paragraph examples."
);
const RESOURCE_ITEMS = [WORKSHEET_RESOURCE, ANSWER_KEY_RESOURCE];
const WORKSHEET_PDF_PATH = path.join(OUT_DIR, WORKSHEET_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to Lesson 23. We are reading Chapter 18 of War Horse today
- Last lesson, Chapter 17 -- the reunion. Albert finally recognised Joey under all the mud and blood
- Chapter 18 picks up with Albert caring for Joey at the veterinary hospital. But something goes wrong

DO:
- Display title slide as students settle
- Have copies of War Horse on desks, bookmarked at Chapter 18 (p. 146)

TEACHER NOTES:
Lesson 23 of 25, third session of Week 5. Chapter 18 is a high-stakes chapter -- Joey contracts tetanus and nearly dies. The emotional investment students have built makes this especially gripping. The topic sentence work is a practical writing skill that complements the reading.

WATCH FOR:
- Students who are anxious about Joey after the reunion -- reassure through framing: "This chapter has a challenge, but keep reading"
- Students who need a brief recap of Chapter 17 -- Albert and Joey reunited at the veterinary hospital

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands again today: reading Chapter 18 and writing stronger topic sentences
- Read the success criteria. A topic sentence is the first sentence of a paragraph -- it tells the reader what the paragraph is about
- Today we are making boring topic sentences more interesting using two strategies

DO:
- Choral read the LI, then the SCs
- Do not explain the strategies yet -- the I Do covers this

TEACHER NOTES:
SC1 targets character analysis through actions and dialogue during a crisis. SC2 targets the literary devices (hyperbole and repetition) that Morpurgo uses in this chapter. SC3 is the sentence-level writing target -- improving topic sentences.

WATCH FOR:
- Students who think "topic sentence" just means the first sentence -- push: it is the first sentence AND it tells the reader the main idea of the paragraph
- Students who are eager to find out what happens to Joey -- channel that into careful reading

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_ZEALOUS = `SAY:
- Our first vocabulary word: zealous. Read it with me: zealous [students repeat]
- Zealous means showing great energy and enthusiasm for something. A zealous person throws themselves into a task with passion and dedication
- In Chapter 18, Albert is zealous in his care for Joey. He does not give up, even when the situation looks hopeless
- Quick check: if a student was zealous about their sport, would they train once a week or every day? [Every day -- zealous means passionate and dedicated]

DO:
- Display word, choral read, define, give example
- Connect to the chapter: Albert's zealous care is what saves Joey's life
- Cold Call 2 students: what is something you are zealous about?

TEACHER NOTES:
Pre-teaching "zealous" before reading primes students to notice Albert's dedication throughout Chapter 18. The word also connects to "zealously" which appears on p.147.

WATCH FOR:
- Students who confuse "zealous" with "jealous" -- they sound similar but mean completely different things. Zealous is about passion, jealous is about envy
- Students who only associate enthusiasm with fun activities -- zealous can also mean passionate about serious, difficult work

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_DREAD = `SAY:
- Second word: dread. Read it with me: dread [students repeat]
- Dread means a deep fear or anxiety about something that might happen. It is stronger than just being worried -- dread is a heavy, sinking feeling
- In Chapter 18, there are moments of real dread -- will Joey survive? The characters feel dread, and Morpurgo makes the reader feel it too
- Ask: What is the difference between being nervous and feeling dread? [Dread is deeper and heavier. Nervous is butterflies. Dread is a weight in your stomach]

DO:
- Display word, choral read, define, example
- Quick synonym check: fear, terror, anxiety -- but dread is specifically about something COMING, not something happening right now
- Ask 1-2 students: when have you felt a sense of dread?

TEACHER NOTES:
"Dread" is a powerful Tier 2 word for descriptive writing. It connects to the emotional tone of Chapter 18 and gives students a precise word for anticipatory fear.

WATCH FOR:
- Students who equate dread with general sadness -- refine: "Dread is specifically about fearing something that MIGHT happen. It is forward-looking"
- Students who may connect to personal anxieties -- handle sensitively, keep examples light

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_READING_INTRO = `SAY:
- Chapter 18. Albert is now Joey's official carer at the veterinary hospital
- Everything seems to be going well -- but then something goes very wrong
- Today's reading is Student Read Aloud again. I will select readers. We have three pause points
- Focus: pay attention to how the characters RESPOND to a crisis. What do their actions and words reveal?
- Pages 146 to 157 -- find your page now

DO:
- Give students 30 seconds to find p. 146
- Select first reader -- choose a confident reader for the opening
- Plan reader rotations at natural breaks

TEACHER NOTES:
Chapter 18 moves from calm to crisis to recovery. Joey contracts tetanus, and Albert and David fight to save him with Major Martin's help. The emotional arc is steep -- students will feel the tension.

WATCH FOR:
- Readers who rush through the tense sections -- coach: "Slow down through the difficult parts. Let the tension build"
- Students who are anxious when Joey becomes ill -- this is a sign of deep engagement with the text

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. "You can't just give up, sir. You can't. Not with Joey." Page 151
- What is going on? [Joey has tetanus. The vet, Major Martin, says he is likely to die. Albert is begging him not to give up]
- Ask: What does this dialogue reveal about Albert's character? [He is desperate, loyal, brave enough to challenge an officer. He refuses to accept that Joey might die]
- Notice the repetition: "You can't... You can't." Why does Albert repeat himself? [He is overwhelmed with emotion. The repetition shows his desperation and determination]

DO:
- Display the quote, read aloud with emotional weight
- Give 10 seconds of think time before taking responses
- Use Think-Pair-Share: 30 seconds think, 30 seconds pair, then share

CFU CHECKPOINT:
Technique: Think-Pair-Share

Script:
- "Think for 30 seconds: what is Albert's state of mind right now, and how do you know?"
- "Share with your partner for 30 seconds"
- "Hands down -- I am selecting pairs. [Name], what did your pair discuss?"
- Scan for: students identifying Albert's desperation, loyalty, and willingness to challenge authority for Joey

PROCEED (>=80%): Most pairs recognise Albert's emotional state and character traits. Continue reading.
PIVOT (<80%): Most likely issue -- students describe WHAT is happening but not what it REVEALS about Albert. Reteach: "Albert is a private soldier talking to a Major. In the army, you do not tell officers what to do. So what does it mean that Albert says 'You can't just give up'? What kind of person does that?" Re-check: "What one word would you use to describe Albert in this moment?"

TEACHER NOTES:
This quote is rich for character analysis (SC1). Albert's repetition also foreshadows the literary device focus. The power imbalance -- a private challenging an officer -- makes Albert's devotion even more remarkable.

WATCH FOR:
- Students who focus only on Albert being "upset" -- push for precision: "Yes, but WHAT KIND of upset? Angry? Desperate? Determined?"
- Students who notice Albert addressing an officer directly -- excellent observation about social hierarchy

[General: Pause Point 1 | VTLM 2.0: Teacher-Led Discussion]`;

const NOTES_PAUSE2 = `SAY:
- Stop here. "You heard what the officer said." Page 154
- How are things looking for Joey now? [Major Martin has agreed to try treating Joey, but it will require constant supervision. Albert and David have promised to do whatever it takes]
- Ask: Why is it significant that Major Martin agrees to help? What changed his mind? [Albert and David's determination. Their willingness to do the hard work of constant care convinced him it was worth trying]

DO:
- Display the quote
- Use Cold Call for responses
- Push for inference -- this is about what convinced the vet, not just what happened

TEACHER NOTES:
The vet's decision to try is a turning point. He did not change his medical opinion -- he changed his willingness to attempt treatment because Albert and David guaranteed the labour-intensive care Joey would need. This connects to the theme of human dedication overcoming impossible odds.

WATCH FOR:
- Students who think the vet just "felt sorry" for Albert -- push deeper: "He is a military vet. He does not waste resources on sentiment. What PRACTICAL reason did Albert give him?"
- Students who notice the collaborative effort between Albert and David -- this partnership is important

[General: Pause Point 2 | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_PAUSE3 = `SAY:
- Final pause. "I can't wait to see the look on Father's face when I bring you back up the lane. I just can't wait." Page 157
- What just happened? [Joey has recovered from tetanus. After weeks of treatment, he has made a full recovery. Albert is talking to Joey about going HOME]
- Ask: Why is this moment so powerful? What has the reader been through to get here? [We have watched Joey through the entire war -- suffering, danger, near death. Now Albert is talking about taking him HOME. It is the promise of the whole novel coming true]
- Notice: "I can't wait... I just can't wait." More repetition. What effect does it create? [Joy, excitement, relief. The repetition builds Albert's overwhelming happiness]

DO:
- Display the quote, read with warmth and relief in your voice
- Let the moment land -- after the tension of tetanus, this is release
- Cold Call 2-3 students for their interpretation

TEACHER NOTES:
This quote bookends the novel's emotional arc. Albert's vision of bringing Joey home completes the promise made in Chapter 1. The repetition mirrors Albert's earlier desperate repetition at Pause Point 1 -- but now the emotion is joy instead of fear.

WATCH FOR:
- Students who connect this to the beginning of the novel -- excellent full-arc thinking
- Students who notice the contrast between this repetition (joy) and the earlier one (desperation) -- this is a sophisticated observation about Morpurgo's technique

[General: Pause Point 3 | VTLM 2.0: Deep Comprehension]`;

const NOTES_PAUSE3_REVEAL = `SAY:
- Morpurgo uses two key literary devices in this chapter
- First: hyperbole. Look at page 147: "the nightmare I had lived through seemed to fade into unreality, and the war itself was suddenly a million miles away." A million miles away is exaggeration -- hyperbole. It shows how completely Joey's recovery erased the trauma of war for Albert
- Second: repetition again. Page 152: "No horse, no guns. No horse, no ammunition. No horse, no cavalry. No horse, no ambulances. No horse, no water for the troops." Five repetitions of "No horse." What is Morpurgo doing? [Showing how essential horses are to the war effort. Each repetition adds another thing that depends on horses. It builds the case that horses are irreplaceable]
- Both devices create EMPHASIS -- they make the reader feel more strongly

DO:
- Reveal the analysis card
- Read the repetition passage aloud -- let students hear the rhythm
- Draw the connection: hyperbole amplifies emotion, repetition amplifies importance
- Transition: "Now let's move to our topic sentence writing"

[General: Pause Point 3 Reveal | VTLM 2.0: Deep Comprehension]`;

const NOTES_TS_IDO = `SAY:
- Now sentence-level writing. Today we are improving boring topic sentences
- A topic sentence is the FIRST sentence of a paragraph. It tells the reader what the paragraph is about. A strong topic sentence makes the reader want to keep reading
- Two strategies for improving a boring topic sentence
- Strategy 1: Change the SENTENCE TYPE. We have four types: statement, command, exclamation, and question. A boring topic sentence is usually a plain statement. Watch what happens when I change the type
- Here is a boring TS: "Joey got sick." That is a flat statement. Now watch:
- Exclamation: "Tetanus! Joey became the latest victim." -- The exclamation grabs attention immediately
- Strategy 2: Use VIVID WORDS. Replace dull adjectives with precise, interesting ones
- Same boring TS: "Joey got sick." Now with vivid words: "Joey began to suffer from the cruel effects of tetanus." -- "Suffer" and "cruel" are much more powerful than "got sick"

DO:
- Display the two strategies clearly
- Model each one with the Joey example from the prompt
- Think aloud: "I am asking myself -- what sentence type will hook the reader? And -- which words are dull and could be replaced with something vivid?"
- Point to the changes in each model sentence

MISCONCEPTIONS:
- Misconception: A topic sentence is just any sentence that starts a paragraph
  Why: Students confuse position with function. A topic sentence does not just sit at the top -- it announces the main idea
  Impact: Without this understanding, students write opening sentences that do not set up the paragraph
  Quick correction: "A topic sentence is a promise to the reader. It says: this is what the paragraph will be about. If the rest of the paragraph does not match the promise, the topic sentence is wrong"

TEACHER NOTES:
The two strategies are complementary. Changing sentence type shifts the FEEL. Using vivid words sharpens the CONTENT. Students can use either or both. The modelling uses the exact examples from the lesson plan.

WATCH FOR:
- Students who think exclamation marks automatically make writing better -- they are one tool, not a universal fix
- Students who confuse vivid words with long words -- vivid means precise and evocative, not complicated

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Let's try one together. Here is a boring topic sentence: "The vet helped Joey."
- That is a flat statement. What could we do to improve it?
- Strategy 1: Change the sentence type. Could we make it a question? "Could Major Martin really save a horse with tetanus?" -- A question draws the reader in
- Strategy 2: Vivid words. What could we use instead of "helped"? [Take suggestions: treated, fought to save, battled, dedicated himself to]
- Let's build one together: "Against all odds, Major Martin dedicated himself to saving the stricken horse."
- What makes that better? [It has vivid words: "against all odds," "dedicated," "stricken." It paints a picture]

DO:
- Display the boring TS: "The vet helped Joey."
- Take student suggestions for sentence type changes first
- Then take suggestions for vivid word substitutions
- Co-construct one improved version as a class
- Write it up clearly so students see the transformation

CFU CHECKPOINT:
Technique: Thumbs Up/Down

Script:
- Show two options: A) "The vet tried to help the horse." B) "Desperately, Major Martin battled to save the fading horse from the grip of tetanus."
- "Which is a stronger topic sentence? Thumbs up for A, thumbs down for B."
- Scan for: thumbs down (B is stronger -- vivid words, specific detail)

PROCEED (>=80%): Most show thumbs down. Move to You Do.
PIVOT (<80%): Most likely issue -- students choose A because it is "clearer" or "simpler." Reteach: "Both sentences are clear. But which one makes you want to read MORE? B creates a picture -- 'battled,' 'fading,' 'grip of tetanus.' A tells you what happened. B makes you FEEL it. A strong topic sentence hooks the reader." Re-check with: "Which topic sentence would make you turn the page?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide a word bank of vivid alternatives (e.g., struggled -> battled, sick -> stricken, helped -> dedicated) and let students choose from the bank rather than generating their own
- Extra Notes: Students can use the I Do model as a template and substitute words from the bank

EXTENDING PROMPT:
- Task: Write two different improved versions of the SAME boring TS -- one using sentence type change, one using vivid words. Then write a brief explanation of which is more effective and why

TEACHER NOTES:
The We Do uses a DIFFERENT sentence from the You Do stems. "The vet helped Joey" is distinct from the three You Do stems (about Albert, David, and Joey's recovery). This ensures genuine transfer in the You Do.

WATCH FOR:
- Students who only suggest adding adjectives -- push: "That is vivid words. But what about changing the sentence TYPE? Could we make it a question or exclamation?"
- Students who suggest overly complex improvements -- guide toward precise, not complicated

[General: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU = `SAY:
- Quick check before you work independently
- I am going to show you a boring topic sentence and two improved versions. You tell me which strategy each one uses
- Boring: "Albert looked after Joey."
- Version A: "Would Joey have survived without Albert's unwavering care?"
- Version B: "Albert devoted every waking moment to nursing the stricken horse back to health."
- On your whiteboards: write which strategy Version A uses and which strategy Version B uses

DO:
- Display all three sentences
- Allow 30 seconds for whiteboard responses
- Signal: boards up

CFU CHECKPOINT:
Technique: Show Me Boards

Script:
- "Write: Version A = [strategy name]. Version B = [strategy name]. Boards up in 30 seconds."
- Scan for: A = sentence type (changed to question), B = vivid words (devoted, unwavering, stricken, nursing)
- Some students may note B also uses sentence structure changes -- accept and affirm

PROCEED (>=80%): Most correctly identify A = sentence type, B = vivid words. Move to You Do.
PIVOT (<80%): Most likely issue -- students cannot distinguish between the two strategies. Reteach: "Look at Version A. What TYPE of sentence is it? A question. The original was a statement. That is Strategy 1 -- changing the sentence type. Now look at Version B. It is still a statement. But which WORDS are different? 'Devoted,' 'unwavering,' 'stricken,' 'nursing' -- those are all more vivid than 'looked after.' That is Strategy 2." Re-check: "If I wrote 'What a remarkable recovery!' -- which strategy is that?" [Sentence type -- exclamation]

TEACHER NOTES:
This CFU checks that students can identify the strategies before applying them independently. The examples use content from Chapter 18 to maintain textual connection.

WATCH FOR:
- Students who write "both" for Version B -- affirm that it does use stronger structure too, but the PRIMARY strategy is vivid words
- Students who cannot name the strategies -- they may understand them but not remember the labels

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Version A: "Would Joey have survived without Albert's unwavering care?" -- this uses Strategy 1, SENTENCE TYPE. The original statement became a question. The question form draws the reader in
- Version B: "Albert devoted every waking moment to nursing the stricken horse back to health." -- this uses Strategy 2, VIVID WORDS. "Devoted," "waking moment," "nursing," "stricken" -- all much more powerful than "looked after"
- Both are strong improvements. You can use either strategy -- or combine them

DO:
- Reveal the answers
- Highlight the key feature of each strategy
- Transition: "Now it's your turn. You have three boring topic sentences to improve"

[General: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- On the worksheet, you have three boring topic sentences from Chapter 18
- First: choose a strategy for each one -- sentence type, vivid words, or both
- Next: rewrite each topic sentence using your chosen strategy
- Then: pick your BEST improved topic sentence and write a full paragraph underneath it. The paragraph needs 3-4 supporting sentences that match your topic sentence
- You have 10 minutes. Work independently

DO:
- Distribute the Session 3 Topic Sentences Worksheet
- Circulate -- check that students are genuinely changing the sentence, not just adding one word
- After 5 minutes, prompt: "You should be starting your paragraph by now"
- Common error: students who improve the TS but then write a paragraph that does not match it

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide the word bank and sentence type examples from the I Do. Students improve just ONE topic sentence (their choice) and write a 2-sentence supporting paragraph
- Extra Notes: Students can work from the models and substitute words

EXTENDING PROMPT:
- Task: After completing all three TS improvements and one paragraph, write a second paragraph from a DIFFERENT improved TS. Compare: how does the topic sentence shape what comes next in the paragraph?

TEACHER NOTES:
The three TS stems are deliberately varied: TS1 is about Albert's care, TS2 about David's advocacy, TS3 about Joey's recovery. Students must think about each one independently, not use a template.

WATCH FOR:
- Students who improve the TS but write a mismatched paragraph -- check: "Does your paragraph deliver what your topic sentence promised?"
- Students who use only one strategy for all three -- encourage variety
- Students who write very short paragraphs -- push for 3-4 supporting sentences

[General: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: how characters respond to Joey's illness through their actions and dialogue -- thumbs? [scan]
- SC2: identifying hyperbole and repetition and explaining their effect -- thumbs? [scan]
- SC3: improving topic sentences using sentence types and vivid words -- thumbs? [scan]
- Turn and talk: Which strategy do you think you will use more in your own writing -- changing sentence type or using vivid words? Why?

DO:
- Run through each SC with thumbs check
- The turn-and-talk invites metacognition about their own writing preferences
- Preview: "Next lesson we continue with the final chapters of War Horse"

TEACHER NOTES:
The closing connects the writing skill to students' ongoing writing practice. Students who prefer sentence type changes tend to be more structural thinkers; those who prefer vivid words tend to be more descriptive writers. Both are valid.

WATCH FOR:
- Students "thumbs down" on SC2 -- check if it is the literary device identification or the explanation of effect that is the gap
- Students who say "both" for the turn-and-talk -- push for a specific preference with reasoning

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two printable resources today
- The ${WORKSHEET_RESOURCE.name} has three boring topic sentences to improve plus space for a paragraph
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference with model improvements

DO:
- Print the worksheet before the lesson (one per student)
- Print the answer key (teacher copy only)
- Click any resource card to open the PDF

TEACHER NOTES:
The worksheet includes the boring TS, space for improved versions, and lined area for a paragraph. The answer key shows multiple valid improvements using both strategies.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "War Horse - Lesson 23 - Chapter 18: Joey's Illness";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "War Horse",
    "Chapter 18 -- Joey's Illness",
    "Lesson 23  |  Week 5  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  const theme = { C, FONT_H, FONT_B, addTopBar, addTitle, addFooter, addCard };
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    theme,
    FOOTER,
    NOTES_RESOURCES
  );

  // =========================================================================
  // SLIDE 2 -- Learning Intention & Success Criteria
  // =========================================================================
  liSlide(
    pres,
    [
      "We are learning to analyse how characters respond to a crisis, and to write stronger topic sentences using sentence types and vivid words",
    ],
    [
      "I can explain how characters reveal their qualities through their actions and dialogue during a crisis",
      "I can identify hyperbole and repetition and explain the effect they create",
      "I can improve a boring topic sentence using a change in sentence type or vivid word choices",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Vocabulary: zealous
  // =========================================================================
  vocabSlide(
    pres,
    "zealous",
    "adjective",
    "Showing great energy, enthusiasm, and dedication for something. A zealous person throws themselves into a task with passion and does not give up easily.",
    "Albert was zealous in his care for Joey, refusing to leave the horse's side even for a moment during the long weeks of treatment.",
    NOTES_VOCAB_ZEALOUS,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Vocabulary: dread
  // =========================================================================
  vocabSlide(
    pres,
    "dread",
    "noun / verb",
    "A deep, heavy feeling of fear or anxiety about something that might happen. Stronger than worry -- dread is a sinking, overwhelming sense that something terrible is coming.",
    "A cold sense of dread filled Albert as Major Martin examined Joey's stiffening body and shook his head slowly.",
    NOTES_VOCAB_DREAD,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Reading Introduction
  // =========================================================================
  contentSlide(
    pres,
    "Read Aloud",
    C.PRIMARY,
    "Chapter 18 -- Pages 146-157",
    [
      "Reading Mode: Student Read Aloud",
      "Albert continues caring for Joey at the veterinary hospital",
      "Something goes seriously wrong -- Joey becomes dangerously ill",
      "Focus: how do the characters RESPOND to a crisis? What do their actions and words reveal?",
    ],
    NOTES_READING_INTRO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Pause Point 1
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 1", "Chapter 18 -- p. 151",
    "You can't just give up, sir. You can't. Not with Joey.",
    "p. 151",
    "What's going on? What does this reveal about Albert's character?",
    NOTES_PAUSE1, FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- Pause Point 2
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 2", "Chapter 18 -- p. 154",
    "You heard what the officer said.",
    "p. 154",
    "How are things looking for Joey now? What convinced Major Martin to try?",
    NOTES_PAUSE2, FOOTER
  );

  // =========================================================================
  // SLIDES 8-9 -- Pause Point 3 (withReveal) + Literary Devices
  // =========================================================================
  withReveal(
    () => quoteSlide(
      pres, "Pause Point 3", "Chapter 18 -- p. 157",
      "I can't wait to see the look on Father's face when I bring you back up the lane. I just can't wait.",
      "p. 157",
      "What just happened? Why does Morpurgo repeat \"I can't wait\"?",
      NOTES_PAUSE3, FOOTER
    ),
    (slide) => {
      const ansY = 3.50;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 1.50, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });

      // Label pill
      slide.addShape("roundRect", {
        x: 0.7, y: ansY + 0.08, w: 2.2, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("Literary Devices", {
        x: 0.7, y: ansY + 0.08, w: 2.2, h: 0.28,
        fontSize: 11, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Hyperbole
      slide.addText("Hyperbole (p.147): \"the war itself was suddenly a million miles away\" -- exaggeration amplifies the relief of Joey's recovery", {
        x: 3.1, y: ansY + 0.06, w: 6.2, h: 0.36,
        fontSize: 11.5, fontFace: FONT_B, color: C.WHITE, margin: 0,
      });

      // Repetition
      slide.addText("Repetition (p.152): \"No horse, no guns. No horse, no ammunition. No horse, no cavalry...\" -- five repetitions of \"No horse\" build the case that horses are irreplaceable in the war effort. Each one adds another dependency.", {
        x: 0.75, y: ansY + 0.48, w: 8.5, h: 0.92,
        fontSize: 12, fontFace: FONT_B, color: C.WHITE, valign: "top", margin: 0,
      });
      slide.addNotes(NOTES_PAUSE3_REVEAL);
    }
  );

  // =========================================================================
  // SLIDE 10 -- Topic Sentences: I Do
  // =========================================================================
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Improving Topic Sentences",
    "Two strategies to improve a boring TS:\n\nStrategy 1: SENTENCE TYPE\nChange a flat statement into a question, exclamation, or command\n\nStrategy 2: VIVID WORDS\nReplace dull words with precise, interesting ones\n\nBoring TS: \"Joey got sick.\"",
    "Strategy 1 -- Sentence Type:\n\"Tetanus! Joey became the latest victim.\"\n(Exclamation grabs attention)\n\nStrategy 2 -- Vivid Words:\n\"Joey began to suffer from the cruel effects of tetanus.\"\n(\"Suffer\" and \"cruel\" create a stronger picture than \"got sick\")",
    NOTES_TS_IDO,
    FOOTER
  );

  // =========================================================================
  // SLIDES 11-12 -- We Do: Improve a Topic Sentence Together (withReveal)
  // =========================================================================
  withReveal(
    () => contentSlide(
      pres,
      "We Do",
      C.SUCCESS,
      "Improve This Topic Sentence",
      [
        "Boring TS: \"The vet helped Joey.\"",
        "Strategy 1: How could we change the sentence type?",
        "Strategy 2: Which words could be more vivid?",
        "With your partner, discuss one improvement",
      ],
      NOTES_WEDO,
      FOOTER
    ),
    (slide) => {
      const ansY = 3.80;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 1.20, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });

      slide.addShape("roundRect", {
        x: 0.7, y: ansY + 0.10, w: 1.7, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("Our Improved TS", {
        x: 0.7, y: ansY + 0.10, w: 1.7, h: 0.28,
        fontSize: 10, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addText("Question: \"Could Major Martin really save a horse with tetanus?\"\n\nVivid words: \"Against all odds, Major Martin dedicated himself to saving the stricken horse.\"", {
        x: 2.6, y: ansY + 0.06, w: 6.7, h: 1.06,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_WEDO);
    }
  );

  // =========================================================================
  // SLIDES 13-14 -- CFU: Identify the Strategy (withReveal)
  // =========================================================================
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Strategy?",
      "Show Me Boards",
      "Boring: \"Albert looked after Joey.\"\n\nA) \"Would Joey have survived without Albert's unwavering care?\"\nB) \"Albert devoted every waking moment to nursing the stricken horse back to health.\"\n\nOn your whiteboard: which strategy does each version use?",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      const ansY = 4.05;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 4.3, h: 0.9, rectRadius: 0.10,
        fill: { color: C.SECONDARY },
      });
      slide.addText("A = Sentence Type\n(Statement -> Question)", {
        x: 0.7, y: ansY + 0.08, w: 3.9, h: 0.74,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true,
        valign: "middle", margin: 0,
      });

      slide.addShape("roundRect", {
        x: 5.0, y: ansY, w: 4.5, h: 0.9, rectRadius: 0.10,
        fill: { color: C.PRIMARY },
      });
      slide.addText("B = Vivid Words\n(devoted, unwavering, stricken, nursing)", {
        x: 5.2, y: ansY + 0.08, w: 4.1, h: 0.74,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true,
        valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // =========================================================================
  // SLIDE 15 -- You Do: Improve Three Topic Sentences
  // =========================================================================
  taskSlide(
    pres,
    "You Do",
    "Improve These Topic Sentences",
    [
      { label: "FIRST", instruction: "Choose a strategy (sentence type, vivid words, or both) for each TS and rewrite it" },
      { label: "NEXT", instruction: "TS1: Albert helped Joey to get better.  |  TS2: David asked Major Martin to treat Joey.  |  TS3: Joey got better." },
      { label: "THEN", instruction: "Pick your BEST improved TS and write a paragraph with 3-4 supporting sentences underneath it" },
    ],
    NOTES_YOUDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 16 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "Which strategy will you use more in your own writing -- changing sentence type or using vivid words? Why?",
    [
      "I can explain how characters reveal their qualities through their actions and dialogue during a crisis",
      "I can identify hyperbole and repetition and explain the effect they create",
      "I can improve a boring topic sentence using a change in sentence type or vivid word choices",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 17 -- Resources
  // =========================================================================

  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: Topic Sentences Worksheet -------------------------------------
  const ws = createPdf({ title: WORKSHEET_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Improving Topic Sentences", {
    color: C.PRIMARY,
    subtitle: "Chapter 18: Joey's Illness",
    lessonInfo: "War Horse | Lesson 23 | Week 5 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "A topic sentence is the FIRST sentence of a paragraph. It tells the reader what the paragraph will be about. A strong topic sentence hooks the reader.\n\nTwo strategies to improve a boring topic sentence:\n- SENTENCE TYPE: Change a flat statement into a question, exclamation, or command\n- VIVID WORDS: Replace dull words with precise, interesting ones that create a picture", wsY, { color: C.PRIMARY });

  // TS 1
  wsY = addSectionHeading(ws, "Topic Sentence 1:", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "\"Albert helped Joey to get better.\"", wsY, { fontSize: 13, bold: true });
  wsY = addBodyText(ws, "Strategy I will use:", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 24 });
  wsY = addBodyText(ws, "My improved topic sentence:", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 8;

  // TS 2
  wsY = addSectionHeading(ws, "Topic Sentence 2:", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "\"David asked Major Martin to treat Joey.\"", wsY, { fontSize: 13, bold: true });
  wsY = addBodyText(ws, "Strategy I will use:", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 24 });
  wsY = addBodyText(ws, "My improved topic sentence:", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 8;

  // TS 3
  wsY = addSectionHeading(ws, "Topic Sentence 3:", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "\"Joey got better.\"", wsY, { fontSize: 13, bold: true });
  wsY = addBodyText(ws, "Strategy I will use:", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 24 });
  wsY = addBodyText(ws, "My improved topic sentence:", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 12;

  // Paragraph
  wsY = addSectionHeading(ws, "My Paragraph:", wsY, { color: C.PRIMARY, fontSize: 14 });
  wsY = addBodyText(ws, "Choose your BEST improved topic sentence. Write it below, then add 3-4 supporting sentences.", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 8, { lineSpacing: 26 });

  addPdfFooter(ws, "War Horse | Lesson 23 | Topic Sentences Worksheet");

  // --- PDF 2: Answer Key ---------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Improving Topic Sentences -- Answer Key", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Chapter 18: Joey's Illness",
    lessonInfo: "War Horse | Lesson 23 | Week 5 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Accept any improvement that genuinely uses one or both strategies (sentence type change or vivid words). The key test: is the improved version more engaging than the original? Model answers and alternatives below.", akY, { color: C.ALERT });

  // TS 1
  akY = addSectionHeading(ak, "TS1: \"Albert helped Joey to get better.\"", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sentence Type: \"How did one soldier's refusal to give up save a dying horse?\" (Question draws reader in)", akY);
  akY = addBodyText(ak, "Vivid Words: \"Albert devoted every waking hour to nursing Joey through the agony of tetanus.\" (devoted, nursing, agony -- all more powerful)", akY);
  akY = addBodyText(ak, "Combined: \"Day and night, Albert fought to drag Joey back from the brink of death.\" (Exclamatory feel + vivid words)", akY, { italic: true });
  akY += 8;

  // TS 2
  akY = addSectionHeading(ak, "TS2: \"David asked Major Martin to treat Joey.\"", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "Sentence Type: \"Would Major Martin listen to a desperate plea from two young soldiers?\" (Question creates tension)", akY);
  akY = addBodyText(ak, "Vivid Words: \"David pleaded with Major Martin to attempt the gruelling treatment that might save Joey's life.\" (pleaded, gruelling, save -- more powerful)", akY);
  akY = addBodyText(ak, "Combined: \"Please, sir! David's voice cracked as he begged the vet to try.\" (Exclamation + vivid detail)", akY, { italic: true });
  akY += 8;

  // TS 3
  akY = addSectionHeading(ak, "TS3: \"Joey got better.\"", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "Sentence Type: \"Against all odds, a miracle unfolded in the veterinary hospital!\" (Exclamation captures the triumph)", akY);
  akY = addBodyText(ak, "Vivid Words: \"Slowly, painfully, Joey clawed his way back to health after weeks of relentless treatment.\" (clawed, relentless -- show the struggle)", akY);
  akY = addBodyText(ak, "Combined: \"Could anyone have predicted this triumphant recovery?\" (Question + vivid word)", akY, { italic: true });
  akY += 12;

  // Paragraph example
  akY = addSectionHeading(ak, "Model Paragraph (from TS1 improvement):", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "\"Albert devoted every waking hour to nursing Joey through the agony of tetanus. He barely slept, keeping watch by Joey's side as the horse's body trembled and stiffened. When Major Martin warned that hope was fading, Albert refused to accept it. With David's help, he followed every instruction -- cleaning, feeding, and talking softly to Joey through the darkest hours. It was this unwavering dedication that pulled Joey back from the edge.\"", akY);
  akY += 10;

  akY = addSectionHeading(ak, "Common Errors to Watch For", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Adding one adjective and calling it improved -- push for a genuine transformation of the sentence", akY);
  akY = addBodyText(ak, "- Changing the meaning of the TS -- the improved version must still be ABOUT the same topic, just expressed more engagingly", akY);
  akY = addBodyText(ak, "- Paragraph that does not match the topic sentence -- the supporting sentences must deliver on the promise the TS makes", akY);
  akY = addBodyText(ak, "- Overcomplicating -- vivid does not mean long or flowery. Precise, evocative word choices are the goal", akY);

  addPdfFooter(ak, "War Horse | Lesson 23 | Answer Key -- TEACHER COPY");

  // --- Write all files ----------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/WH5_Lesson23.pptx` }),
    writePdf(ws, WORKSHEET_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  const pptxPath = `${OUT_DIR}/WH5_Lesson23.pptx`;
  console.log("PPTX written to " + pptxPath);
  console.log(`Done: ${WORKSHEET_RESOURCE.name}.pdf`);
  console.log(`Done: ${ANSWER_KEY_RESOURCE.name}.pdf`);
}

build().catch(console.error);
