"""Polish teacher notes for Lesson 1 (Storm Boy - sentence expansion) per teachernotes.md v2.0.

Reads the existing "with notes" deck, replaces notes per the map below, and saves
a new file "... - with notes v2.pptx" alongside the source.
"""
from pathlib import Path
import sys

# Ensure we can import the helper from the same scripts/ folder.
sys.path.insert(0, str(Path(__file__).parent))

from write_notes import set_notes_text  # noqa: E402
from pptx import Presentation  # noqa: E402

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/1. Storm Boy - sentence expansion - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")


NOTES = {
    1: """TEACHER NOTES:
Title slide. Begin once materials on slide 8 are ready and the novel is open at page 5.""",

    2: """TEACHER NOTES:
Read aloud or summarise before opening the novel for the first time. Lesson 1 introduces the Coorong setting; Fingerbone Bill enters in Lesson 2.

SENSITIVITY ADVISORY:
- What it is: Storm Boy is set on Ngarrindjeri country and includes Aboriginal characters and language from the period.
- Framing language: "The story shows respect for the Coorong and the people who lived there. Some words from older stories sound different to how we speak today."
- Watch for: students affected by names or images of deceased persons, especially Aboriginal and Torres Strait Islander students.
- Protocol: pause if a student is upset, offer a quiet break with a peer or aide, follow up at recess and with your wellbeing lead if needed.""",

    3: """TEACHER NOTES:
Teacher orientation only, not for students. Read once before delivering the unit. The Literature Study Guide names the pause points and queries used through the lesson.""",

    4: """TEACHER NOTES:
Teacher reference for the I Do, We Do, You Do badges and the support and extension icons used through the deck. Not student-facing.""",

    5: """TEACHER NOTES:
Teacher reference for the response routines used through the deck: whiteboards, choral, thumbs, show fingers, pair share, cold call. Not student-facing.""",

    6: """TEACHER NOTES:
Teacher reference for the sentence-element colour coding used in the modelling slides: who, what doing, when, where, why, how. Not student-facing.""",

    7: """SAY:
- "Read the learning intention with me."
- "These are the three things we are practising today."
- "Ask: which one will be on your worksheet? Expected: SC3, expanding a kernel sentence with who, when and where."
- "If 'adverbial' feels new, that is okay. We will build it together."

DO:
- Choral read the LI, then track each success criterion with your finger.
- Ask one student to say SC1 in their own words.
- Park 'adverbial' on the board. It will be unpacked at the I Do.

TEACHER NOTES:
SC1 is the floor. Almost every student should reach this from today's reading. SC3 is what the worksheet assesses.

WATCH FOR:
- Students unsure about adverbial. Quick gloss: a word or group of words that tells when, where or how.""",

    8: """SAY:
- "Boards out, novel out, booklet ready."
- "Texta in your hand. Lid checked."

DO:
- Scan the room for missing items before reading begins.
- Pair up any students who are missing a board or a working texta.

TEACHER NOTES:
Material check. The reading starts on the next slide, so settle this fast.

WATCH FOR:
- Dry textas. Swap before the first show-me, not during it.""",

    9: """DO:
- Print the worksheet, one per student.
- Print 4 to 6 enabling scaffolds.
- Keep the answer key with your copy, not in the student stack.

TEACHER NOTES:
Teacher setup slide, not student-facing. Do not spend class time talking through the resources here. They are referenced again at the You Do.""",

    10: """TEACHER NOTES:
Section divider. Today's reading mode is Teacher Read Aloud. Have your novel pre-marked with the chosen pause points.""",

    11: """SAY:
- "Listen carefully and picture the setting in your mind."
- "Listen for: where does Storm Boy live? Who lives with him?"
- "I will pause to check, so be ready to think."

DO:
- Read aloud, slow and clear.
- Pause at "wet underbelly" (p.5), "like snakes" (p.6) and "but Storm Boy" (p.6).
- After each pause: 10 seconds silent thinking, 20 seconds partner talk, cold call 1-2 pairs.

CFU CHECKPOINT:
Technique: Think-Pair-Share
Script:
- Ask: what is the author describing when they say 'wet underbelly'? Expected: the underside of the land where it meets the sea.
- Scan for: students linking it to land near the water, not a literal animal underside.
PROCEED:
- >=80% can describe the setting in their own words. Continue reading.
PIVOT:
- Most likely: students take 'underbelly' literally as an animal's belly.
- Reteach: hold a hand up, palm down. "Underbelly is the soft side underneath. The land near the sea is like the underbelly of Australia."
- Re-check: "Is the wet underbelly the sky, or the land near the water?"

TEACHER NOTES:
Name 'like snakes' as a simile in passing. Deeper craft work happens later in the unit. Board notes from this read feed the kernel sentences at the I Do and We Do.

WATCH FOR:
- Students worried about real snakes. Clarify: "It is a comparison, not a real snake."
- Students who cannot recall any details by the third pause. Check in during partner talk and re-anchor with the picture on the slide.""",

    12: """TEACHER NOTES:
Section divider. Today's words are seldom and shrieked, both pulled from pages 5 to 9.""",

    13: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is seldom. Say it with me. Seldom."
- "Seldom means rarely or not often. It is an adverb."
- "Watch this. The people seldom saw Hide-Away Tom or Storm Boy. Seldom tells us how often something happens."
- "Another sentence: it seldom rains during a drought."

DO:
- Point to the image as you give the meaning.
- Choral read the word twice.
- Use a low, slow voice on 'seldom' so the meaning matches the feeling.

TEACHER NOTES:
First meeting with the word. Anchor it in the story example before the True/False checks on the next slides.

WATCH FOR:
- Students saying seldom is the same as never. Correct: "Seldom means almost never, not zero. There is a difference."
- Students confusing seldom with always or often. Re-anchor: "Seldom means NOT often." """,

    14: """SAY:
- "Listen: lions seldom climb trees. True or false?"
- "On your whiteboard. T or F. 5 seconds."
- "Show me!"
- "Expected: True. Lions almost never climb trees."

DO:
- Read twice. Wait. Signal. Scan from back to front.
- Reveal the answer only after boards are up.

WATCH FOR:
- Students writing F because some big cats can climb. Re-anchor: "Lions are big. They almost never bother. That is seldom." """,

    15: """SAY:
- "Listen: it seldom rains in the jungle. True or false?"
- "On your whiteboard. Show me!"
- "Expected: False. Jungles are very wet."

DO:
- Read twice. Signal. Scan.
- Cold call one student to explain why.

WATCH FOR:
- Students confusing jungle and desert. Quick gloss: "Desert is dry. Jungle is wet." """,

    16: """SAY:
- "Listen: people seldom walk on the moon. True or false?"
- "Show me!"
- "Expected: True. Only a few astronauts have ever walked there."

DO:
- Read twice. Signal. Scan.

WATCH FOR:
- Students writing F because they have not seen anyone do it. Reframe: "Seldom means almost never. A few astronauts ever have. That is seldom." """,

    17: """SAY:
- "Listen: koalas are seldom awake during the day. True or false?"
- "Show me!"
- "Expected: True. Koalas sleep about 20 hours a day."

DO:
- Read twice. Signal. Scan.

WATCH FOR:
- Students writing F because they saw a koala awake at the zoo. Confirm: "Awake sometimes, but not often. That is seldom." """,

    18: """SAY:
- "Listen: the sun seldom appears at night. True or false?"
- "Show me!"
- "Expected: False. The sun NEVER appears at night."

DO:
- Read twice. Signal. Scan.
- Make the boundary clear: "Seldom means almost never. Never is different."

TEACHER NOTES:
This item exists to sharpen seldom against never. Name the contrast explicitly before moving on.

WATCH FOR:
- Students writing T because they think seldom and never are the same. Correct: "Seldom is almost never. Never is zero." """,

    19: """SAY:
- "Going overseas is something most families do seldom."
- "On your whiteboard: write or draw something YOU seldom do."
- "30 seconds. Show me!"

DO:
- Time 30 seconds.
- Scan boards.
- Cold call 2-3 students to share. Prompt them to start with "I seldom..."

WATCH FOR:
- Students writing something they do every day. Prompt: "That is OFTEN. Try something you almost never do."
- Readiness signal: most students name something that genuinely happens rarely, not 'watch TV'.""",

    20: """SAY:
- "Two columns on your whiteboard: similar to seldom, different from seldom."
- "Sort the four words. 1 minute. Show me!"
- "Expected: SIMILAR are rarely and not often. DIFFERENT are frequently and most of the time."

DO:
- Time 1 minute.
- Signal, scan.
- Pick a board to share with the class.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: "Hold your sort up." Expected: rarely and not often on similar; frequently and most of the time on different.
- Scan for: correct placement on both sides.
PROCEED:
- >=80% sort all four correctly. Move to the next vocab word.
PIVOT:
- Most likely: 'frequently' placed under similar to seldom.
- Reteach: "Frequently means often. Often is the opposite of seldom."
- Re-check: "Is hardly ever similar to seldom or different?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: cover one column and decide each word at a time. Start with frequently.
EXTENDING PROMPT:
- Task: add one more word of your own to each column. Read it to your partner.

WATCH FOR:
- Students placing all four words in one column. Stop and re-read the column headings together.""",

    21: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is shrieked. Say it with me. Shrieked."
- "Shrieked means to make a loud, sharp, high noise. It is a verb."
- "In the story, the wind shrieked as it passed through the bushes outside the humpy."
- "Another sentence: the child shrieked with delight as he opened his birthday present."

DO:
- Point to each part of the slide as you teach it.
- Choral read the word twice. First low, then high and sharp so they hear the meaning.
- Quick word family: shriek, shrieks, shrieking.

TEACHER NOTES:
The word feels like its meaning when said sharply. Use voice to teach as much as the definition.

WATCH FOR:
- Students saying shriek and scream are the same. Clarify: "Shriek is sharper and higher. A scream can be longer and lower." """,

    22: """SAY:
- "A: a person seeing a mouse run across the floor. B: a dog seeing a cat in the backyard."
- "Where does the word shriek fit? Write A or B. Show me!"
- "Expected: A. People shriek with fright. Dogs bark, growl or yelp."

DO:
- Read both options twice.
- Signal, scan.
- Cold call one student to explain.

WATCH FOR:
- Students picking B because dogs do make loud noises. Correct: "Loud, but not high and sharp like a shriek." """,

    23: """SAY:
- "A: the teacher shrieked the door open. B: the girl shrieked with laughter."
- "Which uses shrieked correctly? Show me!"
- "Expected: B. You cannot shriek a door. Shriek is a sound, not a way to open something."

DO:
- Read both options twice.
- Signal, scan.
- Quick reason: "Shriek is the noise the person makes, not what they do to an object."

WATCH FOR:
- Students unsure why A is wrong. Re-anchor: "Shriek must be a sound from a person." """,

    24: """SAY:
- "A: mum shrieked in pain when she touched the hot stove. B: the waves shrieked in the ocean."
- "Which uses shrieked correctly? Show me!"
- "Expected: A. People shriek in pain. Waves crash or roar, but they do not shriek."

DO:
- Read both twice.
- Signal, scan.

WATCH FOR:
- Students picking B because the ocean is loud. Correct: "Loud, but not high and sharp." """,

    25: """SAY:
- "Which of these would actually make a shrieking sound?"
- "Write all the ones that would shriek. 1 minute. Show me!"
- "Expected: brakes on a car, violin, a siren, door hinges. Not: turning a page, bouncing a basketball."

DO:
- 1 minute writing.
- Scan boards.
- Cold call one student to share, then a second to add or correct.

TEACHER NOTES:
Builds the embodied feel of the word: sharp, high, often unpleasant. The contrast with quiet items like a page and a basketball is what makes the word stick.

WATCH FOR:
- Students adding everything. Reframe: "Only sounds that are loud, sharp and high. A bouncing ball is loud but low." """,

    26: """SAY:
- "A: the campers shrieked with fear when they saw Storm Boy stuck on the beach in the storm."
- "B: the lady gently shrieked a lullaby to help the baby fall asleep."
- "Which uses shrieked INCORRECTLY? Show me!"
- "Expected: B. Gently and shrieked do not go together. A lullaby is soft. A shriek is loud."

DO:
- Read both twice.
- Signal, scan.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: A or B, which uses shrieked incorrectly? Expected: B.
- Scan for: B across the room.
PROCEED:
- >=80% pick B. Move to the You Do.
PIVOT:
- Most likely: students pick A because Storm Boy in danger sounds wrong.
- Reteach: "A fits. The campers are afraid, so they shriek. The wrong one is gently plus shrieked."
- Re-check: "Does gently match shrieked?"

WATCH FOR:
- Students unable to explain. Ask them to describe the volume of a lullaby, then of a shriek.""",

    27: """SAY:
- "In your booklet: write the meaning of seldom, then the meaning of shrieked. Use your own words."
- "Add an example sentence if you have time."
- "Three minutes."

DO:
- Set 3 minutes.
- Circulate, check 2-3 students at the start of writing.
- Hand the scaffold to identified students before they get stuck.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: sentence frame: "Seldom means ___. Shrieked means ___."
EXTENDING PROMPT:
- Task: write one sentence using each word in context. Underline the word.

WATCH FOR:
- Students copying the slide definition word for word. Prompt: "Your own words. Try it without looking." """,

    28: """TEACHER NOTES:
Section divider. The next slides explicitly teach sentence expansion using the kernel sentence and the who/when/where table.""",

    29: """SAY:
- "We expand sentences to give more information, make writing more interesting, and write like real writers."
- "Watch this: 'He ran.' Now: 'After lunch, he ran across the playground.'"
- "Ask: which one tells you more? Expected: the second."

DO:
- Read each bullet on the slide.
- Write the quick example on the board so students can see the difference visually.

WATCH FOR:
- Students saying expanding means making it fancy. Clarify: "Expanding adds INFORMATION: who, when, where. Not fancy words." """,

    30: """SAY:
- "Setting answers when and where."
- "Listen: 'The waves came crashing in from thousands of miles away.' Where? From thousands of miles away."
- "Listen: 'All day and all night, the waves tumble and thunder.' When? All day and all night."
- "Ask: which line tells us when? Expected: 'All day and all night'."

DO:
- Read both lines slowly.
- Underline or point to the where detail and the when detail on the slide.
- Make the link to today's task: "This is what we will do with our own sentences."

WATCH FOR:
- Students who cannot identify when or where. Reteach: point to each phrase and label it on the board.""",

    31: """SAY:
- "Watch this. My kernel sentence: 'They lived here.' That is correct, but it tells me nothing."
- "I need to ask: WHO, WHAT DOING, WHEN, WHERE."
- "WHO: Storm Boy and Hide-Away Tom. WHAT DOING: lived. WHEN: after they left Adelaide. WHERE: in a humpy between the Coorong and the sea."
- "I am going to put the WHEN at the front: 'After they left Adelaide, Storm Boy and Hide-Away Tom lived in a little humpy between the Coorong and the sea.'"
- "Notice the comma after the when detail at the front."

DO:
- Display the kernel and the table.
- Fill in each column as you think aloud.
- Build the expanded sentence under the table.
- Point clearly to the comma.

MISCONCEPTIONS:
- Misconception: students think expanding means adding more adjectives.
  Why: prior 'juicy words' approaches.
  Impact: padded sentences with no new information.
  Quick correction: "Expanding answers WHO, WHEN, WHERE, WHY. Not fancy words."

TEACHER NOTES:
Core I Do. Fronted adverbial taught explicitly with the comma rule.

WATCH FOR:
- Students unsure about adverbial. Reframe: "A group of words that tells when, where or how." """,

    32: """SAY:
- "Your turn with support. New kernel: 'He walked along.'"
- "On your whiteboard: write a WHERE detail. Where did Storm Boy walk? 20 seconds. Show me!"
- "Flip your board. Write a WHEN detail. 20 seconds. Show me!"
- "Now we build it together with the when at the front."

DO:
- Display the kernel.
- Run two whiteboard checks: WHERE first, then WHEN.
- Take 2-3 student contributions, build the sentence on the board live.
- Point to the comma after the fronted adverbial.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: write a WHERE detail. Hold up. Expected: phrases like 'along the beach', 'between the sandhills'.
- Scan for: phrases, not single words.
PROCEED:
- >=80% produce a clear WHERE phrase. Continue with WHEN.
PIVOT:
- Most likely: students write a single word like 'beach'.
- Reteach: "Add a preposition. Along the beach. Near the water. Between the sandhills."
- Re-check with a different kernel: "The bird flew. Add a WHERE phrase."

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: use the scaffold word bank. Drop one WHERE word and one WHEN word into the frame: "___, Storm Boy walked along ___."
- Extra Notes: pair with a confident partner.
EXTENDING PROMPT:
- Task: expand the kernel two ways. When at front, and when at end. Read both. Which sounds better and why?

WATCH FOR:
- Students copying the I Do example. Redirect: "Use a different where or when from the story."
- Students forgetting the comma when fronting. Point to the I Do model on the board.""",

    33: """SAY:
- "First: read each kernel on your worksheet. They come from today's reading."
- "Next: plan your who, when and where."
- "Then: write your expanded sentence with the when at the front. Don't forget the comma."

DO:
- Distribute the worksheet.
- Set 10 minutes.
- Circulate, check the first kernel for each student before letting them run.
- Hand the scaffold to identified students.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: scaffold worksheet with word bank and frames. Aim for 2 of 4 kernels.
- Extra Notes: sit with these students for the first kernel, then release.
EXTENDING PROMPT:
- Task: write a brand new kernel from the story and expand it two ways. Choose your best and explain why in one sentence.

TEACHER NOTES:
Watch the comma rule and the use of text-based detail. Sentences with no link to Storm Boy mean students need a quick re-anchor in the reading.

WATCH FOR:
- Students forgetting the comma after the fronted adverbial. Point to the I Do model.
- Students writing rambling sentences. Refocus: "One who, one when, one where. Keep it clear."
- Readiness signal: 3 of 4 kernels with a correct fronted adverbial and comma.""",

    34: """SAY:
- "Read each I can statement with me."
- "SC1: I can describe the setting and characters from pages 5 to 9. Thumbs."
- "SC2: I can identify similes and personification. Thumbs."
- "SC3: I can expand a kernel with who, when and where. Thumbs."
- "Turn and Talk: tell your partner one detail from today that stuck with you. 30 seconds."

DO:
- Run a thumbs check after each SC, scanning the room before moving on.
- Time the Turn and Talk strictly.
- Cold call 1-2 pairs to share.

TEACHER NOTES:
Use the thumbs data to flag students for a small-group reteach. SC2 is the metalinguistic one and is normally the lowest at first exposure.

WATCH FOR:
- Students thumbs down on SC3. They likely need the scaffold again next lesson.""",

    35: """SAY:
- "Read each I can statement with me."
- "Show me on your thumbs: up, sideways or down for each one."
- "Pick the I can statement you feel most confident about today."

DO:
- Choral read each SC, scan thumbs each time.
- Note which SC has the most thumbs sideways or down to plan tomorrow.

TEACHER NOTES:
Use this data to decide tomorrow's launch and any small-group reteach. If most students show thumbs down on SC3, plan to re-model the comma rule next session.

WATCH FOR:
- Students avoiding the rating. Prompt them to commit to one.
- Patterns where SC2 or SC3 is mostly thumbs down. Flag for reteach.""",
}


def main():
    if not SRC.exists():
        raise SystemExit(f"Source not found: {SRC}")
    prs = Presentation(str(SRC))
    written = 0
    for i, slide in enumerate(prs.slides, start=1):
        if i in NOTES:
            set_notes_text(slide, NOTES[i])
            written += 1
    prs.save(str(OUT))
    print(f"Wrote {written} note slides")
    print(f"Saved: {OUT}")


if __name__ == "__main__":
    main()
