"""Deep-pass teacher notes for Lesson 5 (Write an introduction for an information report) per teachernotes.md v2.0."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/5. literature_presentation Write an introduction for an information report 4En06V Storm Boy - Information report - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. Today is the writing day for the introduction. Begin once Writer's Notebooks are open at Lesson 5.""",

    2: """TEACHER NOTES:
Read aloud or summarise before today's lesson.

SENSITIVITY ADVISORY:
- What it is: today's texts cover Ngarrindjeri culture and language.
- Framing language: 'The Ngarrindjeri people are the Traditional Custodians of the Coorong.'
- Watch for: students unsure how to talk about Aboriginal and Torres Strait Islander content.
- Protocol: model respectful language, pause if a student raises a concern, follow up privately.""",

    3: """TEACHER NOTES:
Teacher orientation only, not for students.""",

    4: """TEACHER NOTES:
Teacher reference for I Do, We Do, You Do badges and support and extension icons. Not student-facing.""",

    5: """TEACHER NOTES:
Teacher reference for the response routines used through the deck. Not student-facing.""",

    6: """TEACHER NOTES:
Teacher reference for sentence-element colour coding. Not student-facing.""",

    7: """SAY:
- "Read the learning intention with me."
- "Today is a writing lesson. We plan, write, then edit our introduction."
- "Ask: what does GST stand for? Expected: General, Specific, Topic."
- "If editing feels new, that is okay. We will work through a checklist together."

DO:
- Choral read the LI, then track each success criterion with your finger.
- Take 2-3 student responses on GST.

TEACHER NOTES:
SC2 is the core. SC3 extends to self-editing.

WATCH FOR:
- Students unable to recall GST. Reassure: 'We will revise it on the next slides.'""",

    8: """SAY:
- "Boards out, Writer's Notebook open at Lesson 5, Literacy Book ready."
- "Texta in your hand. Lid checked."

DO:
- Scan the room for missing items.
- Confirm Writer's Notebooks have yesterday's notes intact.

TEACHER NOTES:
Material check. Today's writing happens in both the Writer's Notebook (planning) and the Literacy Book (final introduction).""",

    9: """TEACHER NOTES:
Section divider. The next slides revisit the texts from Lesson 4 and add a short video for visual context.""",

    10: """SAY:
- "Short video about the Coorong. Two minutes 46 seconds."
- "Listen for facts you could use in your introduction."

DO:
- Play the video.
- Pause briefly at especially useful images.
- Ask: 'What did you see that would help describe the Coorong?' Take 2-3 responses.

TEACHER NOTES:
The video gives a visual sense of the setting and is useful for adjective brainstorming on slide 30.

WATCH FOR:
- Students unable to access the link. Have a backup screenshot or alternate source ready before the lesson.""",

    11: """SAY:
- "Same shorthand as Lesson 4."
- "Slash means new idea. Plus or ampersand means and. Equals means same. Arrow means leads to. b/c means because. w/ means with."

DO:
- Glue page 6 of the booklet into the Writer's Notebook.
- Display the key.
- Quick refresh: ask 'Coorong = important. What does that mean?'

WATCH FOR:
- Students who have forgotten symbols. Keep the key visible during note-taking.""",

    12: """TEACHER NOTES:
Reference slide for the 'Where is the Coorong?' text. Move to summary on the next slides.""",

    13: """TEACHER NOTES:
Reference slide for the Coorong National Park resource. Move to summary.""",

    14: """SAY:
- "Read this passage with me."
- "Ask: what facts? Expected: Coorong NP is a special place in SA's southeast, 200 km SE of Adelaide, long and narrow."
- "Add a fact in shorthand. 1 minute. Show me!"

DO:
- Read the passage aloud.
- Time 1 minute.
- Scan boards.

WATCH FOR:
- Students writing whole sentences. Redirect: 'Just the key idea. Use symbols.'""",

    15: """SAY:
- "New passage. Read with me."
- "Ask: what facts? Expected: home to water birds and migratory species. Cared for by many groups working together."
- "Add a fact. 1 minute. Show me!"

DO:
- Read aloud.
- Time 1 minute.
- Scan boards.

WATCH FOR:
- Students unsure about 'migratory'. Quick gloss: 'Animals that move from place to place.'""",

    16: """SAY:
- "New passage. Read with me."
- "Ask: what facts? Expected: 115 species protected, most important waterbird wetland in the Murray-Darling Basin, designated Ramsar Wetland."
- "Add to your notes. 1 minute. Show me!"

DO:
- Read aloud.
- Highlight 'Ramsar Wetland'.
- Time 1 minute.

WATCH FOR:
- Students recording every number. Redirect: 'Pick the most important fact for the introduction.'""",

    17: """SAY:
- "New passage. Read with me."
- "Ask: what facts? Expected: wetlands important for migratory shorebirds, birds travel from Alaska and Siberia, Ramsar Convention is an international agreement."
- "Add to your notes. 1 minute. Show me!"

DO:
- Read aloud.
- Define Ramsar Convention briefly.
- Time 1 minute.

WATCH FOR:
- Students confusing Ramsar. Repeat the spelling: 'R-A-M-S-A-R.'""",

    18: """SAY:
- "New passage. Read with me, slowly and respectfully."
- "Ask: what facts? Expected: Coorong is important to Ngarrindjeri people. Kurangk = 'long narrow neck'. Sites include middens."
- "Add to your notes. 1 minute. Show me!"

DO:
- Read aloud.
- Pronounce Ngarrindjeri and Kurangk clearly.
- Time 1 minute.

TEACHER NOTES:
Cultural significance content for the specific statement.

WATCH FOR:
- Students unsure how to write Ngarrindjeri. Display the word and let them copy.""",

    19: """SAY:
- "New passage. Read with me."
- "Ask: what facts? Expected: 2009, Ngarrindjeri negotiated KNY Agreement with the state government. Recognised Ngarrindjeri ownership of lands and waters."
- "Add to your notes. 1 minute. Show me!"

DO:
- Read aloud, respectfully.
- Brief explanation: 'KNY means Listen to what Ngarrindjeri people have to say.'
- Time 1 minute.

WATCH FOR:
- Students stalled on the long phrase. Encourage abbreviation: 'Write KNY for short.'""",

    20: """SAY:
- "Ask: what is the purpose of the Ramsar Convention?"
- "1 encourage destruction, 2 prevent destruction of habitats, 3 maintain biodiversity, 4 prevent and reverse the loss of wetlands."
- "Show me with fingers."
- "Expected: 4. Ramsar is specifically about wetlands."

DO:
- Read each option.
- Cue finger response.
- Scan, confirm.

CFU CHECKPOINT:
Technique: Show Fingers 1-4
Script:
- Ask: 'Show me 1, 2, 3 or 4.' Expected: 4.
- Scan for: 4 across the room.
PROCEED:
- >=80% pick 4. Move to writing.
PIVOT:
- Most likely: students pick 2 or 3.
- Reteach: 'Ramsar is about wetlands specifically.'
- Re-check: 'Is Ramsar about wetlands, animals or plants?'

WATCH FOR:
- Students copying neighbours. Cover their hand and re-cue.""",

    21: """TEACHER NOTES:
Section divider. The next slides revise GST and walk through writing each part of the introduction.""",

    22: """SAY:
- "We are continuing our information report on the Coorong."
- "Focus today: writing the introduction using GST."
- "Present tense. Adjectives to build description."

DO:
- Read the slide.
- Quick reminder: 'The Coorong IS, not WAS.'
- Quick adjective example: narrow, shallow, expansive.

WATCH FOR:
- Students slipping into past tense. Quick correction: 'Information reports use present tense.'""",

    23: """SAY:
- "Quick recap. The purpose is to inform."
- "Read with me: 'An information report aims to inform the reader by describing factual information about a topic.'"

DO:
- Choral read.
- Connect to today's task: 'We are informing the reader about the Coorong.'

WATCH FOR:
- Students who treat this as new content. Confirm it is a recap from Lesson 4.""",

    24: """SAY:
- "Recap of the features."
- "Today we focus on two: present tense and adjectives."

DO:
- Read each feature.
- Circle present tense and adjectives on the slide if you can.

WATCH FOR:
- Students trying to use all features at once. Refocus: 'Today, just present tense and good adjectives.'""",

    25: """SAY:
- "Ask: what is the purpose of an information report? 1 entertain, 2 inform, 3 persuade."
- "Show me with fingers."
- "Expected: 2."

DO:
- Read each option.
- Cue finger response.
- Scan, confirm.

CFU CHECKPOINT:
Technique: Show Fingers 1-3
Script:
- Ask: 'Show me 1, 2 or 3.' Expected: 2.
- Scan for: 2 across the room.
PROCEED:
- >=80% pick 2. Move to GST.
PIVOT:
- Most likely: students pick 3 (persuade).
- Reteach: 'Persuade means to convince. Inform means to share facts.'
- Re-check: 'Are we sharing facts or convincing?'

WATCH FOR:
- Students who pick 1 from doubt. Anchor: 'Reports are not stories. They share facts.'""",

    26: """SAY:
- "The introduction has three parts using GST."
- "G is GENERAL statement. It introduces the topic."
- "S is SPECIFIC statement. It uses key terms."
- "T is TOPIC statement. It tells the reader what comes next."
- "Together they orient the reader."

DO:
- Read each part.
- Reuse the funnel analogy: 'General is wide, Specific is narrower, Topic points to the body.'

WATCH FOR:
- Students confusing Specific and Topic. Reuse the funnel.""",

    27: """SAY:
- "Look at the model introduction."
- "First sentence is the GENERAL statement (G). It introduces the topic."
- "Second is the SPECIFIC statement (S). It uses key terms."
- "Third is the TOPIC statement (T). It tells the reader what the body paragraphs cover."

DO:
- Read each sentence aloud.
- Underline G, S and T on the slide if you can.
- Highlight key terms in S.
- Highlight the four discussion points in T.

TEACHER NOTES:
Core I Do for the lesson.

WATCH FOR:
- Students who think GST is just three sentences. Confirm: 'Yes, but each sentence has a job.'""",

    28: """SAY:
- "Ask: which part of GST tells us what the rest of the report will be about?"
- "1 General, 2 Specific, 3 Topic."
- "Show me with fingers."
- "Expected: 3."

DO:
- Read each option.
- Cue finger response.
- Scan, confirm.

CFU CHECKPOINT:
Technique: Show Fingers 1-3
Script:
- Ask: 'Show me 1, 2 or 3.' Expected: 3.
- Scan for: 3 across the room.
PROCEED:
- >=80% pick 3. Move to writing the General statement.
PIVOT:
- Most likely: students pick 1 or 2.
- Reteach: 'Topic statement signposts the body paragraphs.'
- Re-check: 'Which one mentions the body paragraphs? Topic.'

WATCH FOR:
- Students unsure which is which. Reuse the funnel: 'Topic is the narrowest end of the funnel.'""",

    29: """SAY:
- "Look at the model General statement."
- "Ask: what two adjectives does the author use to build description? Turn and tell your partner. 30 seconds."
- "Expected: narrow and shallow."

DO:
- Read aloud.
- Time the partner talk.
- Cold call to confirm.
- Highlight the adjectives.
- Quick connection: 'These build description.'

WATCH FOR:
- Students picking 'Limestone' or 'South'. Clarify: 'Those are part of place names. Look for words that describe what the lagoon looks like.'""",

    30: """SAY:
- "Look at the picture of a lagoon."
- "What adjectives could describe it? Think about shape, size, colours, texture."
- "On your whiteboard: write 3 to 5 adjectives. 2 minutes. Show me!"

DO:
- Display the picture.
- Read the four prompt questions.
- Time 2 minutes.
- Scan boards.
- Cold call 4-5 students.
- Build a class list on the board: shape, size, colour, texture.

WATCH FOR:
- Students writing opinions like 'pretty' or 'amazing'. Redirect: 'Use factual adjectives.'""",

    31: """SAY:
- "Look at the words. Which are appropriate adjectives for the Coorong?"
- "On your whiteboard. 1 minute. Show me!"
- "Expected: narrow, natural, shallow, expansive, inland. Not: aerial."

DO:
- Read each option.
- Time 1 minute.
- Scan boards.
- Confirm with reasons.

WATCH FOR:
- Students including 'aerial'. Clarify: 'Aerial means from the air. It does not describe the Coorong.'""",

    32: """SAY:
- "Time to write our General statement."
- "Look at the two examples on the slide."
- "With your partner: use your notes to write your own General statement. Include adjectives."
- "4 minutes. Show me when you have one!"

DO:
- Display the examples.
- Cue partner work.
- Set 4 minutes.
- Circulate.
- Read 3-4 examples aloud at the end.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: frame: 'The Coorong is a ___ ___ ___ located in ___.' Choose adjectives from the class word bank.
- Extra Notes: sit with these students briefly.
EXTENDING PROMPT:
- Task: write two versions using different adjectives. Choose the stronger one and explain why.

WATCH FOR:
- Students skipping adjectives. Redirect: 'Add at least one adjective.'
- Students writing in past tense. Quick correction: 'Present tense.'""",

    33: """SAY:
- "Open your Writer's Notebook to Lesson 5."
- "Question 2: write your own GENERAL statement."
- "Use your partner's idea or improve it."
- "Aim for one clear sentence with one or two adjectives."

DO:
- Open the Writer's Notebook.
- Set 4 minutes.
- Circulate, check the first sentence for each student.
- Provide a frame for those who need it.

WATCH FOR:
- Students copying the model exactly. Affirm and push: 'Make it your own.'""",

    34: """SAY:
- "SPECIFIC statement gives more information using key terms."
- "Model: 'It is an expansive area of wetlands and waterways that is both culturally and ecologically important.'"
- "Read with me."
- "Key terms: expansive, wetlands, waterways, culturally, ecologically."

DO:
- Choral read the model.
- Underline key terms.
- Brainstorm: 'What other key terms could we use? Lagoon, Limestone Coast, Ramsar wetland, Ngarrindjeri.'

WATCH FOR:
- Students reusing the model exactly. Encourage substitution with one or two new key terms.""",

    35: """SAY:
- "Ask: which could be a specific statement for an introduction about pelicans?"
- "1 South Australia is a wonderful place to live."
- "2 I like pelicans because of their oversized beaks."
- "3 Pelicans are a unique species of seabirds with long bills and huge throat pouches."
- "Show me with fingers."
- "Expected: 3."

DO:
- Read each option.
- Cue finger response.
- Scan, confirm.
- Explain: '1 is not about pelicans. 2 is an opinion. 3 names key terms - species, seabirds, bills, throat pouches.'

CFU CHECKPOINT:
Technique: Show Fingers 1-3
Script:
- Ask: 'Show me 1, 2 or 3.' Expected: 3.
- Scan for: 3 across the room.
PROCEED:
- >=80% pick 3. Move to writing the specific statement.
PIVOT:
- Most likely: students pick 2 because it sounds like the topic.
- Reteach: '2 is an opinion (I like). Reports do not use opinions.'
- Re-check: 'Which item uses key terms, not opinions?'

WATCH FOR:
- Students unsure between 2 and 3. Anchor: 'Look for I or me. That is opinion.'""",

    36: """SAY:
- "Open your Writer's Notebook to Question 3."
- "Write a SPECIFIC statement using your dot points from Lesson 4."
- "Use key terms. Stay factual. 5 minutes."

DO:
- Open the Writer's Notebook.
- Set 5 minutes.
- Circulate, check key terms and no opinions.
- Hand the scaffold to identified students.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: frame: 'It is a ___ ___ that is ___ and ___ important.'
- Extra Notes: sit briefly with these students.
EXTENDING PROMPT:
- Task: write two versions using different key terms (Ramsar wetland, lagoon, etc.).

WATCH FOR:
- Students using opinion language. Redirect: 'Specific statements use factual key terms.'
- Students copying the model. Affirm and push: 'Use a different key term.'""",

    37: """SAY:
- "TOPIC statement tells readers what the rest of the report will discuss."
- "Body Paragraph 1: what it is and where it is located."
- "Body Paragraph 2: cultural significance."
- "Body Paragraph 3: flora and fauna."
- "Our topic statement lists all three."

DO:
- Read each subheading aloud.
- Note the typo: the slide says 'recount'. Read it as 'report' aloud.

TEACHER NOTES:
The slide says 'recount'. We are writing a report. Correct briefly so students do not copy the typo.

WATCH FOR:
- Students confused by the subheadings. Display them clearly.""",

    38: """SAY:
- "Suggest a topic statement. Share with a partner."
- "Two examples on the slide."
- "Notice the signal phrase: 'This report will discuss' or 'The following report will provide'."

DO:
- Read both examples aloud.
- Highlight the signal phrase.
- Time partner discussion 2 minutes.
- Cold call 2 pairs.

WATCH FOR:
- Students adding new body paragraphs. Quick reminder: 'Just the three.'""",

    39: """SAY:
- "Open your Writer's Notebook to Question 4."
- "Write a TOPIC statement."
- "You can copy the model, paraphrase it, or write your own."
- "Aim for one clear sentence that lists all three body topics. 4 minutes."

DO:
- Open the Writer's Notebook.
- Set 4 minutes.
- Circulate, check that all three topics are mentioned.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: copy the model and underline the three topics.
- Extra Notes: copying is a valid first step for very anxious writers.
EXTENDING PROMPT:
- Task: write two versions, one starting with 'This report will discuss' and one with a different signal phrase.

WATCH FOR:
- Students mentioning only one or two body paragraphs. Redirect: 'All three.'""",

    40: """SAY:
- "Now we edit. Look at the editing checklist."
- "Capitalisation - sentence starts and proper nouns: Coorong, Ngarrindjeri, South Australia."
- "Full stops at the end of sentences."
- "Spelling."
- "Reader is oriented to the topic."
- "Adjectives describe the topic."
- "Present tense."

DO:
- Read each item.
- Demonstrate one quick check on the model text: e.g. find a proper noun and check the capital.
- Display the checklist for editing.

TEACHER NOTES:
Anchor for SC3: editing.

WATCH FOR:
- Students who think editing is just spell-checking. Highlight bigger items: orientation, present tense, adjectives.""",

    41: """SAY:
- "Combine your three GST statements into your introduction."
- "Write it in your Literacy Book."
- "Then use the editing checklist to check it."
- "10 minutes."

DO:
- Distribute the Literacy Book.
- Display the model and the checklist.
- Set 10 minutes.
- First 3 minutes: check sentence by sentence with as many students as possible.
- Hand the scaffold to identified students.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: use the model as a frame. Substitute words while keeping the structure.
- Extra Notes: pair with a confident editing partner.
EXTENDING PROMPT:
- Task: swap with a partner. Provide one strength and one suggestion using the checklist.

TEACHER NOTES:
This is the You Do plus self-edit. Outcome is a complete introduction in the Literacy Book.

WATCH FOR:
- Students skipping the editing. Prompt: 'Tick each item before you finish.'
- Readiness signal: 3-sentence introduction in present tense with at least 2 adjectives, checked against the checklist.""",

    42: """TEACHER NOTES:
Credits and attribution slide. Not student-facing. End the lesson on the closing reflection slide rather than this one.""",

    43: """SAY:
- "Read each I can statement with me."
- "Show me on your thumbs: up, sideways or down for each one."
- "Pick the I can statement you feel most confident about today."

DO:
- Choral read each SC, scan thumbs each time.
- Note which SC has the most thumbs sideways or down to plan tomorrow.

TEACHER NOTES:
Use this data to decide tomorrow's launch and any small-group reteach.

WATCH FOR:
- Students avoiding the rating. Prompt them to commit to one.
- Patterns where SC2 or SC3 is mostly thumbs down. Flag for reteach.""",
}


def main():
    stats = polish_deck(SRC, OUT, overrides=NOTES)
    print(f"L05 written: {stats}")


if __name__ == "__main__":
    main()
