---
name: humanizer
description: Rewrite text so it no longer reads as AI-written, using the patterns catalogued in Wikipedia's "Signs of AI writing" field guide, while keeping the original meaning and facts and adding no new information. Use this skill whenever the user asks to humanize a piece of text, or says things like "make this sound natural", "less robotic", "this sounds like ChatGPT", "remove the AI tone", "rewrite so it sounds like a person wrote it", or pastes a draft (post, email, caption, script, report, abstract, lecture note) and complains that it feels stiff, generic or machine-made. Also use it for the same request in Khmer or any other language (e.g. "កែឱ្យដូចមនុស្សសរសេរ", "សរសេរឱ្យធម្មជាតិ").
---

# Humanizer

Rewrite the text the user gives you so that the habits typical of AI-written prose are gone, and what is left reads like something a competent person wrote on purpose.

The pattern list below comes from Wikipedia's "Signs of AI writing" (WikiProject AI Cleanup), a field guide built from real examples of chatbot text.

## Two rules that outrank everything else

1. **Keep the meaning and the facts.** Every claim, number, name, date, quote and the author's stance must survive unchanged. If the original hedges, keep the hedge; if it is certain, stay certain.
2. **Add nothing.** No new examples, statistics, anecdotes, opinions, names or details, even when a sentence would be better with one. AI text is often vague, and the honest fix for a vague sentence is a shorter plain sentence or deleting it, not an invented specific.

These matter because the user will put their name on the result. A rewrite that sounds great but says something they never said is worse than the robotic original.

When removing a pattern would leave an empty sentence (pure puffery with no fact inside), delete the sentence. Shorter output is normal and expected.

## How to work

1. Read the whole text first. Note its language, audience, purpose and register (Facebook post, formal letter, academic paragraph, video script). The rewrite stays in that language and register. Do not make a formal report chatty or a casual post formal.
2. List for yourself the facts and claims the text contains. This is your checklist for the end.
3. Rewrite, clearing the patterns in the catalogue below. Rewrite at the level of the sentence and paragraph. Swapping flagged words for synonyms does not work: "pivotal" replaced by "crucial" is the same tell.
4. Check the result against your fact list: nothing missing, nothing new.
5. Reread the result as a stranger would. If a sentence still sounds like a brochure or a motivational poster, say the plain thing instead.

## Catalogue of signs to remove

### 1. Inflated significance

AI text puffs up its subject by claiming it represents or contributes to something larger.

Watch for: stands as / serves as, is a testament to, a reminder of, plays a vital / significant / crucial / pivotal / key role, underscores / highlights its importance, reflects a broader, enduring / lasting legacy, setting the stage for, marks a shift, turning point, evolving landscape, indelible mark, deeply rooted, in today's fast-paced world, in the era of.

Fix: state what the thing is or did. "Founded in 1989, marking a pivotal moment in the evolution of regional statistics" becomes "Founded in 1989."

### 2. Canned claims of notability

Watch for: has been featured in / profiled by, leading expert, widely recognized, renowned, active social media presence, lists of outlets offered as proof that the subject matters.

Fix: keep the concrete fact (who covered it, what was said) if one is there. Drop the praise wrapped around it.

### 3. Superficial analysis tacked on the end

A present-participle phrase that adds a comment and no information: "..., highlighting its importance", "..., ensuring a seamless experience", "..., reflecting the community's values", "..., contributing to growth", "..., fostering collaboration".

Also: valuable insights, aligns with, resonates with, encompassing.

Fix: cut the tail. If it holds a real claim, give that claim its own sentence with a subject and a verb.

### 4. Promotional tone

Watch for: rich (heritage, history, tapestry), vibrant, breathtaking, stunning, nestled, in the heart of, boasts, world-class, seamless, cutting-edge, game-changer, unlock, elevate, empower, harness, revolutionize, take it to the next level, journey.

Fix: use neutral words. Describe, don't sell. If the text is an advertisement, keep it persuasive but in the words a person would say out loud to a customer.

### 5. Vague attribution

Watch for: experts say, studies show, research suggests, observers note, industry reports, many believe, it is widely considered, with no source named.

Fix: if the original names a source, keep the name. If it does not, you cannot supply one. Either state the claim plainly as the author's own or keep a simple honest hedge ("some people think"). Never invent a study or an expert.

### 6. Formula structure

- **Negative parallelism**: "It's not X. It's Y." / "not just X, but Y" / "X isn't about A, it's about B" / "No fluff, no jargon, just results." Say Y directly.
- **Rule of three**: reflexive triplets of adjectives or phrases ("fast, simple and reliable"). Keep the items that carry meaning; do not pad to three.
- **False ranges**: "from X to Y" where X and Y are not ends of any scale ("from email to strategy").
- **Setup and reveal**: "The result?", "Here's the thing:", "The truth is", "And the best part?", "Let's dive in." Just say the point.
- **Rhetorical question answered at once** by the writer.
- **Quotable closing line**: a short punchy sentence at the end of a paragraph that restates what was just said.
- **"Despite challenges..." formula**: "Despite its success, X faces challenges... Looking ahead, the future remains bright."
- **Summary conclusion**: "In conclusion", "Overall", "Ultimately", "In summary" followed by a restatement. End where the content ends.
- **Elegant variation**: cycling synonyms for one thing ("AI... the technology... these tools... artificial intelligence") to avoid repeating a word. People repeat the word. Pick one term and keep it.
- **Uniform rhythm**: every sentence the same length, every paragraph the same shape, or a mechanical short-short-long beat. Let sentence length follow the content.

### 7. Vocabulary that marks AI text

additionally, moreover, furthermore, notably, importantly, crucial, pivotal, vital, delve, tapestry, landscape (figurative), realm, testament, underscore, highlight (as a verb for abstract things), showcase, foster, cultivate, garner, enhance, leverage, navigate (figurative), robust, comprehensive, intricate, multifaceted, nuanced, meticulous, interplay, embark, ever-evolving, It's worth noting that, It's important to note.

Also intensifiers and sincerity markers: truly, really, genuinely, honestly, incredibly, absolutely, simply.

Fix: use the everyday word (also, important, use, show, improve, look at) or restructure so no word is needed. One of these words in a long text is not a problem; a cluster is.

### 8. Avoiding "is" and "has"

AI text swaps plain verbs for heavier ones: serves as, stands as, represents, functions as, acts as, features, offers, boasts, provides.

Fix: "The library serves as a hub for students" becomes "The library is where students gather" or simply "is a hub for students". Use is, are, has.

### 9. Chatbot residue

- Openers and closers addressed to a user: "Certainly!", "Great question!", "Here's the rewritten version:", "I hope this helps!", "Let me know if you'd like...".
- Disclaimers: "As an AI language model", "As of my last knowledge update", "I don't have access to".
- Unfilled placeholders: [Your Name], [insert date].
- Leftover markup and tracking: stray asterisks or pound signs, `utm_source=chatgpt.com`, `turn0search0`, `citeturn`, broken reference markers.

Fix: delete. If a placeholder is unfilled, leave it visibly marked and tell the user in your note; do not fill it with a guess.

### 10. Formatting tells

- Em dashes used for drama or asides, especially several in one paragraph. Use a comma, a full stop, a colon or parentheses.
- Bold scattered over ordinary phrases.
- Bullet lists where each item starts with a bold label and a colon ("**Speed:** ..."). Turn into normal sentences unless the content really is a list.
- Headings on a short text, Title Case Headings, and emoji used as bullets or section markers.
- Hook / Content / Conclusion / CTA labels left in a text meant to be published as continuous prose.
- Curly quotes mixed with straight quotes.

Match the formatting to where the text will be used. A Facebook post is paragraphs. A formal report may keep its headings.

## What "human" means here

It does not mean sloppy. Do not add typos, slang, filler ("you know", "honestly"), forced contractions, jokes or fake personal stories. Those are a different kind of fake, and a personal story would be new information.

It means:

- Plain statements with concrete subjects and active verbs.
- The same word for the same thing throughout.
- Sentences of different lengths because the ideas are different sizes.
- A paragraph can end without a punchline.
- The level of formality the author chose, kept.

Watch for replacement tells. Clearing one pattern often brings in its neighbour: an em dash becomes a semicolon in every sentence, "not X but Y" becomes "rather than X, Y", "pivotal" becomes "key". If the fix has the same shape as the problem, restructure the sentence.

## Text in Khmer and other languages

The catalogue is written with English examples, but the habits are the same in every language. Apply the patterns by what they do, not by matching English words:

- Inflated significance and promotional adjectives (in Khmer, strings like "ដ៏សំខាន់បំផុត", "ដើរតួនាទីយ៉ាងសំខាន់", "ក្នុងយុគសម័យឌីជីថល", "ដ៏អស្ចារ្យ").
- "Not only... but also" frames ("មិនត្រឹមតែ...ប៉ុណ្ណោះទេ ថែមទាំង..."), reflexive triplets, setup-and-reveal lines.
- Stacked connectors at the start of sentences ("លើសពីនេះទៅទៀត", "ជាងនេះទៅទៀត", "សរុបមក").
- Translationese: sentences that follow English word order or English idiom. Rewrite them the way a native speaker would say them.
- Em dashes and bold labels carried over from English formatting.

Reply in the language of the text. Keep the author's level of politeness and formality, and leave English technical terms (Excel, Email, AI) as the author wrote them.

## Output

Give the rewritten text first, ready to copy, with no preamble in front of it.

After it, add a short note of two to four lines saying what you changed (for example: "cut three significance claims, removed the 'not X but Y' frame in the opening, replaced the bullet labels with sentences"). Mention in the note:

- anything you deleted outright because it carried no information, so the user can restore it if they disagree;
- any claim that was vague in the original and that you could not make specific without inventing facts, so the user can supply the detail themselves;
- any unfilled placeholder.

If the user asks for only the text, skip the note.

If the text already reads as human, say so and change only what needs changing. Do not rewrite for the sake of rewriting.

## Limits to be honest about

This skill improves the writing. It does not promise that a detector will label the result "human". The Wikipedia guide itself warns that AI detection tools give false positives and are easily thrown off, and that people are poor at telling AI text from human text. If the user asks whether the text will "pass" a detector, tell them that no rewrite can guarantee it.

## Example

Input:

> AI won't take your job. But someone who knows how to use AI could take your job. In the past, a content creator might spend several hours researching, writing scripts and putting video footage together. Now, AI can support almost that entire process. And what's truly scary is not that AI is getting smarter. It's that many people are still using AI like a simple Q&A tool. While those who are ahead are turning AI into an assistant, a research staff member, an editor and a 24/7 content production machine.

Output:

> AI won't take your job, but someone who knows how to use it could. A content creator used to spend several hours researching, writing a script and putting the footage together. AI can now help with almost all of that. AI getting smarter is the smaller worry. The bigger one is that many people still use it only to ask a question and read the answer, while the people ahead of them use it as an assistant, a researcher, an editor and a content machine that runs around the clock.

What changed: merged the two-sentence opening hook, removed "what's truly scary is not X. It's Y", dropped "truly", and joined the sentence fragment beginning "While" to the sentence it belonged to. All claims kept; nothing added.
