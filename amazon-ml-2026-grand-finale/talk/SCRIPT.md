# Spoken script v3: Team Androids, Grand Finale

Full spoken script for ONE presenter (Pardheev), in natural paragraphs, one per slide, to be pasted word for word into the speaker notes of the v3 deck (`SLIDES_v3.md`). This is the LONG version: about 1,743 spoken words, about 12:27 at 140 words per minute. The talk is stopped at 10:00, so it has to be cut by about 2:27 (suggested cuts are at the end).

## How to read it

- Read it like you are explaining the project to a smart friend: calm, a little slower on the numbers, and with a short pause at every [pause].
- Words in [square brackets] are stage directions, not spoken: [click] = advance the slide, [point to ...] = point at that part of the slide, [pause] = one breath.
- Numbers are written in the short, rounded form you say on stage; the main slides show the same rounded values (exact ones are on the backup slides). Say 0.75 as "point seven five", 0.4 as "point four", 0.82 as "point eight two", 0.9913 as "point nine nine one three", 0.965 as "point nine six five", 0.089 as "point zero eight nine", 0.99994 as "point nine nine nine nine four", 0.9998 as "point nine nine nine eight", 24/11 as "twenty-four slash eleven", F0.5 as "F point five".
- Every section follows the same story: what was hard, the idea, and the number that proves it.

## Timing

| # | Slide | Spoken words | Seconds (140 wpm) | Running total |
|---|---|---|---|---|
| 1 | Team **Androids** | 41 | 18 | 0:18 |
| 2 | The Scale **Challenge** | 105 | 45 | 1:03 |
| 3 | Meet Three **Records** | 121 | 52 | 1:54 |
| 4 | Pipeline **Architecture** | 49 | 21 | 2:15 |
| 5 | Architecture: **Find** | 72 | 31 | 2:46 |
| 6 | Architecture: **Filter** | 62 | 27 | 3:13 |
| 7 | Architecture: **Read** | 60 | 26 | 3:39 |
| 8 | Architecture: **Decide** | 62 | 27 | 4:05 |
| 9 | Architecture: Second **Chances** | 40 | 17 | 4:22 |
| 10 | Blocking: the **Funnel** | 136 | 58 | 5:21 |
| 11 | What **Changed** | 105 | 45 | 6:06 |
| 12 | Second-Chance **Searches** | 104 | 45 | 6:50 |
| 13 | France: Zero **Labels** | 106 | 45 | 7:36 |
| 14 | Training and **Selection** | 113 | 48 | 8:24 |
| 15 | Handling Edge **Cases** | 98 | 42 | 9:06 |
| 16 | Example: the **US** | 87 | 37 | 9:43 |
| 17 | Example: **India** | 94 | 40 | 10:24 |
| 18 | Example: **France** | 113 | 48 | 11:12 |
| 19 | Results and **Learnings** | 129 | 55 | 12:07 |
| 20 | Thank You / **Questions** | 46 | 20 | 12:27 |
| | **Total** | **1,743** | **747** | **12:27** |

Seconds = spoken words (stage directions excluded) at 140 words per minute. Numbers take a little longer to say than ordinary words, so time yourself once with a timer.

---

## Slide 1 · Team **Androids**

[Smile and look at the camera.] Hello, everyone. I'm Pardheev, and together with Monika and Hemachand, we are team Androids. Today I will show you how we matched almost ten million messy records against 1.7 million businesses while checking only a tiny fraction of the possible pairs. [click]

*41 words, about 18 s*

---

## Slide 2 · The Scale **Challenge**

Almost 10 million incoming records arrive from three countries: the US, India and France. Each one belongs to one of 1.7 million known businesses, or to none, and we have to decide which. Comparing every record with every business in its own country would mean 6.7 trillion pairs, far too many for any model. [point to the thin orange bar] Our final candidate list keeps only about 7 per business, fewer than two pairs in every million. And there is a twist: France came with zero training labels. [pause] So this talk answers two questions: how do we find those seven, and how do we decide which one, if any, is right? [click]

*105 words, about 45 s*

---

## Slide 3 · Meet Three **Records**

Let me introduce three real records from the test set; they will come back at the end. [point to each card in turn] The American one, Legacy Marketing Enterprises, has no address, so the name is our only clue. The Indian one is written in Devanagari, and its address is just a house number, 24/11, and Mumbai. The French one, TQ Comite SARL in Nantes, comes from the country with no labels. [pause] And the scoring is strict. The metric, F0.5, weighs precision above recall and is averaged per business, and a business with no true records scores a full point only if we link nothing to it. So one wrong link can turn a perfect one into a zero, which is why we built everything precision first. [click]

*121 words, about 52 s*

---

## Slide 4 · Pipeline **Architecture**

Here is the whole pipeline. It runs separately per country and moves from cheap to expensive in five zones: find, filter and shortlist, read, decide, and second chances. [sweep your hand from left to right] I will light up one zone at a time, and then zoom in on the four ideas that made the difference. [click]

*49 words, about 21 s*

---

## Slide 5 · Architecture: **Find**

Zone one is Find. We first clean the text, so that in France, for example, 15B becomes 15 bis. Then a fast inverted index written in C++, which works like the index at the back of a book, looks up businesses that share rare words with the record. It keeps separate lists for name words, address words, short letter pieces and how a name sounds, so a typo still finds its business. [click]

*72 words, about 31 s*

---

## Slide 6 · Architecture: **Filter**

Zone two is Filter and shortlist. A quick rule-based filter, with no model, drops candidates that are clearly weaker than the best one, but keeps special cases such as an almost perfect address match. Then a CatBoost ranker, a gradient-boosting model, orders the rest, and we keep at most four candidates per record. That shortlist is all the expensive models ever see. [click]

*62 words, about 27 s*

---

## Slide 7 · Architecture: **Read**

Zone three is Read. Here we use cross-encoders, small language models that read the record and the candidate together, side by side, and give one match score; we call them readers. Because they see both texts at once, they notice details like Inc becoming LLC. We fine-tuned two multilingual readers, MiniLM and E5, and France gets two extra French ones. [click]

*60 words, about 26 s*

---

## Slide 8 · Architecture: **Decide**

Zone four is Decide. For every candidate we describe exactly what changed between the two records, and a LightGBM decision model, one per country, turns that into a probability. Then the rule is simple: link the best candidate only if it is confident, at least 0.75, and a clear winner, at least 0.4 ahead of the runner-up. Otherwise the record stays unlinked. [click]

*62 words, about 27 s*

---

## Slide 9 · Architecture: Second **Chances**

Zone five is second chances and France. Two extra searches revisit records that are still unlinked, a per-business consistency check adjusts the bar slightly when a business already has other links, and France gets its own rules, computed by code. [click]

*40 words, about 17 s*

---

## Slide 10 · Blocking: the **Funnel**

Now let's zoom in on the four ideas, starting with the funnel. [point down the bars] The index lookup gives about 130 candidates per business. The quick rule cuts that to about 22 at almost no cost: the true business survived for 99.97 percent of held-out records. The shortlist brings us to 6.75, and the second-chance searches add a few more, so we end at 7.10 per business. The price is small: the shortlist still holds the true business for roughly 99 percent of US and 98 percent of Indian records. [pause] This is also our answer for billions. At most four pairs per record reach the expensive models, so the cost grows in a straight line. Today we split the index by country; at that scale we would split it by state or postcode and run the pieces in parallel. [click]

*136 words, about 58 s*

---

## Slide 11 · What **Changed**

The second idea is the heart of our model. Most matchers describe a candidate only by how similar it is, but we also describe exactly what changed. [point to the chips] Did a number change, and by how much? Were words added, dropped or replaced, and how rare are they? Did the legal form change, say from Inc to LLC? We learned this from the training data: true records of a business tend to lose information, by dropping a digit or a suffix, abbreviating, or transliterating. These features, the readers' scores and each candidate's lead over its rivals go into the decision model, about two hundred features per country. [click]

*105 words, about 45 s*

---

## Slide 12 · Second-Chance **Searches**

The third idea is a second chance. Some records are missed not because the model is unsure, but because the first search never found the right business. [point to the top lane] An Indian name in Devanagari can send the first search to the wrong places. So a word list learned from training matches turns it into Latin spelling, and we search again within the same state, which added 9,759 links. [point to the bottom lane] The second search starts from the address, from a rare house number and a rare street word, and it added 3,137 links in the US and India. Both only touch still-unlinked records, so they never undo a link. [click]

*104 words, about 45 s*

---

## Slide 13 · France: Zero **Labels**

The fourth idea is France, where we had no labels at all. [point to the loop] So we let the readers teach themselves. We kept only the French pairs our models were most sure about, matches at least 97 percent sure and clear non-matches, and fine-tuned two French readers on them, mixed one to one with labelled US and Indian pairs so they would not forget what they knew. France also uses a stricter bar: 0.82 instead of 0.75. [point to the rules box] On top, we added careful rules computed by code. For example, when the address is the same but one name word was swapped for a common word, we do not link. [click]

*106 words, about 45 s*

---

## Slide 14 · Training and **Selection**

How did we avoid fooling ourselves? [point to the ten boxes] We split the businesses into ten folds, and every record travels with its business. The ranker and readers learn on folds zero to five, and the decision models on folds six to nine, so they never learn from scores of a model trained on the same records. To test a change, we hid fold eight or nine, retrained without it, and measured the exact competition metric there. Most settings, like the 0.75 bar, stayed only if they helped on the hidden fold. New readers also faced written pass-or-fail checks. [pause] One honest note: some later choices, and the final pick among our files, also used public leaderboard scores. [click]

*113 words, about 48 s*

---

## Slide 15 · Handling Edge **Cases**

There are three edge cases to cover. First, a business with no true records scores a full point only if we link nothing to it, so we link only on a clear win, and 100,087 businesses end up with no link at all. Second, an unseen country takes a default route built like France, with a model trained without that country's labels and the stricter bar. Honestly, this route is untested, because the test set had no such country. Third, noisy spellings are handled by the cleaning rules, the letter-piece and sound lists and the two second-chance searches. [click]

*98 words, about 42 s*

---

## Slide 16 · Example: the **US**

Now let's go back to our three records. The American one has no address. [point along the steps] The shortlist holds four businesses with nearly the same name: Enterprises Inc, Enterprises LLC, Enterprises PC, and Industries Inc. The features see that LLC and PC changed the legal form, Industries replaced a word, and Inc changed nothing. The decision model gives Inc 0.965 and the best other only 0.089, so it is confident and a clear winner, and we link the record to the Inc in Windsor, Connecticut, and to nothing else. [click]

*87 words, about 37 s*

---

## Slide 17 · Example: **India**

The Indian record is in Devanagari. [point along the steps] The shortlist holds three businesses called Vijay Enterprises Private Limited, in Agra, Satara and Byculla, but nothing ties them to this address, and all three score below one in a hundred thousand. So it stays unlinked, and the second chance begins. The word list turns the name into vijay enterprises, the address gives Maharashtra, and the new search finds ten such businesses there and keeps five. Only one has the same house number, 24/11, in Ghatkopar, and the second-chance model gives it 0.99994. [pause] So the link is added. [click]

*94 words, about 40 s*

---

## Slide 18 · Example: **France**

Finally, here is the French record, TQ Comite SARL in Nantes. [point along the steps] The search finds 16 candidates, and the quick filter keeps two. One has exactly the same name but sits in Pessac, another city; our original reader liked it, but the French readers rejected it. The other, TQ Parents SARL, is at exactly the same address, and the decision model gives it 0.9998. But at the same address, the common word comite has replaced parents, and our French rule vetoes exactly that. [pause] So we link nothing, because under F0.5 abstaining is cheaper than a wrong link. Test labels are private, so these examples show how we decide, not whether each call was right. [click]

*113 words, about 48 s*

---

## Slide 19 · Results and **Learnings**

Where did this get us? Our final file scores 0.9913 on the public leaderboard, up from 0.968 for our first baseline, and on hidden folds close versions of our US and Indian decision models score between 0.990 and 0.992. [pause] What worked best was studying how true records differ from their business before building features. [point to the limits box] Now let me state our limits plainly. We never had a fully untouched test of the final file, because the hidden folds also helped set our thresholds. France relied on self-training and on rules designed after looking at unlabelled test records, and we kept the French E5 reader although it failed two of its four written checks. Next time, we would keep an untouched validation fold from day one and label a small French sample. [click]

*129 words, about 55 s*

---

## Slide 20 · Thank You / **Questions**

To sum up, we keep about seven candidates per business, make one careful decision per record, and link nothing whenever we are not sure. That is how 6.7 trillion possible pairs became a score of 0.9913. Thank you, and we would love to hear your questions.

*46 words, about 20 s*

---

## Where to cut to reach 10 minutes

The talk is stopped at 10:00. This long version runs about 12:27, on purpose, so you can choose what to drop. The cuts below lose no required topic (each dropped sentence is shown on the slide, said again later, or covered in `QA_PREP.md`). All 12 together remove 372 words: 1,371 words, about 9:48 at 140 words per minute, or about 9:08 at 150. Rehearse with a timer and stop cutting once you finish under 9:40.

1. Architecture zones, slides 5-8 (the lit-up diagram and slides 10-13 carry the rest) (82 words). Drop slide 5: "We first clean the text, so that in France, for example, 15B becomes 15 bis."; slide 5: "It keeps separate lists for name words, address words, short letter pieces and how a name sounds, so a typo still finds its business."; slide 6: "That shortlist is all the expensive models ever see."; slide 7: "Because they see both texts at once, they notice details like Inc becoming LLC."; slide 7: "We fine-tuned two multilingual readers, MiniLM and E5, and France gets two extra French ones."; slide 8: "Otherwise the record stays unlinked."
2. Slide 10: the shortlist price (backup B2 and Q&A) and the sharding detail (Q&A 2) (47 words). Drop slide 10: "The price is small: the shortlist still holds the true business for roughly 99 percent of US and 98 percent of Indian records."; slide 10: "Today we split the index by country; at that scale we would split it by state or postcode and run the pieces in parallel."
3. Slide 1: the preview sentence (27 words). Drop slide 1: "Today I will show you how we matched almost ten million messy records against 1.7 million businesses while checking only a tiny fraction of the possible pairs."
4. Slide 2: the two-questions sentence (23 words). Drop slide 2: "So this talk answers two questions: how do we find those seven, and how do we decide which one, if any, is right?"
5. Slide 15: noisy spellings (slides 5 and 12 already cover it) (19 words). Drop slide 15: "Third, noisy spellings are handled by the cleaning rules, the letter-piece and sound lists and the two second-chance searches."
6. Slide 4: the build-up preview (21 words). Drop slide 4: "I will light up one zone at a time, and then zoom in on the four ideas that made the difference."
7. Slide 16: the features sentence (the strip on the slide shows it) (19 words). Drop slide 16: "The features see that LLC and PC changed the legal form, Industries replaced a word, and Inc changed nothing."
8. Slides 12-14: sentences the slides already show (28 words). Drop slide 12: "Both only touch still-unlinked records, so they never undo a link."; slide 13: "France also uses a stricter bar: 0.82 instead of 0.75."; slide 14: "New readers also faced written pass-or-fail checks."
9. Slide 9: the middle sentence (the lit-up zone and slides 12-13 show it) (33 words). Drop slide 9: "Two extra searches revisit records that are still unlinked, a per-business consistency check adjusts the bar slightly when a business already has other links, and France gets its own rules, computed by code."
10. Slides 16, 17 and 19: sentences the slides already show (30 words). Drop slide 16: "The American one has no address."; slide 17: "So it stays unlinked, and the second chance begins."; slide 19: "What worked best was studying how true records differ from their business before building features."
11. Slide 18: the private-labels sentence (keep it for Q&A) (17 words). Drop slide 18: "Test labels are private, so these examples show how we decide, not whether each call was right."
12. Slide 11: where the idea came from (keep it for Q&A) (26 words). Drop slide 11: "We learned this from the training data: true records of a business tend to lose information, by dropping a digit or a suffix, abbreviating, or transliterating."

If you still need time after these, speak a little faster on the architecture slides (4-9): the lit-up zones carry them.

## Number check (every number spoken, and its source)

| Spoken | Exact value | Source |
|---|---|---|
| almost 10 million records; 1.7 million businesses | 9,969,589 incoming records; 1,732,544 businesses | Documentation 2.1 |
| France: zero training labels | France appears only in the test files | Documentation 2.1 |
| 6.7 trillion pairs | 6,724,569,566,212 within-country pairs | Documentation 3 |
| about 7 per business; 7.10 | 12,293,019 pairs, 7.0954 per business | Documentation 3 |
| fewer than two pairs in every million | 1 − 0.99999817 = 1.83 per million (derived) | Documentation 3 |
| a perfect one into a zero; precision first | F0.5 per business, macro-averaged; an empty row scores 1 only for an empty prediction | Documentation 2.1 |
| 15B becomes 15 bis | French normalisation rule | Documentation 2.2 |
| at most four candidates per record | top-ranked pair always, ranks 2-4 if probability ≥ 0.001 | Documentation 3 |
| MiniLM and E5; two extra French readers | MiniLM-L12 and multilingual-e5-base readers; French MiniLM and French E5 readers | Documentation 4, B.5 |
| at least 0.75; at least 0.4 ahead | threshold 0.75 (US, India), margin 0.4 | Documentation 4 |
| about 130 → about 22 → 6.75 → 7.10 | 130.1 (blocking study, 25 Sep) → 21.7107 → 6.7458 (derived: 11,687,317 / 1,732,544) → 7.0954 | Documentation 3 |
| 99.97 percent | owner kept by the filter for 99.974 % of fold-8/9 records | Documentation 3 |
| roughly 99 percent of US and 98 percent of Indian records | shortlist recall on train fold 9: US 98.8748 % to 99.3628 %; India 98.0572 % to 98.4831 % | Documentation B.2 |
| about two hundred features per country | 158 (France), 211 (India), 217 (US) | Documentation 4 |
| true records lose information: drop a digit or a suffix, abbreviate, transliterate | "True copies lose information" | Documentation 2.1 |
| 9,759 links | +9,759 native-script rescue links | Documentation B.1 |
| 3,137 links in the US and India | +192 (US) + 2,945 (India) address-rescue links | Documentation B.1 |
| matches at least 97 percent sure, and clear non-matches | positive confidence ≥ 0.97, negative ≤ 0.03 (French E5 reader); score ≥ 0.97 for positives (French MiniLM reader) | Documentation 4 |
| one to one with labelled US and Indian pairs | 474,808 French + 474,808 US/India pairs (French E5 reader); pseudo-labels + 500,000 US/India pairs (French MiniLM reader) | Documentation 4 |
| 0.82 instead of 0.75 | France and default threshold 0.82 | Documentation 4 |
| rules designed after looking at unlabelled French test records | Appendix C, item 2 | Documentation C |
| ten folds; zero to five; six to nine; hid fold eight or nine | folds per business; ranker and readers folds 0-5; decision models 6-9; held-out protocol | Documentation 4, 5 |
| some later choices, and the final pick, used public leaderboard scores | US/India refit, France refit dropped, second French round dropped, final file chosen among 18 uploads | Documentation B.6, Appendix C item 4 |
| most settings, like the 0.75 bar, stayed only if they helped on the hidden fold | 0.75 chosen on fold 8; consistency check on folds 8 and 9; rescue thresholds on fold 8 | Documentation 4 |
| 100,087 businesses end up with no link | 100,087 empty rows in the output file | Documentation B.1 |
| 0.965; 0.089 | 0.965022; 0.088664 (margin 0.876) | Documentation D.1 |
| below one in a hundred thousand; ten, keeps five; 24/11; 0.99994 | 1.9e-6 / 6.9e-6 / 7.2e-6; 10 in Maharashtra, top 5; house number 24/11; 0.99994 | Documentation D.2 |
| 16 candidates, keeps two; 0.9998 | matcher 16, kept 2; 0.999837 | Documentation D.3 |
| 0.9913; 0.968 | 0.991300 (upload 16); 0.968212 (upload 2, first scored baseline) | Documentation 5, B.6 |
| close versions of our decision models score between 0.990 and 0.992 on hidden folds | India recipe with 3 of its 5 seeds, scored on US rows 0.991212 / 0.991545 and on Indian rows 0.990046 / 0.990298 (fold 9 / fold 8); the final US model trains on US rows only and adds 6 features | Documentation 5 |
| failed two of its four written checks | two passes, two fails | Documentation B.4 |
