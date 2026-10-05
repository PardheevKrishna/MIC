# Presenter's explainer v3: Team Androids, Grand Finale

For Pardheev only, not for the slides. For each slide of the v3 deck: one plain sentence first, then what the slide means in simple words, why we did it this way instead of the obvious alternative, and the questions a senior scientist is likely to ask, with short honest answers. Every number is from the Documentation, the two READMEs or the verified example files. Longer answers are in `QA_PREP.md`.

Three habits for the Q&A: answer the question first, give one number, then stop. If the question touches a weakness, say the weakness plainly first. If you do not know, say "we did not measure that" rather than guess.

## Words you will hear and what they mean

| Word | Meaning in one line |
|---|---|
| Business | A record of the reference list (Source 1): 1,732,544 of them in the test files. |
| Incoming record | A record from Source 2 or Source 3 that we must link to one business, or to none: 9,969,589 of them. |
| Inverted index | A lookup table from each word (or letter piece) to the businesses that contain it, like the index at the back of a book. |
| Shortlist | Up to 4 candidates per record: the top one always, the next three only if the ranker gives them at least 0.001. The only pairs the expensive models read. |
| CatBoost ranker | A gradient-boosting model (many small decision trees added together) that orders the candidates of a record from most to least likely. |
| Cross-encoder ("reader") | A small language model that reads the record and the candidate together in one input and outputs one match score. A bi-encoder, by contrast, embeds each side separately and compares the two vectors. |
| LightGBM decision model | Another gradient-boosting model, one per country, that turns the reader scores and our features into a probability that the candidate is the right business. |
| "What changed" features | Features that describe how a candidate differs from the record: numbers changed, words added, dropped or replaced, legal form changed. |
| Confident and clear winner | Our decision rule: link the best candidate only if its probability is at least 0.75 (0.82 in France) and it leads the second-best by at least 0.4. |
| Second-chance searches | Two extra searches for records the main pass left unlinked: one for names in Indian scripts, one from the address. |
| Per-business consistency check | After the main decision, each record is re-checked with a bar that depends on how many other records were linked to the same business. |
| Folds | We split the businesses into 10 groups; a record always goes with its business. Models learn on some groups and are tested on groups they never saw. |
| F0.5 | A score that mixes precision and recall but weighs precision more. Computed per business and averaged over all businesses. |
| Precision | Of the links we made, the share that are right. |
| Recall | Of the true links, the share we found. |
| Self-training (pseudo-labels) | Using a model's own very confident guesses on unlabelled data as if they were labels, then training on them. |
| Transliteration | Writing a word in another script, for example विजय → vijay. |
| Held-out (hidden) fold | A group of businesses kept out of training and used only for testing. |
| Macro-average | Score each business separately, then take the plain average. |

---

## Slide 1 · Team Androids

**In one sentence.** Who we are, and the promise: we matched about 10 million records while checking only a tiny fraction of the possible pairs.

**What this means.** The title slide: our team, three names and photos. Use the first sentence to set up the problem, not to describe ourselves.

**Why we did it this way.** Judges remember the first 30 seconds; opening on the scale of the task gets them curious faster than an agenda slide would.

**Likely questions.**
- *Who did what in the team?* Answer with your own split of the work (it is not in our documents, so agree on one sentence with Monika and Hemachand before Wednesday).

---

## Slide 2 · The Scale Challenge

**In one sentence.** There are far too many possible pairs, so we keep only about 7 per business to look at, and France has no answers to learn from.

**What this means.** About 10 million incoming records (9,969,589) must each be linked to one of 1,732,544 businesses in the same country, or to none. One business can own several records: the test files hold 5.75 incoming records per business, and our final file links 3.38 per business on average. Comparing everything with everything inside each country would be 6,724,569,566,212 pairs; our candidate list keeps 12,293,019, which is 7.10 per business, or fewer than two pairs in every million. (The ranker itself scores more: 21.7 per business after the quick filter.) France has no labels at all.

**Why we did it this way.** The obvious alternative, scoring all pairs with a good model, is impossible at 6.7 trillion pairs; the other obvious one, a single similarity score with a threshold, cannot tell apart businesses with almost the same name. So we spend almost nothing on most pairs and real effort on very few. The size of the candidate file is also judged (candidate efficiency), so a small file is a goal in itself.

**Likely questions.**
- *Is 7.10 per business or per record?* Per business. Per incoming record it is 1.23 candidates, and 90.59 % of records have exactly one candidate.
- *Why "or none"?* Some records belong to no business in the list; linking them anyway would be a wrong link, which the metric punishes hard (next slide).

---

## Slide 3 · Meet Three Records

**In one sentence.** Three real, messy records we follow through the talk, and a score that punishes wrong links hard.

**What this means.** Three real test records, each hard in a different way: the US one has no address, the Indian one is in Devanagari with a very short address, the French one comes from a country with no labels. The metric is F0.5 per business, averaged over all businesses; a business with no true records scores 1 only if we link nothing to it, so one wrong link there turns a 1 into a 0.

**Why we did it this way.** Showing three concrete records first makes the method easier to follow, and they come back as the examples on slides 16-18. The metric explains every later choice: because wrong links cost more than missed ones, we are precision first everywhere (a high bar, a required lead over the runner-up, and rules that can veto).

**Likely questions.**
- *Were these examples cherry-picked?* Yes, on purpose: we picked them to show three different paths (a link, a second-chance link, a veto). Test labels are private, so we can show how the pipeline decides, not prove each call is right.
- *What is the F0.5 formula?* F0.5 = 1.25 · P · R / (0.25 · P + R), computed per business and averaged (macro) over all businesses.

---

## Slide 4 · Pipeline Architecture (overview)

**In one sentence.** The system is five steps in a row, from cheap and rough to expensive and careful.

**What this means.** The whole system in five zones, run separately for each country: 1 Find, 2 Filter and shortlist, 3 Read, 4 Decide, 5 Second chances and France, then the output files. Cheap steps at the start decide what gets read; expensive models at the end decide what gets linked.

**Why we did it this way.** One end-to-end model would have to score every pair or trust a single retrieval score. A cascade lets each step do the cheapest job it can do well, and we can measure what each step loses. Running per country lets France use its own readers, threshold and rules, while the code stays the same; a country profile only switches the parts.

**Likely questions.**
- *Why not one global model?* The countries differ in format (French addresses, Indian scripts) and in labels (none in France). Where sharing helps we share: the India model trains on US and India rows, and so does the France model.
- *Where do you lose the most true matches?* At the shortlist cap of 4: for the US the true business is in the index results for 99.37 % of held-out records and in the shortlist for at least 98.87 %.

---

## Slide 5 · Architecture: Find (zone 1)

**In one sentence.** Clean the text, then use a word index to pull up businesses that share rare words, letter pieces or sounds.

**What this means.** We clean the text (lower case, "&" to "and", common abbreviations unified, leading zeros dropped; in France, for example, "15B" becomes "15 bis" and "S.A.R.L." becomes "sarl"). Then a C++ inverted index finds businesses that share rare words with the record. It has five lists (channels): name words, address words, 4-letter and 3-letter pieces of the core name (legal words like inc or sarl removed), and a "sound" version of the name (vowels removed, repeated letters collapsed). Rare words count more (IDF weighting).

**Why we did it this way.** Separate lists for name and address keep candidates that only match on one side, and the letter-piece and sound lists catch typos and abbreviations. It is deterministic, fast and easy to audit. On a labelled held-out fold, the index finds the true business for 99.37 % of US records and 98.51 % of Indian records.

**Likely questions.**
- *Why not embedding (vector) search for blocking?* We did not benchmark a dense retriever, so we cannot claim it would be worse. Our index already keeps the true business for about 99 % of US records; vector search would be a natural extra channel for brand or "doing business as" names that share few words.
- *Does the cleaning use outside lists?* No external data or lookups; only small hand-written word lists for normalisation and features, as our documentation lists.

---

## Slide 6 · Architecture: Filter and shortlist (zone 2)

**In one sentence.** A free rule drops obviously weak candidates; a ranking model keeps the best four per record.

**What this means.** A quick rule-based filter with no model keeps a candidate if any of six simple conditions holds: it has the best total score, its score is at least 60 % of the best, it has the best address score, its address covers at least 95 % of the record's address, the record name is very short (3 characters or fewer) with some address overlap, or it has exactly the same core name with some address overlap. Then the CatBoost ranker orders the survivors, and the shortlist keeps the top one always and ranks 2-4 only when their probability is at least 0.001. So at most 4 per record.

**Why we did it this way.** A learned pruner would need a model pass over all of the roughly 130 candidates per business; the rule is nearly free, and its measured cost is tiny (next slide). The special conditions keep candidates that a single combined score would drop: address-only matches, very short names and acronyms, same-name businesses. The cap of 4 protects the expensive readers.

**Likely questions.**
- *Why exactly 4?* It is a cost-recall trade-off. It is our largest recall loss: on held-out fold 9 the shortlist keeps at least 98.87 % of US owners against 99.37 % for the index, at least 98.06 % against 98.51 % for India.
- *What does the ranker see?* 76 features from the index (string similarities, word and number overlaps, scores per list); it trained on a sample of 140,000 records.

---

## Slide 7 · Architecture: Read (zone 3)

**In one sentence.** Small language models (readers) read the record and each candidate together and say how well they match.

**What this means.** Cross-encoders read the raw "name | address" of the record and of the candidate together (up to 128 tokens) and give one score. We fine-tuned two multilingual models, MiniLM-L12 (117.7 million parameters) and multilingual-e5-base (278.0 million), on labelled US and India pairs for one epoch. France adds a French MiniLM reader and a French E5 reader (slide 13); their scores are blended 0.35 MiniLM + 0.15 French MiniLM + 0.5 French E5.

**Why we did it this way.** A bi-encoder compresses each side into one vector before comparing, so it can miss that exactly one word or the legal form changed; a cross-encoder compares the two texts token by token. It is the most accurate reader and the most expensive, so it sees only the shortlist. Both base models are open (Apache-2.0 and MIT) and far below the 8-billion-parameter limit; we also tried a larger XLM-R reader during development and dropped it.

**Likely questions.**
- *Why not a large language model?* The rules cap models at 8 billion parameters, and at 11.7 million pairs the cost matters; small fine-tuned cross-encoders were accurate enough, and the decision model adds what they cannot see.
- *How were they trained?* Labelled pairs from folds 0-5: every "hard" record (where the ranker was unsure or wrong) plus a random quarter of the easy ones; binary cross-entropy, one epoch, bf16.

---

## Slide 8 · Architecture: Decide (zone 4)

**In one sentence.** A decision model gives each candidate a probability; we link only if the best is high and far ahead of the rest.

**What this means.** For each shortlisted candidate we compute "what changed" features (slide 11), add the reader scores and how the candidate compares with its rivals, and a LightGBM decision model per country gives a probability. The rule: link the best candidate only if it is at least 0.75 (0.82 in France) and at least 0.4 ahead of the runner-up; otherwise link nothing.

**Why we did it this way.** A reader scores one pair at a time, but the decision is a competition between candidates, so the decision model also sees rank, margin and the best other score. A threshold without a margin would link records where two candidates are almost tied, which is exactly where wrong links happen.

**Likely questions.**
- *How were 0.75 and 0.4 chosen?* 0.75 was tuned on held-out fold 8 for the exact competition metric, not for AUC. The 0.4 lead is one fixed value used in every country: it makes us link only when one candidate clearly beats the rest. (Do not say it was never tuned: our documents do not record how 0.4 was picked.)
- *Why gradient boosting and not a neural network?* The inputs are about two hundred tabular features; boosted trees are strong, fast and easy to check on such data. We average five seeds for stability.

---

## Slide 9 · Architecture: Second chances and France (zone 5)

**In one sentence.** Records still unlinked get extra searches, France gets extra rules, then we write and validate the files.

**What this means.** Records still unlinked get two extra searches (slide 12). The per-business consistency check re-decides each record with a bar that depends on how many other records the model accepted for the same business: 0.75 with none, 0.68 with one, 0.72 with two, 0.77 with three or more (same 0.4 margin). France gets its own rules computed by code (slide 13). Finally both output files are written, every link is checked to be among the candidates, and the organisers' validator runs.

**Why we did it this way.** These steps fix specific, measured gaps of the main decision instead of lowering its bar everywhere, which would add wrong links.

**Likely questions.**
- *Why does the bar depend on other links?* The four values were tuned on held-out folds 8 and 9 for the exact metric. It moves few links (US +277 / −284, India +247 / −248).
- *Is the validator the official one?* Yes, the organisers' validation script, run on the final files; it passes.

---

## Slide 10 · Blocking: the Funnel (innovation 1, with scaling)

**In one sentence.** From trillions of pairs to about 7 per business, losing about 1 % (US) to 2 % (India) of true matches, with cost growing in a straight line.

**What this means.** From 6.7 trillion possible pairs: the index lookup gives about 130 candidates per business (measured in our blocking study of 25 September), the quick rule cuts to 21.7, the shortlist to 6.75, and the second-chance searches add a few new pairs, to 7.10. The rule costs almost nothing: the true business survived for 99.974 % of held-out records, and held-out F0.5 moved from 0.990117 to 0.990086. At most 4 pairs per record reach the readers, 1.17 on average.

**Why we did it this way.** Blocking sets the recall ceiling, so we measured recall at every step instead of trusting it. Every stage has a cap per record (posting-list caps and budgets in the index, a fixed number of top candidates per list, 4 for the readers), so the work per record is bounded and total cost grows linearly with the number of records. That is our scaling argument: split the index by state or postcode as well as by country, and run the shards in parallel.

**Likely questions.**
- *How long does it run?* A full retrain from raw data takes about 11.5 hours on one 16 GB GPU (RTX 4070 Ti SUPER). At inference, blocking plus the ranker took about 1.3 hours and reader scoring about 1.4 hours for the 9.97 million records.
- *What would a billion records cost?* A rough linear extrapolation, not a measurement: about 1.17 billion reader pairs, on the order of 140 GPU-hours, spread over shards.
- *What do you lose?* Mostly at the cap of 4: about half a percentage point of recall on fold 9 (US 99.37 % at the index, at least 98.87 % in the shortlist).

---

## Slide 11 · What Changed (innovation 2)

**In one sentence.** Instead of only 'how similar', we tell the model exactly what differs: numbers, words, legal form.

**What this means.** Besides "how similar", every candidate is described by "what changed": number features (absolute and relative difference, digit edit distance, shared prefix and suffix), word features (which words were added, dropped or replaced, how rare they are, and how close in meaning, using a pretrained MiniLM word encoder), and the legal-form relation (same, changed, added). With the reader scores and competition features, that is 217 features for the US, 211 for India and 158 for France.

**Why we did it this way.** A similarity score treats different edits as the same if they cost the same, but they mean different things: our data study showed that true records of a business tend to lose information (they drop a digit or the legal suffix, abbreviate, or transliterate). Describing the edit itself lets the model learn which edits are harmless. The features are generic (operations, not word lists), so they carry over to France.

**Likely questions.**
- *Which features matter most?* We did not prepare a feature-importance ranking for this talk. The US example shows the mechanism: the legal-form features and the readers separate the three "Enterprises" names, and the word-replacement features mark "Industries".
- *Did this help on the leaderboard?* We did not run a clean ablation that removes all "what changed" features. The closest evidence: when we replaced data-specific inputs with the final generic "what changed" inputs, the public score moved from 0.990881 to 0.990934, so the generic version cost nothing and helped slightly.

---

## Slide 12 · Second-Chance Searches (innovation 3)

**In one sentence.** When the first search never found the right business, a second search by transliterated name or by address finds it.

**What this means.** Some misses are retrieval failures: the right business never entered the shortlist. Native-script rescue (India): a word list learned from labelled training matches turns each Indian-script word into its Latin spellings; legal words are dropped, and we search businesses with that core name in the same state, keep the top 5 by a fixed address score, and a small LightGBM model links the best one if it scores at least 0.88 and leads by 0.3 (+9,759 links). Address rescue (US and India): businesses sharing one of the record's rarest house numbers and one of its rarest street words are filtered and scored by two LightGBM models (+192 US, +2,945 India links).

**Why we did it this way.** Lowering the main threshold would not help, because the right business is not among the candidates at all. A targeted second search fixes the miss type we found on training data, and it runs only on records still unlinked, so it cannot break an existing link.

**Likely questions.**
- *How did you set the second-chance thresholds?* On held-out fold 8, with each wrong link counted twice, because under F0.5 a wrong link hurts more than a missed one.
- *Does the Indian word list use test data?* The rescue's list is learned from labelled train matches. Separately, we compute an unsupervised word alignment on test records with a very confident first-stage match (probability at least 0.9); our documentation discloses this.

---

## Slide 13 · France: Zero Labels (innovation 4)

**In one sentence.** No French labels, so the readers learned from their own most confident French guesses, plus a stricter bar and rules computed by code.

**What this means.** France has no labels, so the readers taught themselves (self-training on pseudo-labels): we kept only French pairs our models were very sure about (at least 97 % sure of a match; for the non-match examples of the French E5 reader, at most 3 %), and fine-tuned two French readers on them, mixed about one to one with labelled US and Indian pairs (474,808 + 474,808 for the French E5 reader; 500,000 labelled pairs for the French MiniLM reader). The France decision model is trained on US and Indian rows and uses the stricter bar 0.82. Then rules computed by code remove risky links and add safe ones, for example: same address but one name word swapped for a common word means no link; the business's street name missing from the record's address means no link.

**Why we did it this way.** Applying the US/India models as they are ignores the French address format and legal words; and we had no French labels to train on. Self-training only on very confident pairs adds French knowledge with little noise, and mixing in labelled pairs keeps the readers from drifting. The rules encode precise, checkable conditions where the models were unsure. A diagnostic upload showed how much France matters: emptying every French row dropped the public score from 0.970125 to 0.839.

**Likely questions.**
- *Weren't the rules fitted to the test set?* Yes, partly, and we disclose it: they were designed after looking at unlabelled French test records, with AI-assisted review of sampled pairs, and some parameters were chosen after seeing public scores. Each rule is code on model scores and text; no test label and no human judgement of a pair is an input.
- *How do you know self-training helped?* Two written checks passed (US/India quality within tolerance; 0.999206 agreement with held-out French pseudo-labels), two failed (slide 19). We have no measured French F0.5.
- *Why 0.82?* Set without labels, from the number of French links per business together with a threshold sweep on US/India fold 8; more conservative than 0.75 because France has no labels.

---

## Slide 14 · Training and Selection

**In one sentence.** We tested on businesses the models never saw; honest caveat: some later choices and the final pick also used the leaderboard.

**What this means.** All businesses are split into 10 folds by a hash of their id, and each record stays with its business. The ranker and readers learn on folds 0-5. The decision models learn on folds 6-9 (fit on 6-7, early stopping on 8, then for the US and India a refit on 6-9), so they see reader scores on records the readers never trained on. To test a change, we score fold 9 with models fit on folds 6-8 (and fold 8 with models fit on 6, 7 and 9), leaving out every record touching the tested fold, using the exact competition metric. Most settings, like the 0.75 bar, stayed only if they helped on the hidden fold, and new readers also had written pass-or-fail checks. But some later choices (the US/India refit, dropping the France refit and the second French round) and the final pick among our files used public-leaderboard scores, and the French E5 reader was kept although it failed two of its checks. The slide says this in one line.

**Why we did it this way.** A random split by record would put records of the same business on both sides and inflate the score. Training the decision model on the readers' own training folds would teach it to over-trust over-confident scores. Tuning on the exact per-business metric, not AUC, matters because empty rows score 1 and the score is averaged per business.

**Likely questions.**
- *How did you avoid overfitting the public leaderboard?* Not completely, and we say so: held-out folds were the main evidence and set the thresholds, but we made 18 uploads and chose the final file among them using public scores. The late public gains were small (0.991114 to 0.991300).
- *Why five seeds?* Single LightGBM fits vary because of row and feature sampling; averaging five seeds reduces that. Its own effect is not isolated (it entered the final file together with the French E5 reader).

---

## Slide 15 · Handling Edge Cases

**In one sentence.** Businesses with no records, new countries and messy spellings each get a cautious default.

**What this means.** No matching record (the output calls these empty rows): 100,087 businesses (5.8 %) end with no linked record, because a correct empty row scores a full point and we only link on a clear win. Unseen countries: a country without its own profile takes a default route: the France-style decision model (trained without target-country labels), the MiniLM reader only, and the stricter 0.82 bar. Noisy spellings: cleaning rules, letter-piece and sound channels, and the two second-chance searches.

**Why we did it this way.** For empty rows, the precision-first rule is the defence. For unseen countries, refusing them would leave every row empty; the France recipe is the one we built without labels, so it is the natural default.

**Likely questions.**
- *Has the unseen-country route been tested?* No. The test set had only the US, India and France, so it is untested on real data; for a real new market we would add a country normaliser, self-train readers, and label a small sample to set the threshold.
- *What about chains with branches at nearby addresses?* That is one of our known false-positive types; the margin rule and the address features reduce it, but do not remove it.
- *What about singletons?* A singleton is a record or business with no match. 100,087 businesses (5.8 %) end with no link, and with 0.5875 links per incoming record about 41 % of incoming records are linked to nothing (derived from Documentation B.1). A correct empty answer is worth a full point, so we only link on a clear win.

---

## Slide 16 · Example: the US

**In one sentence.** No address: the legal form (Inc vs LLC vs PC) and one word decide, and only the Inc is linked.

**What this means.** "Legacy Marketing Enterprises (Inc)" has no address. The shortlist holds four businesses: Enterprises Inc (Windsor, CT), Enterprises LLC (Irondequoit, NY), Enterprises PC (Joshua Tree, CA), Industries Inc (Louisville, TN). The decision model scores 0.965022, 0.000007, 0.088664 and 0.000088. The best is confident (≥ 0.75) and leads by 0.876 (≥ 0.4), so the record is linked to the Windsor business only.

**Why we did it this way.** With no address, only the name can decide, and the names differ only in the legal form or one word, exactly what the "what changed" features describe.

**Likely questions.**
- *The ranker put the LLC second; why does the model put it last?* The decision model sees the legal-form change (Inc to LLC) and the reader scores; the ranker did not have those features.
- *Is the link correct?* Test labels are private, so we cannot say; it shows how the pipeline decides.

---

## Slide 17 · Example: India

**In one sentence.** The Devanagari name first finds same-name businesses in the wrong places; the second search by state and house number finds the right one.

**What this means.** The name is in Devanagari, the address "24/11, Mumbai, Maharashtra". The first search found three "Vijay Enterprises Private Limited" in Agra, Satara and Byculla; all three scored below 0.00001, so no link. The native-script rescue turned the name into "vijay enterprises", searched Maharashtra, found 10 such businesses, kept 5, and the one at the same house number 24/11 (Ghatkopar, Mumbai) scored 0.99994, next best 0.0000191. Link added.

**Why we did it this way.** The Latin index found businesses with the right name in the wrong places, and nothing tied them to this address. Searching by the transliterated name within the state, then using the house number, finds the right one.

**Likely questions.**
- *What if the house number had been missing?* We did not test that case. The house number is what separated this business from the other four in Mumbai (address token-set ratio 100 against 66.7 for each of the others), so without it the margin rule would decide whether anything is linked.
- *How many such rescues?* 9,759 links in total, from 171,713 records the rescue looked at.

---

## Slide 18 · Example: France

**In one sentence.** The model was 99.98 % sure, but one swapped word at the same address is too risky in France, so no link.

**What this means.** "TQ COMITE SARL", "2 R SUZANNE LENGLEN, NANTES". The index found 16 candidates, the quick filter kept 2. "TQ Comite SARL" in Pessac has the same name but another city: the original MiniLM reader liked it (+7.1), the French readers rejected it (blend −3.2). "TQ Parents SARL" at the same address in Nantes scored 0.999837 and was accepted, but the France rules vetoed it: same house number and street, and the business name lost the word "parents" while the record has the common word "comite" instead. Result: no link.

**Why we did it this way.** Under F0.5 a wrong link costs more than a missed one, so when the only difference is one swapped word at the same address, abstaining is the safer bet.

**Likely questions.**
- *Couldn't TQ Parents be the right business?* It could; the labels are private. The rule is a precision choice: abstaining loses a little recall if it was right, while a wrong link would cost precision on that business.
- *Why not link the exact-name TQ Comite SARL in Pessac?* It is in another city, and both French readers rejected it; the decision model gave it 0.000098, far below the bar.
- *How often do the France rules act?* The two word-swap vetoes together removed 29,021 of 887,630 accepted French links; after all rules France has 867,559 links.

---

## Slide 19 · Results and Learnings

**In one sentence.** 0.9913 public score, hidden-fold scores in the same range, and our honest limits.

**What this means.** The final file scores 0.991300 on the public leaderboard (upload 16 of 18), up from 0.968212 for the first scored baseline. On held-out folds, close versions of our decision models (the India recipe with 3 of its 5 seeds, scored on US and on Indian records) score 0.991212 / 0.991545 (US, fold 9 / fold 8) and 0.990046 / 0.990298 (India), the same range; they exclude the second-chance searches and the France rules. The slide rounds them to 4 decimals. The limits box says plainly what is not proven.

**Why we did it this way.** Judges asked for honesty about limitations; stating them ourselves is more credible than having them found in Q&A, and our documentation already discloses all of them.

**Likely questions.**
- *Will 0.9913 hold on the private leaderboard?* There is some selection bias, because we picked the final file among 18 uploads using public scores. The held-out scores sit in the same range and the late gains were small, which limits the risk; France carries the most risk.
- *What did not work?* A larger XLM-R reader (dropped in development), refitting the France decision model on more folds (public 0.990726 vs 0.990748 without it), and a second French self-training round (0.991277 vs 0.991300).
- *Why keep the French E5 reader if it failed checks?* It failed a generated French test set (best F0.5 0.910516 vs 0.914964 for its base model) and our own "expected gain must be safely positive" rule; we kept it because the file with it had our best public score. That is a leaderboard-driven decision, and we disclose it.

---

## Slide 20 · Thank You / Questions

**In one sentence.** A one-sentence summary, then questions.

**What this means.** The one-sentence summary: about 7 candidates per business, one careful decision per record, and no link when unsure.

**Why we did it this way.** Ending on two numbers (6.7 trillion possible pairs cut to 12.3 million candidates, and the public 0.9913) leaves the judges with the efficiency story and the result.

**Likely questions.** See `QA_PREP.md`: open the backup slides B1-B7 if a question needs the full architecture, the funnel numbers, the held-out table, the written checks or the detailed examples.
