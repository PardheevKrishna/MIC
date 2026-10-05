# Q&A preparation v3: Team Androids, Grand Finale

Five minutes of questions from senior Amazon scientists. Answer in 2-4 sentences: the direct answer first, then one number, then stop. If a question touches a limitation, say the limitation plainly first; the Documentation discloses all of them (Section 5, Appendix C). Use the words of the v3 slides: business, incoming record, shortlist (up to 4 per record), quick rule-based filter, CatBoost ranker, readers (MiniLM reader, E5 reader, French MiniLM reader, French E5 reader), "what changed" features, LightGBM decision model per country, confident and clear winner, second-chance searches, per-business consistency check, France rules computed by code.

Limitation questions: 6, 7, 8, 9, 10, 13. Training and selection: 3-7, 20. Key idea and features: 19, 21. Hard examples in training: 15 (only if asked).

## Numbers to have ready

| Topic | Number | Source |
|---|---|---|
| Public F0.5 | 0.991300 (portal 0.9913), upload 16 of 18 | Doc 5, B.6 |
| Test size | 1,732,544 businesses; 9,969,589 incoming records; France 259,452 / 1,434,993, no labels | Doc 2.1 |
| Comparison space | 6,724,569,566,212 same-country pairs; reduction ratio 0.99999817 | Doc 3 |
| Funnel per business | about 130 (index lookup, blocking study) → 21.71 (quick rule-based filter) → 6.75 (shortlist, derived) → 7.10 (candidate file, 12,293,019 pairs) | Doc 3 |
| Per incoming record | 1.23 candidates; 1.17 pairs reach the readers (derived); 90.59 % have exactly one candidate | Doc B.1 |
| Recall, train fold 9 | index 99.3694 % US / 98.5068 % India; shortlist 98.8748-99.3628 % / 98.0572-98.4831 % | Doc B.2 |
| Cost of the quick filter | true business kept for 99.974 % of held-out records; fold-9 F0.5 0.990117 → 0.990086 | Doc 3 |
| Folds | 10 folds per business (hash of the business id) | Doc 4 |
| Training splits | ranker folds 0-5, early stopping on 6; readers folds 0-5 (1 epoch, bf16); decision models fit on 6-7, early stopping on 8, US/India refit on 6-9, mean of 5 seeds; native-script rescue folds 0-7; address rescue 6-7 | Doc 4 |
| Held-out folds 9 / 8 | US 0.991212 / 0.991545; India 0.990046 / 0.990298 | Doc 5 |
| Decision | best candidate only if score ≥ 0.75 (US, India; tuned on fold 8) or 0.82 (France, default) and lead over the runner-up ≥ 0.4 (one fixed value for every country); consistency check 0.75 / 0.68 / 0.72 / 0.77 for 0 / 1 / 2 / 3+ other accepted records (folds 8 and 9); rescues 0.88 native script, 0.86 US address (fold 8) | Doc 4 |
| French E5 reader checks | US/India quality within tolerance: pass; agreement with held-out French pseudo-labels 0.999206: pass; generated labelled French test set 0.910516 vs 0.914964: fail; "expected gain safely positive" rule, pessimistic −0.0001806: fail | Doc B.4 |
| Kept / dropped (public) | US/India refit 0.990621 → 0.990676 kept; France refit 0.990726 dropped; second French round 0.991277 dropped; final generic inputs (upload 14) 0.990881 → 0.990934 | Doc B.6 |
| Links | 5,856,936 links; 100,087 empty rows | Doc B.1 |
| Second-chance links | native script +9,759 (India); address +192 (US), +2,945 (India) | Doc B.1 |
| France | 887,630 accepted by the decision model → 867,559 links after the France rules | Doc B.3, B.1 |
| Compute | full retrain about 11.5 h, 1× RTX 4070 Ti SUPER 16 GB, 32 GB RAM, about 70 GB disk; inference: blocking + ranker about 1.3 h, reader scoring (all except the French E5 reader) about 1.4 h | READMEs; clean-run report |
| Reproduction | from the trained models the inference stages rewrite both files byte for byte; a full retrain reproduced 99.82 % of the links, validator PASS | READMEs |

---

## 1. Why this cascade instead of embedding (vector) blocking or one end-to-end model?

Each stage does the cheapest job it can do well. The inverted index, with separate name and address lists plus letter-piece and sound lists, keeps the true business for 99.37 % (US) and 98.51 % (India) of held-out fold-9 records, and it is deterministic and easy to audit; a quick rule then removes most pairs before any model runs. Cross-encoders are the most accurate readers and the most expensive, so they see only 6.75 pairs per business, and the decision model exists because the decision is a competition between candidates, which a pair-only reader cannot see. We did not benchmark a dense vector retriever, so we cannot claim it would be worse; it is a natural extra channel for brand or "doing business as" names that share few words.

## 2. How would this scale to billions of records?

The work per record is bounded at every stage: posting-list caps and per-list budgets in the index, a fixed number of top candidates per list, and at most 4 pairs per record for the readers, so cost grows linearly with the number of records and shards run independently. Today the index is split by country; at billions we would split by country and state or postcode, which the native-script rescue already does. As a rough linear extrapolation, not a measurement: scoring our 11.69 million shortlist pairs took about 1.4 hours on one 16 GB GPU (every reader except the French E5 reader), so a billion incoming records (about 1.17 billion pairs) would be on the order of 140 GPU-hours, spread across shards.

## 3. How were the thresholds chosen, and why precision first?

F0.5 weighs precision more than recall, and a business with no true records scores 1 only if its row is empty, so one wrong link can zero a business. Every decision model uses the same rule (best candidate only, score at least the threshold and at least 0.4 ahead of the runner-up), and the operating points are tuned for the exact per-business metric, not AUC: 0.75 on fold 8, the per-business consistency check (0.75 / 0.68 / 0.72 / 0.77 for 0 / 1 / 2 / 3 or more other accepted records) on folds 8 and 9, and the rescues on fold 8 with false positives counted twice. France uses the more conservative 0.82, set without labels from the number of French links per business with a fold-8 threshold sweep on US/India.

## 4. Why build folds at the business level?

Because the business is the unit of the metric and of several features: one business owns several incoming records, the score is computed per business, and the consistency check looks at all records competing for the same business. Folds are assigned per business by a hash of its id, so all its records fall in one fold; a record-level split would put records of the same business on both sides and inflate the held-out score. The held-out protocol goes further: when we score fold 9, the fits also leave out the records that touch fold 9.

## 5. Why different folds for the ranker and readers than for the decision models, and why refit on 6-9?

The decision model's inputs include the ranker probability and the reader scores; if those models had trained on the same records, their scores would be over-confident there and the decision model would learn to over-trust them. So the ranker and readers train on folds 0-5 and the decision models on 6-9. For the US and India we fit on 6-7 with early stopping on 8 to find the number of trees, then refit on all of 6-9 with that number scaled by the row count, since no fold is left for early stopping; on the leaderboard that step moved 0.990621 to 0.990676. The same refit for France scored 0.990726 against 0.990748 without it, so France keeps the 6-7 fit.

## 6. Why five seeds?

Our LightGBM settings sample 80 % of rows every iteration and 90 % of features, so single fits vary; the mean of five seeds reduces that variance for five small model calls per pair. Honestly, its effect is not isolated: the held-out table uses 3 of the 5 seeds, and the 5-seed models entered the final file together with the French E5 reader (public 0.991114 → 0.991300), so we do not claim a separate number for the seeds.

## 7. How did you avoid overfitting the public leaderboard?

Not completely, and we say so. Our main evidence was the held-out per-business F0.5 on folds 8 and 9 from fits that left out the scored fold, and the thresholds were set on those folds, not on the leaderboard. But we made 18 uploads, the final file was chosen among them using public scores, and public feedback also shaped some France rule parameters. What limits the damage: on held-out folds, close versions of our decision models (the India recipe with 3 of its 5 seeds) score US 0.991212 / 0.991545 and India 0.990046 / 0.990298, the same range as the public 0.991300, and the late public gains were small (0.991114 → 0.991300).

## 8. Why keep the French E5 reader when it failed its written checks?

It passed two checks (US/India quality within tolerance, AUC 0.998544 vs 0.998570; 0.999206 agreement with held-out French pseudo-labels) and failed two: a generated labelled French test set (best F0.5 0.910516 vs 0.914964 for its base model) and our own rule that the expected gain must be safely positive (pessimistic estimate −0.0001806). We kept it because the file containing it had our best public score: a leaderboard-driven decision against our own written rule, and we disclose it as such. It changes French links only (−2,228 / +3,396), and the same upload also added the 5-seed models, so its own effect on the public score is not isolated.

## 9. France has no labels. How did you build and check it?

The France decision model is trained on labelled US and India rows, and the French readers are self-trained: the French MiniLM reader on very confident pseudo-labels from an earlier run (best candidate, score at least 0.97, clear lead) plus 500,000 labelled US/India pairs, and the French E5 reader on pairs where the two MiniLM readers together were at least 97 % sure of a match or at most 3 %, 474,808 French pairs plus as many labelled US/India pairs. The checks are label-free: 0.999206 agreement with held-out French pseudo-labels, a generated labelled French test set, and control groups on the French records for the rules. A diagnostic upload shows what is at stake: emptying every French row dropped the public score from 0.970125 to 0.839. But we have no measured French F0.5, and we say so.

## 10. Weren't the France rules fitted to the test set?

Yes, and we disclose it: the rules were designed after analysing the unlabelled French test records, including AI-assisted review of sampled pairs, and some parameters were chosen after seeing test results, including public-leaderboard scores. Each rule is code on model scores, raw text and the address; no test label and no human or AI judgement of any pair is an input to the pipeline. The largest action, the two word-swap vetoes together, removes 29,021 of 887,630 accepted French links. In production we would turn the rule conditions into features and learn them from a small labelled French sample.

## 11. What happens with a country you have never seen?

It takes the default route: the decision model built for France (158 features, trained on labelled US/India rows, a single fit), the multilingual MiniLM reader only, and the conservative threshold 0.82. The test set had no such country, so this route is untested on real data. For a real new market we would add a country normaliser and repeat the France recipe: self-trained readers first, then a small labelled sample to set the threshold.

## 12. How do you know blocking does not lose true matches?

We measure recall at every candidate step on labelled fold 9: 99.37 % (US) and 98.51 % (India) at the index, at least 98.87 % and 98.06 % in the shortlist. The quick filter is almost free: the true business survived for 99.974 % of held-out records and fold-9 F0.5 moved from 0.990117 to 0.990086. The cap of 4 is the largest loss, and the two second-chance searches target the remaining miss types (native-script names, same address with a different name).

## 13. Will 0.9913 hold on the private leaderboard?

There is some selection bias: we chose the final file among 18 uploads using public scores. On held-out folds, close versions of our decision models (the India recipe with 3 of its 5 seeds) score US 0.991212 / 0.991545 and India 0.990046 / 0.990298, the same range, and the late public gains were small increments (0.991114 → 0.991300), which limits how much the selection could have overfitted. The part with the most risk is France, because its rules were tuned with public feedback and it has no labels.

## 14. What are the main failure modes?

Missed matches: true businesses not found or cut by the cap of 4 (the shortlist keeps at least 98.87 % of US owners against 99.37 % at the index), native-script names without a known transliteration, brand or website names with little word overlap, records without a house number on a street with several businesses, and correct candidates rejected because a second one is too close. Wrong links: names at the business's address with one generic word replaced or added or the legal form changed, acronyms that differ from the business name by one initial at another address, and branches of a chain at nearby addresses. The "what changed" features, the margin rule, the second-chance searches and the France rules each target one of these.

## 15. How do the decision models learn to reject near-identical candidates? (answer only if asked)

Mainly through the "what changed" features and the margin rule. In addition, for the US and India only, the decision models were trained with extra hard examples: copies of unlinked training records that sit right next to a candidate business, with small name edits; how many to add was set from label-free counts on the test records, never from labels, and our documentation discloses it. The France decision model uses no such extra examples.

## 16. Why per-country models rather than one global model?

The countries differ in format (French addresses and legal words, Indian scripts) and in labels (none for France). The code path is the same; a country profile only selects the cleaning rules, readers, decision model, threshold, second-chance searches and rules. Sharing happens where it helps: the India decision model trains on US and India rows, and so does the France one.

## 17. What does it cost, and could it run online?

On one machine (RTX 4070 Ti SUPER 16 GB, 32 GB RAM) a full retrain from raw data takes about 11.5 hours and about 70 GB of disk; at inference the two longest stages are blocking plus the ranker (about 1.3 hours) and reader scoring with every reader except the French E5 reader (about 1.4 hours) for 9.97 million incoming records. Per record, the online path would be one index lookup in its shard, on average 1.17 reader pairs and one LightGBM call, short enough for streaming in principle, but we have not measured single-record latency. From the trained models, the inference stages rewrite both submitted files byte for byte.

## 18. What would you do differently?

First, keep an untouched end-to-end validation fold, including the second-chance searches, from day one. Second, add finer blocking keys such as state or postcode: smaller shards, shorter lists, and room for a larger cap than 4 where the cap loses true businesses. Third, label a small French sample to replace hand-written rules with learned features and to set the French threshold. Fourth, fold the second-chance searches into one decision model.

## 19. What is new in your approach?

Three things. Each candidate is described by how it differs from the business record (numbers changed; words added, dropped and replaced, with rarity and meaning similarity; legal-form change), not only by how similar it is. The engineering is validation-first: recall measured at every candidate step, business-level folds with separate folds per model layer, thresholds tuned on the exact per-business metric, and written checks before adopting a model. And the French readers are self-trained on unlabelled French records; for compute, a quick rule with no model cuts about 130 to 21.71 candidates per business at a held-out cost of 0.00003 F0.5.

## 20. Which decisions used held-out folds, and which used public scores?

Held-out or written evidence: the quick rule-based filter (fold-9 F0.5 0.990117 → 0.990086), the 0.75 threshold (fold 8), the consistency check (folds 8 and 9), the rescue thresholds (fold 8, false positives counted twice) and the French E5 reader checks. Public scores: the US/India refit (0.990621 → 0.990676), dropping the France refit (0.990726 vs 0.990748) and the second French self-training round (0.991277 vs 0.991300), some France rule parameters, and the final file choice. The 0.4 margin is one fixed value shared by every country's rule.

## 21. Did the "what changed" features help on the leaderboard?

We did not run a clean ablation that removes all "what changed" features. The closest evidence: when we replaced data-specific inputs with the final generic "what changed" inputs (upload 14), the public score moved from 0.990881 to 0.990934, so the generic version cost nothing and helped slightly.
