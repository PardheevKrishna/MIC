# Vesta Document Data Reconstruction Guide (transcribed from README screenshots)

Source: screenshots of `17_vesta_power_bi/README.md` (2026 export analysis). All values below are the
rules the synthetic data must satisfy. Generate new IDs, users, names — never copy real ones.

## Relationship model

```
vesta_Document.document_id (PK)  1 -> 0..N  vesta_DocumentVersion
vesta_DocumentVersion.document_version_id (PK)
vesta_DocumentVersion.document_id (FK -> vesta_Document.document_id)
vesta_DocumentVersion.based_on_version_id (optional self-reference)
```

Integrity rules
- Every version's document_id exists in vesta_Document (no orphans).
- document_id unique in parent (131 values). document_version_id unique in child (145 values).
- Both identity series have gaps (surrogate keys, not row numbers).
- 123 documents have one version, 3 have two, 4 have three, 1 has four (8 documents = 6.1% have history).
- based_on_version_id, when populated, points to an EARLIER version of the SAME document. Optional.

## vesta_Document (131 rows)

One row = one document associated with a model and a document type. document_id ranges 5..228;
118 distinct model_id values (a model can have multiple documents, usually different types).

| Column | Type | Nulls | Rule |
|---|---|---|---|
| document_id | int | 0 | PK, unique, gapped ascending; never equal to model_id |
| document_type_code | string | 0 | MDD (118, 90.1%) / MVD (13, 9.9%) |
| business_document_id | int | 0 | MDD -> 100110, MVD -> 100101 (fixed mapping) |
| created_by | string | 0 | 80 distinct; short account-like id: leading letter + 5-6 digits |
| created_on | datetime | 0 | equals first version's created_on |
| last_modified_by | string | 131 | entirely blank |
| last_modified_on | datetime | 131 | entirely blank |
| model_id | int | 0 | 118 distinct across 131 docs; separate integer domain; reuse for MDD/MVD pairs |
| is_active | bool | 131 | entirely blank |
| description | string | 0 | 47 distinct; pipe-delimited template label beginning with the doc type |

Parent generation rules
1. Unique document_id from ascending, non-contiguous integer sequence.
2. document_type_code: MDD with p=0.901 else MVD.
3. business_document_id from type mapping.
4. model_id: reuse a model for ~10% of rows, preferably an MDD/MVD pair; else new model.
5. created_by and created_on generated together with the initial version.
6. description template taxonomy, e.g. `MDD created from template: MDD | Quantitative Model | Moderate risk | General`.
7. Leave last_modified_* and is_active blank.

Sample rows
```
document_id,document_type_code,business_document_id,created_by,created_on,last_modified_by,last_modified_on,model_id,is_active,description
301,MDD,100110,u900101,2026-07-08 09:14:22,,,31001,,MDD created from template: MDD | Quantitative Model | Moderate risk | General
319,MVD,100101,u900101,2026-07-15 13:42:05,,,31001,,MVD created from template: MVD | Quantitative Model | Moderate risk | General
326,MDD,100110,k900218,2026-08-03 16:08:39,,,31027,,MDD created from template: MDD | Quantitative Model | High risk | Vendor
```

## vesta_DocumentVersion (145 rows)

One row = a point in the document lifecycle. document_version_id ranges 1004..1257.
All created_on in 2026, 2026-06-24 14:13:21 .. 2026-10-28 10:31:30.
Nonblank last_modified_on span 2026-07-10 .. 2026-10-28 and are never earlier than created_on.

| Column | Type | Nulls | Rule |
|---|---|---|---|
| document_version_id | int | 0 | PK, unique, gapped ascending |
| document_id | int | 0 | FK to parent |
| template_id | int | 0 | 87 distinct; reuse templates across documents |
| template_version | int | 0 | 35 distinct low positive integers |
| version_no | string | 0 | six values: 2026.1, 2026.1.1, 2026.1.2, 2026.1.3, 2026.2.1, 2026.2.2 (dotted label, not decimal) |
| state | string | 0 | Draft, Lock, Retired, Publish, Commit to Vesta |
| based_on_version_id | int | 133 | present on 12 rows; earlier sibling version |
| created_by | string | 0 | 84 distinct; commonly = parent creator for initial version |
| created_on | datetime | 0 | initial version = parent created_on |
| approved_by | string | 140 | populated on 3 published versions (nonblank count = 3) |
| approved_on | datetime | 140 | paired with approved_by |
| last_modified_by | string | 1 | usually populated; login-like; may differ from created_by |
| last_modified_on | datetime | 35 | optional; >= created_on; paired with last_modified_by |
| section_id | int/string | 145 | entirely blank |
| template_driver_mapping | string | 145 | entirely blank |
| mar_id | string | 113 | 27 distinct; numeric strings and ids like CM500120 |
| var_id | string | 105 | 23 distinct; mostly numeric, optionally prefixed with P |
| committed_to_vesta_by | string | 100 | user id |
| committed_to_vesta_on | datetime | 100 | always paired with committed_to_vesta_by |
| var_approverId | string | 143 | only 2 rows |
| var_ReferenceId | string | 142 | only 3 rows |
| published_date | datetime | 140 | 5 nonblank rows |
| published_by | string | 140 | 4 nonblank rows; paired with published_date |
| mar_referenceNumber | string | 103 | 37 distinct; commonly MA + digits |
| created_by_name | string | 48 | display name; blank in ~1/3 rows |
| last_modified_by_name | string | 26 | display name; blank in ~1/5 rows |
| ack_template_version | int | 137 | very sparse small positive integer |

Note: the README's null counts (140) for published_by/approved_by disagree with its text (4 / 3 nonblank);
the most consistent reading is that the export held some empty strings. The generator follows the text
(nonblank counts): published_date 5, published_by 4, approved_by/approved_on 3.

## Lifecycle distribution (145 rows)

| State | Count | Share | MDD | MVD |
|---|---|---|---|---|
| Draft | 94 | 64.8% | 83 | 11 |
| Lock | 31 | 21.4% | 31 | 0 |
| Retired | 11 | 7.6% | 6 | 5 |
| Publish | 6 | 4.1% | 4 | 2 |
| Commit to Vesta | 3 | 2.1% | 3 | 0 |

Lock and Commit to Vesta only on MDD. MVD only Draft, Retired, Publish.

## State-dependent workflow fields (nonblank counts)

| State | based_on_version_id | mar_id | var_id | committed_to_vesta_on | published_date / approval |
|---|---|---|---|---|---|
| Draft (94) | 5 | 0 | 9 | 0 | 0 |
| Lock (31) | 0 | 20 | 15 | 30 | 0 |
| Retired (11) | 4 | 6 | 10 | 6 | 0 |
| Publish (6) | 3 | 4 | 4 | 6 | 5 |
| Commit to Vesta (3) | 0 | 2 | 2 | 3 | 0 |

Practical rules
- Draft rows minimally populated: workflow ids uncommon, publication fields blank.
- Lock rows usually have a Vesta commit timestamp and often MAR/VAR fields.
- Publish rows always have a Vesta commit timestamp; most have publication and approval data.
- Retired rows may retain workflow references from the previous lifecycle stage.
- Never populate approved_by/approved_on without nonblank published_date.

## Version-number and chain rules

Labels all begin 2026. Initial MDD most often 2026.1.1; initial MVD can use 2026.1. Later rows increment
final component (2026.1.2, 2026.1.3) or move to new middle component (2026.2.1).
For every multi-version document:
1. Sort rows by created_on.
2. Each successive row has a later timestamp and a unique document_version_id.
3. Later version label; never reuse a version number within a document.
4. Populate based_on_version_id for some, but not all, successive records.
5. Credible paths: Retired -> Retired -> Publish, Retired -> Draft, Publish -> Draft.
6. One observed four-version chain is Retired -> Publish -> Retired -> Draft.

## Validation checklist
- Parent document_id unique; child document_version_id unique.
- Every child document_id present in parent.
- Type/business id: MDD -> 100110, MVD -> 100101.
- All zero-null required fields populated.
- Parent last_modified_by, last_modified_on, is_active blank.
- Child section_id, template_driver_mapping blank.
- last_modified_on >= created_on whenever populated.
- based_on_version_id references an earlier version of the same document.
- Version labels are strings YYYY.N or YYYY.N.N (not numeric).
- Lock and Commit to Vesta only on MDD.
- Approval data only on published records; paired fields populated together.
- All values synthetic.
