# MCQ Exam Builder

**Turn existing question banks into checked exam versions, with answers that stay in sync.**

[Open the live application](https://mcq-exam-builder.streamlit.app/) · [Supported formats](docs/SOURCE_FORMATS.md) · [Product vision](docs/PRODUCT_VISION.md) · [Architecture](docs/ARCHITECTURE.md)

MCQ Exam Builder helps lecturers prepare multiple-choice examinations from the Word documents, spreadsheets and question files they already use. Import the material, review flagged questions, then create shuffled versions with matching answer keys and editable Word papers in one download.

Built from the practical demands of university teaching and programme coordination, the project focuses on the work between writing a question and delivering an examination: interpreting inconsistent files, checking answer mappings and keeping every version aligned with its key.

## Why this exists

A question bank is rarely a clean dataset. One lecturer uses bold text for correct answers; another appends an answer key. Word numbering, tables and embedded figures add their own conventions. Preparing several versions means checking that a missing choice, a misplaced image or an outdated answer letter has not slipped into the final paper.

MCQ Exam Builder brings those checks into one review workflow. It reports detected structural problems and ambiguous answer evidence before generation. The lecturer remains responsible for academic correctness and final approval; structural validation cannot establish whether an answer is scientifically correct or a question is well written.

## What the product does

| Problem | Product response |
| --- | --- |
| Question banks use inconsistent document structures | Normalises 27 supported source-file extensions into one review model |
| Correct answers may be expressed through Word formatting or an answer key | Records answer evidence, flags ambiguity and blocks unresolved answers |
| Shuffling can separate an answer from its question | Preserves answer identity through deterministic question and option shuffling |
| Word images can be cropped, duplicated or positioned ambiguously | Preserves supported crops and blocks image states that cannot be associated safely |
| A valid-looking export can hide lost content | Rejects detected excess options, malformed Office containers and unsupported structures with preparation guidance |
| Lecturers need several final artefacts | Produces Word examinations, answer keys, mappings, validation reports and supported platform exports in one ZIP |

## Workflow

1. **Upload** one or more supported question files.
2. **Review** normalised questions, answers, images and validation findings.
3. **Configure** versions, shuffling, formatting and export choices.
4. **Download** the approved examination package and its supporting reports.

The optional preview sits inside the configuration stage. Skipping it does not bypass validation or a required review confirmation. Questions must have one correct answer and two to six options. PDF support is limited to suitable selectable-text documents; there is no OCR. See the [format contracts and limits](docs/SOURCE_FORMATS.md) before uploading complex layouts.

## Decisions that matter

**Track the answer, not its displayed letter.** If source option B moves to D, the key must move with it. Generation retains the original option identity, even when two options contain identical text. Recognised order-dependent wording, such as “All of the above”, keeps its answer choices in their original order; lecturers should still check other positional references.

**Treat ambiguity as a review task.** Formatting can suggest an answer without proving it. Imported questions carry source locations and answer evidence so a lecturer can inspect the interpretation. Blocking findings must be resolved before a package can be created.

**Make an arrangement reproducible.** The same reviewed questions, settings and seed produce the same question and option order. Changing the input or configuration invalidates an older preview and package.

**Keep storage out of the default workflow.** The application has no exam database and does not send question content to generative-AI services. Hosted use processes files on the application server, with temporary files for conversion and image export. This is not browser-only processing or a guarantee of immediate deletion. The [privacy notice](PRIVACY.md) explains retention and hosting boundaries.

## Engineering evidence

The public snapshot is validated with synthetic fixtures rather than private examination material:

- automated synthetic regression and compatibility checks run in CI on Python 3.12;
- all supported source extensions are exercised with generated synthetic fixtures;
- correct-answer identity is checked across shuffled versions;
- deterministic Word crop geometries are verified;
- malformed Office containers, archive resource limits and invalid images have rejection tests at the import boundary;
- formula-like spreadsheet values are neutralised before export;
- declared runtime dependencies are checked for known vulnerabilities in CI;
- Streamlit tests cover cold start, upload, review, preview, configuration and download.

See [Quality and verification](docs/QUALITY.md) for the public verification scope, or [GitHub Actions](https://github.com/christosdvm/mcq-exam-builder/actions/workflows/tests.yml) for current results. These tests provide regression evidence, not proof that every document layout will import correctly. Pilot adoption and time-saving results have not yet been measured.

## Architecture

Python and Streamlit provide the review interface. Source adapters normalise questions and report import findings; a UI-independent generation module produces version arrangements; exporters assemble documents and reports into a ZIP.

| Component | Responsibility |
| --- | --- |
| `question_importer.py` and `source_formats.py` | Parsing, document conversion, source evidence and validation |
| `exam_versions.py` | Seeded shuffling and correct-answer identity |
| `app.py` | Review workflow, session state, configuration and package exports |

Export logic still lives in `app.py`; extracting it is the next modularisation step. [Architecture](docs/ARCHITECTURE.md) records the current boundaries, trade-offs and migration plan.

## Run locally

Python 3.12 is recommended. Legacy Word, RTF and OpenDocument conversion also requires LibreOffice Writer on the system path.

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

On Windows PowerShell, activate with `.venv\Scripts\Activate.ps1` instead.

## Test

```bash
python -m unittest discover -v -p 'test_*.py'
```

The format suite uses LibreOffice to create and convert real synthetic fixtures. It therefore takes longer than the pure parsing and generation tests.

## Documentation

| Document | Purpose |
| --- | --- |
| [Product vision](docs/PRODUCT_VISION.md) | User, problem, principles and non-goals |
| [Project profile](docs/PROJECT_PROFILE.md) | Concise product description and maintainer context |
| [Architecture](docs/ARCHITECTURE.md) | Current system and safety boundaries |
| [Supported formats](docs/SOURCE_FORMATS.md) | Input contracts and explicit limits |
| [Quality and verification](docs/QUALITY.md) | Public verification scope |
| [Deployment guide](docs/DEPLOYMENT_GUIDE.md) | Deployment guidance |
| [Development workflow](docs/DEVELOPMENT_WORKFLOW.md) | Branch, review and test conventions |
| [Privacy notice](PRIVACY.md) | Processing, temporary files and hosting boundaries |
| [Security policy](SECURITY.md) | Vulnerability reporting and supported version |

## Roadmap

1. Finish separating export logic from the UI, preserving the tested generation and validation behaviour.
2. Observe lecturers preparing examinations; measure completion, interventions and preparation time against their existing workflow.
3. Use those findings to prioritise additional formats, institutional deployment or a programmatic interface.

The longer-term aim is a reusable assessment-preparation core that can support different delivery interfaces. Question generation, proctoring and student-record management remain outside the current scope.

## Author

**Dr Christos I. Karagiannis, DVM, MSc, Dip. ECAWBM(BM), MRCVS** is an EBVS Specialist in Veterinary Behavioural Medicine, Assistant Professor in Animal Welfare and Veterinary Behavioural Medicine, and veterinary programme coordinator.

MCQ Exam Builder grew from a recurring challenge in university assessment preparation: question banks are usually working documents, but a single answer-mapping error can undermine a final paper. My teaching and coordination roles gave me a close view of where the process breaks down. I designed the project around those failure points—reviewable imports, reproducible exam versions, and safeguards against answer drift and content loss.

This is an independently maintained, practice-led software project. It is not an institutional examination platform and does not imply institutional endorsement.

## Licence

The code is provided for viewing and evaluation under an all-rights-reserved [licence notice](LICENSE.md). It is not distributed under an open-source licence; reuse or deployment requires prior written permission.
