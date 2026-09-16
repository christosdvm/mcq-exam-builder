# Architecture

The application is a Python/Streamlit workflow with three primary boundaries:

- `question_importer.py` and `source_formats.py`: import, conversion, normalization and structural validation.
- `exam_versions.py`: deterministic version generation and answer-identity preservation.
- `app.py`: review UI, configuration, session state and export orchestration.

Inputs are treated as untrusted. The application validates supported containers, blocks unsafe or ambiguous structures, and does not maintain a persistent exam database.
