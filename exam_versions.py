"""Deterministic exam-version generation independent of the Streamlit UI."""

from __future__ import annotations

import random
import re
from typing import Dict, List

import pandas as pd

IMAGE_COLUMN = "Image"
LETTERS = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")

POSITION_DEPENDENT_OPTION_RE = re.compile(
    r"\b(?:all|none)\s+of\s+(?:the\s+)?above\b|"
    r"\b(?:both|either|neither)\s+[A-F]\s+(?:and|or|nor)\s+[A-F]\b|"
    r"\boptions?\s+[A-F]\s+(?:and|or)\s+[A-F]\b|"
    r"\bstatements?\s+(?:I|1)\s+(?:and|or)\s+(?:II|2)\b",
    flags=re.IGNORECASE,
)


def normalise_cell(value) -> str:
    """Return a trimmed string while treating null-like values as empty."""
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def answer_columns_for_count(n: int) -> List[str]:
    """Return the canonical option columns for a question."""
    return [f"Option {index + 1}" for index in range(n)]


def allowed_letters_for_count(n: int) -> List[str]:
    """Return the answer labels available for a question."""
    return LETTERS[:n]


def clean_correct(value) -> str:
    """Normalise the accepted answer-label forms to a bare uppercase letter."""
    return normalise_cell(value).upper().replace(".", "").replace(")", "").strip()


def _split_image_files(value) -> List[str]:
    """Normalise the importer image contract without depending on a source adapter."""
    def image_text(item) -> str:
        if item is None:
            return ""
        try:
            if pd.isna(item):
                return ""
        except (TypeError, ValueError):
            pass
        return str(item).replace("\u00a0", " ").strip()

    if isinstance(value, list):
        return [image_text(item) for item in value if image_text(item)]
    text = image_text(value)
    if not text:
        return []
    return [part.strip() for part in text.split("|") if part.strip()]


def version_label(index: int, naming: str) -> str:
    """Return a one-based number or spreadsheet-style alphabetic label."""
    if naming == "Letters":
        remaining = index
        label = ""
        while True:
            label = chr(ord("A") + (remaining % 26)) + label
            remaining = remaining // 26 - 1
            if remaining < 0:
                break
        return label
    return str(index + 1)


def has_position_dependent_options(option_texts) -> bool:
    """Detect answer wording whose meaning depends on the displayed order."""
    return any(
        POSITION_DEPENDENT_OPTION_RE.search(normalise_cell(text))
        for text in option_texts
    )


def build_version_questions(
    df: pd.DataFrame,
    shuffle_questions: bool,
    shuffle_answers: bool,
    seed: int,
) -> List[Dict]:
    """Build one reproducible version while preserving answer identity."""
    # This is deterministic exam arrangement, not generation of a security token.
    rng = random.Random(seed)  # nosec B311
    records = df.to_dict("records")
    if shuffle_questions:
        rng.shuffle(records)

    output = []
    for record in records:
        option_count = int(record["Option_Count"])
        answer_columns = answer_columns_for_count(option_count)
        active_letters = allowed_letters_for_count(option_count)

        # Preserve the source label so duplicate option text cannot detach the answer.
        options = [
            {
                "original_letter": active_letters[index],
                "text": normalise_cell(record[answer_columns[index]]),
            }
            for index in range(option_count)
        ]
        original_correct = clean_correct(record["Correct"])

        preserve_answer_order = has_position_dependent_options(
            [option["text"] for option in options]
        )
        if shuffle_answers and not preserve_answer_order:
            rng.shuffle(options)

        new_correct = next(
            active_letters[index]
            for index, option in enumerate(options)
            if option["original_letter"] == original_correct
        )
        image_names = _split_image_files(
            record.get("Image_Files") or record.get(IMAGE_COLUMN, "")
        )
        output.append(
            {
                "Internal_ID": normalise_cell(record["Internal_ID"]),
                "No": normalise_cell(record.get("No", "")),
                "Question": normalise_cell(record["Question"]),
                "Image": image_names[0] if image_names else "",
                "Images": image_names,
                "Options": [
                    {
                        "letter": active_letters[index],
                        "text": options[index]["text"],
                        "original_letter": options[index]["original_letter"],
                    }
                    for index in range(option_count)
                ],
                "Correct": new_correct,
                "Original_Correct": original_correct,
                "Answer_Shuffle_Skipped": bool(
                    shuffle_answers and preserve_answer_order
                ),
                "Option_Count": option_count,
                "Source_File": normalise_cell(record["Source_File"]),
            }
        )
    return output


def build_versions(
    df: pd.DataFrame,
    num_versions: int,
    naming: str,
    shuffle_questions: bool,
    shuffle_answers: bool,
    base_seed: int = 10000,
) -> Dict[str, List[Dict]]:
    """Build a labelled set of deterministic exam versions."""
    return {
        version_label(index, naming): build_version_questions(
            df,
            shuffle_questions=shuffle_questions,
            shuffle_answers=shuffle_answers,
            seed=base_seed + index,
        )
        for index in range(num_versions)
    }
