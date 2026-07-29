from __future__ import annotations

from build_focused_qa import QUESTIONS, validate_question_data


def main() -> None:
    validate_question_data()
    assert len(QUESTIONS) == 18
    assert [item["slide"] for item in QUESTIONS] == [
        1,
        1,
        2,
        2,
        3,
        3,
        4,
        4,
        4,
        5,
        5,
        6,
        6,
        6,
        7,
        7,
        8,
        8,
    ]
    expected_keys = {
        "slide",
        "question_en",
        "question_zh",
        "answer_en",
        "answer_zh",
        "keywords",
        "cue",
    }
    assert all(set(item) == expected_keys for item in QUESTIONS)

    joined = "\n".join(
        f"{item['answer_en']}\n{item['cue']}" for item in QUESTIONS
    )
    for token in (
        "0.98",
        "250",
        "50,000",
        "1,000",
        "4.148",
        "2.123",
        "9.053",
        "4.703",
        "4.009",
        "97.911%",
        "0.515",
    ):
        assert token in joined, f"Missing numerical anchor: {token}"

    forbidden = ("event-equivalent", "H = 2", "AI assistance")
    for phrase in forbidden:
        assert phrase.lower() not in joined.lower(), f"Backup-only topic found: {phrase}"

    print("Focused Q&A coverage: PASS")
    print(f"Questions: {len(QUESTIONS)}")
    print("Slide sections: 8")


if __name__ == "__main__":
    main()
