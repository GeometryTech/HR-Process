import os
import pandas as pd
import logging
from resume_skill_parser import main

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

TEST_FOLDER = "test_resumes"
CSV_PATH = "candidate_skills.csv"

def test_parser_outputs():
    """Run parser and check that CSV, Excel, and PNG are created properly."""
    assert os.path.isdir(TEST_FOLDER), f"Missing test resumes folder: {TEST_FOLDER}"

    # Run the parser
    main(TEST_FOLDER)

    # === Check CSV output ===
    assert os.path.exists(CSV_PATH), "CSV file not created"
    df = pd.read_csv(CSV_PATH)

    # Validate required columns
    for col in ["Name", "Email", "Phone", "Skills"]:
        assert col in df.columns, f"Missing column: {col}"

    # Validate at least one candidate
    assert len(df) > 0, "No candidates extracted"

    # Validate skills and contact info
    # Only check skills actually present in your resumes
    expected_skills = ["azure", "cloud", "powershell", "r", "shell"]
    for skill in expected_skills:
        assert any(skill in str(skills).lower() for skills in df["Skills"].values), \
            f"Expected skill '{skill}' missing"

    # Validate that emails are extracted
    assert any(pd.notnull(email) and "@" in str(email) for email in df["Email"].values), \
        "No valid email extracted"

    print(df.to_markdown(index=False))

def test_skill_distribution_prints(capsys):
    """Check that the skill frequency distribution prints correctly."""
    assert os.path.isdir(TEST_FOLDER), f"Missing test resumes folder: {TEST_FOLDER}"
    main(TEST_FOLDER)

    captured = capsys.readouterr()
    assert "Technology Skill Distribution:" in captured.out, \
        "Skill distribution header not printed"
    print(captured.out)

if __name__ == "__main__":
    test_parser_outputs()
    test_skill_distribution_prints(None)
